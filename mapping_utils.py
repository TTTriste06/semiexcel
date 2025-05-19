import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    try:
        df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

        # 记录新旧料号是否匹配成功
        matched = df_merged["新规格"].notna()
        df_merged["是否匹配新旧料号"] = matched

        # 替换条件掩码
        mask_replace = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )

        # 替换为新料号
        df_merged.loc[mask_replace, spec_col] = df_merged.loc[mask_replace, "新规格"]
        df_merged.loc[mask_replace, name_col] = df_merged.loc[mask_replace, "新品名"]
        df_merged.loc[mask_replace, wafer_col] = df_merged.loc[mask_replace, "新晶圆品名"]

        # 清洗无用列
        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        # 统计聚合前数量
        group_cols = [spec_col, name_col, wafer_col]
        pre_group_counts = df_cleaned.groupby(group_cols).size().reset_index(name="合并前行数")

        # 聚合数值列
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        # 合并非数值列
        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols + ["是否匹配新旧料号"]]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        # 补回是否匹配列（取 True 就表示至少有一条是匹配的）
        match_flag = df_cleaned.groupby(group_cols)["是否匹配新旧料号"].any().reset_index()
        df_grouped = pd.merge(df_grouped, match_flag, on=group_cols, how="left")

        # 标记是否为新旧料号合并：合并前有多条记录
        df_grouped = pd.merge(df_grouped, pre_group_counts, on=group_cols, how="left")
        df_grouped["是否新旧料号合并"] = df_grouped["合并前行数"] > 1

        if verbose:
            import streamlit as st
            match_count = matched.sum()
            unmatched_count = (~matched).sum()
            st.info(f"🎯 成功替换 {match_count} 行；未匹配 {unmatched_count} 行")
            if unmatched_count > 0:
                print("⚠️ 未匹配示例（前5行）：")
                print(df_merged[~matched][left_on].head())

        return df_grouped

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return df
