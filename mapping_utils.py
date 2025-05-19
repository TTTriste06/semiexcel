import pandas as pd

import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    """
    将 DataFrame 中的三个主键列替换为新旧料号映射表中的新值，并对重复记录聚合（数值列求和）。
    同时输出所有由新旧料号合并产生的记录主键列表。
    
    返回:
    - 替换并聚合后的 DataFrame
    - List[Tuple]：被合并的新主键行（如 [("新规格A", "新品名A", "新晶圆A"), ...]）
    """

    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    try:
        df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

        # 匹配统计
        matched = df_merged["新规格"].notna()
        match_count = matched.sum()
        unmatched_count = (~matched).sum()

        if verbose:
            msg = f"🎯 成功替换 {match_count} 行；未匹配 {unmatched_count} 行"
            try:
                import streamlit as st
                st.info(msg)
            except:
                print(msg)

        # 显示前几条未匹配记录
        if unmatched_count > 0 and verbose:
            try:
                print("⚠️ 未匹配示例（前 5 行）：")
                print(df_merged[~matched][left_on].head())
            except:
                pass

        # 创建布尔掩码用于替换
        mask_valid = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )

        # 替换三列（注意用原字段名）
        df_merged.loc[mask_valid, spec_col] = df_merged.loc[mask_valid, "新规格"]
        df_merged.loc[mask_valid, name_col] = df_merged.loc[mask_valid, "新品名"]
        df_merged.loc[mask_valid, wafer_col] = df_merged.loc[mask_valid, "新晶圆品名"]

        # 删除映射中间列
        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        # 获取聚合前每个主键组出现次数（用于判断合并）
        group_cols = [spec_col, name_col, wafer_col]
        group_counts = df_cleaned.groupby(group_cols).size().reset_index(name="合并前行数")

        # 标记合并行（三元组）= 聚合前重复行数 > 1
        merged_key_list = group_counts[group_counts["合并前行数"] > 1][group_cols].apply(tuple, axis=1).tolist()

        # 聚合数值列
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        # 保留非数值字段（如单位、类型等）
        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        return df_grouped, merged_key_list

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return df, []
