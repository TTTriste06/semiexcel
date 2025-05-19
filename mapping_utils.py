import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    """
    将 DataFrame 中的三个主键列替换为新旧料号映射表中的新值，并对重复记录聚合（数值列求和）。

    参数:
    - df: 原始 DataFrame
    - mapping_df: 包含 ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
    - field_map: 当前表格中列名与映射字段的对应关系，如 {"规格": "产品型号", ...}
    - verbose: 是否输出替换信息

    返回:
    - 替换并聚合后的 DataFrame
    """

    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    try:
        df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

        # 打印匹配统计
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

        # 显示前几条未匹配记录（调试用）
        if unmatched_count > 0 and verbose:
            try:
                print("⚠️ 未匹配示例（前 5 行）：")
                print(df_merged[~matched][left_on].head())
            except:
                pass

        # 创建一个布尔掩码，表示三列新值都不为 None 且非空字符串
        mask_None = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )
        
        # 对满足条件的行进行整体替换
        df_merged.loc[mask_None, "规格"] = df_merged.loc[mask_None, "新规格"]
        df_merged.loc[mask_None, "品名"] = df_merged.loc[mask_None, "新品名"]
        df_merged.loc[mask_None, "晶圆品名"] = df_merged.loc[mask_None, "新晶圆品名"]

        # 替换三列值
        # df_merged[spec_col] = df_merged["新规格"].combine_first(df_merged[spec_col])
        # df_merged[name_col] = df_merged["新品名"].combine_first(df_merged[name_col])
        # df_merged[wafer_col] = df_merged["新晶圆品名"].combine_first(df_merged[wafer_col])

        # 删除映射中间列
        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        # 聚合：主键列相同的行合并
        group_cols = [spec_col, name_col, wafer_col]
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        # 保留其他字段（如单位、类型等）
        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        return df_grouped

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return df
