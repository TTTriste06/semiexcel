import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True, return_merge_keys=False):
    """
    替换料号并聚合，若 return_merge_keys=True，则额外返回由多个旧料号合并的新主键列表。
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
        if verbose:
            match_count = matched.sum()
            unmatched_count = (~matched).sum()
            msg = f"🎯 成功替换 {match_count} 行；未匹配 {unmatched_count} 行"
            try:
                import streamlit as st
                st.info(msg)
            except:
                print(msg)

        # 替换新值
        mask_valid = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )
        df_merged.loc[mask_valid, spec_col] = df_merged.loc[mask_valid, "新规格"]
        df_merged.loc[mask_valid, name_col] = df_merged.loc[mask_valid, "新品名"]
        df_merged.loc[mask_valid, wafer_col] = df_merged.loc[mask_valid, "新晶圆品名"]

        # 删除中间列
        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        # 聚合前重复行统计
        group_cols = [spec_col, name_col, wafer_col]
        if return_merge_keys:
            group_counts = df_cleaned.groupby(group_cols).size().reset_index(name="合并前行数")
            merged_key_list = list(map(list, group_counts[group_counts["合并前行数"] > 1][group_cols].values))

        # 聚合数值列
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]
        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        # 保留其他字段
        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        if return_merge_keys:
            return df_grouped, merged_key_list
        else:
            return df_grouped

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return (df, []) if return_merge_keys else df
