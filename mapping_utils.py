import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    for col in [spec_col, name_col, wafer_col]:
        df[col] = df[col].astype(str).str.strip()
    for col in ["旧规格", "旧品名", "旧晶圆品名"]:
        mapping_df[col] = mapping_df[col].astype(str).str.strip()

    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    try:
        df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

        matched = df_merged["新规格"].notna()
        unmatched_count = (~matched).sum()

        # 生成布尔掩码：成功被新旧料号替换的行
        mask_None = (
            df_merged["新规格"].notna() & (df_merged["新规格"].astype(str).str.strip() != "") &
            df_merged["新品名"].notna() & (df_merged["新品名"].astype(str).str.strip() != "") &
            df_merged["新晶圆品名"].notna() & (df_merged["新晶圆品名"].astype(str).str.strip() != "")
        )

        df_merged["_由新旧料号映射"] = mask_None  # 标记列 ✅

        # 替换字段
        df_merged.loc[mask_None, spec_col] = df_merged.loc[mask_None, "新规格"]
        df_merged.loc[mask_None, name_col] = df_merged.loc[mask_None, "新品名"]
        df_merged.loc[mask_None, wafer_col] = df_merged.loc[mask_None, "新晶圆品名"]

        # 删除中间列
        drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
        df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

        group_cols = [spec_col, name_col, wafer_col]
        numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
        sum_cols = [col for col in numeric_cols if col not in group_cols]

        df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

        other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
        if other_cols:
            df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
            df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

        # ✅ 返回主键集合
        mapped_keys = set(
            tuple(df_merged.loc[idx, [spec_col, name_col, wafer_col]].values)
            for idx in df_merged.index[df_merged["_由新旧料号映射"]]
        )

        return df_grouped, mapped_keys

    except Exception as e:
        print(f"❌ 替换失败: {e}")
        return df, set()

def apply_extended_substitute_mapping(df, mapping_df, field_map, already_mapped_keys=None, verbose=True):
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    if already_mapped_keys is None:
        already_mapped_keys = set()

    extended_cols = []
    for i in range(1, 5):
        for col in [f"替代规格{i}", f"替代品名{i}", f"替代晶圆{i}"]:
            mapping_df[col] = mapping_df.get(col, "").astype(str).str.strip()
        extended_cols.append((f"替代规格{i}", f"替代品名{i}", f"替代晶圆{i}"))

    def try_substitute(row):
        if (row[spec_col], row[name_col], row[wafer_col]) in already_mapped_keys:
            return row  # 已映射，跳过

        for _, map_row in mapping_df.iterrows():
            for a, b, c in extended_cols:
                if (row[spec_col], row[name_col], row[wafer_col]) == (map_row[a], map_row[b], map_row[c]):
                    row[spec_col] = map_row["新规格"]
                    row[name_col] = map_row["新品名"]
                    row[wafer_col] = map_row["新晶圆品名"]
                    row["_由替代料号映射"] = True
                    return row
        row["_由替代料号映射"] = False
        return row

    df[spec_col] = df[spec_col].astype(str).str.strip()
    df[name_col] = df[name_col].astype(str).str.strip()
    df[wafer_col] = df[wafer_col].astype(str).str.strip()

    df = df.apply(try_substitute, axis=1)

    matched_keys = set(
        tuple(row) for row in df.loc[df["_由替代料号映射"], [spec_col, name_col, wafer_col]].itertuples(index=False, name=None)
    )

    df.drop(columns=["_由替代料号映射"], inplace=True)

    return df, matched_keys

