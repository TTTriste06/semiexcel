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

    # 标准化字段
    df[spec_col] = df[spec_col].astype(str).str.strip()
    df[name_col] = df[name_col].astype(str).str.strip()
    df[wafer_col] = df[wafer_col].astype(str).str.strip()

    # 标准化替代列
    extended_cols = []
    for i in range(1, 5):
        for col in [f"替代规格{i}", f"替代品名{i}", f"替代晶圆{i}"]:
            mapping_df[col] = mapping_df.get(col, "").astype(str).str.strip()
        extended_cols.append((f"替代规格{i}", f"替代品名{i}", f"替代晶圆{i}"))

    matched_keys = set()

    def try_substitute(row):
        original_key = (row[spec_col], row[name_col], row[wafer_col])
        if original_key in already_mapped_keys:
            return row  # 跳过已替换行

        for idx, map_row in mapping_df.iterrows():
            for a, b, c in extended_cols:
                sub_key = (map_row[a], map_row[b], map_row[c])
                
                # ✅ 打印每一组替代键值（用于调试）
                if verbose:
                    st.write(f"🧪 尝试替代组: {sub_key} vs 当前行: {original_key}")

                if original_key == sub_key:
                    row[spec_col] = map_row["新规格"]
                    row[name_col] = map_row["新品名"]
                    row[wafer_col] = map_row["新晶圆品名"]
                    matched_keys.add((row[spec_col], row[name_col], row[wafer_col]))
                    row["_由替代料号映射"] = True
                    return row

        row["_由替代料号映射"] = False
        return row

    df = df.apply(try_substitute, axis=1)

    df.drop(columns=["_由替代料号映射"], inplace=True)

    if verbose:
        st.success(f"🔁 替代匹配成功数: {len(matched_keys)}")

    return df, matched_keys
