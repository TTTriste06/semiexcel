import pandas as pd
import streamlit as st

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    """
    按品名字段替换主料号（新旧料号映射）
    """
    name_col = field_map["品名"]
    df[name_col] = df[name_col].astype(str).str.strip()
    mapping_df["旧品名"] = mapping_df["旧品名"].astype(str).str.strip()
    mapping_df["新品名"] = mapping_df["新品名"].astype(str).str.strip()

    df = df.copy()
    merged = df.merge(mapping_df[["旧品名", "新品名"]], how="left", left_on=name_col, right_on="旧品名")
    mask = merged["新品名"].notna() & (merged["新品名"] != "")
    merged["_由新旧料号映射"] = mask

    if verbose:
        st.write(f"✅ 新旧料号替换成功: {mask.sum()}，未匹配: {(~mask).sum()}")

    merged.loc[mask, name_col] = merged.loc[mask, "新品名"]
    merged = merged.drop(columns=["旧品名", "新品名"], errors="ignore")

    mapped_keys = set(merged.loc[mask, name_col])

    return merged.drop(columns=["_由新旧料号映射"], errors="ignore"), mapped_keys

def apply_extended_substitute_mapping(df, mapping_df, field_map, verbose=True):
    """
    替代料号品名替换（仅品名字段替换，无聚合合并）
    """
    name_col = field_map["品名"]
    df = df.copy()
    df[name_col] = df[name_col].astype(str).str.strip().str.replace("\n", "").str.replace("\r", "")

    # 清洗映射表中所有替代品名及新品名
    substitute_records = []
    for i in range(1, 5):
        sub_name = f"替代品名{i}"
        for col in [sub_name, "新品名"]:
            if col not in mapping_df.columns:
                mapping_df[col] = ""
            mapping_df[col] = mapping_df[col].astype(str).str.strip().str.replace("\n", "").str.replace("\r", "")

        valid_rows = mapping_df[
            mapping_df[[sub_name, "新品名"]].notna().all(axis=1) &
            (mapping_df[sub_name] != "") &
            (mapping_df["新品名"] != "")
        ]

        for _, row in valid_rows.iterrows():
            substitute_records.append({
                "旧品名": row[sub_name],
                "新品名": row["新品名"]
            })

    # 替换品名
    matched_keys = set()
    for sub in substitute_records:
        mask = (df[name_col] == sub["旧品名"])
        if mask.any():
            if verbose:
                st.write(f"🔁 替代品名: {sub['旧品名']} → {sub['新品名']}，行数: {mask.sum()}")
            df.loc[mask, name_col] = sub["新品名"]
            matched_keys.update(df.loc[mask, name_col])

    if verbose:
        st.success(f"✅ 替代品名替换完成，共替换: {len(matched_keys)} 种")

    return df, matched_keys
    


