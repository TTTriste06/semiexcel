import pandas as pd
import streamlit as st

def apply_mapping_and_merge(df, mapping_df, field_map, verbose=True):
    """
    用 key-based 显式映射方式将旧料号替换为新料号，并聚合相同行。
    """

    # 原表字段
    col_spec = field_map["规格"]
    col_name = field_map["品名"]
    col_wafer = field_map["晶圆品名"]

    # 创建唯一 key（保证三个字段一致）
    df["__key__"] = df[col_spec].astype(str) + "||" + df[col_name].astype(str) + "||" + df[col_wafer].astype(str)
    mapping_df["__key__"] = (
        mapping_df["旧规格"].astype(str) + "||" +
        mapping_df["旧品名"].astype(str) + "||" +
        mapping_df["旧晶圆品名"].astype(str)
    )

    # 打印出前几行 key 做对比
    st.write("原始表 Key 示例：", df["__key__"].head().tolist())
    st.write("新旧料号 Key 示例：", mapping_df["__key__"].head().tolist())


    # 构造 key → [新规格, 新品名, 新晶圆品名] 映射字典
    mapping_dict = mapping_df.set_index("__key__")[["新规格", "新品名", "新晶圆品名"]].to_dict(orient="index")

    # 执行替换
    replaced_rows = 0
    new_specs, new_names, new_wafers = [], [], []
    for key in df["__key__"]:
        if key in mapping_dict:
            new_specs.append(mapping_dict[key]["新规格"])
            new_names.append(mapping_dict[key]["新品名"])
            new_wafers.append(mapping_dict[key]["新晶圆品名"])
            replaced_rows += 1
        else:
            new_specs.append(None)
            new_names.append(None)
            new_wafers.append(None)

    # 替换字段（保留原值）
    df[col_spec] = pd.Series(new_specs).combine_first(df[col_spec])
    df[col_name] = pd.Series(new_names).combine_first(df[col_name])
    df[col_wafer] = pd.Series(new_wafers).combine_first(df[col_wafer])

    if verbose:
        try:
            st.info(f"🔁 替换成功 {replaced_rows} 行；保留原值 {len(df) - replaced_rows} 行")
        except:
            print(f"🔁 替换成功 {replaced_rows} 行；保留原值 {len(df) - replaced_rows} 行")

    # 删除 key 列
    df.drop(columns="__key__", inplace=True, errors="ignore")

    # 聚合（数值列求和）
    group_cols = [col_spec, col_name, col_wafer]
    numeric_cols = df.select_dtypes(include="number").columns.difference(group_cols).tolist()

    df_agg = df.groupby(group_cols, as_index=False)[numeric_cols].sum()

    # 处理其他非数值字段（保留第一个）
    non_numeric_cols = df.columns.difference(group_cols + numeric_cols).tolist()
    if non_numeric_cols:
        df_first = df.groupby(group_cols, as_index=False)[non_numeric_cols].first()
        df_agg = pd.merge(df_agg, df_first, on=group_cols, how="left")

    return df_agg
