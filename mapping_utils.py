import pandas as pd

def apply_mapping_and_merge(df, mapping_df, field_map):
    """
    将指定 DataFrame 中的三列（规格、品名、晶圆品名）依据新旧料号对照表替换为新值，并进行聚合合并。

    参数:
    - df: 原始表格（如预测、安全库存）
    - mapping_df: 新旧料号表，必须包含 ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
    - field_map: 当前表中三列的列名映射，例如：
        {
            "规格": "产品型号",
            "品名": "ProductionNO.",
            "晶圆品名": "晶圆品名"
        }

    返回:
    - 替换后的并合并后的 DataFrame
    """

    # 当前表的三列列名
    spec_col = field_map["规格"]
    name_col = field_map["品名"]
    wafer_col = field_map["晶圆品名"]

    # 新旧料号匹配列
    left_on = [spec_col, name_col, wafer_col]
    right_on = ["旧规格", "旧品名", "旧晶圆品名"]

    # 合并映射
    df_merged = df.merge(mapping_df, how="left", left_on=left_on, right_on=right_on)

    # 替换为新值（有匹配就替换，没有就保留原值）
    df_merged[spec_col] = df_merged["新规格"].combine_first(df_merged[spec_col])
    df_merged[name_col] = df_merged["新品名"].combine_first(df_merged[name_col])
    df_merged[wafer_col] = df_merged["新晶圆品名"].combine_first(df_merged[wafer_col])

    # 删除辅助映射列
    drop_cols = ["旧规格", "旧品名", "旧晶圆品名", "新规格", "新品名", "新晶圆品名"]
    df_cleaned = df_merged.drop(columns=[col for col in drop_cols if col in df_merged.columns])

    # 聚合合并：根据三个主键字段，数值列求和
    group_cols = [spec_col, name_col, wafer_col]
    numeric_cols = df_cleaned.select_dtypes(include="number").columns.tolist()
    sum_cols = [col for col in numeric_cols if col not in group_cols]

    # 聚合主表
    df_grouped = df_cleaned.groupby(group_cols, as_index=False)[sum_cols].sum()

    # 如果还有其他非主键非数值列（如单位、类型等），保留第一条记录
    other_cols = [col for col in df_cleaned.columns if col not in group_cols + sum_cols]
    if other_cols:
        df_first = df_cleaned.groupby(group_cols, as_index=False)[other_cols].first()
        df_grouped = pd.merge(df_grouped, df_first, on=group_cols, how="left")

    return df_grouped
