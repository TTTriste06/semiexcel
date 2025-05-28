from openpyxl.styles import PatternFill
import pandas as pd

def standardize(val):
    if val is None:
        return ""
    return str(val).strip().replace("\u3000", " ").strip('\'"“”‘’')


def append_forecast_unmatched_to_summary_by_keys(summary_df: pd.DataFrame, forecast_df: pd.DataFrame) -> pd.DataFrame:
    """
    将未匹配的预测记录中的产品型号、生产料号和预测值添加到汇总表末尾。

    参数：
    - summary_df: 汇总 DataFrame（列名如：规格、品名、晶圆品名等）
    - forecast_df: 原始预测表，包含未匹配记录（列名：产品型号、生产料号）

    返回：
    - summary_df: 添加了未匹配预测项的新 DataFrame
    """
    forecast_cols = [col for col in forecast_df.columns if "预测" in col]
    needed_cols = ["产品型号", "生产料号"] + forecast_cols
    forecast_subset = forecast_df[needed_cols].copy()

    # 只保留在 summary_df["品名"] 中未出现的记录
    unmatched = forecast_subset[~forecast_subset["生产料号"].isin(summary_df["品名"])].copy()

    unmatched["规格"] = unmatched["产品型号"]
    unmatched["品名"] = unmatched["生产料号"]
    unmatched["晶圆品名"] = ""

    # 按 summary_df 的列添加预测列，其余列补空
    summary_cols = summary_df.columns
    new_rows = unmatched[["晶圆品名", "规格", "品名"] + forecast_cols if forecast_cols else []]

    for col in summary_cols:
        if col not in new_rows.columns:
            new_rows[col] = ""

    new_rows = new_rows[summary_cols]
    summary_df = pd.concat([summary_df, new_rows], ignore_index=True)

    return summary_df
