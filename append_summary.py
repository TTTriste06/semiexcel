from openpyxl.styles import PatternFill
import pandas as pd

def standardize(val):
    if val is None:
        return ""
    return str(val).strip().replace("\u3000", " ").strip('\'"“”‘’')


def append_forecast_unmatched_to_summary_by_keys(summary_df: pd.DataFrame, forecast_df: pd.DataFrame) -> pd.DataFrame:
    """
    将未匹配的预测记录（被标红）中的产品料号、生产料号和预测值添加到汇总表末尾。

    参数：
    - summary_df: 汇总 DataFrame
    - forecast_df: 原始预测表，包含未匹配记录（标红）

    返回：
    - summary_df: 追加了未匹配预测项的新 DataFrame
    """

    # 找出预测列（如“5月预测”、“6月预测”...）
    forecast_cols = [col for col in forecast_df.columns if "预测" in col]

    # 只保留品名、生产料号和预测值
    needed_cols = ["产品型号", "生产料号"] + forecast_cols
    forecast_subset = forecast_df[needed_cols].copy()

    # 只保留预测值不为0 且 在 summary_df 中未匹配的（即不在已有的品名中）
    unmatched_forecast = forecast_subset[~forecast_subset["品名"].isin(summary_df["品名"])]

    # 创建新行，填入品名、规格、晶圆品名
    unmatched_forecast["规格"] = ""
    unmatched_forecast["晶圆品名"] = ""
    new_rows = unmatched_forecast.rename(columns={
        "产品型号": "规格",
        "生产料号": "品名"
    })

    # 只保留汇总表中已有的列（避免列不一致）
    summary_cols = summary_df.columns
    new_rows = new_rows[[col for col in summary_cols if col in new_rows.columns]]
    for col in summary_cols:
        if col not in new_rows.columns:
            new_rows[col] = ""

    # 调整顺序并拼接
    new_rows = new_rows[summary_cols]
    summary_df = pd.concat([summary_df, new_rows], ignore_index=True)

    return summary_df
