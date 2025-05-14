import pandas as pd
import re
import streamlit as st
from openpyxl.styles import PatternFill

def merge_safety_inventory(summary_df, safety_df, writer=None):
    """
    将安全库存表中 Wafer 和 Part 信息合并到汇总数据中，并标红未匹配行（可选）。

    参数:
    - summary_df: 汇总后的未交订单表，包含 '晶圆品名'、'规格'、'品名'
    - safety_df: 安全库存表，包含 'WaferID', 'OrderInformation', 'ProductionNO.', ' InvWaf', ' InvPart'
    - writer: openpyxl ExcelWriter（可选，如果提供，则对未匹配行进行标红）

    返回:
    - 合并后的汇总 DataFrame
    """

    # 重命名列用于匹配
    safety_df = safety_df.rename(columns={
        'WaferID': '晶圆品名',
        'OrderInformation': '规格',
        'ProductionNO.': '品名'
    }).copy()

    # 添加标记列
    safety_df['已匹配'] = False

    # 匹配主键集合
    summary_keys = set(
        tuple(str(x).strip() for x in row)
        for row in summary_df[['晶圆品名', '规格', '品名']].dropna().values
    )

    # 标记匹配情况
    safety_df['已匹配'] = safety_df.apply(
        lambda row: (str(row['晶圆品名']).strip(), str(row['规格']).strip(), str(row['品名']).strip()) in summary_keys,
        axis=1
    )

    # 标红未被匹配的行（如果提供 writer）
    if writer and "赛卓-安全库存" in writer.sheets:
        ws = writer.sheets["赛卓-安全库存"]
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        for row in range(3, ws.max_row + 1):
            wafer = str(ws.cell(row=row, column=1).value).strip()
            spec = str(ws.cell(row=row, column=2).value).strip()
            name = str(ws.cell(row=row, column=3).value).strip()
            if (wafer, spec, name) not in summary_keys:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = red_fill

    # 执行合并
    merged = summary_df.merge(
        safety_df[['晶圆品名', '规格', '品名', ' InvWaf', ' InvPart']],
        on=['晶圆品名', '规格', '品名'],
        how='left'
    )

    return merged


def append_unfulfilled_summary_columns(summary_df, pivoted_df):
    """
    提取历史未交订单 + 各未来月份未交订单列，计算总未交订单，并将它们添加到汇总 summary_df 的末尾。

    参数:
    - summary_df: 汇总 sheet（包含晶圆品名、规格、品名）
    - pivoted_df: 已透视后的未交订单表（含列如 未交订单数量_2025-03）

    返回:
    - 增加了新列的 summary_df
    """
    # 匹配所有未交订单列（含历史和各月）
    unfulfilled_cols = [col for col in pivoted_df.columns if "未交订单数量" in col]
    unfulfilled_df = pivoted_df[["晶圆品名", "规格", "品名"] + unfulfilled_cols].copy()

    # 计算总未交订单
    unfulfilled_df["总未交订单"] = unfulfilled_df[unfulfilled_cols].sum(axis=1)

    # 按所需顺序组织列
    ordered_cols = ["晶圆品名", "规格", "品名", "总未交订单"]
    if "历史未交订单数量" in pivoted_df.columns:
        ordered_cols.append("历史未交订单数量")
    ordered_cols += [col for col in unfulfilled_cols if col != "历史未交订单数量"]

    unfulfilled_df = unfulfilled_df[ordered_cols]

    # 合并到 summary_df
    merged = summary_df.merge(unfulfilled_df, on=["晶圆品名", "规格", "品名"], how="left")

    return merged

def append_forecast_to_summary(summary_df, forecast_df):
    """
    从预测表中提取与 summary_df 匹配的预测记录，仅提取一行预测（每组主键）。
    """

    # Debug: 显示原始预测表列
    st.write("原始预测表列名：", forecast_df.columns.tolist())

    # 重命名主键列
    forecast_df = forecast_df.rename(columns={
        "产品型号": "规格",
        "ProductionNO.": "品名"
    })

    # 主键列
    key_cols = ["晶圆品名", "规格", "品名"]

    # 找出预测月份列（如“5月预测”、“6月预测”...）
    month_cols = [col for col in forecast_df.columns if isinstance(col, str) and "预测" in col]
    st.write("识别到的预测列：", month_cols)

    if not month_cols:
        st.warning("⚠️ 没有识别到任何预测列，请检查列名是否包含'预测'")
        return summary_df

    # 去重：每组主键保留第一行
    forecast_df = forecast_df[key_cols + month_cols].drop_duplicates(subset=key_cols)

    # 合并进 summary
    merged = summary_df.merge(forecast_df, on=key_cols, how="left")
    st.write("合并后的汇总示例：", merged.head(3))
    return merged

def merge_finished_inventory(summary_df, finished_df):
    # 确保列名干净
    finished_df.columns = finished_df.columns.str.strip()

    # 主键列转换
    finished_df = finished_df.rename(columns={"WAFER品名": "晶圆品名"})

    key_cols = ["晶圆品名", "规格", "品名"]
    value_cols = ["数量_HOLD仓", "数量_成品仓", "数量_半成品仓"]

    # 验证是否都存在
    for col in key_cols + value_cols:
        if col not in finished_df.columns:
            st.error(f"❌ 缺失列：{col}")
            return summary_df

    st.write("✅ 正在按主键合并以下列：", value_cols)
    merged = summary_df.merge(finished_df[key_cols + value_cols], on=key_cols, how="left")

    return merged

def append_product_in_progress(summary_df, product_in_progress_df, mapping_df):
    """
    将成品在制与半成品在制信息合并到 summary_df 中。
    
    参数：
    - summary_df: 汇总表（含“晶圆品名”，“规格”，“品名”）
    - product_in_progress_df: 透视后的“赛卓-成品在制”数据
    - mapping_df: 新旧料号映射表，包含“半成品”列

    返回：
    - summary_df: 添加了“成品在制”与“半成品在制”的 DataFrame
    """
    numeric_cols = product_in_progress_df.select_dtypes(include='number').columns.tolist()
    summary_df = summary_df.copy()
    summary_df["成品在制"] = 0
    summary_df["半成品在制"] = 0

    # 填充成品在制
    for idx, row in product_in_progress_df.iterrows():
        mask = (
            (summary_df["晶圆品名"] == row["晶圆型号"]) &
            (summary_df["规格"] == row["产品规格"]) &
            (summary_df["品名"] == row["产品品名"])
        )
        if mask.any():
            summary_df.loc[mask, "成品在制"] = row[numeric_cols].sum()

    # 半成品逻辑
    semi_rows = mapping_df[mapping_df["半成品"].notna() & (mapping_df["半成品"] != "")]
    semi_info_table = semi_rows[["新规格", "新品名", "新晶圆品名", "半成品"]].copy()
    semi_info_table["未交数据和"] = 0

    for idx, row in semi_info_table.iterrows():
        matched = product_in_progress_df[
            (product_in_progress_df["产品规格"] == row["新规格"]) &
            (product_in_progress_df["晶圆型号"] == row["新晶圆品名"]) &
            (product_in_progress_df["产品品名"] == row["半成品"])
        ]
        semi_info_table.at[idx, "未交数据和"] = matched[numeric_cols].sum().sum() if not matched.empty else 0

    for idx, row in semi_info_table.iterrows():
        mask = (
            (summary_df["晶圆品名"] == row["新晶圆品名"]) &
            (summary_df["规格"] == row["新规格"]) &
            (summary_df["品名"] == row["新品名"])
        )
        if mask.any():
            summary_df.loc[mask, "半成品在制"] = row["未交数据和"]

    return summary_df

