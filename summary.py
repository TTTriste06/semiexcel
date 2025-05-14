import pandas as pd
import re
import streamlit as st

def merge_safety_inventory(summary_df, safety_df):
    """
    将安全库存表中 Wafer 和 Part 信息合并到汇总数据中。
    
    参数:
    - summary_df: 汇总后的未交订单表，包含 '晶圆品名'、'规格'、'品名'
    - safety_df: 安全库存表，包含 'WaferID', 'OrderInformation', 'ProductionNO.', ' InvWaf', ' InvPart'
    
    返回:
    - 合并后的汇总 DataFrame，增加了 ' InvWaf' 和 ' InvPart' 两列
    """

    # 重命名列用于匹配
    safety_df = safety_df.rename(columns={
        'WaferID': '晶圆品名',
        'OrderInformation': '规格',
        'ProductionNO.': '品名'
    }).copy()

    # 添加标记列（可选，用于调试或统计）
    safety_df['已匹配'] = False

    # 合并：left join 确保 summary_df 保留所有行
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
    """
    将成品库存的数量列（HOLD仓、成品仓、半成品仓）合并进汇总表
    """
    import streamlit as st

    try:
        # ✅ 重命名主键列（用于匹配）
        finished_df = finished_df.rename(columns={
            "WAFER品名": "晶圆品名",
        })

        # ✅ 匹配字段
        key_cols = ["晶圆品名", "规格", "品名"]

        # ✅ 选择需要的列
        value_cols = [col for col in finished_df.columns if col.startswith("数量_")]

        st.write("🔍 成品库存合并字段：", value_cols)

        merged = summary_df.merge(
            finished_df[key_cols + value_cols],
            on=key_cols,
            how="left"
        )
        return merged
    except Exception as e:
        st.error(f"❌ 成品库存合并失败: {e}")
        return summary_df
