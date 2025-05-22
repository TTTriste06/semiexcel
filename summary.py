import pandas as pd
import re
import streamlit as st
from openpyxl.styles import PatternFill


def merge_safety_inventory(summary_df, safety_df):
    """
    将安全库存表中 Wafer 和 Part 信息合并到汇总数据中，并返回未匹配的主键列表。

    参数:
    - summary_df: 汇总后的未交订单表，包含 '晶圆品名'、'规格'、'品名'
    - safety_df: 安全库存表，包含 'WaferID', 'OrderInformation', 'ProductionNO.', ' InvWaf', ' InvPart'

    返回:
    - merged: 合并后的汇总 DataFrame
    - unmatched_keys: list of (晶圆品名, 规格, 品名) 未匹配主键
    """

    # 重命名列统一主键
    safety_df = safety_df.rename(columns={
        'WaferID': '晶圆品名',
        'OrderInformation': '规格',
        'ProductionNO.': '品名'
    }).copy()

    key_cols = ['晶圆品名', '规格', '品名']

    # 所有主键（包括缺失值的行，缺失值视为空字符串）
    safety_df['主键'] = safety_df.apply(
        lambda row: tuple(str(row[col]).strip() if pd.notnull(row[col]) else '' for col in key_cols),
        axis=1
    )
    all_keys = set(safety_df['主键'])

    # 合并到 summary
    merged = summary_df.merge(
        safety_df[key_cols + [' InvWaf', ' InvPart']],
        on=key_cols,
        how='left'
    )

    # 记录实际被使用到的主键（在合并结果中出现了非空 InvWaf 或 InvPart 的行）
    used_keys = set(
        merged[~merged[[' InvWaf', ' InvPart']].isna().all(axis=1)]
        .apply(lambda row: tuple(str(row[col]).strip() if pd.notnull(row[col]) else '' for col in key_cols), axis=1)
    )

    # 剩下的就是未被使用的
    unmatched_keys = list(all_keys - used_keys)

    # st.write("用到的：")
    # st.write(used_keys)

    return merged, unmatched_keys





def append_unfulfilled_summary_columns(summary_df, pivoted_df):
    """
    提取历史未交订单 + 各未来月份未交订单列，计算总未交订单，并将它们添加到汇总 summary_df 的末尾。
    返回合并后的 summary_df 和未匹配的主键列表。
    """

    # 匹配所有未交订单列
    unfulfilled_cols = [col for col in pivoted_df.columns if "未交订单数量" in col]
    unfulfilled_df = pivoted_df[["晶圆品名", "规格", "品名"] + unfulfilled_cols].copy()

    # 计算总未交订单
    unfulfilled_df["总未交订单"] = unfulfilled_df[unfulfilled_cols].sum(axis=1)

    # 整理列顺序
    ordered_cols = ["晶圆品名", "规格", "品名", "总未交订单"]
    if "历史未交订单数量" in pivoted_df.columns:
        ordered_cols.append("历史未交订单数量")
    ordered_cols += [col for col in unfulfilled_cols if col != "历史未交订单数量"]
    unfulfilled_df = unfulfilled_df[ordered_cols]

    # 查找未匹配主键
    summary_keys = set(
        tuple(str(x).strip() for x in row)
        for row in summary_df[["晶圆品名", "规格", "品名"]].dropna().values
    )
    unmatched_keys = []
    for _, row in unfulfilled_df.iterrows():
        key = (str(row["晶圆品名"]).strip(), str(row["规格"]).strip(), str(row["品名"]).strip())
        if key not in summary_keys:
            unmatched_keys.append(key)

    # 合并
    merged = summary_df.merge(unfulfilled_df, on=["晶圆品名", "规格", "品名"], how="left")

    return merged, unmatched_keys



def append_forecast_to_summary(summary_df, forecast_df):
    """
    从预测表中提取与 summary_df 匹配的预测记录，并返回未匹配的主键列表。

    参数:
    - summary_df: 汇总表（含主键）
    - forecast_df: 原始预测表

    返回:
    - merged: 合并后的 summary_df
    - unmatched_keys: list of (晶圆品名, 规格, 品名) 未被匹配的主键
    """

    # Debug: 显示原始预测表列
    # st.write("原始预测表列名：", forecast_df.columns.tolist())

    # 重命名主键列
    forecast_df = forecast_df.rename(columns={
        "产品型号": "规格",
        "ProductionNO.": "品名"
    })

    # 主键列
    key_cols = ["晶圆品名", "规格", "品名"]

    # 找出预测月份列（如“5月预测”、“6月预测”...）
    month_cols = [col for col in forecast_df.columns if isinstance(col, str) and "预测" in col]
    # st.write("识别到的预测列：", month_cols)

    if not month_cols:
        st.warning("⚠️ 没有识别到任何预测列，请检查列名是否包含'预测'")
        return summary_df, []

    # 去重：每组主键保留第一行
    forecast_df = forecast_df[key_cols + month_cols].drop_duplicates(subset=key_cols)

    # 查找未匹配的主键
    summary_keys = set(
        tuple(str(x).strip() for x in row)
        for row in summary_df[key_cols].dropna().values
    )
    unmatched_keys = []
    for _, row in forecast_df.iterrows():
        key = tuple(str(row[col]).strip() for col in key_cols)
        if key not in summary_keys:
            unmatched_keys.append(key)

    # 合并进 summary
    merged = summary_df.merge(forecast_df, on=key_cols, how="left")
    # st.write("合并后的汇总示例：", merged.head(3))

    return merged, unmatched_keys



def merge_finished_inventory(summary_df, finished_df):
    """
    合并成品库存表进 summary_df，并返回未匹配的主键。

    参数:
    - summary_df: 汇总数据
    - finished_df: 透视后的成品库存表

    返回:
    - merged: 合并后的 DataFrame
    - unmatched_keys: list of (晶圆品名, 规格, 品名) 未匹配的键
    """

    # 确保列名干净
    finished_df.columns = finished_df.columns.str.strip()

    # 主键列转换
    finished_df = finished_df.rename(columns={"WAFER品名": "晶圆品名"})

    key_cols = ["晶圆品名", "规格", "品名"]
    value_cols = ["数量_HOLD仓", "数量_成品仓", "数量_半成品仓"]

    for col in key_cols + value_cols:
        if col not in finished_df.columns:
            st.error(f"❌ 缺失列：{col}")
            return summary_df, []

    # 提取未匹配主键
    summary_keys = set(
        tuple(str(x).strip() for x in row)
        for row in summary_df[key_cols].dropna().values
    )

    unmatched_keys = []
    for _, row in finished_df.iterrows():
        key = tuple(str(row[col]).strip() for col in key_cols)
        if key not in summary_keys:
            unmatched_keys.append(key)

    # st.write("✅ 正在按主键合并以下列：", value_cols)
    merged = summary_df.merge(finished_df[key_cols + value_cols], on=key_cols, how="left")

    return merged, unmatched_keys



def append_product_in_progress(summary_df, product_in_progress_df, mapping_df):
    """
    将成品在制与半成品在制信息合并到 summary_df 中，并返回未匹配的主键。

    参数：
    - summary_df: 汇总表（含“晶圆品名”，“规格”，“品名”）
    - product_in_progress_df: 透视后的“赛卓-成品在制”数据
    - mapping_df: 新旧料号映射表，包含“半成品”列

    返回：
    - summary_df: 添加了“成品在制”与“半成品在制”的 DataFrame
    - unmatched_keys: list of (晶圆品名, 规格, 品名) 未被使用的原始行主键
    """
    numeric_cols = product_in_progress_df.select_dtypes(include='number').columns.tolist()
    summary_df = summary_df.copy()
    summary_df["成品在制"] = 0
    summary_df["半成品在制"] = 0

    used_keys = set()
    unmatched_keys = []

    # 填充成品在制
    for idx, row in product_in_progress_df.iterrows():
        key = (str(row["晶圆型号"]).strip(), str(row["产品规格"]).strip(), str(row["产品品名"]).strip())
        mask = (
            (summary_df["晶圆品名"] == key[0]) &
            (summary_df["规格"] == key[1]) &
            (summary_df["品名"] == key[2])
        )
        if mask.any():
            used_keys.add(key)
            summary_df.loc[mask, "成品在制"] = row[numeric_cols].sum()
        else:
            unmatched_keys.append(key)

    # 半成品逻辑
    semi_rows = mapping_df[mapping_df["半成品"].notna() & (mapping_df["半成品"] != "")]
    semi_info_table = semi_rows[[
        "新规格", "新品名", "新晶圆品名",
        "旧规格", "旧品名", "旧晶圆品名",
        "半成品"
    ]].copy()
    semi_info_table["未交数据和"] = 0

    # 日志记录每一次匹配尝试
    check_log = []

    for idx, row in semi_info_table.iterrows():
        matched = product_in_progress_df[
            (product_in_progress_df["产品规格"] == row["新规格"]) &
            (product_in_progress_df["晶圆型号"] == row["新晶圆品名"]) &
            (product_in_progress_df["产品品名"] == row["半成品"])
        ]
        if not matched.empty:
            value = matched[numeric_cols].sum().sum()
            source = "新规格/新晶圆匹配"
        else:
            matched = product_in_progress_df[
                (product_in_progress_df["产品规格"] == row["旧规格"]) &
                (product_in_progress_df["晶圆型号"] == row["旧晶圆品名"]) &
                (product_in_progress_df["产品品名"] == row["半成品"])
            ]
            if not matched.empty:
                value = matched[numeric_cols].sum().sum()
                source = "旧规格/旧晶圆回退匹配"
            else:
                value = 0
                source = "未匹配"

        semi_info_table.at[idx, "未交数据和"] = value
        check_log.append({
            "半成品": row["半成品"],
            "新品名": row["新品名"],
            "新规格": row["新规格"],
            "新晶圆品名": row["新晶圆品名"],
            "旧规格": row["旧规格"],
            "旧晶圆品名": row["旧晶圆品名"],
            "匹配来源": source,
            "匹配值": value
        })

    # 写入 summary_df
    for idx, row in semi_info_table.iterrows():
        key = (row["新晶圆品名"], row["新规格"], row["新品名"])
        mask = (
            (summary_df["晶圆品名"] == key[0]) &
            (summary_df["规格"] == key[1]) &
            (summary_df["品名"] == key[2])
        )
        if mask.any():
            summary_df.loc[mask, "半成品在制"] = row["未交数据和"]
            used_keys.add(key)
        else:
            unmatched_keys.append(key)

    # 打印匹配日志
    st.write("【半成品匹配日志】")
    for log in check_log:
        st.write(log)

    return summary_df, unmatched_keys
