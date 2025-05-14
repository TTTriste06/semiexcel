import streamlit as st
import pandas as pd
from month_selector import extract_months_from_columns
from pivot_processor import process_date_column
from config import CONFIG

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎使用数据汇总工具")
        st.markdown("---")
        st.markdown("### 功能简介：")
        st.markdown("- 上传 5 个主数据表")
        st.markdown("- 上传辅助数据（预测、安全库存、新旧料号）")
        st.markdown("- 自动生成汇总 Excel 文件")

def get_uploaded_files():
    st.header("📤 上传数据文件")
    uploaded_files = st.file_uploader(
        "请上传 5 个主数据文件（未交订单、成品在制、成品库存、晶圆库存、CP在制）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="main_files"
    )

    uploaded_dict = {}
    for file in uploaded_files:
        uploaded_dict[file.name] = file

    # 输出上传文件名调试
    st.write("✅ 已上传文件名：", list(uploaded_dict.keys()))

    # 额外上传的 3 个文件
    st.subheader("📁 上传额外文件（可选）")
    forecast_file = st.file_uploader("赛卓-预测.xlsx", type="xlsx", key="forecast")
    safety_file = st.file_uploader("赛卓-安全库存.xlsx", type="xlsx", key="safety")
    mapping_file = st.file_uploader("赛卓-新旧料号.xlsx", type="xlsx", key="mapping")

    # 动态生成未交订单的月份选择框（如果已上传）
    if "赛卓-未交订单.xlsx" in uploaded_dict:
        try:
            df_unfulfilled = pd.read_excel(uploaded_dict["赛卓-未交订单.xlsx"])
            st.write("✅ 未交订单数据加载成功，前几列：")
            st.write(df_unfulfilled.head())

            pivot_config = CONFIG["pivot_config"].get("赛卓-未交订单.xlsx")
            st.write("✅ 未交订单配置：", pivot_config)

            if pivot_config and "date_format" in pivot_config:
                df_unfulfilled = process_date_column(df_unfulfilled, pivot_config["columns"], pivot_config["date_format"])
                st.write("✅ 处理后列名：", df_unfulfilled.columns.tolist())

                months = extract_months_from_columns(df_unfulfilled.columns)
                st.write("📅 提取到的月份有：", months)

                selected = st.selectbox("📅 选择历史数据截止月份（不选视为全部）", ["全部"] + months, index=0)
                CONFIG["selected_month"] = None if selected == "全部" else selected
            else:
                st.warning("⚠️ 未交订单配置中缺少 date_format 字段")
        except Exception as e:
            st.error(f"❌ 无法处理未交订单月份：{e}")

    start = st.button("🚀 生成汇总 Excel")
    return uploaded_dict, forecast_file, safety_file, mapping_file, start
