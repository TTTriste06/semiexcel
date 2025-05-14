import streamlit as st

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎来到我的应用")
        st.markdown('---')
        st.markdown('### 功能简介：')
        st.markdown('- 上传多个 Excel 表格')
        st.markdown('- 实时生成透视汇总表')
        st.markdown('- 一键导出 Excel 汇总报告')

def get_uploaded_files():
    st.markdown("### 📤 请上传以下 5 个 Excel 文件：")
    expected_files = {
        "赛卓-未交订单.xlsx",
        "赛卓-成品在制.xlsx",
        "赛卓-CP在制.xlsx",
        "赛卓-成品库存.xlsx",
        "赛卓-晶圆库存.xlsx"
    }

    uploaded_file_list = st.file_uploader(
        "上传 5 个 Excel 文件",
        type=["xlsx"],
        accept_multiple_files=True
    )

    uploaded_files = {f.name: f for f in uploaded_file_list} if uploaded_file_list else {}

    missing_files = expected_files - uploaded_files.keys()
    if missing_files:
        st.warning(f"⚠️ 缺少文件: {', '.join(missing_files)}")

    st.markdown("### 📊 上传辅助文件（可选，若不上传则使用历史版本）")
    forecast_file = st.file_uploader("📈 赛卓-预测.xlsx", type=["xlsx"], key="forecast")
    safety_stock_file = st.file_uploader("🛡️ 赛卓-安全库存.xlsx", type=["xlsx"], key="safety")
    mapping_file = st.file_uploader("🔁 赛卓-新旧料号.xlsx", type=["xlsx"], key="mapping")

    # ✅ 动态生成未交订单的月份选择框
    if "赛卓-未交订单.xlsx" in uploaded_dict:
        try:
            df_unfulfilled = pd.read_excel(uploaded_dict["赛卓-未交订单.xlsx"])
            pivot_config = CONFIG["pivot_config"].get("赛卓-未交订单.xlsx")
            if pivot_config and "date_format" in pivot_config:
                from pivot_processor import process_date_column
                from month_selector import extract_months_from_columns
    
                df_unfulfilled = process_date_column(df_unfulfilled, pivot_config["columns"], pivot_config["date_format"])
                months = extract_months_from_columns(df_unfulfilled.columns)
    
                selected = st.selectbox("📅 选择历史数据截止月份（不选视为全部）", ["全部"] + months, index=0)
                CONFIG["selected_month"] = None if selected == "全部" else selected
        except Exception as e:
            st.error(f"❌ 无法识别未交订单中的月份列：{e}")


    st.markdown("---")
    start = st.button("🚀 生成汇总报告")
    return uploaded_files, forecast_file, safety_stock_file, mapping_file, start
