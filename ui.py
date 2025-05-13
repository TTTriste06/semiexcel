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

    st.markdown("---")
    start = st.button("🚀 生成汇总报告")
    return uploaded_files, start
