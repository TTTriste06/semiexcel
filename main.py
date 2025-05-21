import streamlit as st
from io import BytesIO
from datetime import datetime
import pandas as pd
from pivot_processor import PivotProcessor
from ui import setup_sidebar, get_uploaded_files
from github_utils import upload_to_github, download_from_github, load_or_fallback_from_github
from urllib.parse import quote

def main():
    st.set_page_config(page_title="Excel数据透视汇总工具", layout="wide")
    setup_sidebar()

    if "started" not in st.session_state:
        st.session_state.started = False

    if not st.session_state.started:
        uploaded_files, forecast_file, safety_file, mapping_file, start = get_uploaded_files()

        if start:
            if len(uploaded_files) < 5:
                st.error("❌ 请上传所有 5 个主要文件后再点击生成！")
                return

            st.session_state.started = True
            st.session_state.uploaded_files = uploaded_files
            st.session_state.forecast_file = forecast_file
            st.session_state.safety_file = safety_file
            st.session_state.mapping_file = mapping_file
            st.rerun()

    else:
        uploaded_files = st.session_state.uploaded_files
        forecast_file = st.session_state.forecast_file
        safety_file = st.session_state.safety_file
        mapping_file = st.session_state.mapping_file

        additional_sheets = {}

        load_or_fallback_from_github("新旧料号", "mapping_file", "赛卓-新旧料号.xlsx", additional_sheets)
        load_or_fallback_from_github("安全库存", "safety_file", "赛卓-安全库存.xlsx", additional_sheets)
        load_or_fallback_from_github("预测", "forecast_file", "赛卓-预测.xlsx", additional_sheets)

        buffer = BytesIO()
        processor = PivotProcessor()
        processor.process(uploaded_files, buffer, additional_sheets)

        file_name = f"运营数据订单-在制-库存汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.success("✅ 汇总完成！你可以下载结果文件：")
        st.download_button(
            label="📥 下载 Excel 汇总报告",
            data=buffer.getvalue(),
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
