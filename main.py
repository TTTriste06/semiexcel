import streamlit as st
import pandas as pd
from pivot_processor import PivotProcessor
from ui import setup_sidebar, get_uploaded_files
from io import BytesIO
from datetime import datetime

def main():
    st.set_page_config(page_title="Excel数据透视汇总工具", layout="wide")
    setup_sidebar()

    uploaded_files, start = get_uploaded_files()

    if start:
        if len(uploaded_files) < 5:
            st.error("❌ 请上传所有 5 个文件后再点击生成！")
            return

        processor = PivotProcessor()
        buffer = BytesIO()
        processor.process(uploaded_files, buffer)


        file_ts_name = f"运营数据订单-在制-库存汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.success("✅ 汇总完成！你可以下载结果文件：")
        st.download_button(
            label="📥 下载 Excel 汇总报告",
            data=buffer.getvalue(),
            file_name=file_ts_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()

