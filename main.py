import streamlit as st
from io import BytesIO
from datetime import datetime
import pandas as pd
from pivot_processor import PivotProcessor
from ui import setup_sidebar, get_uploaded_files
from github_utils import upload_to_github, download_from_github
from urllib.parse import quote

def main():
    st.set_page_config(page_title="Excel数据透视汇总工具", layout="wide")
    setup_sidebar()

    # 获取上传文件
    uploaded_files, forecast_file, safety_file, mapping_file, start = get_uploaded_files()

    if start:
        if len(uploaded_files) < 5:
            st.error("❌ 请上传所有 5 个主要文件后再点击生成！")
            return

        # GitHub 辅助文件名称
        github_files = {
            "赛卓-预测.xlsx": forecast_file,
            "赛卓-安全库存.xlsx": safety_file,
            "赛卓-新旧料号.xlsx": mapping_file
        }

        additional_sheets = {}

        for name, file in github_files.items():
            if file:  # 如果上传了新文件，则保存到 GitHub
                file_bytes = file.read()
                file_io = BytesIO(file_bytes)
                
                # 对中文文件名进行 URL 编码，避免 GitHub 报 400
                safe_name = quote(name)

                # 上传使用编码后的文件名
                upload_to_github(BytesIO(file_bytes), safe_name)

                # 保留原始名字作为字典 key
                df = pd.read_excel(file_io)
                additional_sheets[name.replace(".xlsx", "")] = df
            else:
                try:
                    # 下载时也编码文件名
                    safe_name = quote(name)
                    content = download_from_github(safe_name)

                    df = pd.read_excel(BytesIO(content))
                    additional_sheets[name.replace(".xlsx", "")] = df
                    st.info(f"📂 使用了 GitHub 上存储的历史版本：{name}")
                except FileNotFoundError:
                    st.warning(f"⚠️ 未提供且未在 GitHub 找到历史文件：{name}")
              
        # 生成 Excel 汇总
        buffer = BytesIO()
        processor = PivotProcessor()
        processor.process(uploaded_files, buffer, additional_sheets)

        # 提供下载按钮
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

