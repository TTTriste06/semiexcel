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

        # GitHub 上传/下载文件的目标名称
        github_targets = {
            forecast_file.name if forecast_file else "赛卓-预测.xlsx": "赛卓-预测.xlsx",
            safety_file.name if safety_file else "赛卓-安全库存.xlsx": "赛卓-安全库存.xlsx",
            mapping_file.name if mapping_file else "赛卓-新旧料号.xlsx": "赛卓-新旧料号.xlsx"
        }

        # 文件对象
        local_files = {
            forecast_file.name if forecast_file else "赛卓-预测.xlsx": forecast_file,
            safety_file.name if safety_file else "赛卓-安全库存.xlsx": safety_file,
            mapping_file.name if mapping_file else "赛卓-新旧料号.xlsx": mapping_file
        }

        additional_sheets = {}

        for upload_name, github_name in github_targets.items():
            file = local_files[upload_name]

            st.write(f"处理文件: {upload_name}（GitHub 名: {github_name}）")

            if file:  # 用户上传了新文件
                file_bytes = file.read()
                file_io = BytesIO(file_bytes)

                upload_to_github(BytesIO(file_bytes), github_name)

                df = pd.read_excel(file_io)
                additional_sheets[upload_name.replace(".xlsx", "")] = df
            else:
                try:
                    content = download_from_github(github_name)
                    df = pd.read_excel(BytesIO(content))
                    additional_sheets[upload_name.replace(".xlsx", "")] = df
                    st.info(f"📂 使用了 GitHub 上的历史版本：{github_name}")
                except FileNotFoundError:
                    st.warning(f"⚠️ 未提供且 GitHub 上找不到：{github_name}")

        # 生成 Excel 汇总
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
