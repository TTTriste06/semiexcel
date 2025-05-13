import streamlit as st
import pandas as pd

st.set_page_config(page_title='数据汇总自动化工具', layout='wide')
st.write("✅ 已加载主页面配置")

from openpyxl import load_workbook
from config import (
    GITHUB_TOKEN_KEY, REPO_NAME, BRANCH,
    CONFIG, OUTPUT_FILE, PIVOT_CONFIG,
    FULL_MAPPING_COLUMNS, COLUMN_MAPPING
)
st.write("✅ 已导入配置")

from github_utils import upload_to_github, download_excel_from_repo
from ui import setup_sidebar, get_user_inputs
st.write("✅ 已导入 github_utils 和 ui")

def main():
    st.set_page_config(page_title='数据汇总自动化工具', layout='wide')
    setup_sidebar()

    # 获取用户上传
    uploaded_files, pred_file, safety_file, mapping_file = get_user_inputs()

    # 加载文件
    mapping_df = None
    safety_df = None
    pred_df = None
    if safety_file:
        safety_df = pd.read_excel(safety_file)
        upload_to_github(safety_file, "safety_file.xlsx", "上传安全库存文件")
    else:
        safety_df = download_excel_from_repo("safety_file.xlsx")
    st.write(safety_df)
    if pred_file:
        pred_df = pd.read_excel(pred_file)
        upload_to_github(pred_file, "pred_file.xlsx", "上传预测文件")
    else:
        pred_df = download_excel_from_repo("pred_file.xlsx")
    st.write(pred_df)
    if mapping_file:
        mapping_df = pd.read_excel(mapping_file)
        upload_to_github(mapping_file, "mapping_file.xlsx", "上传新旧料号文件")
    else:
        mapping_df = download_excel_from_repo("mapping_file.xlsx")
    st.write(mapping_df)
    


if __name__ == '__main__':
    main()
