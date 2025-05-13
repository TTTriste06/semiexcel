import streamlit as st
import pandas as pd
from openpyxl import load_workbook

from ui import setup_sidebar, get_user_inputs
from config import (
    GITHUB_TOKEN_KEY, REPO_NAME, BRANCH,
    CONFIG, OUTPUT_FILE, PIVOT_CONFIG,
    FULL_MAPPING_COLUMNS, COLUMN_MAPPING
)
from github_utils import upload_to_github, download_excel_from_repo
from prepare import apply_full_mapping

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
    if pred_file:
        pred_df = pd.read_excel(pred_file)
        upload_to_github(pred_file, "pred_file.xlsx", "上传预测文件")
    else:
        pred_df = download_excel_from_repo("pred_file.xlsx")
    if mapping_file:
        mapping_df = pd.read_excel(mapping_file)
        upload_to_github(mapping_file, "mapping_file.xlsx", "上传新旧料号文件")
    else:
        mapping_df = download_excel_from_repo("mapping_file.xlsx")

    if st.button('🚀 提交并生成报告') and uploaded_files:
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            summary_df = pd.DataFrame()
            pending_df = None

            # 处理 uploaded_files
            for f in uploaded_files:
                filename = f.name
                if filename not in PIVOT_CONFIG:
                    st.warning(f"跳过未配置的文件: {filename}")
                    continue

                df = pd.read_excel(f)

                # 替换新旧料号
                if filename in COLUMN_MAPPING:
                    mapping = COLUMN_MAPPING[filename]
                    spec_col, prod_col, wafer_col = mapping["规格"], mapping["品名"], mapping["晶圆品名"]
                    if all(col in df.columns for col in [spec_col, prod_col, wafer_col]):
                        df = apply_full_mapping(df, mapping_df, spec_col, prod_col, wafer_col)
                    else:
                        st.warning(f"⚠️ 文件 {filename} 缺少字段: {spec_col}, {prod_col}, {wafer_col}")
                else:
                    st.info(f"📂 文件 {filename} 未定义映射字段，跳过 apply_full_mapping")

                # 透视表处理
                pivot_config = PIVOT_CONFIG[filename]
                pivoted = create_pivot(df, pivot_config, filename, mapping_df)
    
                # 写入 Excel（sheet name 去掉 .xlsx 后缀）
                sheet_name = filename.replace(".xlsx", "")
                pivoted.to_excel(writer, sheet_name=sheet_name)
    
                st.success(f"📊 已处理并写入: {sheet_name}")
    
            st.success("✅ 所有文件处理完毕，正在生成报告...")
    
        # 下载按钮
        with open(OUTPUT_FILE, "rb") as f:
            st.download_button("📥 下载汇总报告", f, file_name=OUTPUT_FILE)


    


if __name__ == '__main__':
    main()
