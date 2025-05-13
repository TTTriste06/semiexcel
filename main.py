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
        wrote_any_sheet = False  # 标志：是否至少写入了一个有效 sheet
    
        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            st.write("🚧 uploaded_files =", uploaded_files)
    
            for f in uploaded_files:
                filename = f.name
                st.write(f"📂 正在处理文件: {filename}")
    
                if filename not in PIVOT_CONFIG:
                    st.write("✅ 上传的文件名：", [f.name for f in uploaded_files])
                    st.write("✅ PIVOT_CONFIG keys:", list(PIVOT_CONFIG.keys()))
    
                    st.warning(f"⚠️ 跳过未配置的文件: {filename}")
                    continue
    
                try:
                    df = pd.read_excel(f)
                except Exception as e:
                    st.error(f"❌ 无法读取 {filename}: {e}")
                    continue
    
                # 映射料号替换
                if filename in COLUMN_MAPPING:
                    mapping = COLUMN_MAPPING[filename]
                    spec_col = mapping["规格"]
                    prod_col = mapping["品名"]
                    wafer_col = mapping["晶圆品名"]
    
                    missing_cols = [col for col in [spec_col, prod_col, wafer_col] if col not in df.columns]
                    if missing_cols:
                        st.warning(f"⚠️ 文件 {filename} 缺少必要列: {missing_cols}")
                        st.write("实际列:", df.columns.tolist())
                        st.write("映射要求:", spec_col, prod_col, wafer_col)
    
                        continue
    
                    df = apply_full_mapping(df, mapping_df, spec_col, prod_col, wafer_col)
                else:
                    st.info(f"ℹ️ 文件 {filename} 未定义映射字段，跳过 apply_full_mapping")
    
                # 创建透视表
                pivot_config = PIVOT_CONFIG[filename]
                pivoted = create_pivot(df, config, filename, mapping_df)
    
                if pivoted is not None and not pivoted.empty:
                    st.write(f"✅ {filename} 透视表生成成功，行数: {pivoted.shape[0]}")
                    pivoted.to_excel(writer, sheet_name=sheet_name)
                    wrote_any_sheet = True
                else:
                    st.warning(f"⚠️ {filename} 的透视表为空，未写入")
    
    
                if pivoted is None or pivoted.empty:
                    st.warning(f"⚠️ 文件 {filename} 的透视结果为空，未写入 Excel")
                    continue
    
                sheet_name = filename.replace(".xlsx", "")[:31]  # Excel 限制 sheet 名最多 31 字符
                pivoted.to_excel(writer, sheet_name=sheet_name)
                wrote_any_sheet = True
                st.success(f"✅ 写入 sheet: {sheet_name}，共 {pivoted.shape[0]} 行")
    
            # 如果一个有效 sheet 都没有写入，添加保底空页防止崩溃
            if not wrote_any_sheet:
                st.warning("⚠️ 所有文件都未处理成功，写入空白页避免报错")
                pd.DataFrame({"提示": ["未处理任何有效数据"]}).to_excel(writer, sheet_name="无数据")
    
        # 下载按钮
        with open(OUTPUT_FILE, "rb") as f:
            st.download_button("📥 下载汇总报告", f, file_name=OUTPUT_FILE)
    
    
        


if __name__ == '__main__':
    main()
