import streamlit as st
import pandas as pd
import base64
from config import CONFIG
from memory_manager import clean_memory, display_debug_memory_stats
from io import BytesIO

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎使用数据汇总工具")
        st.markdown("---")
        st.markdown("### 功能简介：")
        st.markdown("- 上传 5 个主数据表（将自动重命名）")
        st.markdown("- 上传辅助数据（预测、安全库存、新旧料号）")
        st.markdown("- 自动生成汇总 Excel 文件")

        st.markdown("### 🧹 内存与资源管理")
        if st.button("清理内存"):
            clean_memory()
        if st.button("查看内存使用排行"):
            display_debug_memory_stats()

def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")

    manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    CONFIG["selected_month"] = manual_month.strip() if manual_month.strip() else None

    st.markdown("### 🔽 上传 5 个核心 Excel 文件（中文文件名将自动重命名）")

    uploaded_core_files = st.file_uploader(
        "上传 5 个核心文件",
        type=["xlsx"],
        accept_multiple_files=True
    )

    processed_files = []
    if uploaded_core_files:
        for idx, file in enumerate(uploaded_core_files):
            try:
                content = file.read()
                fake_name = f"core_{idx+1}.xlsx"  # 避免中文名
                processed_files.append((fake_name, content))
                st.write(f"📄 文件 {idx+1}: 原名 `{file.name}` → 存储名 `{fake_name}`")
            except Exception as e:
                st.error(f"❌ 无法读取文件 {file.name}: {e}")

    st.markdown("### 🔁 上传辅助数据文件（预测、安全库存、新旧料号）")
    forecast_file = st.file_uploader("📈 上传预测文件", type=["xlsx"])
    safety_file = st.file_uploader("🛡️ 上传安全库存文件", type=["xlsx"])
    mapping_file = st.file_uploader("🔁 上传新旧料号文件", type=["xlsx"])

    start = st.button("🚀 点击生成汇总 Excel 文件")

    return processed_files, forecast_file, safety_file, mapping_file, start
