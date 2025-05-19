import streamlit as st
import pandas as pd
from config import CONFIG
from memory_manager import clean_memory, display_debug_memory_stats
from pypinyin import lazy_pinyin
import uuid
import re

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎使用数据汇总工具")
        st.markdown("---")
        st.markdown("### 功能简介：")
        st.markdown("- 上传 5 个主数据表")
        st.markdown("- 上传辅助数据（预测、安全库存、新旧料号）")
        st.markdown("- 自动生成汇总 Excel 文件")

        st.markdown("### 🧹 内存与资源管理")
        if st.button("清理内存"):
            clean_memory()
        if st.button("查看内存使用排行"):
            display_debug_memory_stats()


def convert_filename_to_ascii(original_name: str) -> str:
    # 去除路径和扩展名，只取主名部分
    name_part = re.sub(r'\.[^.]+$', '', original_name)
    safe_part = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '_', name_part)
    pinyin = "_".join(lazy_pinyin(safe_part)).strip("_")
    if not pinyin:  # fallback 防止空字符串
        pinyin = uuid.uuid4().hex
    return pinyin + ".xlsx"


def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")

    manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    if manual_month.strip():
        CONFIG["selected_month"] = manual_month.strip()
        st.write(f"✅ 当前选定月份：{CONFIG['selected_month']}")
    else:
        CONFIG["selected_month"] = None

    uploaded_files = st.file_uploader(
        "📂 上传 5 个核心 Excel 文件（未交订单/成品在制/成品库存/晶圆库存/CP在制）",
        type=["xlsx"],
        accept_multiple_files=True,
        key="main_files"
    )

    file_mapping = {}
    if uploaded_files:
        for file in uploaded_files:
            ascii_name = convert_filename_to_ascii(file.name)
            file_mapping[ascii_name] = {
                "file": file,
                "original_name": file.name
            }

        st.write("✅ 文件名替换映射如下：")
        for ascii_name, info in file_mapping.items():
            st.write(f"- 原始: {info['original_name']} → 替换: {ascii_name}")

    # 上传额外辅助文件
    st.subheader("📁 上传额外文件（可用储存的文件）")
    forecast_file = st.file_uploader("📈 上传预测文件", type="xlsx", key="forecast")
    safety_file = st.file_uploader("🔐 上传安全库存文件", type="xlsx", key="safety")
    mapping_file = st.file_uploader("🔁 上传新旧料号对照表", type="xlsx", key="mapping")

    start = st.button("🚀 生成汇总 Excel")

    return file_mapping, forecast_file, safety_file, mapping_file, start
