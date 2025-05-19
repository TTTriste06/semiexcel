### ✅ ui.py — 支持中文文件名上传（适配 Streamlit Cloud）

import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import base64
import json
from config import CONFIG
from memory_manager import clean_memory, display_debug_memory_stats
from io import BytesIO

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎使用数据汇总工具")
        st.markdown("---")
        st.markdown("### 功能简介：")
        st.markdown("- 上传 5 个主数据表（支持中文文件名）")
        st.markdown("- 上传辅助数据（预测、安全库存、新旧料号）")
        st.markdown("- 自动生成汇总 Excel 文件")

        st.markdown("### 🧹 内存与资源管理")
        if st.button("清理内存"):
            clean_memory()
        if st.button("查看内存使用排行"):
            display_debug_memory_stats()

def get_uploaded_files():
    st.header("📄 Excel 数据处理与汇总")

    manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    CONFIG["selected_month"] = manual_month.strip() if manual_month.strip() else None

    st.markdown("### 🔽 上传 5 个核心 Excel 文件（支持中文文件名）")

    uploaded_json = components.html("""
    <!DOCTYPE html>
    <html>
    <body>
      <input type=\"file\" id=\"uploader\" multiple />
      <p id=\"status\"></p>
      <script>
        window.Streamlit = window.parent.Streamlit;
        const uploader = document.getElementById('uploader');
        const status = document.getElementById('status');

        uploader.addEventListener('change', () => {
          const files = uploader.files;
          const results = [];
          let completed = 0;

          for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const reader = new FileReader();
            reader.onload = () => {
              const base64 = reader.result.split(',')[1];
              results.push({ name: file.name, content: base64 });
              completed++;
              if (completed === files.length) {
                const payload = JSON.stringify(results);
                Streamlit.setComponentValue(payload);
              }
            };
            reader.readAsDataURL(file);
          }
        });
      </script>
    </body>
    </html>
    """, height=220, key="custom-uploader")

    core_files = []
    if isinstance(uploaded_json, str):
        try:
            file_objs = json.loads(uploaded_json)
            core_files = [(f["name"], base64.b64decode(f["content"])) for f in file_objs]
            st.success(f"✅ 成功上传 {len(core_files)} 个核心文件")
        except Exception as e:
            st.error(f"❌ 上传失败：{e}")
    else:
        st.info("📅 请上传 Excel 文件...")

    for i, (fname, _) in enumerate(core_files):
        st.write(f"📄 文件 {i+1}: `{fname}`")

    st.markdown("### 🔁 上传辅助数据文件（预测、安全库存、新旧料号）")
    forecast_file = st.file_uploader("📈 上传预测文件", type=["xlsx"])
    safety_file = st.file_uploader("🛡️ 上传安全库存文件", type=["xlsx"])
    mapping_file = st.file_uploader("🔀 上传新旧料号文件", type=["xlsx"])

    start = st.button("🚀 点击生成汇总 Excel 文件")

    return core_files, forecast_file, safety_file, mapping_file, start
