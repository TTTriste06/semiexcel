import streamlit as st
import pandas as pd
import streamlit.components.v1 as components
import base64
import json
from config import CONFIG
from memory_manager import clean_memory, display_debug_memory_stats

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎使用数据汇总工具")
        st.markdown("---")
        st.markdown("### 功能简介：")
        st.markdown("- 上传 5 个主数据表")
        st.markdown("- 上传辅助数据（预测、安全库存、新旧料号）")
        st.markdown("- 自动生成汇总 Excel 文件")

    with st.sidebar:
        st.markdown("### 🧹 内存与资源管理")
        if st.button("清理内存"):
            clean_memory()
        if st.button("查看内存使用排行"):
            display_debug_memory_stats()

def get_uploaded_files():
    st.header("📤 Excel 数据处理与汇总")

    manual_month = st.text_input("📅 输入历史数据截止月份（格式: YYYY-MM，可留空表示不筛选）")
    CONFIG["selected_month"] = manual_month.strip() if manual_month.strip() else None

    st.markdown("### 🔽 上传 5 个核心 Excel 文件（未交订单/成品在制/成品库存/晶圆库存/CP在制）")
    
    if "uploaded_core_files" not in st.session_state:
        st.session_state.uploaded_core_files = []

    # 自定义上传组件（支持中文）
    components.html("""
        <script>
          window.Streamlit = window.parent.Streamlit;
        </script>
        <input type="file" id="uploader" multiple />
        <p id="status"></p>
        <script>
          const uploader = document.getElementById('uploader');
          const status = document.getElementById('status');
    
          uploader.onchange = () => {
            const files = uploader.files;
            const results = [];
    
            const readFile = (file, index) => {
              const reader = new FileReader();
              reader.onload = () => {
                const base64 = reader.result.split(',')[1];
                results.push({ name: file.name, content: base64 });
                if (results.length === files.length) {
                  const payload = JSON.stringify(results);
                  Streamlit.setComponentValue(payload);
                }
              };
              reader.readAsDataURL(file);
            };
    
            for (let i = 0; i < files.length; i++) {
              readFile(files[i], i);
            }
          };
        </script>
    """, height=150, key="core-uploader")


    uploaded = st._legacy_get_component_value("core-uploader")
    if uploaded:
        try:
            decoded_files = []
            for item in json.loads(uploaded):
                filename = item["name"]
                content = base64.b64decode(item["content"])
                decoded_files.append((filename, content))
            st.session_state.uploaded_core_files = decoded_files
            st.success(f"✅ 成功上传 {len(decoded_files)} 个文件")
        except Exception as e:
            st.error(f"❌ 上传失败：{e}")

    # 显示上传内容
    for i, (fname, _) in enumerate(st.session_state.uploaded_core_files):
        st.write(f"📄 文件 {i+1}: `{fname}`")

    # 辅助文件上传（预测、安全库存、新旧料号）
    forecast_file = st.file_uploader("📈 上传预测文件", type=["xlsx"])
    safety_file = st.file_uploader("🛡️ 上传安全库存文件", type=["xlsx"])
    mapping_file = st.file_uploader("🔁 上传新旧料号文件", type=["xlsx"])

    start = st.button("🚀 点击生成汇总 Excel 文件")
    return st.session_state.uploaded_core_files, forecast_file, safety_file, mapping_file, start
