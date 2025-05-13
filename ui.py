import streamlit as st

def setup_sidebar():
    with st.sidebar:
        st.title("欢迎来到我的应用")
        st.markdown('---')
        st.markdown('### 功能简介：')
        st.markdown('- 生成汇总报告')

