import streamlit as st
import pandas as pd
from openpyxl import load_workbook


def main():
    
        # 下载按钮
        with open(OUTPUT_FILE, 'rb') as f:
            st.download_button('📥 下载汇总报告', f, OUTPUT_FILE)

if __name__ == '__main__':
    main()
