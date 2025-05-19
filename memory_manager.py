import gc
import streamlit as st
import tracemalloc

def clean_memory(variables: list = []):
    """
    手动清理指定变量 + 执行垃圾回收。
    """
    for var in variables:
        if var in globals():
            del globals()[var]
        elif var in locals():
            del locals()[var]
    gc.collect()
    st.success("🧹 内存清理完毕！")


def memory_debug_top_stats(n=5):
    """
    显示当前内存占用最多的前 n 行代码（用于调试）
    """
    tracemalloc.start()
    snapshot = tracemalloc.take_snapshot()
    top_stats = snapshot.statistics('lineno')
    return top_stats[:n]


def display_debug_memory_stats(n=5):
    """
    Streamlit 输出内存使用情况
    """
    st.markdown("### 💾 当前内存使用排行")
    top_stats = memory_debug_top_stats(n)
    for stat in top_stats:
        st.code(str(stat))
