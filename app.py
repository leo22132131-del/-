import streamlit as st
import pandas as pd

st.set_page_config(page_title="FA300 欄位診斷工具", layout="wide")
st.title("🔍 FA300 欄位結構與資料診斷")

file_fa300 = st.file_uploader("請上傳您的 FA300 (Excel)", type=["xlsx", "xls"])

if file_fa300:
    try:
        df_fa300 = pd.read_excel(file_fa300)
        
        st.success(f"成功讀取 FA300！總行數：{len(df_fa300)} 行")
        
        st.subheader("1. FA300 的所有欄位名稱：")
        st.write(list(df_fa300.columns))
        
        st.subheader("2. FA300 前 10 筆原始資料預覽：")
        st.dataframe(df_fa300.head(10), use_container_width=True)
        
    except Exception as e:
        st.error(f"讀取失敗：{e}")
