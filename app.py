import streamlit as st
import pandas as pd

st.set_page_config(page_title="FA300 專用排錯器", layout="wide")
st.title("🕵️‍♂️ FA300 抓鬼除錯診斷")

file_fa300 = st.file_uploader("請上傳您的 FA300 檔案 (Excel)", type=["xlsx", "xls"])

if file_fa300:
    try:
        # 1. 檢查有哪些 Sheet
        excel_file = pd.ExcelFile(file_fa300)
        st.write(f"📄 **檔案中的頁籤 (Sheets)**：{excel_file.sheet_names}")
        
        # 讀取第一個 Sheet
        df = pd.read_excel(file_fa300, sheet_name=0)
        
        st.success(f"成功讀取！總共有 **{len(df)}** 行資料，欄位數為 **{len(df.columns)}**")
        
        st.subheader("1. 欄位名稱列表：")
        st.write(list(df.columns))
        
        st.subheader("2. 檔案前 10 列實際內容：")
        st.dataframe(df.head(10))
        
        st.subheader("3. 『服務項目』欄位裡面的前 15 筆獨特值：")
        # 尋找可能包含項目的欄位
        item_col = None
        for col in df.columns:
            if '項目' in str(col) or '碼' in str(col):
                item_col = col
                break
        if item_col:
            st.write(f"抓取欄位 **【{item_col}】** 的前 15 種不同內容：")
            st.write(df[item_col].dropna().unique()[:15].tolist())
        else:
            st.warning("找不到名稱帶有 '項目' 或 '碼' 的欄位！")

    except Exception as e:
        st.error(f"❌ 讀取檔案時發生致命錯誤：{e}")
