import streamlit as st
import pandas as pd

st.set_page_config(page_title="超級診斷工具", layout="wide")
st.title("🔍 FA300 欄位與資料透視診斷器")

file_fa300 = st.file_uploader("請上傳 FA300 檔案 (Excel)", type=["xlsx", "xls"])

if file_fa300:
    try:
        df = pd.read_excel(file_fa300)
        st.success(f"讀取成功！共 {len(df)} 筆資料，{len(df.columns)} 個欄位。")
        
        st.subheader("1. 實際欄位清單 (包含隱藏字元)：")
        col_info = [{"欄位索引": i, "欄位名稱": f"'{col}'", "資料型態": str(df[col].dtype)} for i, col in enumerate(df.columns)]
        st.table(pd.DataFrame(col_info))
        
        st.subheader("2. 前 5 筆實際內容範例：")
        st.dataframe(df.head(5))
        
        st.subheader("3. 各欄位前 3 個不重複值 (觀測資料長怎樣)：")
        sample_dict = {}
        for col in df.columns:
            sample_dict[str(col)] = list(df[col].dropna().unique()[:3])
        st.json(sample_dict)

    except Exception as e:
        st.error(f"讀取失敗：{e}")
