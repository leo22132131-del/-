import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="FA300 深入診斷工具", layout="wide")
st.title("🔍 FA300 資料轉換與比對全流程診斷")

file_fa300 = st.file_uploader("請上傳 FA300 (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text): return ""
    s = str(text).upper().strip()
    match = re.search(r'([A-Z]{2}\d{2})', s)
    if match: return match.group(1)
    return ""

def clean_date(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == 'nan': 
        return ""
    s = str(val).strip().split(' ')[0].split('.')[0]
    
    if s.isdigit() and len(s) == 5:
        try: return pd.to_datetime(int(s), unit='D', origin='1899-12-30').strftime('%Y-%m-%d')
        except: pass
        
    if s.isdigit() and len(s) == 7:
        y = int(s[:3]) + 1911
        return f"{y:04d}-{int(s[3:5]):02d}-{int(s[5:7]):02d}"
        
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 1000: y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        except: pass
        
    try:
        dt = pd.to_datetime(s)
        if dt.year < 1920: dt = dt.replace(year=dt.year + 1911)
        return dt.strftime('%Y-%m-%d')
    except: return s

if file_fa300:
    try:
        df = pd.read_excel(file_fa300)
        st.write(f"📊 **1. 原始檔案總行數**：{len(df)} 筆")
        
        # 顯示原始前 5 筆
        st.subheader("2. 原始資料前 5 筆：")
        st.dataframe(df.head(5), use_container_width=True)
        
        # 轉換
        df['轉換後日期'] = df['服務日期'].apply(clean_date)
        df['提取後碼別'] = df['服務項目'].apply(extract_ba_code)
        
        st.subheader("3. 轉換後的日期與碼別前 10 筆：")
        st.dataframe(df[['個案姓名', '服務日期', '轉換後日期', '服務項目', '提取後碼別']].head(10), use_container_width=True)
        
        # 測試過濾
        df_filtered = df[df['提取後碼別'].str.contains(r'^(BA|BB|BC|BD|GA|GB)', na=False)]
        st.write(f"📊 **4. 成功提取到長照碼別 (BA/BB/BD等) 的筆數**：{len(df_filtered)} 筆")
        
        if '服務數量' in df.columns:
            df['qty'] = pd.to_numeric(df['服務數量'], errors='coerce').fillna(1)
            g = df_filtered.groupby(['個案姓名', '轉換後日期', '提取後碼別'])['qty'].sum().reset_index()
            st.subheader("5. 分組加總後的成果前 10 筆：")
            st.dataframe(g.head(10), use_container_width=True)

    except Exception as e:
        st.error(f"診斷出錯：{e}")
