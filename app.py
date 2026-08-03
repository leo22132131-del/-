import streamlit as st
import pandas as pd
import numpy as np
import re

st.title("長照居家服務核對系統（精準按日期核對）")
st.write("請上傳三個 Excel 報表，系統將自動按【姓名 + 服務日期 + 服務代碼】精準比對。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text):
        return ""
    match = re.search(r'([A-Z]{2}\d{2})', str(text).upper())
    return match.group(1) if match else str(text).strip()

def clean_date_universal(val):
    """全通用日期轉換函數，解析各種長照系統日期格式"""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    
    # 1. 先處理 pandas/datetime 原生物件
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime('%Y-%m-%d')
        
    s = str(val).strip().split(' ')[0].split('.')[0] # 拿掉時間與小數點
    
    # 2. 如果是 Excel 序列號 (如 45505 這種五位數字)
    if s.isdigit() and len(s) == 5:
        try:
            dt = pd.to_datetime(int(s), unit='D', origin='1899-12-30')
            return dt.strftime('%Y-%m-%d')
        except:
            pass

    # 3. 民國年純數字：7碼 (如 1130801 -> 2024-08-01)
    if s.isdigit() and len(s) == 7:
        y = int(s[:3]) + 1911
        return f"{y}-{s[3:5]}-{s[5:7]}"

    # 4. 西元年純數字：8碼 (如 20240801 -> 2024-08-01)
    if s.isdigit() and len(s) == 8:
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"

    # 5. 帶分隔符號 (如 113/8/1, 113.8.1, 113-08-01, 2024/08/01)
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 1000:  # 民國年
                y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        except:
            pass

    # 6. 其他標準文字日期
    try:
        dt = pd.to_datetime(s)
        if dt.year < 1920:
            dt = dt.replace(year=dt.year + 1911)
        return dt.strftime('%Y-%m-%d')
    except:
        return s

from datetime import datetime

def clean_name(name):
    if pd.isna(name):
        return ""
    return re.sub(r'\s+', '', str(name))

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns:
            return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 支審資料
        df_支審 = pd.read_excel(file_支審)
        col_date_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_code_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_name_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        df_支審['date'] = df_支審[col_date_支審].apply(clean_date_universal)
        df_支審['code'] = df_支審[col_code_支審].apply(extract_ba_code)
        df_支審['clean_name'] = df_支審[col_name_支審].apply(clean_name)
        df_支審 = df_支審[~df_支審[col_code_支審].astype(str).str.contains('QA1385', na=False)]
        
        g_支審 = df_支審.groupby(['clean_name', 'date', 'code']).size().reset_index(name='支審次數')

        # 2. FA300
        df_fa300 = pd.read_excel(file_fa300)
        col_date_fa300 = find_column(df_fa300, ['服務日期', '執行日期', '日期'])
        col_code_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_name_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        df_fa300['date'] = df_fa300[col_date_fa300].apply(clean_date_universal)
        df_fa300['code'] = df_fa300[col_code_fa300].apply(extract_ba_code)
        df_fa300['clean_name'] = df_fa300[col_name_fa300].apply(clean_name)
        
        g_fa300 = df_fa300.groupby(['clean_name', 'date', 'code']).size().reset_index(name='FA300次數')

        # 3. dmaker
        df_dmaker = pd.read_excel(file_dmaker)
        col_date_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_code_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_name_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        df_dmaker['date'] = df_dmaker[col_date_dmaker].apply(clean_date_universal)
        df_dmaker['code'] = df_dmaker[col_code_dmaker].apply(extract_ba_code)
        df_dmaker['clean_name'] = df_dmaker[col_name_dmaker].apply(clean_name)

        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby(['clean_name', 'date', 'code'])['數量'].sum().reset_index(name='dmaker數量')
        else:
            g_dmaker = df_dmaker.groupby(['clean_name', 'date', 'code']).size().reset_index(name='dmaker數量')

        # 4. 合併比對
        merged = pd.merge(g_支審, g_fa300, on=['clean_name', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['clean_name', 'date', 'code'], how='outer')

        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker數量'] = merged['dmaker數量'].fillna(0).astype(int)

        result = merged.rename(columns={'clean_name': '個案姓名', 'date': '服務日期', 'code': '代碼(BA/BB)'})
        result = result[['個案姓名', '服務日期', '代碼(BA/BB)', '支審次數', 'FA300次數', 'dmaker數量']]

        # 排除日期無效的欄位
        result = result[~result['服務日期'].isin(['1970-01-01', ''])]
        diff = result[(result['支審次數'] != result['FA300次數']) | (result['支審次數'] != result['dmaker數量'])]

        st.subheader("📊 核對結果總覽")
        if len(diff) == 0:
            st.success("🎉 太棒了！所有日期的 BA/BB 碼數量完全符合！")
        else:
            st.error(f"⚠️ 發現 {len(diff)} 筆不符合的紀錄：")
            st.dataframe(diff)

            @st.cache_data
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8-sig')

            csv = convert_df(diff)
            st.download_button(
                label="📥 下載異常明細表 (CSV)",
                data=csv,
                file_name='長照BA_BB碼異常明細.csv',
                mime='text/csv',
            )

        # 頁面下方排錯小工具：顯示前 5 筆解析出來的日期
        with st.expander("🔍 點此檢查：系統解析出來的日期格式"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.write("**支審解析日期範例**")
                st.write(df_支審[['clean_name', 'date', 'code']].head())
            with col2:
                st.write("**FA300解析日期範例**")
                st.write(df_fa300[['clean_name', 'date', 'code']].head())
            with col3:
                st.write("**dmaker解析日期範例**")
                st.write(df_dmaker[['clean_name', 'date', 'code']].head())

    except Exception as e:
        st.error(f"資料處理時發生錯誤：{e}")
