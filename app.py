import streamlit as st
import pandas as pd
import re

st.title("長照居家服務核對系統（含日期核對）")
st.write("請上傳三個 Excel 報表，系統將自動按【姓名 + 服務日期 + BA碼】精準比對。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text):
        return ""
    # 抓取 BA 碼格式 (例如 BA01, BA02)
    match = re.search(r'([A-Z]{2}\d{2})', str(text))
    return match.group(1) if match else str(text).strip()

def clean_date(date_val):
    if pd.isna(date_val):
        return ""
    # 將日期統一轉為 YYYY-MM-DD 格式，方便精準比對
    try:
        return pd.to_datetime(date_val).strftime('%Y-%m-%d')
    except:
        return str(date_val).strip()

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 讀取與處理 支審資料
        df_支審 = pd.read_excel(file_支審)
        df_支審['date'] = df_支審['服務日期'].apply(clean_date)
        df_支審['code'] = df_支審['服務項目名稱'].apply(extract_ba_code)
        # 排除 QA1385 相關項目（若有）
        df_支審 = df_支審[~df_支審['服務項目名稱'].astype(str).str.contains('QA1385', na=False)]
        g_支審 = df_支審.groupby(['個案姓名', 'date', 'code']).size().reset_index(name='支審次數')

        # 2. 讀取與處理 FA300
        df_fa300 = pd.read_excel(file_fa300)
        df_fa300['date'] = df_fa300['服務日期'].apply(clean_date)
        df_fa300['code'] = df_fa300['服務項目'].apply(extract_ba_code)
        g_fa300 = df_fa300.groupby(['個案姓名', 'date', 'code']).size().reset_index(name='FA300次數')

        # 3. 讀取與處理 dmaker
        df_dmaker = pd.read_excel(file_dmaker)
        df_dmaker['date'] = df_dmaker['使用日期'].apply(clean_date)
        df_dmaker['code'] = df_dmaker['品名'].apply(extract_ba_code)
        # 如果 dmaker 有數量欄位就累加，沒有就計算筆數
        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby(['客戶名', 'date', 'code'])['數量'].sum().reset_index(name='dmaker數量')
        else:
            g_dmaker = df_dmaker.groupby(['客戶名', 'date', 'code']).size().reset_index(name='dmaker數量')

        # 4. 三表合併比對 (按 姓名 + 日期 + BA碼)
        merged = pd.merge(g_支審, g_fa300, left_on=['個案姓名', 'date', 'code'], right_on=['個案姓名', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, left_on=['個案姓名', 'date', 'code'], right_on=['客戶名', 'date', 'code'], how='outer')

        # 補 0 與欄位整理
        merged['個案姓名'] = merged['個案姓名'].fillna(merged['客戶名'])
        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker數量'] = merged['dmaker數量'].fillna(0).astype(int)

        # 欄位重新命名與篩選不符項目
        result = merged.rename(columns={'date': '服務日期', 'code': 'BA碼'})
        result = result[['個案姓名', '服務日期', 'BA碼', '支審次數', 'FA300次數', 'dmaker數量']]

        # 找出三個來源數量不一致的資料
        diff = result[(result['支審次數'] != result['FA300次數']) | (result['支審次數'] != result['dmaker數量'])]

        st.subheader("📊 核對結果總覽")
        if len(diff) == 0:
            st.success("🎉 太棒了！所有日期的 BA 碼數量完全符合！")
        else:
            st.error(f"⚠️ 發現 {len(diff)} 筆不符合的日期紀錄：")
            st.dataframe(diff)

            # 提供下載 Excel 異常報告
            @st.cache_data
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8-sig')

            csv = convert_df(diff)
            st.download_button(
                label="📥 下載日期異常明細表 (CSV)",
                data=csv,
                file_name='長照居家BA碼日期不符明細.csv',
                mime='text/csv',
            )

    except Exception as e:
        st.error(f"資料處理時發生錯誤，請檢查檔案格式是否正確。錯誤訊息：{e}")
