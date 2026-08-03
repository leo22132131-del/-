import streamlit as st
import pandas as pd
import re

st.title("長照居家服務核對系統（含日期核對）")
st.write("請上傳三個 Excel 報表，系統將自動按【姓名 + 服務日期 + 服務代碼】精準比對。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text):
        return ""
    # 抓取兩個大寫字母+兩位數字 (例如 BA01, BB01)
    match = re.search(r'([A-Z]{2}\d{2})', str(text))
    return match.group(1) if match else str(text).strip()

def clean_date(date_val):
    if pd.isna(date_val):
        return ""
    try:
        return pd.to_datetime(date_val).strftime('%Y-%m-%d')
    except:
        return str(date_val).strip()

# 自動尋找符合條件的欄位名稱（去除前後空格）
def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]  # 自動去除欄位名稱的前後空格
    for name in possible_names:
        if name in df.columns:
            return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 讀取與處理 支審資料
        df_支審 = pd.read_excel(file_支審)
        # 加上「服務日期(請輸入7碼)」
        col_date_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_code_支審 = find_column(df_支審, ['服務項目名稱', '服務項目代碼', '服務項目', '項目代碼'])
        col_name_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        if not col_date_支審 or not col_code_支審 or not col_name_支審:
            st.error(f"❌ 支審資料欄位比對失敗！請確認檔案中是否有姓名、日期、服務項目等欄位。現有欄位：{list(df_支審.columns)}")
            st.stop()

        df_支審['date'] = df_支審[col_date_支審].apply(clean_date)
        df_支審['code'] = df_支審[col_code_支審].apply(extract_ba_code)
        df_支審 = df_支審[~df_支審[col_code_支審].astype(str).str.contains('QA1385', na=False)]
        g_支審 = df_支審.groupby([col_name_支審, 'date', 'code']).size().reset_index(name='支審次數')
        g_支審.rename(columns={col_name_支審: '個案姓名'}, inplace=True)

        # 2. 讀取與處理 FA300
        df_fa300 = pd.read_excel(file_fa300)
        col_date_fa300 = find_column(df_fa300, ['服務日期', '執行日期', '日期'])
        col_code_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_name_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        if not col_date_fa300 or not col_code_fa300 or not col_name_fa300:
            st.error(f"❌ FA300 欄位比對失敗！現有欄位：{list(df_fa300.columns)}")
            st.stop()

        df_fa300['date'] = df_fa300[col_date_fa300].apply(clean_date)
        df_fa300['code'] = df_fa300[col_code_fa300].apply(extract_ba_code)
        g_fa300 = df_fa300.groupby([col_name_fa300, 'date', 'code']).size().reset_index(name='FA300次數')
        g_fa300.rename(columns={col_name_fa300: '個案姓名'}, inplace=True)

        # 3. 讀取與處理 dmaker
        df_dmaker = pd.read_excel(file_dmaker)
        col_date_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_code_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_name_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        if not col_date_dmaker or not col_code_dmaker or not col_name_dmaker:
            st.error(f"❌ dmaker 欄位比對失敗！現有欄位：{list(df_dmaker.columns)}")
            st.stop()

        df_dmaker['date'] = df_dmaker[col_date_dmaker].apply(clean_date)
        df_dmaker['code'] = df_dmaker[col_code_dmaker].apply(extract_ba_code)

        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby([col_name_dmaker, 'date', 'code'])['數量'].sum().reset_index(name='dmaker數量')
        else:
            g_dmaker = df_dmaker.groupby([col_name_dmaker, 'date', 'code']).size().reset_index(name='dmaker數量')
        g_dmaker.rename(columns={col_name_dmaker: '個案姓名'}, inplace=True)

        # 4. 三表合併比對
        merged = pd.merge(g_支審, g_fa300, on=['個案姓名', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['個案姓名', 'date', 'code'], how='outer')

        # 補 0 與欄位整理
        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker數量'] = merged['dmaker數量'].fillna(0).astype(int)

        result = merged.rename(columns={'date': '服務日期', 'code': '代碼(BA/BB)'})
        result = result[['個案姓名', '服務日期', '代碼(BA/BB)', '支審次數', 'FA300次數', 'dmaker數量']]

        # 篩選不一致
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

    except Exception as e:
        st.error(f"資料處理時發生錯誤，請檢查檔案格式。錯誤細節：{e}")
