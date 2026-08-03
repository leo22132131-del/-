import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="長照居家服務核對系統", layout="wide")
st.title("長照居家服務核對系統（全月總次數比對）")
st.write("請上傳三個 Excel 報表，系統將按【個案姓名 + 服務代碼】自動比對整個月的總次數/數量。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text):
        return ""
    # 擷取字母代碼，例如 BA01, BB01, GA01 等
    match = re.search(r'([A-Z]{2}\d{2})', str(text).upper())
    return match.group(1) if match else str(text).strip()

def clean_name(name):
    if pd.isna(name):
        return ""
    # 清理姓名（移除所有空格）
    return re.sub(r'\s+', '', str(name))

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns:
            return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 讀取 Excel 檔案
        df_支審 = pd.read_excel(file_支審)
        df_fa300 = pd.read_excel(file_fa300)
        df_dmaker = pd.read_excel(file_dmaker)

        # 2. 自動對應欄位名稱
        col_c_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_n_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        col_c_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_n_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        col_c_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        # 3. 處理 支審資料
        df_支審['code'] = df_支審[col_c_支審].apply(extract_ba_code)
        df_支審['clean_name'] = df_支審[col_n_支審].apply(clean_name)
        df_支審 = df_支審[~df_支審[col_c_支審].astype(str).str.contains('QA1385', na=False)]
        g_支審 = df_支審.groupby(['clean_name', 'code']).size().reset_index(name='支審總次數')

        # 4. 處理 FA300
        df_fa300['code'] = df_fa300[col_c_fa300].apply(extract_ba_code)
        df_fa300['clean_name'] = df_fa300[col_n_fa300].apply(clean_name)
        g_fa300 = df_fa300.groupby(['clean_name', 'code']).size().reset_index(name='FA300總次數')

        # 5. 處理 dmaker
        df_dmaker['code'] = df_dmaker[col_c_dmaker].apply(extract_ba_code)
        df_dmaker['clean_name'] = df_dmaker[col_n_dmaker].apply(clean_name)

        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby(['clean_name', 'code'])['數量'].sum().reset_index(name='dmaker總數量')
        else:
            g_dmaker = df_dmaker.groupby(['clean_name', 'code']).size().reset_index(name='dmaker總數量')

        # 6. 三表合併比對（只按 姓名 + 代碼）
        merged = pd.merge(g_支審, g_fa300, on=['clean_name', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['clean_name', 'code'], how='outer')

        # 補 0 轉成整數
        merged['支審總次數'] = merged['支審總次數'].fillna(0).astype(int)
        merged['FA300總次數'] = merged['FA300總次數'].fillna(0).astype(int)
        merged['dmaker總數量'] = merged['dmaker總數量'].fillna(0).astype(int)

        result = merged.rename(columns={'clean_name': '個案姓名', 'code': '服務代碼(BA/BB等)'})
        result = result[['個案姓名', '服務代碼(BA/BB等)', '支審總次數', 'FA300總次數', 'dmaker總數量']]

        # 7. 篩選不一致項目
        diff = result[(result['支審總次數'] != result['FA300總次數']) | (result['支審總次數'] != result['dmaker總數量'])]

        st.subheader("📊 月總數核對結果")
        if len(diff) == 0:
            st.success("🎉 太棒了！整個月所有個案的服務代碼總數量完全吻合！")
        else:
            st.warning(f"⚠️ 發現 {len(diff)} 筆不相符的個案項目：")
            st.dataframe(diff, use_container_width=True)

            @st.cache_data
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8-sig')

            csv = convert_df(diff)
            st.download_button(
                label="📥 下載月總數異常明細表 (CSV)",
                data=csv,
                file_name='長照月總數異常明細.csv',
                mime='text/csv',
            )

    except Exception as e:
        st.error(f"資料處理時發生錯誤，請檢查檔案格式。錯誤訊息：{e}")
