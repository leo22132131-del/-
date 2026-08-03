import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="長照居家服務核對系統", layout="wide")
st.title("長照居家服務核對系統")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

# 正則匹配提取服務碼 (如 BA01, BB01, GA03, SC03 等)
VALID_CODE_PATTERN = r'([A-Z]{2}\d{2})'

def extract_code(text):
    if pd.isna(text): return ""
    match = re.search(VALID_CODE_PATTERN, str(text).upper().strip())
    return match.group(1) if match else ""

def clean_name(name):
    if pd.isna(name): return ""
    return re.sub(r'\s+', '', str(name)).replace('鳯', '鳳').strip()

def clean_date(val):
    if pd.isna(val): return ""
    s = str(val).strip().split(' ')[0].split('T')[0]
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 1000: y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        except: pass
    if len(s) == 7 and s.isdigit():
        return f"{int(s[:3])+1911:04d}-{int(s[3:5]):02d}-{int(s[5:7]):02d}"
    return s

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns: return name
    for col in df.columns:
        for p in possible_names:
            if p in col: return col
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        df_支審 = pd.read_excel(file_支審)
        df_fa300 = pd.read_excel(file_fa300)
        df_dmaker = pd.read_excel(file_dmaker)

        # 1. 整理 支審
        col_d_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_c_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_n_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        df_支審['date'] = df_支審[col_d_支審].apply(clean_date)
        df_支審['code'] = df_支審[col_c_支審].apply(extract_code)
        df_支審['name'] = df_支審[col_n_支審].apply(clean_name)
        g_支審 = df_支審[df_支審['code'] != ''].groupby(['name', 'date', 'code']).size().reset_index(name='支審次數')

        # 2. 整理 FA300
        col_d_fa = find_column(df_fa300, ['服務日期', '日期'])
        col_c_fa = find_column(df_fa300, ['服務項目', '項目代碼', '項目'])
        col_n_fa = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        df_fa300['date'] = df_fa300[col_d_fa].apply(clean_date)
        df_fa300['code'] = df_fa300[col_c_fa].apply(extract_code)
        df_fa300['name'] = df_fa300[col_n_fa].apply(clean_name)
        g_fa300 = df_fa300[df_fa300['code'] != ''].groupby(['name', 'date', 'code']).size().reset_index(name='FA300次數')

        # 3. 整理 dmaker
        col_d_dm = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_c_dm = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dm = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        df_dmaker['date'] = df_dmaker[col_d_dm].apply(clean_date)
        df_dmaker['code'] = df_dmaker[col_c_dm].apply(extract_code)
        df_dmaker['name'] = df_dmaker[col_n_dm].apply(clean_name)

        def calculate_dmaker_count(group):
            is_self = group[col_c_dm].astype(str).str.contains('自費|全額|超過').any()
            raw_len = len(group)
            return raw_len if is_self else max(1, round(raw_len / 2))

        g_dmaker = df_dmaker[df_dmaker['code'] != ''].groupby(['name', 'date', 'code']).apply(calculate_dmaker_count).reset_index(name='dmaker次數')

        # 合併比對
        merged = pd.merge(g_支審, g_fa300, on=['name', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['name', 'date', 'code'], how='outer').fillna(0)

        merged['支審次數'] = merged['支審次數'].astype(int)
        merged['FA300次數'] = merged['FA300次數'].astype(int)
        merged['dmaker次數'] = merged['dmaker次數'].astype(int)

        diff = merged[
            (merged['支審次數'] != merged['FA300次數']) | 
            (merged['支審次數'] != merged['dmaker次數'])
        ].sort_values(by=['date', 'name'])

        diff_show = diff.rename(columns={
            'date': '服務日期',
            'name': '個案姓名',
            'code': '服務碼別'
        })[['服務日期', '個案姓名', '服務碼別', '支審次數', 'FA300次數', 'dmaker次數']]

        st.subheader("📌 每日異常項目明細")
        if len(diff_show) == 0:
            st.success("🎉 三表數據完全吻合！")
        else:
            st.warning(f"⚠️ 發現 {len(diff_show)} 筆差異紀錄：")
            st.dataframe(diff_show, use_container_width=True)

            @st.cache_data
            def convert_df(df): return df.to_csv(index=False).encode('utf-8-sig')

            st.download_button("📥 下載異常明細 (CSV)", convert_df(diff_show), "每日服務項目異常明細.csv", "text/csv")

    except Exception as e:
        st.error(f"處理失敗：{e}")
