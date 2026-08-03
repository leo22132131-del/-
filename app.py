import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="長照居家服務每日項目核對系統", layout="wide")
st.title("長照居家服務核對系統（每日項目精準核對）")
st.write("上傳三個 Excel 報表後，系統將自動按【個案姓名 + 服務日期 + 服務代碼】核對數量是否一致。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    """抓取項目代碼 (如 BA01, BB01, GA01 等)"""
    if pd.isna(text):
        return ""
    match = re.search(r'([A-Z]{2}\d{2})', str(text).upper())
    return match.group(1) if match else str(text).strip()

def clean_name(name):
    """去除姓名中的所有空格"""
    if pd.isna(name):
        return ""
    return re.sub(r'\s+', '', str(name))

def clean_date(val):
    """統一日期格式為 YYYY-MM-DD (相容民國年與各種分隔符號)"""
    if pd.isna(val) or str(val).strip() == "":
        return ""
    
    s = str(val).strip().split(' ')[0].split('.')[0]
    
    # Excel 序列號 (5位數)
    if s.isdigit() and len(s) == 5:
        try:
            return pd.to_datetime(int(s), unit='D', origin='1899-12-30').strftime('%Y-%m-%d')
        except:
            pass

    # 7 碼純數字民國年 (例如: 1150702 -> 2026-07-02)
    if s.isdigit() and len(s) == 7:
        y = int(s[:3]) + 1911
        return f"{y}-{s[3:5]}-{s[5:7]}"

    # 8 碼純數字西元年 (例如: 20260702 -> 2026-07-02)
    if s.isdigit() and len(s) == 8:
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"

    # 帶有斜線或點號 (例如: 115/7/2 或 2026/07/02)
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 1000:  # 民國年自動轉西元
                y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        except:
            pass

    try:
        dt = pd.to_datetime(s)
        if dt.year < 1920:
            dt = dt.replace(year=dt.year + 1911)
        return dt.strftime('%Y-%m-%d')
    except:
        return s

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns:
            return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 讀取資料
        df_支審 = pd.read_excel(file_支審)
        df_fa300 = pd.read_excel(file_fa300)
        df_dmaker = pd.read_excel(file_dmaker)

        # 2. 自動匹配欄位
        col_d_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_c_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_n_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        col_d_fa300 = find_column(df_fa300, ['服務日期', '執行日期', '日期'])
        col_c_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_n_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        col_d_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_c_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        # 3. 資料整理與清理
        # 支審
        df_支審['date'] = df_支審[col_d_支審].apply(clean_date)
        df_支審['code'] = df_支審[col_c_支審].apply(extract_ba_code)
        df_支審['name'] = df_支審[col_n_支審].apply(clean_name)
        df_支審 = df_支審[~df_支審[col_c_支審].astype(str).str.contains('QA1385', na=False)] # 排除特殊項目
        g_支審 = df_支審.groupby(['name', 'date', 'code']).size().reset_index(name='支審次數')

        # FA300
        df_fa300['date'] = df_fa300[col_d_fa300].apply(clean_date)
        df_fa300['code'] = df_fa300[col_c_fa300].apply(extract_ba_code)
        df_fa300['name'] = df_fa300[col_n_fa300].apply(clean_name)
        g_fa300 = df_fa300.groupby(['name', 'date', 'code']).size().reset_index(name='FA300次數')

        # dmaker
        df_dmaker['date'] = df_dmaker[col_d_dmaker].apply(clean_date)
        df_dmaker['code'] = df_dmaker[col_c_dmaker].apply(extract_ba_code)
        df_dmaker['name'] = df_dmaker[col_n_dmaker].apply(clean_name)
        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby(['name', 'date', 'code'])['數量'].sum().reset_index(name='dmaker數量')
        else:
            g_dmaker = df_dmaker.groupby(['name', 'date', 'code']).size().reset_index(name='dmaker數量')

        # 4. 三表完全比對 (Outer Merge)
        merged = pd.merge(g_支審, g_fa300, on=['name', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['name', 'date', 'code'], how='outer')

        # 空白補 0 並轉為數字
        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker數量'] = merged['dmaker數量'].fillna(0).astype(int)

        # 整理欄位名稱
        result = merged.rename(columns={'name': '個案姓名', 'date': '服務日期', 'code': '服務代碼'})
        result = result[['個案姓名', '服務日期', '服務代碼', '支審次數', 'FA300次數', 'dmaker數量']]

        # 5. 篩選出「三表數量不完全一致」的項目
        diff = result[
            (result['支審次數'] != result['FA300次數']) | 
            (result['支審次數'] != result['dmaker數量'])
        ].sort_values(by=['服務日期', '個案姓名'])

        # 過濾掉日期解析失敗的極少數空筆數
        diff = diff[diff['服務日期'] != '']

        st.subheader("📊 每日不正確項目清單")
        if len(diff) == 0:
            st.success("🎉 太棒了！所有日期的服務項目數量完全吻合！")
        else:
            st.error(f"⚠️ 發現 {len(diff)} 筆不相符的每日紀錄：")
            st.dataframe(diff, use_container_width=True)

            @st.cache_data
            def convert_df(df):
                return df.to_csv(index=False).encode('utf-8-sig')

            st.download_button(
                label="📥 下載不正確項目明細 (CSV)",
                data=convert_df(diff),
                file_name='長照每日不正確項目明細.csv',
                mime='text/csv',
            )

    except Exception as e:
        st.error(f"資料處理時發生錯誤，請檢查檔案格式。錯誤訊息：{e}")
