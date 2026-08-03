import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.title("長照居家服務核對系統（月總數 + 每日明細 雙核對）")
st.write("系統會先比對【全月總次數】，若發現總數不符，再自動列出【每日詳細比對】，讓漏打卡、少上傳無所遁形！")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text):
        return ""
    match = re.search(r'([A-Z]{2}\d{2})', str(text).upper())
    return match.group(1) if match else str(text).strip()

def clean_date_universal(val):
    if pd.isna(val) or str(val).strip() == "":
        return ""
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip().split(' ')[0].split('.')[0]
    if s.isdigit() and len(s) == 5:
        try:
            return pd.to_datetime(int(s), unit='D', origin='1899-12-30').strftime('%Y-%m-%d')
        except:
            pass
    if s.isdigit() and len(s) == 7:
        return f"{int(s[:3])+1911}-{s[3:5]}-{s[5:7]}"
    if s.isdigit() and len(s) == 8:
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 1000: y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        except:
            pass
    try:
        dt = pd.to_datetime(s)
        if dt.year < 1920: dt = dt.replace(year=dt.year + 1911)
        return dt.strftime('%Y-%m-%d')
    except:
        return s

def clean_name(name):
    if pd.isna(name): return ""
    return re.sub(r'\s+', '', str(name))

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns: return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 讀取資料
        df_支審 = pd.read_excel(file_支審)
        df_fa300 = pd.read_excel(file_fa300)
        df_dmaker = pd.read_excel(file_dmaker)

        # 欄位對應
        col_d_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_c_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_n_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        col_d_fa300 = find_column(df_fa300, ['服務日期', '執行日期', '日期'])
        col_c_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_n_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        col_d_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_c_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        # 資料清理
        for df, col_d, col_c, col_n in [(df_支審, col_d_支審, col_c_支審, col_n_支審),
                                        (df_fa300, col_d_fa300, col_c_fa300, col_n_fa300),
                                        (df_dmaker, col_d_dmaker, col_c_dmaker, col_n_dmaker)]:
            df['date'] = df[col_d].apply(clean_date_universal)
            df['code'] = df[col_c].apply(extract_ba_code)
            df['clean_name'] = df[col_n].apply(clean_name)

        df_支審 = df_支審[~df_支審[col_c_支審].astype(str).str.contains('QA1385', na=False)]

        # --- 階段一：月總數比對 ---
        g_支審_m = df_支審.groupby(['clean_name', 'code']).size().reset_index(name='支審總數')
        g_fa300_m = df_fa300.groupby(['clean_name', 'code']).size().reset_index(name='FA300總數')
        if '數量' in df_dmaker.columns:
            g_dmaker_m = df_dmaker.groupby(['clean_name', 'code'])['數量'].sum().reset_index(name='dmaker總數')
        else:
            g_dmaker_m = df_dmaker.groupby(['clean_name', 'code']).size().reset_index(name='dmaker總數')

        m_monthly = pd.merge(g_支審_m, g_fa300_m, on=['clean_name', 'code'], how='outer')
        m_monthly = pd.merge(m_monthly, g_dmaker_m, on=['clean_name', 'code'], how='outer').fillna(0)
        
        diff_monthly = m_monthly[(m_monthly['支審總數'] != m_monthly['FA300總數']) | (m_monthly['支審總數'] != m_monthly['dmaker總數'])]

        st.subheader("📌 第一階段：月總次數異常清單")
        if len(diff_monthly) == 0:
            st.success("🎉 全月總次數完全吻合！")
        else:
            st.warning(f"⚠️ 共有 {len(diff_monthly)} 個【個案+項目】月總數不符！")
            diff_m_show = diff_monthly.rename(columns={'clean_name': '個案姓名', 'code': '代碼'})
            st.dataframe(diff_m_show)

        # --- 階段二：每日明細比對 ---
        st.markdown("---")
        st.subheader("📅 第二階段：每日明細比對（精確定位日期）")

        g_支審_d = df_支審.groupby(['clean_name', 'date', 'code']).size().reset_index(name='支審')
        g_fa300_d = df_fa300.groupby(['clean_name', 'date', 'code']).size().reset_index(name='FA300')
        if '數量' in df_dmaker.columns:
            g_dmaker_d = df_dmaker.groupby(['clean_name', 'date', 'code'])['數量'].sum().reset_index(name='dmaker')
        else:
            g_dmaker_d = df_dmaker.groupby(['clean_name', 'date', 'code']).size().reset_index(name='dmaker')

        m_daily = pd.merge(g_支審_d, g_fa300_d, on=['clean_name', 'date', 'code'], how='outer')
        m_daily = pd.merge(m_daily, g_dmaker_d, on=['clean_name', 'date', 'code'], how='outer').fillna(0)
        
        # 轉成整數
        for col in ['支審', 'FA300', 'dmaker']:
            m_daily[col] = m_daily[col].astype(int)

        diff_daily = m_daily[(m_daily['支審'] != m_daily['FA300']) | (m_daily['支審'] != m_daily['dmaker'])]
        diff_daily = diff_daily[~diff_daily['date'].isin(['1970-01-01', ''])]

        diff_d_show = diff_daily.rename(columns={'clean_name': '個案姓名', 'date': '服務日期', 'code': '代碼'})
        st.dataframe(diff_d_show)

        @st.cache_data
        def convert_df(df): return df.to_csv(index=False).encode('utf-8-sig')

        st.download_button("📥 下載每日異常明細表 (CSV)", convert_df(diff_d_show), "每日長照核對異常.csv", "text/csv")

    except Exception as e:
        st.error(f"處理失敗：{e}")
