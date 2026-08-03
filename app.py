import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="長照居家服務核對系統", layout="wide")
st.title("長照居家服務核對系統（終極修復版）")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

# 正則匹配：抓取所有英文2碼+數字2碼 (如 GA04, SC03, BB01, BD03, BA01 等)
VALID_CODE_PATTERN = r'([A-Z]{2}\d{2})'

def extract_code(text):
    if pd.isna(text): return ""
    s = str(text).upper().strip()
    match = re.search(VALID_CODE_PATTERN, s)
    if match: return match.group(1)
    return ""

def clean_name(name):
    if pd.isna(name): return ""
    # 清理所有不可見字元與空白
    s = re.sub(r'[\s\u3000\t\r\n]+', '', str(name))
    s = s.replace('鳯', '鳳')
    return s.strip()

def clean_date(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == 'nan': 
        return ""
    
    if isinstance(val, (pd.Timestamp, pd.DatetimeIndex)):
        y = val.year
        if y < 1920: y += 1911
        return f"{y:04d}-{val.month:02d}-{val.day:02d}"

    s = str(val).strip().split(' ')[0].split('T')[0].split('.')[0]
    
    # 支援 115/07/31 或 2026-07-31
    parts = re.split(r'[/.-]', s)
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y < 1000: y += 1911
            return f"{y:04d}-{m:02d}-{d:02d}"
        except: pass

    if s.isdigit() and len(s) == 7:
        y = int(s[:3]) + 1911
        return f"{y:04d}-{int(s[3:5]):02d}-{int(s[5:7]):02d}"
    
    if s.isdigit() and len(s) == 8: 
        return f"{s[:4]}-{int(s[4:6]):02d}-{int(s[6:8]):02d}"
        
    try:
        dt = pd.to_datetime(s)
        if dt.year < 1920: dt = dt.replace(year=dt.year + 1911)
        return dt.strftime('%Y-%m-%d')
    except: 
        return ""

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns: return name
    for col in df.columns:
        for p in possible_names:
            if p in col: return col
    return None

def process_fa300_final(df):
    df.columns = [str(c).strip() for c in df.columns]
    
    # 依據您提供的截圖與 JSON 欄位結構：
    # "個案姓名", "身分證", "服務日期", "申報年月", "服務項目", "類型", "服務數量", "服務人員", "狀態"
    col_n = '個案姓名' if '個案姓名' in df.columns else df.columns[0]
    col_d = '服務日期' if '服務日期' in df.columns else df.columns[2]
    col_c = '服務項目' if '服務項目' in df.columns else df.columns[4]
    col_q = '服務數量' if '服務數量' in df.columns else (df.columns[6] if len(df.columns) > 6 else None)
    col_s = '狀態' if '狀態' in df.columns else (df.columns[8] if len(df.columns) > 8 else None)

    df_work = df.copy()

    # 1. 狀態過濾 (排除作廢/取消)
    if col_s and col_s in df_work.columns:
        df_work = df_work[~df_work[col_s].astype(str).str.contains('取消|作廢|刪除|不核銷|錯誤', na=False)]

    df_clean = pd.DataFrame()
    df_clean['name'] = df_work[col_n].apply(clean_name)
    df_clean['date'] = df_work[col_d].apply(clean_date)
    df_clean['code'] = df_work[col_c].apply(extract_code)
    
    # 2. 數量處理 (確保 > 0)
    if col_q and col_q in df_work.columns:
        df_clean['qty'] = pd.to_numeric(df_work[col_q], errors='coerce').fillna(0)
    else:
        df_clean['qty'] = 1

    # 只留有效日期、有效碼別、數量 > 0 的紀錄
    df_clean = df_clean[
        (df_clean['date'].str.match(r'^\d{4}-\d{2}-\d{2}$', na=False)) & 
        (df_clean['code'] != '') & 
        (df_clean['qty'] > 0)
    ]
    
    return df_clean.groupby(['name', 'date', 'code'])['qty'].sum().reset_index(name='FA300次數')

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
        df_支審 = df_支審[(df_支審['code'] != '') & (df_支審['date'].str.match(r'^\d{4}-\d{2}-\d{2}$', na=False))]
        g_支審 = df_支審.groupby(['name', 'date', 'code']).size().reset_index(name='支審次數')

        # 2. 整理 FA300
        g_fa300 = process_fa300_final(df_fa300)

        # 3. 整理 dmaker
        col_d_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_c_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        df_dmaker['date'] = df_dmaker[col_d_dmaker].apply(clean_date)
        df_dmaker['code'] = df_dmaker[col_c_dmaker].apply(extract_code)
        df_dmaker['name'] = df_dmaker[col_n_dmaker].apply(clean_name)
        df_dmaker = df_dmaker[(df_dmaker['code'] != '') & (df_dmaker['date'].str.match(r'^\d{4}-\d{2}-\d{2}$', na=False))]

        def calculate_dmaker_count(group):
            is_self = group[col_c_dmaker].astype(str).str.contains('自費|全額|超過').any()
            raw_len = len(group)
            if is_self:
                return raw_len
            else:
                return max(1, round(raw_len / 2))

        g_dmaker = df_dmaker.groupby(['name', 'date', 'code']).apply(calculate_dmaker_count).reset_index(name='dmaker次數')

        # 顯示讀取統計，方便確認三表資料量
        st.info(f"📊 成功解析數據數量：支審 {len(g_支審)} 筆組 | FA300 {len(g_fa300)} 筆組 | dmaker {len(g_dmaker)} 筆組")

        # 三表外連結
        merged = pd.merge(g_支審, g_fa300, on=['name', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['name', 'date', 'code'], how='outer')

        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker次數'] = merged['dmaker次數'].fillna(0).astype(int)

        # 篩選不相符紀錄
        diff = merged[
            (merged['支審次數'] != merged['FA300次數']) | 
            (merged['支審次數'] != merged['dmaker次數'])
        ].sort_values(by=['date', 'name'])

        diff_show = diff.rename(columns={
            'date': '服務日期',
            'name': '個案姓名',
            'code': '服務碼別',
            '支審次數': '支審次數',
            'FA300次數': 'FA300次數',
            'dmaker次數': 'dmaker次數'
        })[['服務日期', '個案姓名', '服務碼別', '支審次數', 'FA300次數', 'dmaker次數']]

        st.subheader("📌 每日異常項目明細")
        if len(diff_show) == 0:
            st.success("🎉 太棒了！三表每日服務項目與次數完全吻合！")
        else:
            st.warning(f"⚠️ 發現 {len(diff_show)} 筆差異紀錄：")
            st.dataframe(diff_show, use_container_width=True)

            @st.cache_data
            def convert_df(df): return df.to_csv(index=False).encode('utf-8-sig')

            st.download_button("📥 下載異常明細 (CSV)", convert_df(diff_show), "每日服務項目異常明細.csv", "text/csv")

    except Exception as e:
        st.error(f"資料處理失敗，詳細錯誤訊息：{e}")
