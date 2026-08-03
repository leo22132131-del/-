import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="長照居家服務核對系統", layout="wide")
st.title("長照居家服務核對系統（極簡日期+項目核對）")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text): return ""
    s = str(text).upper().strip()
    match = re.search(r'([A-Z]{2}\d{2})', s)
    if match: return match.group(1)
    return ""

def clean_name(name):
    if pd.isna(name): return ""
    s = re.sub(r'[\s\u3000\t\r\n]+', '', str(name))
    s = s.replace('鳯', '鳳')
    return s.strip()

def clean_date(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() == 'nan': 
        return ""
    
    s = str(val).strip().split(' ')[0].split('T')[0].split('.')[0]
    
    # 處理帶斜線/短線/點的日期 (如 115/07/01 或 2026-07-01)
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

def read_excel_smart(file):
    df_first = pd.read_excel(file, nrows=5)
    cols_str = "".join([str(c) for c in df_first.columns])
    if re.search(r'[A-Z]\d{9}', cols_str) or re.search(r'\d{3}/', cols_str):
        df = pd.read_excel(file, header=None)
    else:
        df = pd.read_excel(file)
    return df

def process_fa300(df):
    name_col, date_col, code_col, qty_col = None, None, None, None
    
    # 針對沒有標頭的 FA300 進行欄位精準掃瞄
    for col in df.columns:
        col_str = str(col).strip()
        vals = df[col].dropna().astype(str).head(10).tolist()
        
        # 1. 真正的日期欄：裡面必須有 '/' 或 '-'，且絕對不能含有英文字母
        if not date_col:
            if any('/' in v or '-' in v for v in vals) and not any(re.search(r'[A-Za-z]', v) for v in vals):
                date_col = col
                
        # 2. 姓名欄：2-4個字的純中文
        if not name_col:
            if any(len(v) in [2, 3, 4] and not re.search(r'[\d\w/.-]', v) for v in vals):
                name_col = col

        # 3. 服務項目欄：含有 BA/BB/BD/GA 等項目代碼
        if not code_col:
            if any(re.search(r'[A-Z]{2}\d{2}', v) for v in vals):
                code_col = col

        # 4. 數量欄：標頭有數量，或是純數字1
        if '數量' in col_str or '次數' in col_str:
            qty_col = col

    # 如果沒找到標頭，預設 FA300 的欄位順序：0:姓名, 2:日期, 3:項目, 4:數量
    if not date_col: date_col = 2 if 2 in df.columns else df.columns[2]
    if not name_col: name_col = 0 if 0 in df.columns else df.columns[0]
    if not code_col: code_col = 3 if 3 in df.columns else df.columns[3]

    df['date'] = df[date_col].apply(clean_date)
    df['code'] = df[code_col].apply(extract_ba_code)
    df['name'] = df[name_col].apply(clean_name)
    
    df = df[df['code'].str.contains(r'^(BA|BB|BC|BD|GA|GB)', na=False)]
    
    if qty_col:
        df['qty'] = pd.to_numeric(df[qty_col], errors='coerce').fillna(1)
        return df.groupby(['name', 'date', 'code'])['qty'].sum().reset_index(name='FA300次數')
    else:
        return df.groupby(['name', 'date', 'code']).size().reset_index(name='FA300次數')

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
        df_支審 = read_excel_smart(file_支審)
        df_fa300 = read_excel_smart(file_fa300)
        df_dmaker = read_excel_smart(file_dmaker)

        # 1. 整理 支審
        col_d_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_c_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_n_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        df_支審['date'] = df_支審[col_d_支審].apply(clean_date)
        df_支審['code'] = df_支審[col_c_支審].apply(extract_ba_code)
        df_支審['name'] = df_支審[col_n_支審].apply(clean_name)
        df_支審 = df_支審[df_支審['code'].str.contains(r'^(BA|BB|BC|BD|GA|GB)', na=False)]
        g_支審 = df_支審.groupby(['name', 'date', 'code']).size().reset_index(name='支審次數')

        # 2. 整理 FA300
        g_fa300 = process_fa300(df_fa300)

        # 3. 整理 dmaker
        col_d_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_c_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        df_dmaker['date'] = df_dmaker[col_d_dmaker].apply(clean_date)
        df_dmaker['code'] = df_dmaker[col_c_dmaker].apply(extract_ba_code)
        df_dmaker['name'] = df_dmaker[col_n_dmaker].apply(clean_name)
        df_dmaker = df_dmaker[df_dmaker['code'].str.contains(r'^(BA|BB|BC|BD|GA|GB)', na=False)]

        def calculate_dmaker_count(group):
            is_self = group[col_c_dmaker].astype(str).str.contains('自費|全額|超過').any()
            raw_len = len(group)
            if is_self:
                return raw_len
            else:
                return max(1, round(raw_len / 2))

        g_dmaker = df_dmaker.groupby(['name', 'date', 'code']).apply(calculate_dmaker_count).reset_index(name='dmaker次數')

        # 合併三表
        merged = pd.merge(g_支審, g_fa300, on=['name', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['name', 'date', 'code'], how='outer')

        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker次數'] = merged['dmaker次數'].fillna(0).astype(int)

        # 篩選異常資料
        diff = merged[
            (merged['支審次數'] != merged['FA300次數']) | 
            (merged['支審次數'] != merged['dmaker次數'])
        ].sort_values(by=['date', 'name'])

        diff = diff[diff['date'] != '']

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
            st.success("🎉 太棒了！每日服務項目與次數完全吻合！")
        else:
            st.warning(f"⚠️ 發現 {len(diff_show)} 筆差異紀錄：")
            st.dataframe(diff_show, use_container_width=True)

            @st.cache_data
            def convert_df(df): return df.to_csv(index=False).encode('utf-8-sig')

            st.download_button("📥 下載異常明細 (CSV)", convert_df(diff_show), "每日服務項目異常明細.csv", "text/csv")

    except Exception as e:
        st.error(f"資料處理失敗：{e}")
