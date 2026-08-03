import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="長照居家服務核對系統", layout="wide")
st.title("長照居家服務核對系統（極簡異常版）")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

dmaker_mode = st.radio(
    "dmaker 計算方式：",
    ["按『簽到筆數』計算 (推薦：避免簽到退重複算2次)", "按『數量欄位』加總"],
    index=0
)

def extract_ba_code(text):
    if pd.isna(text): return ""
    s = str(text).upper().strip()
    match = re.search(r'([A-Z]{2}[-_\s]?\d{2})', s)
    if match: return re.sub(r'[-_\s]', '', match.group(1))
    return s

def detect_is_self_pay(row, col_code, col_type=None):
    code_text = str(row[col_code]) if col_code and not pd.isna(row[col_code]) else ""
    if any(k in code_text for k in ['自費', '超過額度', '全額自費', '自費項']): return "自費"
    if col_type and not pd.isna(row[col_type]):
        type_text = str(row[col_type])
        if any(k in type_text for k in ['自費', '全額', '超過']): return "自費"
    return "公費"

def clean_name(name):
    if pd.isna(name): return ""
    return re.sub(r'\s+', '', str(name))

def clean_date(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    s = str(val).strip().split(' ')[0].split('.')[0]
    if s.isdigit() and len(s) == 5:
        try: return pd.to_datetime(int(s), unit='D', origin='1899-12-30').strftime('%Y-%m-%d')
        except: pass
    if s.isdigit() and len(s) == 7:
        y = int(s[:3]) + 1911
        return f"{y}-{s[3:5]}-{s[5:7]}"
    if s.isdigit() and len(s) == 8: return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
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

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns: return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        df_支審 = pd.read_excel(file_支審)
        df_fa300 = pd.read_excel(file_fa300)
        df_dmaker = pd.read_excel(file_dmaker)

        col_d_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_c_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_n_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])
        col_t_支審 = find_column(df_支審, ['費用類別', '身分別', '公自費', '付款方式'])

        col_d_fa300 = find_column(df_fa300, ['服務日期', '執行日期', '日期'])
        col_c_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_n_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])
        col_t_fa300 = find_column(df_fa300, ['費用類別', '身分別', '公自費', '付款方式'])

        col_d_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_c_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_n_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])
        col_q_dmaker = find_column(df_dmaker, ['數量', '服務數量', '次數'])
        col_t_dmaker = find_column(df_dmaker, ['費用類別', '身分別', '公自費', '類別'])

        # 1. 支審
        df_支審['date'] = df_支審[col_d_支審].apply(clean_date)
        df_支審['code'] = df_支審[col_c_支審].apply(extract_ba_code)
        df_支審['name'] = df_支審[col_n_支審].apply(clean_name)
        df_支審['pay_type'] = df_支審.apply(lambda r: detect_is_self_pay(r, col_c_支審, col_t_支審), axis=1)
        df_支審 = df_支審[~df_支審[col_c_支審].astype(str).str.contains('QA1385', na=False)]
        g_支審 = df_支審.groupby(['name', 'date', 'code', 'pay_type']).size().reset_index(name='支審次數')

        # 2. FA300
        df_fa300['date'] = df_fa300[col_d_fa300].apply(clean_date)
        df_fa300['code'] = df_fa300[col_c_fa300].apply(extract_ba_code)
        df_fa300['name'] = df_fa300[col_n_fa300].apply(clean_name)
        df_fa300['pay_type'] = df_fa300.apply(lambda r: detect_is_self_pay(r, col_c_fa300, col_t_fa300), axis=1)
        g_fa300 = df_fa300.groupby(['name', 'date', 'code', 'pay_type']).size().reset_index(name='FA300次數')

        # 3. dmaker
        df_dmaker['date'] = df_dmaker[col_d_dmaker].apply(clean_date)
        df_dmaker['code'] = df_dmaker[col_c_dmaker].apply(extract_ba_code)
        df_dmaker['name'] = df_dmaker[col_n_dmaker].apply(clean_name)
        df_dmaker['pay_type'] = df_dmaker.apply(lambda r: detect_is_self_pay(r, col_c_dmaker, col_t_dmaker), axis=1)
        
        if "按『簽到筆數』計算" in dmaker_mode:
            g_dmaker = df_dmaker.groupby(['name', 'date', 'code', 'pay_type']).size().reset_index(name='dmaker數量')
        else:
            if col_q_dmaker:
                df_dmaker['qty'] = pd.to_numeric(df_dmaker[col_q_dmaker], errors='coerce').fillna(1)
                g_dmaker = df_dmaker.groupby(['name', 'date', 'code', 'pay_type'])['qty'].sum().reset_index(name='dmaker數量')
            else:
                g_dmaker = df_dmaker.groupby(['name', 'date', 'code', 'pay_type']).size().reset_index(name='dmaker數量')

        # 合併
        merged = pd.merge(g_支審, g_fa300, on=['name', 'date', 'code', 'pay_type'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['name', 'date', 'code', 'pay_type'], how='outer')

        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker數量'] = merged['dmaker數量'].fillna(0).astype(int)

        # 組合簡單的說明欄位
        def make_status(r):
            status = []
            z = r['支審次數']
            f = r['FA300次數']
            d = r['dmaker數量']
            if z != f:
                status.append(f"FA300少{z-f}次" if z > f else f"FA300多{f-z}次")
            if z != d:
                status.append(f"dmaker少{z-d}次" if z > d else f"dmaker多{d-z}次")
            return "；".join(status)

        merged['狀態說明'] = merged.apply(make_status, axis=1)

        result = merged.rename(columns={
            'date': '服務日期', 
            'name': '個案姓名', 
            'code': '服務代碼', 
            'pay_type': '費用類別'
        })
        
        # 只保留簡單的欄位
        result = result[['服務日期', '個案姓名', '服務代碼', '費用類別', '狀態說明']]

        # 篩選有異常的
        diff = result[result['狀態說明'] != ''].sort_values(by=['服務日期', '個案姓名'])
        diff = diff[diff['服務日期'] != '']

        st.subheader("📌 異常服務項目清單")
        if len(diff) == 0:
            st.success("🎉 太棒了！沒有發現任何異常！")
        else:
            st.warning(f"⚠️ 共有 {len(diff)} 筆資料存在紀錄不相符：")
            st.dataframe(diff, use_container_width=True)

            @st.cache_data
            def convert_df(df): return df.to_csv(index=False).encode('utf-8-sig')

            st.download_button("📥 下載異常清單 (CSV)", convert_df(diff), "服務項目異常清單.csv", "text/csv")

    except Exception as e:
        st.error(f"資料處理失敗：{e}")
