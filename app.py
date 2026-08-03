import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta

st.set_page_config(page_title="長照居家服務精準核對系統", layout="wide")
st.title("長照居家服務核對系統（以支審為基準）")
st.write("以【支審資料】為申報基準，精準比對 FA300 與 dmaker 的核對狀況。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

date_tolerance = st.checkbox("開啟「日期 ±1 天彈性核對」（若居服員常有跨夜打卡或隔天補登情況建議勾選）", value=False)

def extract_ba_code(text):
    if pd.isna(text): return ""
    match = re.search(r'([A-Z]{2}\d{2})', str(text).upper())
    return match.group(1) if match else str(text).strip()

def clean_date_universal(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    if isinstance(val, (pd.Timestamp, datetime)): return val.strftime('%Y-%m-%d')
    s = str(val).strip().split(' ')[0].split('.')[0]
    if s.isdigit() and len(s) == 5:
        try: return pd.to_datetime(int(s), unit='D', origin='1899-12-30').strftime('%Y-%m-%d')
        except: pass
    if s.isdigit() and len(s) == 7: return f"{int(s[:3])+1911}-{s[3:5]}-{s[5:7]}"
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
        df_支審 = pd.read_excel(file_支審)
        df_fa300 = pd.read_excel(file_fa300)
        df_dmaker = pd.read_excel(file_dmaker)

        # 找尋關鍵欄位
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

        # 排除 QA1385 等無效代碼
        df_支審 = df_支審[~df_支審[col_c_支審].astype(str).str.contains('QA1385', na=False)]
        df_支審 = df_支審[df_支審['code'] != ""]

        # 分組計算數量
        g_支審 = df_支審.groupby(['clean_name', 'date', 'code']).size().reset_index(name='支審申報數')
        g_fa300 = df_fa300.groupby(['clean_name', 'date', 'code']).size().reset_index(name='FA300紀錄數')
        
        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby(['clean_name', 'date', 'code'])['數量'].sum().reset_index(name='dmaker刷卡數')
        else:
            g_dmaker = df_dmaker.groupby(['clean_name', 'date', 'code']).size().reset_index(name='dmaker刷卡數')

        # 以支審資料為主體開始比對
        results = []
        for _, row in g_支審.iterrows():
            name = row['clean_name']
            date_str = row['date']
            code = row['code']
            target_qty = row['支審申報數']

            # 比對 FA300
            fa_match = g_fa300[(g_fa300['clean_name'] == name) & (g_fa300['date'] == date_str) & (g_fa300['code'] == code)]
            fa_qty = fa_match['FA300紀錄數'].sum() if len(fa_match) > 0 else 0

            # 比對 dmaker (支援 ±1 天彈性比對)
            if date_tolerance and date_str != "":
                try:
                    dt = datetime.strptime(date_str, '%Y-%m-%d')
                    valid_dates = [(dt + timedelta(days=i)).strftime('%Y-%m-%d') for i in [-1, 0, 1]]
                    dm_match = g_dmaker[(g_dmaker['clean_name'] == name) & (g_dmaker['date'].isin(valid_dates)) & (g_dmaker['code'] == code)]
                except:
                    dm_match = g_dmaker[(g_dmaker['clean_name'] == name) & (g_dmaker['date'] == date_str) & (g_dmaker['code'] == code)]
            else:
                dm_match = g_dmaker[(g_dmaker['clean_name'] == name) & (g_dmaker['date'] == date_str) & (g_dmaker['code'] == code)]
            
            dm_qty = dm_match['dmaker刷卡數'].sum() if len(dm_match) > 0 else 0

            # 判斷結果狀態
            status = []
            if fa_qty != target_qty: status.append(f"FA300不符(差{fa_qty - target_qty:+}次)")
            if dm_qty != target_qty: status.append(f"dmaker不符(差{dm_qty - target_qty:+}次)")

            results.append({
                '個案姓名': name,
                '申報日期': date_str,
                '服務代碼': code,
                '支審申報數': target_qty,
                'FA300次數': fa_qty,
                'dmaker次數': dm_qty,
                '異常說明': '；'.join(status) if status else '完全符合'
            })

        df_result = pd.DataFrame(results)
        df_diff = df_result[df_result['異常說明'] != '完全符合']

        st.subheader("📊 核對結果分析")
        col1, col2 = st.columns(2)
        col1.metric("支審申報總筆數", len(df_result))
        col2.metric("發現異常筆數", len(df_diff), delta_color="inverse")

        if len(df_diff) == 0:
            st.success("🎉 太棒了！支審申報的所有服務項目在 FA300 與 dmaker 均完全吻合！")
        else:
            st.warning(f"⚠️ 發現 {len(df_diff)} 筆申報資料存在紀錄不相符的情況：")
            st.dataframe(df_diff, use_container_width=True)

            @st.cache_data
            def convert_df(df): return df.to_csv(index=False).encode('utf-8-sig')

            st.download_button("📥 下載支審異常對照明細表 (CSV)", convert_df(df_diff), "長照支審申報異常核對表.csv", "text/csv")

    except Exception as e:
        st.error(f"資料核對失敗，請檢查檔案格式與欄位名稱：{e}")
