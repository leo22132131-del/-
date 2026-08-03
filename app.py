import streamlit as st
import pandas as pd
import re

st.title("長照居家服務核對系統（精準按日期核對）")
st.write("請上傳三個 Excel 報表，系統將自動修正日期格式並按【姓名 + 服務日期 + 服務代碼】比對。")

file_支審 = st.file_uploader("1. 上傳 支審資料 (Excel)", type=["xlsx", "xls"])
file_fa300 = st.file_uploader("2. 上傳 FA300 (Excel)", type=["xlsx", "xls"])
file_dmaker = st.file_uploader("3. 上傳 dmaker (Excel)", type=["xlsx", "xls"])

def extract_ba_code(text):
    if pd.isna(text):
        return ""
    match = re.search(r'([A-Z]{2}\d{2})', str(text).upper())
    return match.group(1) if match else str(text).strip()

def clean_date_smart(date_val):
    """精準處理台灣長照各種日期格式（包含 1150702, 115/07/02, 2026/07/02, Excel 時間戳記等）"""
    if pd.isna(date_val):
        return ""
    
    val_str = str(date_val).strip().split(' ')[0] # 去掉時間部分 (如 00:00:00)
    val_str = val_str.split('.')[0] # 去掉小數點 (如 .0)
    
    # 情況 A：純數字 7 碼民國年 (例如: 1130702 或 1150702)
    if len(val_str) == 7 and val_str.isdigit():
        y = int(val_str[:3]) + 1911
        m = val_str[3:5]
        d = val_str[5:7]
        return f"{y}-{m}-{d}"
        
    # 情況 B：純數字 6 碼民國年 (例如: 990702)
    if len(val_str) == 6 and val_str.isdigit():
        y = int(val_str[:2]) + 1911
        m = val_str[2:4]
        d = val_str[4:6]
        return f"{y}-{m}-{d}"

    # 情況 C：純數字 8 碼西元年 (例如: 20240702)
    if len(val_str) == 8 and val_str.isdigit():
        return f"{val_str[:4]}-{val_str[4:6]}-{val_str[6:8]}"

    # 情況 D：帶有符號的民國年 (例如: 113/7/2, 115-07-02, 113.07.02)
    m = re.match(r'^(\d{2,3})[/.-](\d{1,2})[/.-](\d{1,2})', val_str)
    if m:
        y, month, d = m.groups()
        if int(y) < 1000:  # 代表是民國年
            y = int(y) + 1911
        return f"{int(y):04d}-{int(month):02d}-{int(d):02d}"

    # 情況 E：一般西元年
    try:
        dt = pd.to_datetime(val_str)
        if dt.year < 1920: # 避免轉成 0115年 等異常年份
            dt = dt.replace(year=dt.year + 1911)
        return dt.strftime('%Y-%m-%d')
    except:
        return val_str

def clean_name(name):
    if pd.isna(name):
        return ""
    return re.sub(r'\s+', '', str(name))

def find_column(df, possible_names):
    df.columns = [str(c).strip() for c in df.columns]
    for name in possible_names:
        if name in df.columns:
            return name
    return None

if file_支審 and file_fa300 and file_dmaker:
    try:
        # 1. 支審資料
        df_支審 = pd.read_excel(file_支審)
        col_date_支審 = find_column(df_支審, ['服務日期(請輸入7碼)', '服務日期', '費用日期', '日期'])
        col_code_支審 = find_column(df_支審, ['服務項目代碼', '服務項目名稱', '服務項目', '項目代碼'])
        col_name_支審 = find_column(df_支審, ['個案姓名', '姓名', '客戶名'])

        df_支審['date'] = df_支審[col_date_支審].apply(clean_date_smart)
        df_支審['code'] = df_支審[col_code_支審].apply(extract_ba_code)
        df_支審['clean_name'] = df_支審[col_name_支審].apply(clean_name)
        df_支審 = df_支審[~df_支審[col_code_支審].astype(str).str.contains('QA1385', na=False)]
        
        g_支審 = df_支審.groupby(['clean_name', 'date', 'code']).size().reset_index(name='支審次數')

        # 2. FA300
        df_fa300 = pd.read_excel(file_fa300)
        col_date_fa300 = find_column(df_fa300, ['服務日期', '執行日期', '日期'])
        col_code_fa300 = find_column(df_fa300, ['服務項目', '服務項目名稱', '服務項目代碼'])
        col_name_fa300 = find_column(df_fa300, ['個案姓名', '姓名', '客戶名'])

        df_fa300['date'] = df_fa300[col_date_fa300].apply(clean_date_smart)
        df_fa300['code'] = df_fa300[col_code_fa300].apply(extract_ba_code)
        df_fa300['clean_name'] = df_fa300[col_name_fa300].apply(clean_name)
        
        g_fa300 = df_fa300.groupby(['clean_name', 'date', 'code']).size().reset_index(name='FA300次數')

        # 3. dmaker
        df_dmaker = pd.read_excel(file_dmaker)
        col_date_dmaker = find_column(df_dmaker, ['使用日期', '服務日期', '刷卡日期', '日期'])
        col_code_dmaker = find_column(df_dmaker, ['品名', '服務項目', '項目代碼'])
        col_name_dmaker = find_column(df_dmaker, ['客戶名', '個案姓名', '姓名'])

        df_dmaker['date'] = df_dmaker[col_date_dmaker].apply(clean_date_smart)
        df_dmaker['code'] = df_dmaker[col_code_dmaker].apply(extract_ba_code)
        df_dmaker['clean_name'] = df_dmaker[col_name_dmaker].apply(clean_name)

        if '數量' in df_dmaker.columns:
            g_dmaker = df_dmaker.groupby(['clean_name', 'date', 'code'])['數量'].sum().reset_index(name='dmaker數量')
        else:
            g_dmaker = df_dmaker.groupby(['clean_name', 'date', 'code']).size().reset_index(name='dmaker數量')

        # 4. 合併比對
        merged = pd.merge(g_支審, g_fa300, on=['clean_name', 'date', 'code'], how='outer')
        merged = pd.merge(merged, g_dmaker, on=['clean_name', 'date', 'code'], how='outer')

        merged['支審次數'] = merged['支審次數'].fillna(0).astype(int)
        merged['FA300次數'] = merged['FA300次數'].fillna(0).astype(int)
        merged['dmaker數量'] = merged['dmaker數量'].fillna(0).astype(int)

        result = merged.rename(columns={'clean_name': '個案姓名', 'date': '服務日期', 'code': '代碼(BA/BB)'})
        result = result[['個案姓名', '服務日期', '代碼(BA/BB)', '支審次數', 'FA300次數', 'dmaker數量']]

        # 排除日期無效或空白的項目
        result = result[result['服務日期'] != '1970-01-01']
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
        st.error(f"資料處理時發生錯誤：{e}")
