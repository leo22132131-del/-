import pandas as pd
import streamlit as st
import io
import re

st.set_page_config(page_title="長照服務費用三方核對系統", layout="wide")
st.title("📊 長照服務費用三方核對系統")
st.write("請選擇上傳 **支審資料（日照/居家可單獨或同時上傳）**、**FA300報表** 與 **dmaker報表**，並點擊下方「確定送出交叉比對」按鈕。")

# 1. 檔案上傳區塊
col1, col2, col3, col4 = st.columns(4)
with col1:
    file_支審_日照 = st.file_uploader("1. 支審資料-日照 (可選)", type=["xls", "xlsx"])
with col2:
    file_支審_居家 = st.file_uploader("2. 支審資料-居家 (可選)", type=["xls", "xlsx"])
with col3:
    file_FA300 = st.file_uploader("3. FA300 報表 (必填)", type=["xls", "xlsx"])
with col4:
    file_dmaker = st.file_uploader("4. dmaker 報表 (必填)", type=["xls", "xlsx"])

st.markdown("---")

# 2. 確定送出比對按鈕
submit_btn = st.button("🚀 確定送出交叉比對", type="primary", use_container_width=True)

if submit_btn:
    has_支審 = (file_支審_日照 is not None) or (file_支審_居家 is not None)
    
    if not has_支審:
        st.error("❌ 請至少上傳「支審資料-日照」或「支審資料-居家」其中一份！")
    elif not file_FA300:
        st.error("❌ 請上傳「FA300 報表」！")
    elif not file_dmaker:
        st.error("❌ 請上傳「dmaker 報表」！")
    else:
        st.success("✅ 檔案檢查無誤，開始執行交叉比對...")

        # 1. 讀取並合併支審資料
        list_df_支審 = []
        if file_支審_日照:
            list_df_支審.append(pd.read_excel(file_支審_日照))
        if file_支審_居家:
            list_df_支審.append(pd.read_excel(file_支審_居家))

        df1 = pd.concat(list_df_支審, ignore_index=True)
        df2 = pd.read_excel(file_FA300)
        df3 = pd.read_excel(file_dmaker)

        # 2. 清理代碼（自動去除尾綴 -1, -2）
        def clean_code(text):
            if pd.isna(text): return ""
            s = str(text).strip()
            return re.sub(r'-\d+$', '', s)

        df1["name"] = df1["個案姓名"].astype(str).str.strip().str.replace("鳯", "鳳")
        df1["code"] = df1["服務項目代碼"].apply(clean_code)

        df2["name"] = df2["個案姓名"].astype(str).str.strip().str.replace("鳯", "鳳")
        df2["code"] = df2["服務項目"].astype(str).str.extract(r"([A-Z]{2}\d{2})")[0].apply(clean_code)

        df3["name"] = df3["客戶名"].astype(str).str.strip().str.replace("鳯", "鳳")
        df3["code"] = df3["品名"].astype(str).str.extract(r"([A-Z]{2}\d{2}|QA1385)")[0].apply(clean_code)

        # 3. 分類 dmaker 服務類型（避免公費與部分負擔重複採計）
        df3["is_self_pay"] = df3["品名"].str.contains("自費", na=False)
        # 若品名有「公費」，直接採計為公費；若無「公費」但有「部分負擔」（如 BA 居家碼），採計為公費
        df3["is_public"] = df3["品名"].str.contains("公費", na=False)
        df3["is_copay"] = df3["品名"].str.contains("部分負擔", na=False)

        # 若同時有公費記錄，優先使用公費記錄；否則使用部分負擔記錄
        df3_public_only = df3[df3["is_public"]]
        df3_copay_only = df3[df3["is_copay"] & (~df3["is_public"])]
        df3_public_combined = pd.concat([df3_public_only, df3_copay_only], ignore_index=True)

        c3_public = df3_public_combined.groupby(["name", "code"])["數量"].sum().reset_index(name="dmaker_公費次數")
        c3_self = df3[df3["is_self_pay"]].groupby(["name", "code"])["數量"].sum().reset_index(name="dmaker_自費次數")

        dmaker_summary = pd.merge(c3_public, c3_self, on=["name", "code"], how="outer").fillna(0)

        # 4. 彙整 支審 與 FA300
        c1 = df1.groupby(["name", "code"]).size().reset_index(name="支審次數")
        c2 = df2.groupby(["name", "code"]).size().reset_index(name="FA300次數")

        # 5. 三方 Cross Merge
        final = pd.merge(c1, c2, on=["name", "code"], how="outer")
        final = pd.merge(final, dmaker_summary, on=["name", "code"], how="outer").fillna(0)

        for col in ["支審次數", "FA300次數", "dmaker_公費次數", "dmaker_自費次數"]:
            final[col] = final[col].astype(int)

        final["dmaker_公自費合計"] = final["dmaker_公費次數"] + final["dmaker_自費次數"]
        final["公費異常"] = final["FA300次數"] != final["dmaker_公費次數"]
        final["總數異常"] = (final["支審次數"] != final["dmaker_公自費合計"]) & (final["code"] != "QA1385")

        diff_df = final[final["公費異常"] | final["總數異常"]]

        # 6. 結果面板
        st.subheader("📌 比對結果總覽")
        m1, m2, m3 = st.columns(3)
        m1.metric("總核對長照項目組數", len(final))
        m2.metric("完全吻合組數", len(final) - len(diff_df))
        m3.metric("不吻合/待確認組數", len(diff_df), delta_color="inverse")

        if len(diff_df) == 0:
            st.balloons()
            st.success("🎉 太棒了！所有長照項目的公費與自費數量 100% 完全吻合！")
        else:
            st.warning(f"⚠️ 發現 {len(diff_df)} 筆項目不吻合，請查看下方明細：")
            st.dataframe(diff_df[["name", "code", "支審次數", "FA300次數", "dmaker_公費次數", "dmaker_自費次數", "dmaker_公自費合計"]], use_container_width=True)

        # 7. 下載報表匯出
        st.subheader("📥 下載完整核對結果")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            final.to_excel(writer, sheet_name="完整比對表", index=False)
            if len(diff_df) > 0:
                diff_df.to_excel(writer, sheet_name="異常明細表", index=False)

        st.download_button(
            label="下載核對結果 Excel 報表",
            data=buffer.getvalue(),
            file_name="長照費用核對報告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
