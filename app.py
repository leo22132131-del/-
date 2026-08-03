echo 'import pandas as pd
import streamlit as st
import io

st.set_page_config(page_title="長照服務費用三方核對系統", layout="wide")
st.title("📊 長照服務費用三方核對系統")
st.write("請分別上傳 **支審資料**、**FA300報表** 與 **dmaker報表**，系統將自動進行交叉核對。")

col1, col2, col3 = st.columns(3)
with col1:
    file_支審 = st.file_uploader("1. 上傳 支審資料 (.xls/.xlsx)", type=["xls", "xlsx"])
with col2:
    file_FA300 = st.file_uploader("2. 上傳 FA300 (.xls/.xlsx)", type=["xls", "xlsx"])
with col3:
    file_dmaker = st.file_uploader("3. 上傳 dmaker (.xls/.xlsx)", type=["xls", "xlsx"])

if file_支審 and file_FA300 and file_dmaker:
    st.success("三個檔案皆已上傳，開始進行自動比對...")
    df1 = pd.read_excel(file_支審)
    df2 = pd.read_excel(file_FA300)
    df3 = pd.read_excel(file_dmaker)

    df1["name"] = df1["個案姓名"].astype(str).str.strip().str.replace("鳯", "鳳")
    df1["code"] = df1["服務項目代碼"].astype(str).str.strip()

    df2["name"] = df2["個案姓名"].astype(str).str.strip().str.replace("鳯", "鳳")
    df2["code"] = df2["服務項目"].astype(str).str.extract(r"([A-Z]{2}\d{2})")[0]

    df3["name"] = df3["客戶名"].astype(str).str.strip().str.replace("鳯", "鳳")
    df3["code"] = df3["品名"].astype(str).str.extract(r"([A-Z]{2}\d{2}|QA1385)")[0]

    df3["is_self_pay"] = df3["品名"].str.contains("自費", na=False)
    df3["is_public"] = df3["品名"].str.contains("公費", na=False)
    df3["is_copay"] = df3["品名"].str.contains("部分負擔", na=False)

    c3_public = df3[df3["is_public"]].groupby(["name", "code"])["數量"].sum().reset_index(name="dmaker_公費次數")
    c3_self = df3[df3["is_self_pay"]].groupby(["name", "code"])["數量"].sum().reset_index(name="dmaker_自費次數")
    c3_copay = df3[df3["is_copay"]].groupby(["name", "code"])["數量"].sum().reset_index(name="dmaker_部分負擔次數")

    dmaker_summary = pd.merge(c3_public, c3_self, on=["name", "code"], how="outer")
    dmaker_summary = pd.merge(dmaker_summary, c3_copay, on=["name", "code"], how="outer").fillna(0)

    c1 = df1.groupby(["name", "code"]).size().reset_index(name="支審次數")
    c2 = df2.groupby(["name", "code"]).size().reset_index(name="FA300次數")

    final = pd.merge(c1, c2, on=["name", "code"], how="outer")
    final = pd.merge(final, dmaker_summary, on=["name", "code"], how="outer").fillna(0)

    for col in ["支審次數", "FA300次數", "dmaker_公費次數", "dmaker_自費次數", "dmaker_部分負擔次數"]:
        final[col] = final[col].astype(int)

    final["dmaker_公自費合計"] = final["dmaker_公費次數"] + final["dmaker_自費次數"]
    final["公費異常"] = final["FA300次數"] != final["dmaker_公費次數"]
    final["總數異常"] = (final["支審次數"] != final["dmaker_公自費合計"]) & (final["code"] != "QA1385")

    diff_df = final[final["公費異常"] | final["總數異常"]]

    st.subheader("📌 比對結果總覽")
    m1, m2, m3 = st.columns(3)
    m1.metric("總核對長照項目組數", len(final))
    m2.metric("完全吻合組數", len(final) - len(diff_df))
    m3.metric("不吻合/待確認組數", len(diff_df), delta_color="inverse")

    if len(diff_df) == 0:
        st.balloons()
        st.success("🎉 太棒了！所有長照項目的公費與自費數量 100% 完全吻合！")
    else:
        st.error(f"⚠️ 發現 {len(diff_df)} 筆項目不吻合，請查看下方明細：")
        st.dataframe(diff_df[["name", "code", "支審次數", "FA300次數", "dmaker_公費次數", "dmaker_自費次數", "dmaker_公自費合計"]])

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
    )' > ~/Downloads/app.py
