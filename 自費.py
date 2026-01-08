import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="醫師自費分流-對帳版", layout="wide")

st.title("📊 醫師看診自費資料自動分流 (含網頁對帳表)")
st.info("💡 提示：分流後的 Excel 「總計表」會出現在每個分頁的最下方，請記得往下捲動。")

uploaded_file = st.file_uploader("請上傳原始 Excel 總表", type=["xlsx"])

if uploaded_file:
    try:
        # 1. 讀取資料
        df_all = pd.read_excel(uploaded_file, sheet_name="ALL")
        
        # 2. 資料清洗 (處理金額符號)
        df_all['自費'] = pd.to_numeric(df_all['自費'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        # 3. 提取月份與排序
        df_all['日期'] = df_all['日期'].astype(str)
        df_all['月份'] = df_all['日期'].str[3:5] + "月"
        df_all = df_all.sort_values(by='日期')

        # --- 網頁預覽對帳表 ---
        st.divider()
        st.subheader("📋 網頁即時對帳 (各醫師每月自費總計)")
        
        df_cleaned = df_all.dropna(subset=['醫'])
        doctor_codes = sorted(df_cleaned['醫'].unique())
        
        # 在網頁上用分頁顯示各醫師總計
        tabs = st.tabs([f"醫師 {str(int(c)).zfill(2)}" for c in doctor_codes])
        
        for i, code in enumerate(doctor_codes):
            with tabs[i]:
                doc_data = df_cleaned[df_cleaned['醫'] == code]
                doc_summary = doc_data.groupby('月份')['自費'].sum().reset_index()
                doc_summary.columns = ['月份', '該月自費總計']
                
                # 顯示該醫師的總計表
                st.table(doc_summary)
                st.write(f"**年度總和：${doc_summary['該月自費總計'].sum():,.0f}**")

        # --- 執行 Excel 下載 ---
        if st.button("🚀 下載完整 Excel (含底部統計表)"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_all.to_excel(writer, sheet_name="ALL", index=False)
                
                for code in doctor_codes:
                    sheet_name = str(int(code)).zfill(2)
                    doctor_data = df_cleaned[df_cleaned['醫'] == code].copy()
                    
                    # 寫入明細
                    doctor_data.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # 計算總計並寫在明細下方 (間隔兩行)
                    summary = doctor_data.groupby('月份')['自費'].sum().reset_index()
                    summary.columns = ['月份', '自費總計']
                    
                    start_row = len(doctor_data) + 3
                    summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
                    
                    # 標註標題
                    writer.sheets[sheet_name].cell(row=start_row, column=1, value="【每月總計統計表】")

            st.download_button(
                label="📥 點我下載報表",
                data=output.getvalue(),
                file_name="醫師自費總計分流表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"發生錯誤：{e}")
