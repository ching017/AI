import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="醫師自費資料分流系統-進階結算版", layout="wide")

st.title("📊 醫師看診自費資料自動分流 (含每月總計)")
st.write("說明：系統會自動根據「日期」排序，並在每個醫師分頁底部計算每個月的「自費」總額。")

# --- 1. 檔案上傳 ---
uploaded_file = st.file_uploader("請上傳原始 Excel 總表", type=["xlsx"])

if uploaded_file:
    try:
        # 讀取原始 ALL 頁面
        df_all = pd.read_excel(uploaded_file, sheet_name="ALL")
        
        # --- 資料預處理 ---
        # A. 清洗「自費」欄位：轉為數字並處理千分號
        df_all['自費'] = pd.to_numeric(df_all['自費'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        # B. 提取月份：從「1140101」提取出「01月」
        df_all['日期'] = df_all['日期'].astype(str)
        df_all['月份'] = df_all['日期'].str[3:5] + "月"
        
        # C. 依日期排序
        df_all = df_all.sort_values(by='日期')

        if st.button("🚀 執行分流並計算每月總計"):
            # 移除「醫」欄位為空的資料
            df_cleaned = df_all.dropna(subset=['醫'])
            
            output = BytesIO()
            doctor_codes = df_cleaned['醫'].unique()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 保留原始總表
                df_all.to_excel(writer, sheet_name="ALL", index=False)
                
                # 根據「醫」代碼（3.0, 4.0...）分流
                for code in sorted(doctor_codes):
                    sheet_name = str(int(float(code))).zfill(2)
                    
                    # 篩選出該位醫師的資料
                    doctor_data = df_cleaned[df_cleaned['醫'] == code].copy()
                    
                    # 1. 寫入看診明細資料
                    doctor_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                    
                    # 2. 計算「每月自費總計」
                    # 分組計算每個月的自費總和
                    summary = doctor_data.groupby('月份')['自費'].sum().reset_index()
                    summary.columns = ['月份', '自費總計金額']
                    
                    # 3. 將總計表寫在明細下方 (間隔 3 行)
                    start_row = len(doctor_data) + 3
                    summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
                    
                    # 在統計表上方加上標題
                    worksheet = writer.sheets[sheet_name]
                    worksheet.cell(row=start_row, column=1, value="【每月自費結算表】")
            
            st.success(f"✅ 分流與結算完成！")
            
            # --- 下載按鈕 ---
            st.download_button(
                label="📥 下載分類與結算完成版 Excel",
                data=output.getvalue(),
                file_name="114年自費_醫師分流結算版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"執行出錯：{e}")
