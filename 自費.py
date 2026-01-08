import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="醫師自費資料分流系統-進階版", layout="wide")

st.title("📊 醫師看診自費資料自動分流 (含每月小計)")
st.write("操作說明：上傳 Excel 總表後，系統會自動按日期排序並計算每月自費總和。")

# --- 1. 檔案上傳 ---
uploaded_file = st.file_uploader("請上傳您的 114年自費.xlsx", type=["xlsx"])

if uploaded_file:
    try:
        # 讀取原始 ALL 頁面
        df_all = pd.read_excel(uploaded_file, sheet_name="ALL")
        
        # --- 資料預處理 ---
        # A. 確保「自費」是數字格式
        df_all['自費'] = pd.to_numeric(df_all['自費'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        
        # B. 提取月份 (從 1140101 提取出 01)
        df_all['日期'] = df_all['日期'].astype(str)
        df_all['月份'] = df_all['日期'].str[3:5] + "月"
        
        # C. 依日期排序
        df_all = df_all.sort_values(by='日期')

        st.subheader("原始資料預覽 (已排序)")
        st.dataframe(df_all.head(10), use_container_width=True)

        if st.button("🚀 執行分流與計算小計"):
            # 移除「醫」欄位為空的列
            df_cleaned = df_all.dropna(subset=['醫'])
            
            output = BytesIO()
            doctor_codes = df_cleaned['醫'].unique()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 保留原始總表
                df_all.to_excel(writer, sheet_name="ALL", index=False)
                
                # 根據「醫」代碼分流
                for code in sorted(doctor_codes):
                    sheet_name = str(int(float(code))).zfill(2)
                    
                    # 篩選該醫師資料
                    doctor_data = df_cleaned[df_cleaned['醫'] == code].copy()
                    
                    # 1. 寫入該醫師的所有看診明細
                    doctor_data.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                    
                    # 2. 計算該醫師的每月小計
                    summary = doctor_data.groupby('月份')['自費'].sum().reset_index()
                    summary.columns = ['月份', '該月自費總計']
                    
                    # 3. 將小計表格寫在明細資料的下方 (空兩行)
                    start_row = len(doctor_data) + 3
                    summary.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)
                    
                    # 在 Excel 裡標註這是小計表
                    worksheet = writer.sheets[sheet_name]
                    worksheet.cell(row=start_row, column=1, value="--- 每月自費金額統計 ---")
            
            st.success(f"✅ 分流與小計計算完成！")
            
            # --- 2. 下載按鈕 ---
            st.download_button(
                label="📥 下載含每月小計的 Excel 檔案",
                data=output.getvalue(),
                file_name="114年自費_醫師明細與每月小計.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"錯誤：{e}")
        st.info("提示：請確保 Excel 中包含 'ALL' 頁面，且有 '日期'、'醫'、'自費' 欄位。")
