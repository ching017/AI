import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="診所自費資料分流系統", layout="wide")

st.title("📊 醫師看診自費資料自動分流工具")
st.info("說明：系統會讀取 'ALL' 頁面，並根據 '醫' 欄位（如 3.0, 4.0）自動分類至分頁 '03', '04' 等。")

# --- 1. 檔案上傳 ---
uploaded_file = st.file_uploader("請上傳您的 114年自費.xlsx 原始檔案", type=["xlsx"])

if uploaded_file:
    try:
        # 讀取原始總表
        df_all = pd.read_excel(uploaded_file, sheet_name="ALL")
        
        # 清洗數據：移除「醫」欄位為空的列
        df_all = df_all.dropna(subset=['醫'])
        
        # --- 2. 執行分流運算 ---
        if st.button("開始執行分流與生成報表"):
            output = BytesIO()
            
            # 取得所有醫師代碼 (例如 3.0, 4.0...)
            doctor_codes = df_all['醫'].unique()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 首先保留原始 ALL 頁面
                df_all.to_excel(writer, sheet_name="ALL", index=False)
                
                # 根據「醫」代碼分流
                for code in sorted(doctor_codes):
                    # 格式化代碼：3.0 -> "03", 10.0 -> "10"
                    sheet_name = str(int(code)).zfill(2)
                    
                    # 篩選該醫師的資料
                    doctor_data = df_all[df_all['醫'] == code]
                    
                    # 寫入對應分頁
                    doctor_data.to_excel(writer, sheet_name=sheet_name, index=False)
            
            st.success(f"✅ 分流處理完成！共處理了 {len(doctor_codes)} 位醫師的資料。")
            
            # --- 3. 提供下載 ---
            st.download_button(
                label="📥 下載分類完成的 Excel 檔案",
                data=output.getvalue(),
                file_name="114年自費_各診自動分類版.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"發生錯誤：{e}")
        st.warning("請確保您的 Excel 檔案中確實有名為 'ALL' 的分頁，且包含 '醫' 欄位。")

# --- 顯示資料預覽 ---
if uploaded_file:
    st.divider()
    st.subheader("原始資料預覽 (ALL)")
    st.dataframe(pd.read_excel(uploaded_file, sheet_name="ALL").head(10))
