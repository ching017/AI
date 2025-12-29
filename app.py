import streamlit as st
import pulp
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="診所/單位自動排班系統", layout="wide")

st.title("🏥 專屬自動排班系統")
st.info("規則：一三五早3人、隔週六早3人、二晚及四午2人、其餘時段1-2人。")

# 1. 基本設定
nurses = ["莊欣蓓", "陳思伶", "王靜怡", "黃馨榆", "陳菁萱", "楊詠淳", "蔡宜軒"]
days = list(range(1, 29))  # 設定排 4 週 (28天)
shifts = ["早", "午", "晚"]
day_names = ["一", "二", "三", "四", "五", "六", "日"]

# 2. 定義人力需求函式
def get_requirement(day_index, shift):
    weekday = (day_index - 1) % 7  # 0=Mon, 6=Sun
    week_num = (day_index - 1) // 7 + 1
    
    # 週六晚上、週日午晚不排班
    if (weekday == 5 and shift == "晚") or (weekday == 6 and (shift == "午" or shift == "晚")):
        return 0
    
    if shift == "早":
        if weekday in [0, 2, 4]: # 一三五
            return 3
        if weekday == 5: # 週六
            return 3 if week_num % 2 == 1 else 2 # 隔週六(第1,3週)3人，其餘2人
        return 2 # 其他早上 (二、四、日)
    
    if shift == "午":
        if weekday == 3: # 週四下午
            return 2
        return 1
    
    if shift == "晚":
        if weekday == 1: # 週二晚上
            return 2
        return 1
    
    return 1

# 3. 開始計算
if st.button("開始生成 4 週班表"):
    prob = pulp.LpProblem("NurseScheduling", pulp.LpMinimize)
    
    # 變數：x[n, d, s] = 1 代表護理師 n 在第 d 天上 s 班
    x = pulp.LpVariable.dicts("x", (nurses, days, shifts), cat="Binary")
    
    # 目標函數：盡量讓每個人總班數平均 (軟約束)
    total_shifts = pulp.LpVariable.dicts("total_shifts", nurses, lowBound=0)
    for n in nurses:
        prob += total_shifts[n] == pulp.lpSum([x[n][d][s] for d in days for s in shifts])
    
    # 限制條件
    for d in days:
        # 每班人力需求
        for s in shifts:
            prob += pulp.lpSum([x[n][d][s] for n in nurses]) == get_requirement(d, s)
        
        # 每人每天只能上一個班 (避免連上)
        for n in nurses:
            prob += pulp.lpSum([x[n][d][s] for s in shifts]) <= 1

    # 求解
    prob.solve(pulp.PULP_CBC_CMD(msg=0))
    
    if pulp.LpStatus[prob.status] == "Optimal":
        # 整理結果
        schedule_data = []
        for d in days:
            day_info = {"日期": f"第{d}天(週{day_names[(d-1)%7]})"}
            for s in shifts:
                assigned = [n for n in nurses if pulp.value(x[n][d][s]) == 1]
                day_info[s] = ", ".join(assigned) if assigned else "---"
            schedule_data.append(day_info)
        
        df = pd.DataFrame(schedule_data)
        st.success("班表生成成功！")
        st.dataframe(df, height=800)
        
        # 下載功能
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button("下載 Excel 班表", data=output.getvalue(), file_name="schedule.xlsx")
    else:
        st.error("無法找到符合規則的解，請放寬限制。")
