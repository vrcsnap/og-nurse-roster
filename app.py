import streamlit as st
import datetime
import calendar
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model
import io

# 頁面配置
st.set_page_config(page_title="O&G 產房手術室護士排班系統 (r3)", layout="wide", page_icon="🏥")

# 常數定義 (0=Mon, ..., 6=Sun)
MON, TUE, WED, THU, FRI, SAT, SUN = 0, 1, 2, 3, 4, 5, 6
HANDOVER_DAYS = (MON, TUE, THU)
SPECIAL_DAYS = (TUE, WED, FRI)
WEEKEND_DAYS = (SAT, SUN)

DEFAULT_NURSES = [
    {"id": 1, "name": "Wong Choi Yu", "cap": "C(OTIC)"},
    {"id": 2, "name": "Jennifer Tsang", "cap": "C(OTIC)"},
    {"id": 3, "name": "Lee Ming-kwan", "cap": "C(OTIC)"},
    {"id": 4, "name": "Stella Ng", "cap": "C(OTIC)"},
    {"id": 5, "name": "Phoebe Sham", "cap": "C(OTIC)"},
    {"id": 6, "name": "Wong Sin Yi", "cap": "C(OTIC)"},
    {"id": 7, "name": "Kwok Lai Ming", "cap": "C(OTIC)"},
    {"id": 8, "name": "Tam Nga Chi Janette", "cap": "C(OTIC)"},
    {"id": 9, "name": "Lam Yuen Kiu", "cap": "C(OTIC)"},
    {"id": 10, "name": "Fan Sui Ying", "cap": "C(OTIC)"},
    {"id": 11, "name": "Liu Yuen Man", "cap": "C(OTIC)"},
    {"id": 12, "name": "Yu Chung Yin", "cap": "C(OTIC)"},
    {"id": 13, "name": "Ng Wing Yan", "cap": "C(OTIC)"},
    {"id": 14, "name": "Leung Hoi Tai", "cap": "C(OTIC)"},
    {"id": 15, "name": "Pang Chui Ting", "cap": "C(OTIC)"},
    {"id": 16, "name": "Au Cheuk Kiu", "cap": "C(OTIC)"},
    {"id": 17, "name": "Cheng Yuen Hing", "cap": "C(OTIC)"},
    {"id": 18, "name": "Wong Chung Man", "cap": "C(OTIC)"},
    {"id": 19, "name": "Poon Tsz Ki", "cap": "C(Scrub)"},
    {"id": 20, "name": "Yiu Wing Sheung", "cap": "C(Scrub)"},
    {"id": 21, "name": "Chan Sze Wai", "cap": "C(Scrub)"},
    {"id": 22, "name": "Chan Wing Yee", "cap": "C(Runner)"},
    {"id": 23, "name": "Tsui Tsz Ching", "cap": "C(Recovery)"},
    {"id": 24, "name": "Yam Ka Ki", "cap": "C(Recovery)"},
    {"id": 25, "name": "Cheng Ka Yan", "cap": "C(Runner)"},
    {"id": 26, "name": "Tsoi Ka Ying", "cap": "C(Runner)"},
    {"id": 27, "name": "Chung Sin Kei Sandy", "cap": "C(Runner)"},
    {"id": 28, "name": "Kong Tsz Sin Sarah", "cap": "C(Runner)"}
]

SHIFT_OPTIONS = {
    "想放假 Day Off (O)": "O",
    "想返早班 AM (A)": "A",
    "想返午班 PM (P)": "P",
    "想返夜班 Night (N)": "N"
}

# 初始化 Session 狀態
if "requests_db" not in st.session_state:
    st.session_state.requests_db = []
if "generated_schedule" not in st.session_state:
    st.session_state.generated_schedule = None

# ==========================================
# 介面 Page 1: 人員身份選擇 (下拉選單)
# ==========================================
st.sidebar.title("🏥 O&G 排班系統 (r3)")
nurse_names = [n["name"] for n in DEFAULT_NURSES]
user_options = ["請選擇...", "Admin (Ward Manager)"] + nurse_names

selected_user = st.sidebar.selectbox("請選擇使用者身份：", user_options)

# ==========================================
# 介面分支 A：護士專用操作頁面 (Page 2 - 4)
# ==========================================
if selected_user not in ("請選擇...", "Admin (Ward Manager)"):
    st.header(f"👋 您好，{selected_user}")
    
    tab1, tab2 = st.tabs(["📝 填寫特別申請 (Make Special Request)", "📅 查看已公佈班表 (See Duty Schedule)"])
    
    with tab1:
        st.subheader("15 號前登記下月特別更次 / 放假申請")
        st.info("提示：若下月無特別偏好則無須填寫。系統會自動以「P-A-N-O」及「雙連休 O-O」的人性化優選模式為您排班。")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            req_year = st.selectbox("年份", (2026, 2027), index=0, key="n_yr")
        with col2:
            current_month = datetime.date.today().month
            default_next_month = (current_month % 12) + 1
            req_month = st.selectbox("月份", list(range(1, 13)), index=default_next_month - 1, key="n_mo")
        with col3:
            _, max_day = calendar.monthrange(req_year, req_month)
            req_day = st.number_input("日期", min_value=1, max_value=max_day, value=1, step=1, key="n_da")
            
        selected_shift_label = st.selectbox("偏好班別：", list(SHIFT_OPTIONS.keys()))
        shift_code = SHIFT_OPTIONS[selected_shift_label]
        
        req_reason = st.text_input("備註原因 (選填)：", "")
        
        if st.button("提交申請 (Submit Request)"):
            new_entry = {
                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                "name": selected_user,
                "year": req_year,
                "month": req_month,
                "day": int(req_day),
                "shift": shift_code,
                "shift_label": selected_shift_label,
                "reason": req_reason,
                "status": "待審核"
            }
            st.session_state.requests_db.append(new_entry)
            st.success(f"✅ 已成功登記：{req_year}年{req_month}月{req_day}日 — {selected_shift_label}！")
            
        st.markdown("---")
        st.subheader("您已提交的申請記錄：")
        my_reqs = [r for r in st.session_state.requests_db if r["name"] == selected_user]
        if my_reqs:
            df_my = pd.DataFrame(my_reqs)
            st.dataframe(df_my[["year", "month", "day", "shift_label", "reason", "status"]], use_container_width=True)
            if st.button("撤回我的所有申請"):
                st.session_state.requests_db = [r for r in st.session_state.requests_db if r["name"] != selected_user]
                st.rerun()
        else:
            st.write("目前尚無申請記錄。")

    with tab2:
        st.subheader("最新排班表預覽")
        if st.session_state.generated_schedule is not None:
            st.dataframe(st.session_state.generated_schedule, use_container_width=True)
        else:
            st.info("管理員尚未發布最新月份的排班表。")

# ==========================================
# 介面分支 B：管理員 Admin (WM) 操作頁面 (Page 2 - 4)
# ==========================================
elif selected_user == "Admin (Ward Manager)":
    st.header("⚙️ 護士長管理後台 (Ward Manager Workstation) — Revision 3")
    
    admin_pin = st.sidebar.text_input("請輸入管理員密碼：", type="password")
    if admin_pin != "og2026":
        st.warning("請在左側輸入管理員密碼（預設：`og2026`）以解鎖管理權限。")
    else:
        st.success("✅ 管理員權限已驗證")
        
        col1, col2 = st.columns(2)
        with col1:
            plan_year = st.selectbox("排班年份：", (2026, 2027), index=0)
        with col2:
            plan_month = st.selectbox("排班月份：", list(range(1, 13)), index=9)
            
        st.subheader("📋 步驟一：20 號前審查各護士之特別申請 (Review & Settle Crashes)")
        curr_reqs = [r for r in st.session_state.requests_db if r["year"] == plan_year and r["month"] == plan_month]
        
        approved_reqs = []
        if curr_reqs:
            st.markdown("請勾選批准或駁回申請（若同日同更人數超額，建議依申請時間優先批准）：")
            for idx, r in enumerate(curr_reqs):
                col_a, col_b = st.columns((3, 1))
                with col_a:
                    shift_disp = r.get("shift_label", r["shift"])
                    st.write(f"**{r['name']}** 申請 **{r['month']}/{r['day']}**：{shift_disp} (原因: {r['reason']} | 提交時間: {r['timestamp']})")
                with col_b:
                    is_app = st.checkbox("批准", value=True, key=f"app_{idx}")
                    if is_app:
                        approved_reqs.append(r)
        else:
            st.info("該月份目前暫無護士登記特別申請（全員按 Revision 3 人性化標準規則自動排班）。")
            
        st.markdown("---")
        st.subheader("⚡ 步驟二：自動生成全月排班表 (Generate Duty Schedule)")
        st.caption("✨ Revision 3 人性化引擎：自動最大化連休時數、聚攏雙連休 O-O、極力消除 O-P-O-P 碎假、成對落實 A-N-O 閉環、優先引導 P-A-N-O 黃金班型，並均攤全體夜班負擔。")
        
        if st.button("🚀 開始自動排班運算 (Run OR-Tools Solver)"):
            with st.spinner("求解器正在運算四級資歷、P-A-N-O黃金更次、O-O連休聚攏與夜更公平輪替，請稍候約 20 秒..."):
                _, num_days = calendar.monthrange(plan_year, plan_month)
                first_weekday = datetime.date(plan_year, plan_month, 1).weekday()
                days = list(range(num_days))
                shifts = ['O', 'A', 'P', 'N', 'Day']
                roles = ['None', 'OTIC', 'Scrub', 'Recovery', 'Runner', 'Room']
                
                model = cp_model.CpModel()
                x = {}
                for n in DEFAULT_NURSES:
                    for d in days:
                        for s in shifts:
                            for r in roles:
                                x[n['id'], d, s, r] = model.NewBoolVar(f"x_{n['id']}_{d}_{s}_{r}")
                                
                def shift_assigned(n_id, d, s):
                    return sum(x[n_id, d, s, r] for r in roles if r != 'None')

                # 1. 每日每人唯一性
                for n in DEFAULT_NURSES:
                    for d in days:
                        model.Add(sum(x[n['id'], d, s, r] for s in shifts for r in roles) == 1)
                        for r in roles:
                            if r != 'None':
                                model.Add(x[n['id'], d, 'O', r] == 0)
                        for s in shifts:
                            if s != 'O':
                                model.Add(x[n['id'], d, s, 'None'] == 0)
                                
                # 2. 四級能力階梯
                for n in DEFAULT_NURSES:
                    cap = n['cap']
                    for d in days:
                        for s in shifts:
                            if cap == 'C(Runner)':
                                model.Add(x[n['id'], d, s, 'OTIC'] == 0)
                                model.Add(x[n['id'], d, s, 'Scrub'] == 0)
                                model.Add(x[n['id'], d, s, 'Recovery'] == 0)
                            elif cap == 'C(Recovery)':
                                model.Add(x[n['id'], d, s, 'OTIC'] == 0)
                                model.Add(x[n['id'], d, s, 'Scrub'] == 0)
                            elif cap == 'C(Scrub)':
                                model.Add(x[n['id'], d, s, 'OTIC'] == 0)

                # 3. 每 7 天 2 日 Off
                off_penalties = []
                for n in DEFAULT_NURSES:
                    for start in range(0, num_days - 6, 7):
                        window = range(start, start + 7)
                        w_offs = sum(x[n['id'], d, 'O', 'None'] for d in window)
                        diff = model.NewIntVar(-7, 7, f"d_{n['id']}_{start}")
                        model.Add(diff == w_offs - 2)
                        ad = model.NewIntVar(0, 7, f"ad_{n['id']}_{start}")
                        model.AddAbsEquality(ad, diff)
                        off_penalties.append(ad)

                # 4. 交接連續性 (一P接二A, 二P接三A, 四P接五A)
                for d in days[:-1]:
                    w = (first_weekday + d) % 7
                    if w in HANDOVER_DAYS:
                        for n in DEFAULT_NURSES:
                            model.Add(x[n['id'], d, 'P', 'OTIC'] == x[n['id'], d + 1, 'A', 'OTIC'])

                # 5. 夜班過渡管制
                for n in DEFAULT_NURSES:
                    for d in days:
                        if d > 0:
                            model.Add(shift_assigned(n['id'], d, 'N') <= shift_assigned(n['id'], d - 1, 'A'))
                            model.Add(shift_assigned(n['id'], d - 1, 'N') + shift_assigned(n['id'], d, 'A') <= 1)
                            model.Add(shift_assigned(n['id'], d - 1, 'N') + shift_assigned(n['id'], d, 'Day') <= 1)
                            model.Add(shift_assigned(n['id'], d - 1, 'N') + shift_assigned(n['id'], d, 'N') <= 1)
                        if d > 1:
                            model.Add(shift_assigned(n['id'], d - 2, 'P') + x[n['id'], d - 1, 'O', 'None'] + shift_assigned(n['id'], d, 'N') <= 2)

                # 6. 人手配置
                for d in days:
                    w = (first_weekday + d) % 7
                    # 夜更 (N更) 全組 6 人 (1 APN LWIC + 1 OTIC + 1 Scrub + 2 Room護士)
                    model.Add(sum(x[n['id'], d, 'N', 'OTIC'] for n in DEFAULT_NURSES) == 1)
                    model.Add(sum(x[n['id'], d, 'N', 'Scrub'] for n in DEFAULT_NURSES) == 1)
                    model.Add(sum(x[n['id'], d, 'N', 'Room'] for n in DEFAULT_NURSES) >= 2)
                    
                    if w in SPECIAL_DAYS: # 二、三、五
                        model.Add(sum(x[n['id'], d, 'A', 'OTIC'] for n in DEFAULT_NURSES) == 1)
                        model.Add(sum(x[n['id'], d, 'A', 'Scrub'] for n in DEFAULT_NURSES) == 1)
                        model.Add(sum(x[n['id'], d, 'A', 'Runner'] for n in DEFAULT_NURSES) == 1)
                        model.Add(sum(x[n['id'], d, 'A', 'Room'] for n in DEFAULT_NURSES) >= 4)
                        model.Add(sum(x[n['id'], d, 'A', 'Room'] for n in DEFAULT_NURSES) <= 5)
                        
                        model.Add(sum(x[n['id'], d, 'Day', 'Recovery'] for n in DEFAULT_NURSES) == 1)
                        
                        model.Add(sum(x[n['id'], d, 'P', 'OTIC'] for n in DEFAULT_NURSES) == 1)
                        model.Add(sum(x[n['id'], d, 'P', 'Scrub'] for n in DEFAULT_NURSES) == 1)
                        model.Add(sum(x[n['id'], d, 'P', 'Room'] for n in DEFAULT_NURSES) >= 4)
                        model.Add(sum(x[n['id'], d, 'P', 'Room'] for n in DEFAULT_NURSES) <= 5)
                    else: # 一、四、六、日
                        model.Add(sum(shift_assigned(n['id'], d, 'Day') for n in DEFAULT_NURSES) == 0)
                        for s in ['A', 'P']:
                            model.Add(sum(x[n['id'], d, s, 'OTIC'] for n in DEFAULT_NURSES) == 1)
                            model.Add(sum(x[n['id'], d, s, 'Scrub'] for n in DEFAULT_NURSES) == 1)
                            model.Add(sum(x[n['id'], d, s, 'Room'] for n in DEFAULT_NURSES) >= 4)
                            model.Add(sum(x[n['id'], d, s, 'Room'] for n in DEFAULT_NURSES) <= 5)

                # ==========================================
                # 7. Revision 3 人性化排班模型增強
                # ==========================================

                # (A) 夜班均攤公平性 (每位護士月夜班次數在 3 至 6 班之間)
                night_fairness_penalties = []
                for n in DEFAULT_NURSES:
                    total_nights = sum(shift_assigned(n['id'], d, 'N') for d in days)
                    model.Add(total_nights >= 3)
                    model.Add(total_nights <= 6)
                    ndiff = model.NewIntVar(-3, 3, f"ndiff_{n['id']}")
                    model.Add(ndiff == total_nights - 4)
                    abs_ndiff = model.NewIntVar(0, 3, f"abs_ndiff_{n['id']}")
                    model.AddAbsEquality(abs_ndiff, ndiff)
                    night_fairness_penalties.append(abs_ndiff)

                # (B) 聚攏雙連休 (O - O) 獎勵
                oo_rewards = []
                for n in DEFAULT_NURSES:
                    for d in days[:-1]:
                        is_oo = model.NewBoolVar(f"oo_{n['id']}_{d}")
                        model.Add(is_oo <= x[n['id'], d, 'O', 'None'])
                        model.Add(is_oo <= x[n['id'], d + 1, 'O', 'None'])
                        oo_rewards.append(is_oo)

                # (C) 嚴懲碎片化單日交替休假 (防止 O-P-O-P / O-A-O-A 等碎假)
                isolated_off_penalties = []
                for n in DEFAULT_NURSES:
                    for d in range(1, num_days - 1):
                        is_iso = model.NewBoolVar(f"iso_{n['id']}_{d}")
                        model.Add(is_iso >= x[n['id'], d, 'O', 'None'] + 
                                            (1 - x[n['id'], d - 1, 'O', 'None']) + 
                                            (1 - x[n['id'], d + 1, 'O', 'None']) - 2)
                        isolated_off_penalties.append(is_iso)

                # (D) 成對落實 A - N - O 閉環獎勵
                ano_rewards = []
                for n in DEFAULT_NURSES:
                    for d in days[:-1]:
                        is_ano = model.NewBoolVar(f"ano_{n['id']}_{d}")
                        model.Add(is_ano <= shift_assigned(n['id'], d, 'N'))
                        model.Add(is_ano <= x[n['id'], d + 1, 'O', 'None'])
                        ano_rewards.append(is_ano)

                # (E) P - A - N - O 極佳黃金組合加分
                pano_rewards = []
                for n in DEFAULT_NURSES:
                    for d in range(2, num_days - 1):
                        is_pano = model.NewBoolVar(f"pano_{n['id']}_{d}")
                        model.Add(is_pano <= shift_assigned(n['id'], d - 2, 'P'))
                        model.Add(is_pano <= shift_assigned(n['id'], d, 'N'))
                        model.Add(is_pano <= x[n['id'], d + 1, 'O', 'None'])
                        pano_rewards.append(is_pano)

                # 8. 套用 WM 批准之申請
                for r in approved_reqs:
                    n_match = next((n for n in DEFAULT_NURSES if n["name"] == r["name"]), None)
                    if n_match:
                        d_idx = int(r["day"]) - 1
                        s_target = r["shift"]
                        if s_target == "O":
                            model.Add(x[n_match['id'], d_idx, 'O', 'None'] == 1)
                        else:
                            model.Add(shift_assigned(n_match['id'], d_idx, s_target) == 1)

                # 目標函數綜合評分
                scores = [
                    -50 * sum(off_penalties),
                    -25 * sum(night_fairness_penalties),
                    -35 * sum(isolated_off_penalties),
                    30 * sum(oo_rewards),
                    40 * sum(ano_rewards),
                    50 * sum(pano_rewards)
                ]
                for n in DEFAULT_NURSES:
                    for d in days:
                        w = (first_weekday + d) % 7
                        if w in WEEKEND_DAYS:
                            scores.append(10 * x[n['id'], d, 'O', 'None'])
                model.Maximize(sum(scores))

                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 25.0
                status = solver.Solve(model)

                if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                    st.success("🎉 全月排班成功！已成功套用 Revision 3 人性化連休與 P-A-N-O 黃金模式。")
                    
                    # 建立格式化 Excel 檔案
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = f"{plan_year}-{plan_month:02d} 班表"
                    
                    shift_fills = {
                        'A': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),   # 淺綠
                        'P': PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),   # 暖橙
                        'N': PatternFill(start_color="E1D5E7", end_color="E1D5E7", fill_type="solid"),   # 粉紫
                        'Day': PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"), # 柔藍
                        'O': PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),   # 淺灰
                    }
                    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                    weekend_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                    border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                    
                    ws.cell(row=1, column=1, value=f"O&G 手術室及病房護士排班表 ({plan_year}年{plan_month}月 - Revision 3)").font = Font(name="Arial", size=14, bold=True, color="1F497D")
                    weekday_cn = ['一', '二', '三', '四', '五', '六', '日']
                    headers = ["編號", "資歷能力", "護士姓名"] + [f"{plan_month}/{d+1}\n({weekday_cn[(first_weekday+d)%7]})" for d in days] + ["總上班日", "總放假日"]
                    
                    header_row_index = 3
                    ws.row_dimensions[header_row_index].height = 28
                    for col_idx, h in enumerate(headers, 1):
                        cell = ws.cell(row=3, column=col_idx, value=h)
                        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        is_wk = (4 <= col_idx <= len(days) + 3) and ((first_weekday + col_idx - 4) % 7 in WEEKEND_DAYS)
                        cell.fill = weekend_fill if is_wk else header_fill
                        
                    res_rows = []
                    for r_idx, n in enumerate(DEFAULT_NURSES, 4):
                        ws.row_dimensions[r_idx].height = 22
                        ws.cell(row=r_idx, column=1, value=n['id']).font = Font(name="Arial", size=9, bold=True)
                        ws.cell(row=r_idx, column=1).alignment = Alignment(horizontal='center', vertical='center')
                        ws.cell(row=r_idx, column=1).border = border
                        
                        ws.cell(row=r_idx, column=2, value=n['cap']).font = Font(name="Arial", size=9, bold=True)
                        ws.cell(row=r_idx, column=2).alignment = Alignment(horizontal='center', vertical='center')
                        ws.cell(row=r_idx, column=2).border = border
                        
                        ws.cell(row=r_idx, column=3, value=n['name']).font = Font(name="Arial", size=9)
                        ws.cell(row=r_idx, column=3).alignment = Alignment(horizontal='left', vertical='center')
                        ws.cell(row=r_idx, column=3).border = border
                        
                        row_dict = {"編號": n["id"], "資格": n["cap"], "姓名": n["name"]}
                        work_c, off_c = 0, 0
                        for d in days:
                            cell = ws.cell(row=r_idx, column=d + 4)
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                            cell.font = Font(name="Arial", size=9)
                            cell.border = border
                            
                            s_assigned, r_assigned = 'O', 'None'
                            for s in shifts:
                                for r in roles:
                                    if solver.Value(x[n['id'], d, s, r]) == 1:
                                        s_assigned, r_assigned = s, r
                                        break
                            if s_assigned == 'O':
                                cell.value, cell.fill = "O", shift_fills['O']
                                row_dict[f"{plan_month}/{d+1}"] = "O"
                                off_c += 1
                            else:
                                role_tag = {'OTIC':'(OTIC)', 'Scrub':'(Scrub)', 'Recovery':'(Recovery)', 'Runner':'(Runner)', 'Room':'(Room)'}.get(r_assigned, '')
                                disp_val = f"{s_assigned}\n{role_tag}" if role_tag else s_assigned
                                cell.value, cell.fill = disp_val, shift_fills.get(s_assigned, shift_fills['O'])
                                row_dict[f"{plan_month}/{d+1}"] = f"{s_assigned}{role_tag}"
                                work_c += 1
                                
                        ws.cell(row=r_idx, column=len(days)+4, value=work_c).alignment = Alignment(horizontal='center', vertical='center')
                        ws.cell(row=r_idx, column=len(days)+4).font = Font(name="Arial", size=9, bold=True)
                        ws.cell(row=r_idx, column=len(days)+4).border = border
                        
                        ws.cell(row=r_idx, column=len(days)+5, value=off_c).alignment = Alignment(horizontal='center', vertical='center')
                        ws.cell(row=r_idx, column=len(days)+5).font = Font(name="Arial", size=9, bold=True)
                        ws.cell(row=r_idx, column=len(days)+5).border = border
                        res_rows.append(row_dict)

                    ws.column_dimensions['A'].width = 6
                    ws.column_dimensions['B'].width = 12
                    ws.column_dimensions['C'].width = 22
                    for d in days:
                        ws.column_dimensions[get_column_letter(d + 4)].width = 8
                    ws.column_dimensions[get_column_letter(len(days) + 4)].width = 10
                    ws.column_dimensions[get_column_letter(len(days) + 5)].width = 10
                    
                    df_out = pd.DataFrame(res_rows)
                    st.session_state.generated_schedule = df_out
                    st.dataframe(df_out, use_container_width=True)
                    
                    output = io.BytesIO()
                    wb.save(output)
                    
                    st.download_button(
                        label="📥 下載格式化 Excel 班表 (Download Excel)",
                        data=output.getvalue(),
                        file_name=f"OG_Duty_Schedule_{plan_year}_{plan_month:02d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 運算未能找到可行解，請檢查是否有過多護士請假衝突。")
