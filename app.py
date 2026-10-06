import streamlit as st
import datetime
import calendar
import os
import json
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from ortools.sat.python import cp_model
import io

# 頁面配置
st.set_page_config(page_title="O&G 產房手術室護士排班系統", layout="wide", page_icon="🏥")

# 常數定義 (0=Mon, ..., 6=Sun)
MON, TUE, WED, THU, FRI, SAT, SUN = 0, 1, 2, 3, 4, 5, 6
HANDOVER_DAYS = (MON, TUE, THU)
SPECIAL_DAYS = (TUE, WED, FRI)
WEEKEND_DAYS = (SAT, SUN)

# 持久化檔案路徑 (Persistent Storage Files)
REQUESTS_DB_FILE = "requests_db.csv"
SAVED_SCHEDULE_FILE = "saved_schedule.json"

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

# 護士偏好選項
NURSE_SHIFT_OPTIONS = {
    "想放假 Day Off (O)": "WANT_O",
    "想返早班 AM (A)": "WANT_A",
    "想返午班 PM (P)": "WANT_P",
    "想返夜班 Night (N)": "WANT_N",
    "想返日間常規班 (Day)": "WANT_DAY",
    "想返早班或午班 (A or P)": "WANT_AP",
    "不想返夜班 (No Night)": "AVOID_N",
    "不想返早班 (No AM)": "AVOID_A",
    "不想返午班 (No PM)": "AVOID_P",
    "不想返日間常規班 (No Day)": "AVOID_DAY"
}

# 管理員偏好選項 (包含連續長夜班 Long Night)
ADMIN_SHIFT_OPTIONS = {
    "指定連續長夜班 (Long Night)": "LONG_NIGHT",
    "想放假 Day Off (O)": "WANT_O",
    "想返早班 AM (A)": "WANT_A",
    "想返午班 PM (P)": "WANT_P",
    "想返夜班 Night (N)": "WANT_N",
    "想返日間常規班 (Day)": "WANT_DAY",
    "想返早班或午班 (A or P)": "WANT_AP",
    "不想返夜班 (No Night)": "AVOID_N",
    "不想返早班 (No AM)": "AVOID_A",
    "不想返午班 (No PM)": "AVOID_P",
    "不想返日間常規班 (No Day)": "AVOID_DAY"
}

# ==========================================
# 試算表數據庫持久化讀寫函式 (Persistent Storage Handlers)
# ==========================================
def load_requests_from_db():
    if os.path.exists(REQUESTS_DB_FILE):
        try:
            df = pd.read_csv(REQUESTS_DB_FILE, encoding="utf-8-sig")
            df["reason"] = df["reason"].fillna("")
            df["status"] = df["status"].fillna("待審核")
            if "admin_created" not in df.columns:
                df["admin_created"] = False
            else:
                df["admin_created"] = df["admin_created"].fillna(False).astype(bool)
            if "end_day" not in df.columns:
                df["end_day"] = df["day"]
            else:
                df["end_day"] = df["end_day"].fillna(df["day"]).astype(int)
            df["day"] = df["day"].astype(int)
            df["month"] = df["month"].astype(int)
            df["year"] = df["year"].astype(int)
            return df.to_dict("records")
        except Exception as e:
            st.warning(f"讀取申請試算表時發生錯誤：{e}")
            return []
    return []

def save_requests_to_db(reqs):
    cols = ["timestamp", "name", "year", "month", "day", "end_day", "shift", "shift_label", "reason", "status", "admin_created"]
    if not reqs:
        df = pd.DataFrame(columns=cols)
    else:
        df = pd.DataFrame(reqs)
        for c in cols:
            if c not in df.columns:
                df[c] = False if c == "admin_created" else ("" if c == "reason" else df.get("day", 1))
        df = df[cols]
    df.to_csv(REQUESTS_DB_FILE, index=False, encoding="utf-8-sig")

def load_saved_schedule():
    if os.path.exists(SAVED_SCHEDULE_FILE):
        try:
            with open(SAVED_SCHEDULE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            df_sched = pd.DataFrame(data["df_out"])
            daily_details = {int(k): v for k, v in data["daily_details"].items()}
            meta = data["plan_meta"]
            return df_sched, daily_details, meta
        except Exception:
            return None, None, None
    return None, None, None

def save_schedule_to_db(df_out, daily_details, plan_meta):
    try:
        data = {
            "df_out": df_out.to_dict("records"),
            "daily_details": {str(k): v for k, v in daily_details.items()},
            "plan_meta": plan_meta
        }
        with open(SAVED_SCHEDULE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"儲存班表至持久化記憶體時發生錯誤：{e}")

# 初始化 Session 狀態 (優先自持久化試算表/檔案讀取)
if "requests_db" not in st.session_state:
    st.session_state.requests_db = load_requests_from_db()

if "generated_schedule" not in st.session_state or st.session_state.generated_schedule is None:
    loaded_df, loaded_dd, loaded_meta = load_saved_schedule()
    st.session_state.generated_schedule = loaded_df
    st.session_state.daily_details = loaded_dd
    st.session_state.plan_meta = loaded_meta

# ==========================================
# 介面頂部: 直接於主頁面選擇身份
# ==========================================
st.title("🏥 O&G 產房手術室護士排班系統")

nurse_names = [n["name"] for n in DEFAULT_NURSES]
user_options = ["請選擇您的身份...", "Admin (Ward Manager)"] + nurse_names

selected_user = st.selectbox("👤 請選擇使用者身份 (Select User)：", user_options, index=0)

# ==========================================
# 輔助函式：解析班別申請代碼 (Parse Shift Request)
# ==========================================
def parse_shift_request(shift_str):
    """
    解析班別申請代碼為 (intent, shift_list)
    支援新格式 (如 WANT:A/P, AVOID:N) 與舊格式 (如 WANT_A, WANT_AP, AVOID_N, LONG_NIGHT)
    """
    if not shift_str:
        return "WANT", []
    if shift_str == "LONG_NIGHT":
        return "LONG_NIGHT", ["N"]
    if ":" in shift_str:
        intent_raw, combo_raw = shift_str.split(":", 1)
        intent = intent_raw.strip().upper()
        combo = combo_raw.strip()
        shifts = [s.strip() for s in combo.split("/") if s.strip()]
        return intent, shifts
    
    legacy_map = {
        "WANT_O": ("WANT", ["O"]),
        "WANT_A": ("WANT", ["A"]),
        "WANT_P": ("WANT", ["P"]),
        "WANT_N": ("WANT", ["N"]),
        "WANT_DAY": ("WANT", ["Day"]),
        "WANT_AP": ("WANT", ["A", "P"]),
        "AVOID_N": ("AVOID", ["N"]),
        "AVOID_A": ("AVOID", ["A"]),
        "AVOID_P": ("AVOID", ["P"]),
        "AVOID_DAY": ("AVOID", ["Day"]),
    }
    if shift_str in legacy_map:
        return legacy_map[shift_str]
        
    if shift_str.startswith("WANT_"):
        return "WANT", [shift_str.replace("WANT_", "")]
    if shift_str.startswith("AVOID_"):
        return "AVOID", [shift_str.replace("AVOID_", "")]
    return "WANT", [shift_str]

def is_requesting_off(shift_code):
    intent, shifts = parse_shift_request(shift_code)
    return intent == "WANT" and "O" in shifts

# ==========================================
# 輔助函式：護士申請限制驗證 (每週最多2項、週末最多1個O)
# ==========================================
def validate_nurse_request(existing_reqs, nurse_name, year, month, day, shift_code):
    dt = datetime.date(year, month, day)
    w_cal = dt.isocalendar()
    w_key = (w_cal.year, w_cal.week)
    iso_weekday = dt.isoweekday()  # 1=Mon, ..., 6=Sat, 7=Sun
    
    nurse_reqs = [r for r in existing_reqs if r["name"] == nurse_name and r["year"] == year and r["month"] == month]
    
    # 計算該週 (週一至週日) 已有的申請數量 (扣除同日更替)
    same_week_reqs = [
        r for r in nurse_reqs 
        if (datetime.date(r["year"], r["month"], r["day"]).isocalendar().year, 
            datetime.date(r["year"], r["month"], r["day"]).isocalendar().week) == w_key 
        and r["day"] != day
    ]
    
    if len(same_week_reqs) >= 2:
        return False, f"每位護士每週（星期一至星期日）最多只可提交 2 項特別申請！該週您已有 {len(same_week_reqs)} 項登記。"
        
    # 週末只可申請 1 個 O (星期六及日不能同時要求包含放假 O)
    if is_requesting_off(shift_code) and iso_weekday in (6, 7):
        for r in same_week_reqs:
            r_dt = datetime.date(r["year"], r["month"], r["day"])
            if r_dt.isoweekday() in (6, 7) and is_requesting_off(r["shift"]):
                return False, "每週週末（星期六及日）最多只可申請 1 天例假 (Day Off)，不可同時申請星期六與星期日放假！"
                
    return True, ""

# ==========================================
# 介面分支 A：護士專用操作頁面
# ==========================================
if selected_user not in ("請選擇您的身份...", "Admin (Ward Manager)"):
    st.markdown(f"### 👋 您好，**{selected_user}**")
    
    tab1, tab2 = st.tabs(["📝 填寫特別申請 (Make Special Request)", "📅 查看已公佈班表 (See Duty Schedule)"])
    
    with tab1:
        st.subheader("15 號前登記下月特別更次 / 放假申請")
        st.info("📌 規則提示：每週（星期一至日）最多申請 2 項；週末（星期六及日）最多只可申請 1 天例假 (Day Off)。申請將自動儲存於雲端資料庫，重整頁面不會丟失。")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            req_year = st.selectbox("年份 (Year)", (2026, 2027), index=0, key="n_yr")
        with col2:
            current_month = datetime.date.today().month
            default_next_month = (current_month % 12) + 1
            req_month = st.selectbox("月份 (Month)", list(range(1, 13)), index=default_next_month - 1, key="n_mo")
        with col3:
            _, max_day = calendar.monthrange(req_year, req_month)
            req_day = st.selectbox("日期 (Date)", list(range(1, max_day + 1)), index=0, key="n_da")
            
        st.markdown("#### 選擇更次申請意向與班別組合 (Select Shift Preference)")
        req_intent = st.radio(
            "1️⃣ 選擇申請意向 (Select Intent)：",
            ("🟢 想返 / 想放 (Want - 必須排所選其中一個班別)", "🔴 不想返 / 避開 (Don't Want - 絕不排所選班別)"),
            horizontal=True,
            key="n_intent"
        )
        
        st.markdown("**2️⃣ 選擇班別組合 (Select Shift Combination — 可自由組合 A, P, N, O, Day)：**")
        
        if "n_combo_widget" not in st.session_state:
            st.session_state["n_combo_widget"] = ["A"]

        st.caption("⚡ 常用快捷預設（點擊後直接套用，亦可於下方多選框自由勾選或增減）：")
        b1, b2, b3, b4, b5, b6, b7 = st.columns(7)
        with b1:
            if st.button("A 班", key="btn_p_a"):
                st.session_state["n_combo_widget"] = ["A"]
                st.rerun()
        with b2:
            if st.button("P 班", key="btn_p_p"):
                st.session_state["n_combo_widget"] = ["P"]
                st.rerun()
        with b3:
            if st.button("N 班", key="btn_p_n"):
                st.session_state["n_combo_widget"] = ["N"]
                st.rerun()
        with b4:
            if st.button("放假 O", key="btn_p_o"):
                st.session_state["n_combo_widget"] = ["O"]
                st.rerun()
        with b5:
            if st.button("A 或 P", key="btn_p_ap"):
                st.session_state["n_combo_widget"] = ["A", "P"]
                st.rerun()
        with b6:
            if st.button("A 或 O", key="btn_p_ao"):
                st.session_state["n_combo_widget"] = ["A", "O"]
                st.rerun()
        with b7:
            if st.button("A/P/O", key="btn_p_apo"):
                st.session_state["n_combo_widget"] = ["A", "P", "O"]
                st.rerun()

        shift_choices = ("A", "P", "N", "O", "Day")
        shift_desc = {
            "A": "A (早班 07:30-15:30)",
            "P": "P (午班 13:30-21:15)",
            "N": "N (夜班 21:00-07:45)",
            "O": "O (例假/放假)",
            "Day": "Day (日間常規班 09:00-17:00)"
        }
        
        selected_shifts = st.multiselect(
            "班別多選清單 (點選下拉加入或點 ✕ 移除班別)：",
            options=shift_choices,
            format_func=lambda s: shift_desc.get(s, s),
            key="n_combo_widget"
        )
        
        is_want = req_intent.startswith("🟢")
        if not selected_shifts:
            st.warning("⚠️ 請至少選擇一個班別！")
            shift_code = ""
            selected_shift_label = ""
        else:
            intent_type = "WANT" if is_want else "AVOID"
            combo_code = "/".join(selected_shifts)
            shift_code = f"{intent_type}:{combo_code}"
            
            if is_want:
                selected_shift_label = f"🟢 想返/放：{' 或 '.join(selected_shifts)}"
                st.success(f"📋 **申請意願確認**：{req_year}年{req_month}月{req_day}日【必須排入】**{' 或 '.join(selected_shifts)}** 其中之一（不排其他班別）。")
            else:
                selected_shift_label = f"🔴 不想返：{' 及 '.join(selected_shifts)}"
                st.error(f"📋 **申請意願確認**：{req_year}年{req_month}月{req_day}日【避開】**{' 及 '.join(selected_shifts)}**（當天絕不可排這些班別）。")
        
        req_reason = st.text_input("備註原因 (選填)：", "")
        
        if st.button("提交申請 (Submit Request)"):
            if not shift_code:
                st.error("❌ 登記失敗：請至少選擇一個班別！")
            else:
                is_valid, err_msg = validate_nurse_request(st.session_state.requests_db, selected_user, req_year, req_month, int(req_day), shift_code)
                if not is_valid:
                    st.error(f"❌ 登記失敗：{err_msg}")
                else:
                    # 移除同一天的舊申請 (如果存在)
                    st.session_state.requests_db = [
                        r for r in st.session_state.requests_db
                        if not (r["name"] == selected_user and r["year"] == req_year and r["month"] == req_month and r["day"] == int(req_day))
                    ]
                    new_entry = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "name": selected_user,
                        "year": req_year,
                        "month": req_month,
                        "day": int(req_day),
                        "end_day": int(req_day),
                        "shift": shift_code,
                        "shift_label": selected_shift_label,
                        "reason": req_reason,
                        "status": "待審核",
                        "admin_created": False
                    }
                    st.session_state.requests_db.append(new_entry)
                    save_requests_to_db(st.session_state.requests_db)
                    st.success(f"✅ 已成功登記並儲存至試算表：{req_year}年{req_month}月{req_day}日 — {selected_shift_label}！")
                    st.rerun()
            
        st.markdown("---")
        st.subheader("您已提交的申請記錄：")
        my_reqs = [r for r in st.session_state.requests_db if r["name"] == selected_user]
        if my_reqs:
            my_reqs.sort(key=lambda r: (r["year"], r["month"], r["day"], r["timestamp"]))
            
            for idx, r in enumerate(my_reqs):
                col_r1, col_r2 = st.columns((5, 1))
                with col_r1:
                    reason_str = f" | 原因: {r['reason']}" if r.get('reason') else ""
                    st.write(f"• **{r['year']}年{r['month']}月{r['day']}日** — `{r['shift_label']}`{reason_str} （狀態: {r['status']}）")
                with col_r2:
                    if st.button("🗑️ 刪除", key=f"del_my_{idx}"):
                        st.session_state.requests_db.remove(r)
                        save_requests_to_db(st.session_state.requests_db)
                        st.rerun()
                        
            if st.button("撤回我的所有申請"):
                st.session_state.requests_db = [r for r in st.session_state.requests_db if r["name"] != selected_user]
                save_requests_to_db(st.session_state.requests_db)
                st.rerun()
        else:
            st.write("目前尚無申請記錄。")

    with tab2:
        st.subheader("最新排班表預覽")
        if st.session_state.generated_schedule is not None:
            # 支援篩選特定班別
            col_v1, col_v2 = st.columns((1, 2))
            with col_v1:
                nurse_view_mode = st.radio("班表顯示模式：", ["全部班別 (All Shifts)", "只顯示特定班別 (Filter by Shift)"], index=0, horizontal=True, key="n_vm")
            
            df_curr = st.session_state.generated_schedule.copy()
            if nurse_view_mode == "只顯示特定班別 (Filter by Shift)":
                with col_v2:
                    n_filt = st.selectbox("選擇要單獨檢視的班別：", ["A 更 (早班)", "P 更 (午班)", "N 更 (夜班)", "Day 更 (日間常規班)", "O 更 (例假/休假)"], key="n_flt_s")
                prefix_map = {"A 更 (早班)": "A", "P 更 (午班)": "P", "N 更 (夜班)": "N", "Day 更 (日間常規班)": "Day", "O 更 (例假/休假)": "O"}
                target_code = prefix_map[n_filt]
                day_cols = [c for c in df_curr.columns if "/" in c]
                for c in day_cols:
                    df_curr[c] = df_curr[c].apply(lambda v: v if (str(v).startswith(target_code) and (target_code != "A" or not str(v).startswith("AVOID"))) else "-")
            
            st.dataframe(df_curr, use_container_width=True)
            
            # 當日執勤人員詳細職責表
            if st.session_state.daily_details is not None and st.session_state.plan_meta is not None:
                st.markdown("---")
                st.subheader("📅 當日執勤人員與各崗位職責詳情 (Daily Role Breakdown)")
                p_yr = st.session_state.plan_meta["year"]
                p_mo = st.session_state.plan_meta["month"]
                _, p_days = calendar.monthrange(p_yr, p_mo)
                p_fwd = datetime.date(p_yr, p_mo, 1).weekday()
                weekday_cn = ['一', '二', '三', '四', '五', '六', '日']
                
                sel_day = st.selectbox("選擇查看日期：", list(range(1, p_days + 1)), index=0, 
                                       format_func=lambda d: f"{p_mo}月{d}日 (星期{weekday_cn[(p_fwd + d - 1) % 7]})", key="n_sel_day")
                
                day_list = st.session_state.daily_details.get(sel_day, [])
                if day_list:
                    shift_order_map = {'A': 1, 'Day': 2, 'P': 3, 'N': 4, 'O': 5}
                    role_order_map = {'OTIC': 1, 'Scrub': 2, 'Recovery': 3, 'Runner': 4, 'Room': 5, 'None': 6}
                    sorted_day = sorted(day_list, key=lambda x: (shift_order_map.get(x['shift'], 9), role_order_map.get(x['role'], 9), x['nurse_id']))
                    
                    shift_name_map = {
                        'A': 'A 更 (早班 07:00-15:00)',
                        'Day': 'Day 更 (日間常規 09:00-17:00)',
                        'P': 'P 更 (午班 13:00-21:00)',
                        'N': 'N 更 (夜班 21:00-07:00)',
                        'O': 'O 更 (例假 Day Off)'
                    }
                    role_label_map = {
                        'OTIC': 'OTIC (手術室主管 / In-Charge)',
                        'Scrub': 'Scrub (洗手護士)',
                        'Recovery': 'Recovery (復甦室護士)',
                        'Runner': 'Runner (巡迴護士)',
                        'Room': 'Room (手術室/產房護士)',
                        'None': 'Day Off (例假休息)'
                    }
                    hours_map = {
                        'A': '07:00 - 15:00',
                        'Day': '09:00 - 17:00',
                        'P': '13:00 - 21:00',
                        'N': '21:00 - 07:00 (+1)',
                        'O': '全日休假'
                    }
                    
                    table_rows = []
                    for itm in sorted_day:
                        table_rows.append({
                            "班別 (Shift)": shift_name_map.get(itm['shift'], itm['shift']),
                            "執勤時段 (Hours)": hours_map.get(itm['shift'], ''),
                            "崗位職責 (Role / Duty in Charge)": role_label_map.get(itm['role'], itm['role']),
                            "護士姓名 (Nurse Name)": itm['name'],
                            "資歷能力 (Capability)": itm['cap']
                        })
                    st.dataframe(pd.DataFrame(table_rows), use_container_width=True)
        else:
            st.info("管理員尚未發布最新月份的排班表。")

# ==========================================
# 介面分支 B：管理員 Admin (WM) 操作頁面
# ==========================================
elif selected_user == "Admin (Ward Manager)":
    st.markdown("### ⚙️ 護士長管理後台 (Ward Manager Workstation)")
    
    admin_pin = st.text_input("請輸入管理員密碼：", type="password")
    if admin_pin != "og2026":
        st.warning("請在上方輸入管理員密碼（預設：`og2026`）以解鎖管理權限。")
    else:
        st.success("✅ 管理員權限已驗證")
        
        col1, col2 = st.columns(2)
        with col1:
            plan_year = st.selectbox("排班年份：", (2026, 2027), index=0)
        with col2:
            plan_month = st.selectbox("排班月份：", list(range(1, 13)), index=9)
            
        _, max_day_admin = calendar.monthrange(plan_year, plan_month)

        # ----------------------------------------------------
        # 記憶體與試算表數據庫管理面板 (Persistent Database & Spreadsheet Management)
        # ----------------------------------------------------
        with st.expander("💾 申請試算表數據庫管理 (Spreadsheet Database & Memory)", expanded=False):
            st.markdown(f"本系統已啟用**持久化記憶功能**，所有護士及管理員申請均自動儲存於伺服器試算表檔案 (`{REQUESTS_DB_FILE}`) 中。")
            st.write(f"目前試算表中共有 **{len(st.session_state.requests_db)}** 筆已儲存申請記錄。")
            
            col_db1, col_db2, col_db3 = st.columns(3)
            with col_db1:
                # 導出當前申請試算表
                if st.session_state.requests_db:
                    df_export_db = pd.DataFrame(st.session_state.requests_db)
                    csv_export_db = df_export_db.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        label="📥 導出申請清單試算表 (CSV)",
                        data=csv_export_db,
                        file_name=f"Roster_Requests_Database_{datetime.date.today()}.csv",
                        mime="text/csv",
                        key="btn_dl_reqs_db"
                    )
            with col_db2:
                # 上傳試算表備份還原
                uploaded_db = st.file_uploader("📤 匯入申請試算表 (CSV)", type=["csv"], key="uploader_db_csv")
                if uploaded_db is not None:
                    try:
                        df_uploaded = pd.read_csv(uploaded_db, encoding="utf-8-sig")
                        st.session_state.requests_db = df_uploaded.to_dict("records")
                        save_requests_to_db(st.session_state.requests_db)
                        st.success(f"✅ 成功匯入並同步 {len(df_uploaded)} 筆申請！")
                        st.rerun()
                    except Exception as err:
                        st.error(f"匯入失敗：{err}")
            with col_db3:
                # 清空試算表 (危險操作，防手滑)
                if st.button("🗑️ 清空所有申請記錄", key="btn_clear_db"):
                    st.session_state.requests_db = []
                    save_requests_to_db([])
                    st.warning("⚠️ 試算表申請記錄已全數清空！")
                    st.rerun()

        # ----------------------------------------------------
        # 管理員手動指派功能 (支援單日各更次及連續長夜班 Long Night)
        # ----------------------------------------------------
        with st.expander("➕ 管理員手動指派護士更次 / 連續長夜班 (Admin Manual Request & Long Night)", expanded=False):
            st.markdown("管理員可在此直接為任何護士指定班別、自訂班別組合或長夜班安排：")
            col_ad1, col_ad2 = st.columns(2)
            with col_ad1:
                adm_target_nurse = st.selectbox("指定護士姓名：", nurse_names, key="adm_n_sel")
            with col_ad2:
                adm_req_type = st.radio("指派更次類型 (Assignment Type)：", ["🌙 指定連續長夜班 (Long Night)", "🟢 指定想返/想放組合 (Want)", "🔴 指定避開/不排組合 (Avoid)"], horizontal=True, key="adm_rtype_sel")
                
            if adm_req_type.startswith("🌙"):
                st.info("🌙 您選擇了【指定連續長夜班 (Long Night)】，請設定該護士連續值夜更的日期區間：")
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    adm_ln_start = st.selectbox("起始日期 (From Date)：", list(range(1, max_day_admin + 1)), index=0, key="adm_ln_st")
                with col_d2:
                    adm_ln_end = st.selectbox("結束日期 (To Date)：", list(range(adm_ln_start, max_day_admin + 1)), index=min(6, max_day_admin - adm_ln_start), key="adm_ln_ed")
                adm_day_val = adm_ln_start
                adm_end_val = adm_ln_end
                adm_shift_code = "LONG_NIGHT"
                label_disp = f"連續長夜班 Long Night ({plan_month}/{adm_ln_start} 至 {plan_month}/{adm_ln_end})"
            else:
                adm_day_val = st.selectbox("指定日期 (Date)：", list(range(1, max_day_admin + 1)), index=0, key="adm_single_d")
                adm_end_val = adm_day_val
                
                shift_choices = ("A", "P", "N", "O", "Day")
                shift_desc = {
                    "A": "A (早班)",
                    "P": "P (午班)",
                    "N": "N (夜班)",
                    "O": "O (例假/放假)",
                    "Day": "Day (日間常規班)"
                }
                
                is_adm_want = adm_req_type.startswith("🟢")
                adm_default = ["A"] if is_adm_want else ["N"]
                if "adm_combo_widget" not in st.session_state:
                    st.session_state["adm_combo_widget"] = adm_default
                    
                adm_shifts = st.multiselect(
                    "指定班別組合 (可複選 A, P, N, O, Day 任意組合)：",
                    options=shift_choices,
                    format_func=lambda s: shift_desc.get(s, s),
                    key="adm_combo_widget"
                )
                
                if not adm_shifts:
                    st.warning("⚠️ 請至少選擇一個班別！")
                    adm_shift_code = ""
                    label_disp = ""
                else:
                    intent_type = "WANT" if is_adm_want else "AVOID"
                    combo_str = "/".join(adm_shifts)
                    adm_shift_code = f"{intent_type}:{combo_str}"
                    if is_adm_want:
                        label_disp = f"🟢 管理員指定想返/放：{' 或 '.join(adm_shifts)}"
                    else:
                        label_disp = f"🔴 管理員指定避開：{' 及 '.join(adm_shifts)}"
                
            adm_reason = st.text_input("備註 / 指派原因 (選填)：", "管理員手動指定", key="adm_rsn")
            
            if st.button("確認加入指派清單 (Add Manual Assignment)", key="btn_adm_add"):
                if not adm_shift_code:
                    st.error("❌ 請先選擇有效的班別組合！")
                else:
                    new_adm_entry = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "name": adm_target_nurse,
                        "year": plan_year,
                        "month": plan_month,
                        "day": int(adm_day_val),
                        "end_day": int(adm_end_val),
                        "shift": adm_shift_code,
                        "shift_label": label_disp,
                        "reason": adm_reason,
                        "status": "已批准 (管理員指定)",
                        "admin_created": True
                    }
                    st.session_state.requests_db.append(new_adm_entry)
                    save_requests_to_db(st.session_state.requests_db)
                    st.success(f"✅ 已成功加入並儲存管理員指派：{adm_target_nurse} — {label_disp}")
                    st.rerun()

        # ----------------------------------------------------
        st.subheader("📋 步驟一：審查申請與衝突調解 (Review & Settle Crashes)")
        curr_reqs = [r for r in st.session_state.requests_db if r["year"] == plan_year and r["month"] == plan_month]
        
        approved_reqs = []
        if curr_reqs:
            curr_reqs.sort(key=lambda r: (int(r["day"]), r["name"], r["shift"], r.get("reason", ""), r["timestamp"]))
            
            st.markdown("依**日期**排列之申請與指定清單（請勾選批准或駁回）：")
            for idx, r in enumerate(curr_reqs):
                col_a, col_b, col_c = st.columns((4, 1, 1))
                with col_a:
                    shift_disp = r.get("shift_label", r["shift"])
                    reason_disp = f" | 原因: {r['reason']}" if r.get('reason') else ""
                    admin_tag = " `[管理員指派]`" if r.get("admin_created") else ""
                    if r.get("shift") == "LONG_NIGHT":
                        st.write(f"📅 **{r['month']}/{r['day']}日 至 {r['month']}/{r.get('end_day', r['day'])}日** — **{r['name']}** ：`{shift_disp}`{admin_tag}{reason_disp}")
                    else:
                        st.write(f"📅 **{r['month']}/{r['day']}日** — **{r['name']}** ：`{shift_disp}`{admin_tag}{reason_disp}")
                with col_b:
                    is_app = st.checkbox("批准", value=True, key=f"app_{idx}")
                    if is_app:
                        approved_reqs.append(r)
                with col_c:
                    if st.button("刪除", key=f"adm_del_{idx}"):
                        st.session_state.requests_db.remove(r)
                        save_requests_to_db(st.session_state.requests_db)
                        st.rerun()
        else:
            st.info("該月份目前暫無任何護士登記申請或管理員指定更次。")
            
        st.markdown("---")
        # ----------------------------------------------------
        # 步驟二：生成全月排班表
        # ----------------------------------------------------
        st.subheader("⚡ 步驟二：生成全月排班表 (Generate Duty Schedule)")
        
        if st.button("🚀 生成排班表 (Generate Schedule)"):
            with st.spinner("正在生成排班表，請稍候... (Generating duty schedule...)"):
                _, num_days = calendar.monthrange(plan_year, plan_month)
                first_weekday = datetime.date(plan_year, plan_month, 1).weekday()
                days = list(range(num_days))
                shifts = ['O', 'A', 'P', 'N', 'Day']
                roles = ['None', 'OTIC', 'Scrub', 'Recovery', 'Runner', 'Room']
                
                # 解析所有指定連續長夜班 (Long Night)
                long_nights = {}
                for r in approved_reqs:
                    if r.get("shift") == "LONG_NIGHT":
                        n_match = next((n for n in DEFAULT_NURSES if n["name"] == r["name"]), None)
                        if n_match:
                            s_d = int(r["day"]) - 1
                            e_d = int(r.get("end_day", r["day"])) - 1
                            long_nights.setdefault(n_match["id"], set()).update(range(s_d, e_d + 1))

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

                # 3. 指定連續長夜班 (Long Night) 硬性鎖定
                for n_id, ln_days in long_nights.items():
                    for d in ln_days:
                        model.Add(shift_assigned(n_id, d, 'N') == 1)

                # 4. 每週工時約 44 小時與每週 2 日 Off，每週盡量最多 1 次 Night
                off_penalties = []
                weekly_hour_penalties = []
                night_per_week_penalties = []
                night_fairness_penalties = []

                # 全月休假均衡 (非長夜護士約 8-9 天例假，上班 22-23 天)
                min_offs = int(num_days * 2 / 7)
                max_offs = min_offs + (1 if (num_days * 2) % 7 != 0 else 0)
                for n in DEFAULT_NURSES:
                    ln_days = long_nights.get(n['id'], set())
                    tot_o = sum(x[n['id'], d, 'O', 'None'] for d in days)
                    total_nights = sum(shift_assigned(n['id'], d, 'N') for d in days)
                    
                    if not ln_days:
                        model.Add(tot_o >= min_offs)
                        model.Add(tot_o <= max_offs)
                        model.Add(total_nights >= 3)
                        model.Add(total_nights <= 5)
                        ndiff = model.NewIntVar(-1, 1, f"ndiff_{n['id']}")
                        model.Add(ndiff == total_nights - 4)
                        abs_ndiff = model.NewIntVar(0, 1, f"abs_ndiff_{n['id']}")
                        model.AddAbsEquality(abs_ndiff, ndiff)
                        night_fairness_penalties.append(abs_ndiff)
                    else:
                        model.Add(total_nights >= len(ln_days))
                        model.Add(tot_o >= 6)

                    for start in range(0, num_days - 6, 7):
                        window = range(start, start + 7)
                        w_offs = sum(x[n['id'], d, 'O', 'None'] for d in window)
                        diff_o = model.NewIntVar(-7, 7, f"d_o_{n['id']}_{start}")
                        model.Add(diff_o == w_offs - 2)
                        ad_o = model.NewIntVar(0, 7, f"ad_o_{n['id']}_{start}")
                        model.AddAbsEquality(ad_o, diff_o)
                        off_penalties.append(ad_o)
                        
                        w_hours = model.NewIntVar(0, 70, f"wh_{n['id']}_{start}")
                        model.Add(w_hours == sum(
                            8 * (shift_assigned(n['id'], d, 'A') + shift_assigned(n['id'], d, 'P') + shift_assigned(n['id'], d, 'Day')) +
                            10 * shift_assigned(n['id'], d, 'N')
                            for d in window
                        ))
                        h_diff = model.NewIntVar(-44, 44, f"hdiff_{n['id']}_{start}")
                        model.Add(h_diff == w_hours - 44)
                        abs_hdiff = model.NewIntVar(0, 44, f"abs_hdiff_{n['id']}_{start}")
                        model.AddAbsEquality(abs_hdiff, h_diff)
                        weekly_hour_penalties.append(abs_hdiff)
                        
                        if not ln_days:
                            w_nights = sum(shift_assigned(n['id'], d, 'N') for d in window)
                            extra_nights = model.NewIntVar(0, 7, f"extra_n_{n['id']}_{start}")
                            model.Add(extra_nights >= w_nights - 1)
                            night_per_week_penalties.append(extra_nights)

                # 5. 交接連續性 (一P接二A, 二P接三A, 四P接五A)
                for d in days[:-1]:
                    w = (first_weekday + d) % 7
                    if w in HANDOVER_DAYS:
                        for n in DEFAULT_NURSES:
                            model.Add(x[n['id'], d, 'P', 'OTIC'] == x[n['id'], d + 1, 'A', 'OTIC'])

                # 6. 夜班過渡管制 (O shall always be AN / 嚴格成對落實 A -> N -> O，連續長夜除外)
                for n in DEFAULT_NURSES:
                    ln_days = long_nights.get(n['id'], set())
                    for d in days:
                        if d > 0:
                            if d in ln_days and (d - 1) in ln_days:
                                pass
                            elif d in ln_days:
                                pass
                            else:
                                model.Add(shift_assigned(n['id'], d, 'N') <= shift_assigned(n['id'], d - 1, 'A'))
                            
                            if (d - 1) in ln_days and d in ln_days:
                                pass
                            else:
                                model.Add(shift_assigned(n['id'], d - 1, 'N') <= x[n['id'], d, 'O', 'None'])

                        if d > 1:
                            if d not in ln_days:
                                model.Add(shift_assigned(n['id'], d - 2, 'P') + x[n['id'], d - 1, 'O', 'None'] + shift_assigned(n['id'], d, 'N') <= 2)

                # 7. 人手配置
                for d in days:
                    w = (first_weekday + d) % 7
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

                # 8. 人性化偏好聚攏與連續順暢班別優化 (Ergonomic Continuous Shift Optimization)
                oo_rewards = []
                for n in DEFAULT_NURSES:
                    for d in days[:-1]:
                        is_oo = model.NewBoolVar(f"oo_{n['id']}_{d}")
                        model.Add(is_oo <= x[n['id'], d, 'O', 'None'])
                        model.Add(is_oo <= x[n['id'], d + 1, 'O', 'None'])
                        oo_rewards.append(is_oo)

                isolated_off_penalties = []
                for n in DEFAULT_NURSES:
                    for d in range(1, num_days - 1):
                        is_iso = model.NewBoolVar(f"iso_{n['id']}_{d}")
                        model.Add(is_iso >= x[n['id'], d, 'O', 'None'] 
                                            - x[n['id'], d - 1, 'O', 'None'] 
                                            - x[n['id'], d + 1, 'O', 'None'] 
                                            - shift_assigned(n['id'], d - 1, 'N'))
                        isolated_off_penalties.append(is_iso)

                # 連續順暢班別獎勵 (PA, PAN, PAO, PAPA 等) 及 避免連續相同班別 (AAA, PPP)
                pa_rewards = []
                pan_rewards = []
                pao_rewards = []
                papa_rewards = []
                aaa_penalties = []
                ppp_penalties = []

                for n in DEFAULT_NURSES:
                    nid = n['id']
                    for d in range(num_days - 1):
                        # PA 模式: Day d 為 P 且 Day d+1 為 A (高優先順序連續班次)
                        is_pa = model.NewBoolVar(f"pa_{nid}_{d}")
                        p_d = shift_assigned(nid, d, 'P')
                        a_d1 = shift_assigned(nid, d + 1, 'A')
                        model.Add(is_pa <= p_d)
                        model.Add(is_pa <= a_d1)
                        model.Add(is_pa >= p_d + a_d1 - 1)
                        pa_rewards.append(is_pa)
                        
                    for d in range(num_days - 2):
                        # PAN 模式: Day d 為 P, Day d+1 為 A, Day d+2 為 N (接續 N->O 自然循環)
                        is_pan = model.NewBoolVar(f"pan_{nid}_{d}")
                        p_d = shift_assigned(nid, d, 'P')
                        a_d1 = shift_assigned(nid, d + 1, 'A')
                        n_d2 = shift_assigned(nid, d + 2, 'N')
                        model.Add(is_pan <= p_d)
                        model.Add(is_pan <= a_d1)
                        model.Add(is_pan <= n_d2)
                        model.Add(is_pan >= p_d + a_d1 + n_d2 - 2)
                        pan_rewards.append(is_pan)
                        
                        # PAO 模式: Day d 為 P, Day d+1 為 A, Day d+2 為 O
                        is_pao = model.NewBoolVar(f"pao_{nid}_{d}")
                        o_d2 = x[nid, d + 2, 'O', 'None']
                        model.Add(is_pao <= p_d)
                        model.Add(is_pao <= a_d1)
                        model.Add(is_pao <= o_d2)
                        model.Add(is_pao >= p_d + a_d1 + o_d2 - 2)
                        pao_rewards.append(is_pao)
                        
                        # AAA 懲罰: 避免連續 3 天相同早班 AAA
                        is_aaa = model.NewBoolVar(f"aaa_{nid}_{d}")
                        a_d = shift_assigned(nid, d, 'A')
                        a_d2 = shift_assigned(nid, d + 2, 'A')
                        model.Add(is_aaa <= a_d)
                        model.Add(is_aaa <= a_d1)
                        model.Add(is_aaa <= a_d2)
                        model.Add(is_aaa >= a_d + a_d1 + a_d2 - 2)
                        aaa_penalties.append(is_aaa)
                        
                        # PPP 懲罰: 避免連續 3 天相同午班 PPP
                        is_ppp = model.NewBoolVar(f"ppp_{nid}_{d}")
                        p_d1 = shift_assigned(nid, d + 1, 'P')
                        p_d2 = shift_assigned(nid, d + 2, 'P')
                        model.Add(is_ppp <= p_d)
                        model.Add(is_ppp <= p_d1)
                        model.Add(is_ppp <= p_d2)
                        model.Add(is_ppp >= p_d + p_d1 + p_d2 - 2)
                        ppp_penalties.append(is_ppp)

                    for d in range(num_days - 3):
                        # PAPA 模式: Day d: P, d+1: A, d+2: P, d+3: A
                        is_papa = model.NewBoolVar(f"papa_{nid}_{d}")
                        p_d = shift_assigned(nid, d, 'P')
                        a_d1 = shift_assigned(nid, d + 1, 'A')
                        p_d2 = shift_assigned(nid, d + 2, 'P')
                        a_d3 = shift_assigned(nid, d + 3, 'A')
                        model.Add(is_papa <= p_d)
                        model.Add(is_papa <= a_d1)
                        model.Add(is_papa <= p_d2)
                        model.Add(is_papa <= a_d3)
                        model.Add(is_papa >= p_d + a_d1 + p_d2 + a_d3 - 3)
                        papa_rewards.append(is_papa)

                # 9. 套用已批准之單日申請 (支援任意 A, P, N, O, Day 組合之想返 / 避開)
                for r in approved_reqs:
                    s_code = r.get("shift", "")
                    if s_code != "LONG_NIGHT":
                        n_match = next((n for n in DEFAULT_NURSES if n["name"] == r["name"]), None)
                        if n_match:
                            d_idx = int(r["day"]) - 1
                            intent, req_shifts = parse_shift_request(s_code)
                            
                            chosen_vars = []
                            for s in req_shifts:
                                if s == 'O':
                                    chosen_vars.append(x[n_match['id'], d_idx, 'O', 'None'])
                                elif s in ('A', 'P', 'N', 'Day'):
                                    chosen_vars.append(shift_assigned(n_match['id'], d_idx, s))
                            
                            if chosen_vars:
                                if intent == "WANT":
                                    # 當天必須排在所選組合中的其中一個班別 (例如 A 或 P, 或 A 或 O)
                                    model.Add(sum(chosen_vars) == 1)
                                elif intent == "AVOID":
                                    # 當天絕不可排所選組合中的任何班別 (例如 避開 N)
                                    for v in chosen_vars:
                                        model.Add(v == 0)

                # 目標評分 (整合連續順暢班別獎勵與相同連班懲罰)
                scores = [
                    -50 * sum(off_penalties),
                    -10 * sum(weekly_hour_penalties),
                    -60 * sum(night_per_week_penalties),
                    -30 * sum(night_fairness_penalties),
                    -40 * sum(isolated_off_penalties),
                    30 * sum(oo_rewards),
                    25 * sum(pa_rewards),
                    50 * sum(pan_rewards),
                    35 * sum(pao_rewards),
                    40 * sum(papa_rewards),
                    -80 * sum(aaa_penalties),
                    -80 * sum(ppp_penalties)
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
                    st.success("✅ 班表已成功生成！ (Duty schedule successfully generated!)")
                    pa_cnt = sum(solver.Value(v) for v in pa_rewards)
                    pan_cnt = sum(solver.Value(v) for v in pan_rewards)
                    pao_cnt = sum(solver.Value(v) for v in pao_rewards)
                    papa_cnt = sum(solver.Value(v) for v in papa_rewards)
                    aaa_cnt = sum(solver.Value(v) for v in aaa_penalties)
                    ppp_cnt = sum(solver.Value(v) for v in ppp_penalties)
                    st.info(f"✨ **排班人體工學模式統計**：PA 順暢接更 **{pa_cnt}** 次 | PAN 優質接更 **{pan_cnt}** 次 | PAO 連更接假 **{pao_cnt}** 次 | PAPA 規律輪替 **{papa_cnt}** 次 | 連續相同班別 (AAA / PPP) 成功壓減至 **{aaa_cnt} / {ppp_cnt}** 次。")
                    
                    # 建立格式化 Excel 檔案
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = f"{plan_year}-{plan_month:02d} 班表"
                    
                    shift_fills = {
                        'A': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                        'P': PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
                        'N': PatternFill(start_color="E1D5E7", end_color="E1D5E7", fill_type="solid"),
                        'Day': PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"),
                        'O': PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),
                    }
                    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                    weekend_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                    summary_hdr_fill = PatternFill(start_color="274E13", end_color="274E13", fill_type="solid")
                    border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                                    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                    
                    ws.cell(row=1, column=1, value=f"O&G 手術室及病房護士排班表 ({plan_year}年{plan_month}月)").font = Font(name="Arial", size=14, bold=True, color="1F497D")
                    weekday_cn = ['一', '二', '三', '四', '五', '六', '日']
                    day_headers = [f"{plan_month}/{d+1}\n({weekday_cn[(first_weekday+d)%7]})" for d in days]
                    summary_headers = ["總上班日", "總 A 班", "總 P 班", "總 N 班", "總 Day 班", "總 O 班"]
                    headers = ["編號", "資歷能力", "護士姓名"] + day_headers + summary_headers
                    
                    header_row_index = 3
                    ws.row_dimensions[header_row_index].height = 28
                    for col_idx, h in enumerate(headers, 1):
                        cell = ws.cell(row=3, column=col_idx, value=h)
                        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        is_wk = (4 <= col_idx <= len(days) + 3) and ((first_weekday + col_idx - 4) % 7 in WEEKEND_DAYS)
                        is_summary = col_idx > len(days) + 3
                        if is_summary:
                            cell.fill = summary_hdr_fill
                        elif is_wk:
                            cell.fill = weekend_fill
                        else:
                            cell.fill = header_fill
                        
                    res_rows = []
                    daily_details_dict = {}
                    for d in days:
                        daily_details_dict[d + 1] = []

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
                        work_c, a_c, p_c, n_c, day_c, off_c = 0, 0, 0, 0, 0, 0
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
                            
                            daily_details_dict[d + 1].append({
                                "nurse_id": n["id"],
                                "name": n["name"],
                                "cap": n["cap"],
                                "shift": s_assigned,
                                "role": r_assigned
                            })

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
                                if s_assigned == 'A':
                                    a_c += 1
                                elif s_assigned == 'P':
                                    p_c += 1
                                elif s_assigned == 'N':
                                    n_c += 1
                                elif s_assigned == 'Day':
                                    day_c += 1
                                
                        row_dict["總上班日"] = work_c
                        row_dict["總 A 班"] = a_c
                        row_dict["總 P 班"] = p_c
                        row_dict["總 N 班"] = n_c
                        row_dict["總 Day 班"] = day_c
                        row_dict["總 O 班"] = off_c

                        summary_vals = [work_c, a_c, p_c, n_c, day_c, off_c]
                        for s_idx, val in enumerate(summary_vals):
                            col_num = len(days) + 4 + s_idx
                            c_cell = ws.cell(row=r_idx, column=col_num, value=val)
                            c_cell.alignment = Alignment(horizontal='center', vertical='center')
                            c_cell.font = Font(name="Arial", size=9, bold=True)
                            c_cell.border = border

                        res_rows.append(row_dict)

                    # 底部每日加總列
                    bot_row = len(DEFAULT_NURSES) + 4
                    ws.row_dimensions[bot_row].height = 24
                    ws.cell(row=bot_row, column=1, value="").border = border
                    ws.cell(row=bot_row, column=2, value="").border = border
                    lbl_cell = ws.cell(row=bot_row, column=3, value="每日當值人數總計")
                    lbl_cell.font = Font(name="Arial", size=9, bold=True, color="1F497D")
                    lbl_cell.alignment = Alignment(horizontal='center', vertical='center')
                    lbl_cell.border = border

                    for d in days:
                        col_ltr = get_column_letter(d + 4)
                        day_cell = ws.cell(row=bot_row, column=d + 4, value=f'=COUNTIF({col_ltr}4:{col_ltr}{bot_row-1}, "<>O")')
                        day_cell.font = Font(name="Arial", size=9, bold=True)
                        day_cell.alignment = Alignment(horizontal='center', vertical='center')
                        day_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                        day_cell.border = border

                    for s_idx in range(len(summary_headers)):
                        col_num = len(days) + 4 + s_idx
                        col_ltr = get_column_letter(col_num)
                        tot_cell = ws.cell(row=bot_row, column=col_num, value=f'=SUM({col_ltr}4:{col_ltr}{bot_row-1})')
                        tot_cell.font = Font(name="Arial", size=9, bold=True)
                        tot_cell.alignment = Alignment(horizontal='center', vertical='center')
                        tot_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                        tot_cell.border = border

                    ws.column_dimensions['A'].width = 6
                    ws.column_dimensions['B'].width = 12
                    ws.column_dimensions['C'].width = 22
                    for d in days:
                        ws.column_dimensions[get_column_letter(d + 4)].width = 8
                    ws.column_dimensions[get_column_letter(len(days) + 4)].width = 10
                    ws.column_dimensions[get_column_letter(len(days) + 5)].width = 9
                    ws.column_dimensions[get_column_letter(len(days) + 6)].width = 9
                    ws.column_dimensions[get_column_letter(len(days) + 7)].width = 9
                    ws.column_dimensions[get_column_letter(len(days) + 8)].width = 10
                    ws.column_dimensions[get_column_letter(len(days) + 9)].width = 9
                    
                    df_out = pd.DataFrame(res_rows)
                    st.session_state.generated_schedule = df_out
                    st.session_state.daily_details = daily_details_dict
                    st.session_state.plan_meta = {"year": plan_year, "month": plan_month, "first_weekday": first_weekday, "num_days": num_days}
                    
                    # 儲存 Excel 與持久化檔案
                    output = io.BytesIO()
                    wb.save(output)
                    st.session_state.excel_data = output.getvalue()
                    save_schedule_to_db(df_out, daily_details_dict, st.session_state.plan_meta)
                    
                else:
                    st.error("❌ 運算未能找到可行解，請檢查是否有過多護士請假或長夜安排衝突。")

        # ----------------------------------------------------
        # 班表展示區 (含特定班別篩選 & 第二張每日崗位職責詳情表)
        # ----------------------------------------------------
        if st.session_state.generated_schedule is not None:
            st.markdown("---")
            st.subheader("📊 總排班表檢視 (Monthly Schedule View)")
            
            # 篩選特定班別選項
            col_f1, col_f2 = st.columns((1, 2))
            with col_f1:
                adm_view_mode = st.radio("班表顯示模式 (Display Mode)：", ["全部班別 (All Shifts)", "只顯示特定班別 (Filter by Shift)"], index=0, horizontal=True, key="adm_vm")
            
            df_display_admin = st.session_state.generated_schedule.copy()
            if adm_view_mode == "只顯示特定班別 (Filter by Shift)":
                with col_f2:
                    adm_target_shift = st.selectbox("選擇要單獨檢視的班別：", ["A 更 (早班)", "P 更 (午班)", "N 更 (夜班)", "Day 更 (日間常規班)", "O 更 (例假/休假)"], key="adm_flt_s")
                prefix_map_adm = {"A 更 (早班)": "A", "P 更 (午班)": "P", "N 更 (夜班)": "N", "Day 更 (日間常規班)": "Day", "O 更 (例假/休假)": "O"}
                t_code = prefix_map_adm[adm_target_shift]
                day_cols_adm = [c for c in df_display_admin.columns if "/" in c]
                for c in day_cols_adm:
                    df_display_admin[c] = df_display_admin[c].apply(lambda v: v if (str(v).startswith(t_code) and (t_code != "A" or not str(v).startswith("AVOID"))) else "-")

            st.dataframe(df_display_admin, use_container_width=True)

            if "excel_data" in st.session_state:
                st.download_button(
                    label="📥 下載格式化 Excel 完整班表 (Download Excel)",
                    data=st.session_state.excel_data,
                    file_name=f"OG_Duty_Schedule_{plan_year}_{plan_month:02d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # ----------------------------------------------------
            # 第二張表格：選擇指定日期的每日執勤人員與各崗位職責詳情
            # ----------------------------------------------------
            st.markdown("---")
            st.subheader("📋 第二張表：指定日期執勤人員與各崗位職責詳情 (Daily Roster & Role In-Charge Details)")
            
            meta = st.session_state.plan_meta
            weekday_cn = ['一', '二', '三', '四', '五', '六', '日']
            
            selected_inspect_day = st.selectbox(
                "📅 請選擇欲檢視的日期 (Select Date to View Details)：",
                list(range(1, meta["num_days"] + 1)),
                index=0,
                format_func=lambda d: f"{meta['month']}月{d}日 (星期{weekday_cn[(meta['first_weekday'] + d - 1) % 7]})",
                key="adm_sel_inspect_day"
            )
            
            day_records = st.session_state.daily_details.get(selected_inspect_day, [])
            if day_records:
                shift_order_map = {'A': 1, 'Day': 2, 'P': 3, 'N': 4, 'O': 5}
                role_order_map = {'OTIC': 1, 'Scrub': 2, 'Recovery': 3, 'Runner': 4, 'Room': 5, 'None': 6}
                sorted_records = sorted(day_records, key=lambda x: (shift_order_map.get(x['shift'], 9), role_order_map.get(x['role'], 9), x['nurse_id']))
                
                shift_name_map = {
                    'A': 'A 更 (早班 07:00-15:00)',
                    'Day': 'Day 更 (日間常規 09:00-17:00)',
                    'P': 'P 更 (午班 13:00-21:00)',
                    'N': 'N 更 (夜班 21:00-07:00)',
                    'O': 'O 更 (例假 Day Off)'
                }
                role_label_map = {
                    'OTIC': '🌟 OTIC (產房手術室主管 / Duty In Charge)',
                    'Scrub': '🩺 Scrub (洗手手術護士)',
                    'Recovery': '🛏️ Recovery (復甦室監護護士)',
                    'Runner': '🏃 Runner (巡迴護士)',
                    'Room': '🚪 Room (手術室/產房護士)',
                    'None': '🏖️ Day Off (例假休息)'
                }
                hours_map = {
                    'A': '07:00 - 15:00',
                    'Day': '09:00 - 17:00',
                    'P': '13:00 - 21:00',
                    'N': '21:00 - 07:00 (+1)',
                    'O': '全日休假'
                }
                
                a_nurses = [r['name'] for r in day_records if r['shift'] == 'A']
                p_nurses = [r['name'] for r in day_records if r['shift'] == 'P']
                n_nurses = [r['name'] for r in day_records if r['shift'] == 'N']
                day_nurses = [r['name'] for r in day_records if r['shift'] == 'Day']
                o_nurses = [r['name'] for r in day_records if r['shift'] == 'O']
                
                kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                kpi1.metric("早班 (A 更)", f"{len(a_nurses)} 人")
                kpi2.metric("午班 (P 更)", f"{len(p_nurses)} 人")
                kpi3.metric("夜班 (N 更)", f"{len(n_nurses)} 人")
                kpi4.metric("日間常規 (Day)", f"{len(day_nurses)} 人")
                kpi5.metric("例假 (Off)", f"{len(o_nurses)} 人")
                
                inspect_table = []
                for itm in sorted_records:
                    inspect_table.append({
                        "班別 (Shift)": shift_name_map.get(itm['shift'], itm['shift']),
                        "執勤時段 (Hours)": hours_map.get(itm['shift'], ''),
                        "崗位職責 (Role / Duty in Charge)": role_label_map.get(itm['role'], itm['role']),
                        "護士姓名 (Nurse Name)": itm['name'],
                        "資歷能力 (Capability)": itm['cap']
                    })
                
                df_day_inspect = pd.DataFrame(inspect_table)
                st.dataframe(df_day_inspect, use_container_width=True)
                
                csv_day = df_day_inspect.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label=f"📥 下載 {meta['month']}月{selected_inspect_day}日 當日執勤人員詳情 (CSV)",
                    data=csv_day,
                    file_name=f"OG_Duty_Detail_{meta['year']}_{meta['month']:02d}_{selected_inspect_day:02d}.csv",
                    mime="text/csv",
                    key="btn_dl_day_csv"
                )

