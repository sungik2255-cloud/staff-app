import streamlit as st
import pandas as pd
import time
from datetime import date, datetime
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from supabase import create_client, Client

# ── 1. Page Config ────────────────────────────────────────────
st.set_page_config(page_title="Staff Leave Management", layout="wide")

# ── 2. Supabase 연결 ──────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    url  = st.secrets["SUPABASE_URL"]
    key  = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def read_table(table_name):
    try:
        sb = get_supabase()
        res = sb.table(table_name).select("*").execute()
        df = pd.DataFrame(res.data)
        return df
    except Exception as e:
        st.error(f"❌ 시트 '{table_name}' 접근 실패: {e}")
        return pd.DataFrame()

def upsert_table(table_name, df):
    try:
        sb = get_supabase()
        sb.table(table_name).delete().neq("id", 0).execute()
        if not df.empty:
            df2 = df.drop(columns=["id"], errors="ignore").copy()
            for col in df2.columns:
                df2[col] = df2[col].apply(lambda x: x.isoformat() if hasattr(x, 'isoformat') else x)
            data = df2.fillna("").to_dict(orient="records")
            sb.table(table_name).insert(data).execute()
        return True
    except Exception as e:
        st.error(f"❌ 저장 실패: {e}")
        return False

# ✅ employees 전용 삭제 함수 — 오직 이 함수로만 employees 삭제 가능
def delete_employees_by_name(names_to_delete):
    try:
        emp_data = load_employees()
        updated = emp_data[~emp_data["Name"].isin(names_to_delete)].reset_index(drop=True)
        return upsert_table("employees", updated)
    except Exception as e:
        st.error(f"❌ 직원 삭제 실패: {e}")
        return False

# ── 3. Leave Detail Modal ─────────────────────────────────────
@st.dialog("📋 Leave Detail")
def show_leave_modal(record):
    emp        = record.get("Employee", "")
    dt         = record.get("Date", "")
    vac        = float(record.get("Vacation_Used", 0))
    sick       = float(record.get("Sick_Used", 0))
    leave_type = "🏖️ Vacation" if vac > 0 else "🤒 Sick Leave"
    hours      = vac if vac > 0 else sick
    status     = record.get("Status", "")
    status_icon = "✅ Used" if status == "Used" else "📌 Plan"
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**👤 Employee**")
        st.markdown("**📅 Date**")
        st.markdown("**🗂️ Leave Type**")
        st.markdown("**⏱️ Hours**")
        st.markdown("**📊 Status**")
    with col2:
        st.markdown(emp)
        st.markdown(dt)
        st.markdown(leave_type)
        st.markdown(f"{hours}h")
        st.markdown(status_icon)
    st.divider()
    st.caption("📋 Copy용 텍스트 (Ctrl+A → Ctrl+C)")
    copy_text = f"Employee: {emp}\nDate: {dt}\nLeave Type: {leave_type.split(' ')[-1]}\nHours: {hours}h\nStatus: {status}"
    st.code(copy_text, language=None)

# ── 4. Login System ───────────────────────────────────────────
def check_login():
    params = st.query_params
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.username = None
    if not st.session_state.logged_in:
        if params.get("auth") == "ok" and params.get("role") and params.get("user"):
            st.session_state.logged_in = True
            st.session_state.role = params.get("role")
            st.session_state.username = params.get("user")
    if not st.session_state.logged_in:
        col1, col2, col3 = st.columns([1, 1.2, 1])
        with col2:
            st.markdown("<br><br>", unsafe_allow_html=True)
            st.markdown("<h2 style='text-align:center;'>🏢 Staff Leave Management</h2>", unsafe_allow_html=True)
            st.markdown("<h4 style='text-align:center; color:gray;'>Please log in to continue</h4>", unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            with st.container(border=True):
                username = st.text_input("👤 Username")
                password = st.text_input("🔒 Password", type="password")
                if st.button("Login", type="primary", use_container_width=True):
                    try:
                        admin_user  = st.secrets["admin_username"]
                        admin_pw    = st.secrets["admin_password"]
                        viewer_user = st.secrets["viewer_username"]
                        viewer_pw   = st.secrets["viewer_password"]
                    except:
                        admin_user, admin_pw   = "admin", "admin123"
                        viewer_user, viewer_pw = "viewer", "viewer123"
                    if username == admin_user and password == admin_pw:
                        st.session_state.logged_in = True
                        st.session_state.role = "admin"
                        st.session_state.username = username
                        st.query_params["auth"] = "ok"
                        st.query_params["role"] = "admin"
                        st.query_params["user"] = username
                        st.rerun()
                    elif username == viewer_user and password == viewer_pw:
                        st.session_state.logged_in = True
                        st.session_state.role = "viewer"
                        st.session_state.username = username
                        st.query_params["auth"] = "ok"
                        st.query_params["role"] = "viewer"
                        st.query_params["user"] = username
                        st.rerun()
                    else:
                        st.error("❌ Incorrect username or password.")
        st.stop()

def show_sidebar_user():
    role_label = "🔑 Admin" if st.session_state.role == "admin" else "👁️ Viewer"
    st.sidebar.markdown(f"**{role_label}** ({st.session_state.username})")
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.username = None
        st.query_params.clear()
        st.rerun()

def is_admin():
    return st.session_state.get("role") == "admin"

check_login()
show_sidebar_user()

# ── 5. Smart Rules Database ───────────────────────────────────
SICK_LEAVE_RULES = {
    "Manhattan":    {"text": "⚖️ **NYC Law**: 1 hour for every 30 hours worked (Up to 40-56h/yr).", "rate": 30, "max": 40},
    "Flushing":     {"text": "⚖️ **NYC Law**: 1 hour for every 30 hours worked (Up to 40-56h/yr).", "rate": 30, "max": 40},
    "Philadelphia": {"text": "⚖️ **Philadelphia Law**: 1 hour for every 40 hours worked (Up to 40h/yr).", "rate": 40, "max": 40},
    "Orlando":      {"text": "⚠️ **Florida Law**: No state-mandated paid sick leave.", "rate": None, "max": None},
}

# ── 6. Data Loading Functions ─────────────────────────────────
def load_locations():
    df = read_table("locations")
    if df.empty:
        return pd.DataFrame([{"company_name": "Amlotus", "city_name": "Manhattan"}])
    df = df.dropna(subset=["company_name", "city_name"]).reset_index(drop=True)
    return df

def load_employees():
    df = read_table("employees")
    if df.empty:
        return pd.DataFrame(columns=["Name","Email","Location","Type","Vacation_Limit","Sick_Rate","Sick_Max"])
    df = df.drop_duplicates()
    df["Email"] = df["Email"].astype(str).replace("nan", "")
    df = df.sort_values("Name", ascending=True).reset_index(drop=True)
    return df

def load_work_logs():
    df = read_table("work_log")
    if df.empty:
        return pd.DataFrame(columns=["Employee","Status","Start_Date","End_Date","Hours_Worked"])
    if "Status" not in df.columns: df["Status"] = "Employed"
    df["Start_Date"] = pd.to_datetime(df["Start_Date"], errors="coerce").dt.date
    df["End_Date"]   = pd.to_datetime(df["End_Date"],   errors="coerce").dt.date
    df["Hours_Worked"] = pd.to_numeric(df["Hours_Worked"], errors="coerce").fillna(0)
    return df.reset_index(drop=True)

def load_resigned():
    cols = ["Name","Resigned_Date","Location","Total_Worked","Retained_Vacation","Retained_Sick","Paid_Date","Paid_Amount","Email","Type","Vacation_Limit","Sick_Rate","Sick_Max"]
    df = read_table("resigned_employees")
    if df.empty:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = "" if any(x in c for x in ["Email","Type","Location","Name"]) else 0.0
    df["Email"] = df["Email"].astype(str).replace("nan", "")
    return df

def load_leave_usage():
    df = read_table("leave_usage")
    if df.empty:
        return pd.DataFrame(columns=["Employee","Date","Vacation_Used","Sick_Used","Note","Status"])
    if "Status" not in df.columns: df["Status"] = "Used"
    df["Vacation_Used"] = pd.to_numeric(df["Vacation_Used"], errors="coerce").fillna(0)
    df["Sick_Used"]     = pd.to_numeric(df["Sick_Used"],     errors="coerce").fillna(0)
    return df

# ── 7. Session Initializer ────────────────────────────────────
if "city_df"  not in st.session_state: st.session_state.city_df  = load_locations()
if "emp_df"   not in st.session_state: st.session_state.emp_df   = load_employees()
if "log_df"   not in st.session_state: st.session_state.log_df   = load_work_logs()
if "leave_df" not in st.session_state: st.session_state.leave_df = load_leave_usage()
if "leave_modal_record" not in st.session_state: st.session_state.leave_modal_record = None
# ✅ Add Employee 폼 초기화용 카운터
if "emp_form_key" not in st.session_state: st.session_state.emp_form_key = 0
# ✅ 2↔3페이지 company 연동
if "selected_company" not in st.session_state:
    cdf = st.session_state.city_df
    st.session_state.selected_company = f"{cdf.iloc[0]['company_name']} - {cdf.iloc[0]['city_name']}" if not cdf.empty else ""

# ── 8. Sidebar Menu ───────────────────────────────────────────
menu_options_admin  = ["1. Employee Setup", "2. Log Worked Hours", "3. Plan/Submit Leave", "4. Dashboard & Email"]
menu_options_viewer = ["2. Log Worked Hours", "3. Plan/Submit Leave", "4. Dashboard & Email"]
menu_options = menu_options_admin if is_admin() else menu_options_viewer
if not is_admin(): st.sidebar.info("👁️ View-only mode")

saved_menu = st.query_params.get("menu", menu_options[0])
if saved_menu not in menu_options: saved_menu = menu_options[0]
menu = st.sidebar.radio("Go to", menu_options, index=menu_options.index(saved_menu))
if st.query_params.get("menu") != menu: st.query_params["menu"] = menu

# ═══════════════════════════════════════════════════════════════
# 1. Employee Setup
# ═══════════════════════════════════════════════════════════════
if menu == "1. Employee Setup":
    if not is_admin():
        st.warning("⛔ Admin access required."); st.stop()
    st.title("🏢 Staff Leave Management System")
    col1, col2 = st.columns([1, 1.2], gap="large")
    with col1:
        st.subheader("📍 City & Company Setup")
        with st.container(border=True):
            nc, ny = st.text_input("Company Name"), st.text_input("City Name")
            if st.button("Add to List", use_container_width=True):
                if nc and ny:
                    new_row = pd.DataFrame([{"company_name": nc, "city_name": ny}])
                    updated = pd.concat([st.session_state.city_df, new_row], ignore_index=True)
                    if upsert_table("locations", updated):
                        st.session_state.city_df = load_locations(); st.success("✅ Saved!"); time.sleep(1); st.rerun()
        disp_c = st.session_state.city_df.copy(); disp_c.insert(0, "Select", False)
        ed_c = st.data_editor(disp_c, use_container_width=True, key="ced")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Save Changes", use_container_width=True):
                if upsert_table("locations", ed_c.drop(columns=["Select"]).reset_index(drop=True)):
                    st.session_state.city_df = load_locations(); st.success("✅ Saved!"); time.sleep(1); st.rerun()
        with c2:
            if st.button("🗑️ Delete Selected Company", use_container_width=True):
                if ed_c["Select"].any(): st.session_state.loc_del_conf = True
        if st.session_state.get("loc_del_conf", False):
            st.error("⚠️ Delete selected company?"); cy, cn = st.columns([1, 4])
            with cy:
                if st.button("Yes", key="ly"):
                    updated = st.session_state.city_df[~ed_c["Select"]].reset_index(drop=True)
                    if upsert_table("locations", updated):
                        st.session_state.city_df = load_locations(); st.session_state.loc_del_conf = False; st.success("✅ Saved!"); time.sleep(1); st.rerun()
            with cn:
                if st.button("No", key="ln"): st.session_state.loc_del_conf = False; st.rerun()
    with col2:
        st.subheader("👤 Add Employee")
        v_df = st.session_state.city_df.dropna()
        loc_opts = ["Select Location"] + [f"{r['company_name']} - {r['city_name']}" for _, r in v_df.iterrows()]
        with st.container(border=True):
            # ✅ form_key로 저장 후 폼 완전 초기화
            fk = st.session_state.emp_form_key
            nm = st.text_input("Name", key=f"emp_name_{fk}", value="")
            em = st.text_input("Email", key=f"emp_email_{fk}", value="")
            lo = st.selectbox("Location", options=loc_opts, key=f"emp_loc_{fk}")
            ar, ms = 40, 40
            if lo != "Select Location":
                city_part = lo.split(" - ")[-1].strip()
                rule = SICK_LEAVE_RULES.get(city_part, {"text": "ℹ️ No law found.", "rate": None, "max": None})
                st.info(rule["text"])
                if rule["rate"] is None:
                    ar = st.number_input("Sick Rate (1hr per X)", value=40, key=f"emp_sr_{fk}")
                    ms = st.number_input("Sick Max", value=40, key=f"emp_sm_{fk}")
                else: ar, ms = rule["rate"], rule["max"]
            et = st.selectbox("Type", ["Full-Time", "Part-Time (with vacation)", "Part-Time (No vacation)"], key=f"emp_type_{fk}")
            vl = st.number_input("Annual Vacation Limit", value=(80.0 if "Full" in et else (40.0 if "with" in et else 0.0)), key=f"emp_vl_{fk}")
            if st.button("Add Employee", type="primary", use_container_width=True):
                if nm and lo != "Select Location":
                    new_p = pd.DataFrame([{"Name": nm, "Email": em, "Location": lo, "Type": et, "Vacation_Limit": vl, "Sick_Rate": ar, "Sick_Max": ms}])
                    updated = pd.concat([st.session_state.emp_df, new_p], ignore_index=True)
                    if upsert_table("employees", updated):
                        st.session_state.emp_df = load_employees()
                        # ✅ 폼 초기화: key 카운터 증가
                        st.session_state.emp_form_key += 1
                        st.success("✅ Saved!")
                        time.sleep(1)
                        st.rerun()
                else:
                    st.warning("⚠️ Name과 Location은 필수입니다.")

# ═══════════════════════════════════════════════════════════════
# 2. Log Worked Hours
# ═══════════════════════════════════════════════════════════════
elif menu == "2. Log Worked Hours":
    st.session_state.emp_df = load_employees()
    st.session_state.log_df = load_work_logs()
    st.markdown("### ⏳ Log Worked Hours")

    rc_list = [f"{r['company_name']} - {r['city_name']}" for _, r in st.session_state.city_df.iterrows()]
    cur_idx = rc_list.index(st.session_state.selected_company) if st.session_state.selected_company in rc_list else 0
    sc = st.selectbox("Select Company", options=rc_list, index=cur_idx, key="sc_page2")
    if sc != st.session_state.selected_company:
        st.session_state.selected_company = sc

    dr = st.date_input("Select Work Period", value=(date.today(), date.today()))
    ce = st.session_state.emp_df[st.session_state.emp_df["Location"] == sc].copy()

    if not ce.empty:
        idat = pd.DataFrame({
            "Select": [False]*len(ce),
            "Employee Name": ce["Name"].values,
            "Email": ce["Email"].values,
            "Employee Status": ["Employed"]*len(ce),
            "Hours Worked": [0.0]*len(ce)
        })
        ed_in = st.data_editor(idat,
            column_config={
                "Select": st.column_config.CheckboxColumn("Select"),
                "Employee Status": st.column_config.SelectboxColumn(options=["Employed", "Resigned"], required=True)
            },
            use_container_width=True, height=400, key="win_ed", disabled=not is_admin())

        if not is_admin():
            st.info("👁️ View-only mode: saving is disabled.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🚀 Save Worked Hours", type="primary", use_container_width=True, disabled=not is_admin()):
                if isinstance(dr, (list, tuple)) and len(dr) == 2:
                    nl = []
                    # ✅ Resigned는 별도 테이블로만 이동 — employees 삭제 안 함
                    resigned_records = []
                    full_emp = load_employees(); cu = load_leave_usage()
                    for _, r in ed_in.iterrows():
                        if r["Hours Worked"] > 0:
                            nl.append({"Employee": r["Employee Name"], "Status": r["Employee Status"], "Start_Date": str(dr[0]), "End_Date": str(dr[1]), "Hours_Worked": r["Hours Worked"]})
                        if r["Employee Status"] == "Resigned":
                            inf_rows = full_emp[full_emp["Name"] == r["Employee Name"]]
                            if not inf_rows.empty:
                                inf = inf_rows.iloc[0]
                                logs = load_work_logs()
                                tw = logs[logs["Employee"] == r["Employee Name"]]["Hours_Worked"].sum()
                                uv = cu[cu["Employee"] == r["Employee Name"]]["Vacation_Used"].sum()
                                us = cu[cu["Employee"] == r["Employee Name"]]["Sick_Used"].sum()
                                sr = float(inf["Sick_Rate"]) if float(inf.get("Sick_Rate", 0)) > 0 else 40
                                resigned_records.append({"Name": r["Employee Name"], "Resigned_Date": str(dr[1]), "Location": inf["Location"], "Total_Worked": tw, "Retained_Vacation": round(float(inf["Vacation_Limit"]) - uv, 2), "Retained_Sick": round((tw / sr) - us, 2), "Paid_Date": str(dr[1]), "Paid_Amount": 0.0, "Email": str(inf["Email"]), "Type": inf["Type"], "Vacation_Limit": inf["Vacation_Limit"], "Sick_Rate": inf["Sick_Rate"], "Sick_Max": inf["Sick_Max"]})
                    if nl:
                        existing = load_work_logs()
                        updated_log = pd.concat([existing, pd.DataFrame(nl)], ignore_index=True)
                        upsert_table("work_log", updated_log)
                    if resigned_records:
                        updated_res = pd.concat([load_resigned(), pd.DataFrame(resigned_records)], ignore_index=True)
                        upsert_table("resigned_employees", updated_res)
                        # ✅ Resigned 처리 시 employees에서 제거 — 이것만 허용
                        r_names = [r["Name"] for r in resigned_records]
                        delete_employees_by_name(r_names)
                    st.success("✅ Saved!"); time.sleep(1); st.rerun()
        with c2:
            # ✅ 오직 이 버튼만 employees 삭제 가능
            if st.button("🗑️ Delete Selected Employees", use_container_width=True, disabled=not is_admin()):
                if ed_in["Select"].any(): st.session_state.emp_del_conf = True
        if st.session_state.get("emp_del_conf", False):
            names_to_del = ed_in[ed_in["Select"] == True]["Employee Name"].tolist()
            st.error(f"⚠️ 아래 직원을 명단에서 삭제할까요?\n{', '.join(names_to_del)}")
            dy, dn = st.columns([1, 4])
            with dy:
                if st.button("Yes", key="ey_emp"):
                    if delete_employees_by_name(names_to_del):
                        st.session_state.emp_df = load_employees()
                        st.session_state.emp_del_conf = False
                        st.success("✅ 삭제 완료!"); time.sleep(1); st.rerun()
            with dn:
                if st.button("No", key="en_emp"): st.session_state.emp_del_conf = False; st.rerun()
    else:
        st.info("등록된 직원이 없습니다. 1페이지에서 직원을 먼저 등록해주세요.")

    st.markdown("---")
    st.markdown("### 📤 Bulk Upload Work Hours (CSV)")
    with st.expander("📋 CSV 일괄 업로드 — 180개도 한 번에!", expanded=False):
        st.info("""
**📌 CSV 형식 안내**
- 필수 컬럼: `Employee`, `Start_Date`, `End_Date`, `Hours_Worked`
- 날짜 형식: `MM/DD/YYYY` 또는 `YYYY-MM-DD`
        """)
        template_df = pd.DataFrame([
            {"Employee": "John Kim",  "Start_Date": "01/01/2026", "End_Date": "01/15/2026", "Hours_Worked": 80.0},
            {"Employee": "Jane Lee",  "Start_Date": "01/01/2026", "End_Date": "01/15/2026", "Hours_Worked": 72.0},
        ])
        st.download_button(label="📥 Download Template CSV", data=template_df.to_csv(index=False).encode("utf-8-sig"), file_name="bulk_upload_template.csv", mime="text/csv", use_container_width=True)

        uploaded_csv = st.file_uploader("CSV 파일 선택", type=["csv"], key="bulk_csv")
        if uploaded_csv and is_admin():
            try:
                udf = pd.read_csv(uploaded_csv)
                required_cols = {"Employee", "Start_Date", "End_Date", "Hours_Worked"}
                missing = required_cols - set(udf.columns)
                if missing:
                    st.error(f"❌ 필수 컬럼 없음: {missing}")
                else:
                    udf["Start_Date"] = pd.to_datetime(udf["Start_Date"], format='mixed').dt.strftime("%Y-%m-%d")
                    udf["End_Date"]   = pd.to_datetime(udf["End_Date"],   format='mixed').dt.strftime("%Y-%m-%d")
                    udf["Hours_Worked"] = pd.to_numeric(udf["Hours_Worked"], errors="coerce").fillna(0)
                    udf["Status"] = "Employed"
                    valid_names = st.session_state.emp_df["Name"].tolist()
                    udf["_valid"] = udf["Employee"].isin(valid_names)
                    invalid_names = udf[~udf["_valid"]]["Employee"].unique().tolist()
                    valid_udf = udf[udf["_valid"]].drop(columns=["_valid"]).reset_index(drop=True)
                    col_info1, col_info2 = st.columns(2)
                    with col_info1: st.success(f"✅ 유효한 레코드: **{len(valid_udf)}건**")
                    with col_info2:
                        if invalid_names: st.warning(f"⚠️ 미등록 직원 (제외됨): {', '.join(invalid_names)}")
                    if not valid_udf.empty:
                        valid_udf = valid_udf.sort_values("Employee").reset_index(drop=True)
                        st.dataframe(valid_udf[["Employee","Start_Date","End_Date","Hours_Worked"]], use_container_width=True, hide_index=True)
                        if st.button("🚀 Bulk Save to Work Log", type="primary", use_container_width=True):
                            existing = load_work_logs()
                            if not existing.empty:
                                existing_keys = set(zip(existing["Employee"].astype(str), existing["Start_Date"].astype(str)))
                                to_add = valid_udf[~valid_udf.apply(lambda r: (r["Employee"], r["Start_Date"]) in existing_keys, axis=1)]
                            else:
                                to_add = valid_udf.copy()
                            to_add = to_add[["Employee","Status","Start_Date","End_Date","Hours_Worked"]]
                            if to_add.empty:
                                st.warning("⚠️ 이미 동일한 기간의 데이터가 저장되어 있습니다.")
                            else:
                                combined = pd.concat([existing, to_add], ignore_index=True)
                                if upsert_table("work_log", combined):
                                    st.success(f"✅ {len(to_add)}건 저장 완료!")
                                    time.sleep(1); st.rerun()
            except Exception as e:
                st.error(f"❌ CSV 읽기 오류: {e}")
        elif uploaded_csv and not is_admin():
            st.warning("⛔ Admin만 업로드할 수 있습니다.")

    st.markdown("---")
    st.session_state.log_df = load_work_logs()

    if not st.session_state.log_df.empty:
        with st.expander("🔍 Filter & Download Logs", expanded=True):
            emp_in_company = st.session_state.emp_df[st.session_state.emp_df["Location"] == sc]["Name"].tolist()
            emp_opts = ["All Employees"] + sorted(emp_in_company)
            sfn = st.selectbox("Select Employee", options=emp_opts, key="filter_emp")
            fdf = st.session_state.log_df.copy()
            fdf = fdf[fdf["Employee"].isin(emp_in_company)]
            if sfn != "All Employees": fdf = fdf[fdf["Employee"] == sfn]
            if isinstance(dr, (list, tuple)) and len(dr) == 2:
                fdf = fdf[(fdf["Start_Date"] >= dr[0]) & (fdf["Start_Date"] <= dr[1])]
            if not fdf.empty:
                st.download_button(label="📥 Download Logs", data=fdf.reset_index(drop=True).to_csv(index=False).encode("utf-8-sig"), file_name="WorkLog.csv", mime="text/csv")

        ldat = fdf.sort_values(["Employee", "Start_Date"]).reset_index(drop=True).copy()
        ldat.insert(0, "No.", range(1, len(ldat) + 1))
        ldat.insert(1, "Select", False)
        ed_l = st.data_editor(ldat, use_container_width=True, key="led", hide_index=True)
        col_del, col_save = st.columns(2)
        with col_del:
            if st.button("🗑️ Delete Selected Logs", use_container_width=True, disabled=not is_admin()):
                if ed_l["Select"].any(): st.session_state.log_del_conf = True
        with col_save:
            if st.button("💾 Save Log Changes", type="primary", use_container_width=True, disabled=not is_admin()):
                full_log = load_work_logs()
                for _, row in ed_l.iterrows():
                    idx = int(row["No."]) - 1
                    if idx < len(fdf):
                        orig_idx = fdf.index[idx]
                        full_log.at[orig_idx, "Hours_Worked"] = row.get("Hours_Worked", 0)
                if upsert_table("work_log", full_log.reset_index(drop=True)):
                    st.success("✅ Saved!"); time.sleep(1); st.rerun()
        if st.session_state.get("log_del_conf", False):
            st.error("⚠️ Delete selected log entries?")
            ly, ln = st.columns([1, 4])
            with ly:
                if st.button("Yes", key="ly_log"):
                    full_log = load_work_logs()
                    del_emp   = ed_l[ed_l["Select"] == True]["Employee"].tolist()
                    del_start = ed_l[ed_l["Select"] == True]["Start_Date"].tolist()
                    mask = ~(full_log["Employee"].isin(del_emp) & full_log["Start_Date"].isin(del_start))
                    if upsert_table("work_log", full_log[mask].reset_index(drop=True)):
                        st.session_state.log_del_conf = False; st.success("✅ Saved!"); time.sleep(1); st.rerun()
            with ln:
                if st.button("No", key="ln_log"): st.session_state.log_del_conf = False; st.rerun()

# ═══════════════════════════════════════════════════════════════
# 3. Plan & Submit Leave
# ═══════════════════════════════════════════════════════════════
elif menu == "3. Plan/Submit Leave":
    st.title("📅 Plan & Submit Leave")

    if st.session_state.leave_modal_record is not None:
        show_leave_modal(st.session_state.leave_modal_record)
        st.session_state.leave_modal_record = None

    rc_list = [f"{r['company_name']} - {r['city_name']}" for _, r in st.session_state.city_df.iterrows()]
    cur_idx = rc_list.index(st.session_state.selected_company) if st.session_state.selected_company in rc_list else 0
    sc = st.selectbox("Select Company", options=rc_list, index=cur_idx, key="sc_page3")
    if sc != st.session_state.selected_company:
        st.session_state.selected_company = sc

    ce = load_employees()
    ce = ce[ce["Location"] == sc].copy()
    cl = load_work_logs()
    cu = load_leave_usage()
    emp_names_in_company = ce["Name"].tolist()
    cu_filtered = cu[cu["Employee"].isin(emp_names_in_company)] if not cu.empty else cu

    cy1, cm1, _ = st.columns([1, 1, 3])
    with cy1: today = date.today(); c_year = st.selectbox("Year", range(today.year-1, today.year+2), index=1)
    with cm1: c_month = st.selectbox("Month", range(1, 13), index=today.month-1)

    cal = calendar.monthcalendar(c_year, c_month)
    month_name_str = calendar.month_name[c_month]

    st.markdown("""
    <style>
    .cal-header-sun  { text-align:center; font-weight:bold; font-size:14px; background:#c0392b; color:white; padding:8px 4px; border-radius:4px; margin-bottom:4px; }
    .cal-header-sat  { text-align:center; font-weight:bold; font-size:14px; background:#1a6fc4; color:white; padding:8px 4px; border-radius:4px; margin-bottom:4px; }
    .cal-header-week { text-align:center; font-weight:bold; font-size:14px; background:#4a4a4a; color:white; padding:8px 4px; border-radius:4px; margin-bottom:4px; }
    .cal-day-num-sun { font-weight:bold; font-size:15px; color:#c0392b; padding:2px 4px; }
    .cal-day-num-sat { font-weight:bold; font-size:15px; color:#1a6fc4; padding:2px 4px; }
    .cal-day-num     { font-weight:bold; font-size:15px; color:#333;    padding:2px 4px; }
    .cal-empty { min-height:110px; border:1px solid #eee; background:#fafafa; border-radius:4px; }
    </style>
    """, unsafe_allow_html=True)

    st.markdown(f"<div style='text-align:center; font-weight:bold; font-size:24px; margin:8px 0 12px 0;'>{month_name_str} {c_year}</div>", unsafe_allow_html=True)

    hcols = st.columns(7)
    day_names = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
    for i, d in enumerate(day_names):
        if i == 0:   css = "cal-header-sun"
        elif i == 6: css = "cal-header-sat"
        else:        css = "cal-header-week"
        hcols[i].markdown(f"<div class='{css}'>{d}</div>", unsafe_allow_html=True)

    for week in cal:
        wcols = st.columns(7)
        for i, day in enumerate(week):
            with wcols[i]:
                if day == 0:
                    st.markdown("<div class='cal-empty'></div>", unsafe_allow_html=True)
                else:
                    curr_date_str = f"{c_year}-{c_month:02d}-{day:02d}"
                    day_data = cu_filtered[cu_filtered["Date"] == curr_date_str] if "Date" in cu_filtered.columns else pd.DataFrame()
                    if i == 0:   num_css = "cal-day-num-sun"
                    elif i == 6: num_css = "cal-day-num-sat"
                    else:        num_css = "cal-day-num"
                    st.markdown(f"<div class='{num_css}'>{day}</div>", unsafe_allow_html=True)
                    for idx, row in day_data.iterrows():
                        vac_h  = float(row["Vacation_Used"])
                        sick_h = float(row["Sick_Used"])
                        hrs    = vac_h if vac_h > 0 else sick_h
                        s_char = "P" if row["Status"] == "Plan" else "U"
                        dot = "🔵" if row["Status"] == "Plan" else "🟢"
                        btn_label = f"{dot} {row['Employee']} ({s_char}, {hrs}h)"
                        if st.button(btn_label, key=f"badge_{c_year}_{c_month}_{day}_{idx}", use_container_width=True, help="클릭하면 상세 정보를 볼 수 있습니다"):
                            st.session_state.leave_modal_record = row.to_dict()
                            st.rerun()

    st.markdown("---")

    with st.container(border=True):
        i1, i2, i3, i4, i5 = st.columns(5)
        emp_list = ce["Name"].tolist() if not ce.empty else ["No employees"]
        with i1: p_emp = st.selectbox("Employee Name", emp_list)
        with i2: p_date = st.date_input("Date", value=date.today())
        with i3: p_type = st.selectbox("Leave Type", ["Vacation", "Sick Leave"])
        with i4: p_status = st.selectbox("Status", ["Plan", "Used"])
        with i5: p_hours = st.number_input("Hours", min_value=0.5, step=0.5, value=8.0)
        if st.button("Submit to Calendar", type="primary", use_container_width=True, disabled=not is_admin()):
            new_l = pd.DataFrame([{"Employee": p_emp, "Date": p_date.strftime("%Y-%m-%d"), "Vacation_Used": p_hours if p_type == "Vacation" else 0.0, "Sick_Used": p_hours if p_type == "Sick Leave" else 0.0, "Status": p_status, "Note": ""}])
            updated = pd.concat([cu, new_l], ignore_index=True)
            if upsert_table("leave_usage", updated):
                st.success("✅ Saved!"); time.sleep(1); st.rerun()

    st.subheader("📊 Current Balances")
    th = cl.groupby("Employee")["Hours_Worked"].sum().reset_index() if not cl.empty else pd.DataFrame(columns=["Employee","Hours_Worked"])
    tu = cu_filtered[cu_filtered["Status"] == "Used"].groupby("Employee").agg({"Vacation_Used": "sum", "Sick_Used": "sum"}).reset_index() if not cu_filtered.empty else pd.DataFrame(columns=["Employee","Vacation_Used","Sick_Used"])
    summ = []
    for _, e in ce.iterrows():
        w = th[th["Employee"] == e["Name"]]["Hours_Worked"].values; tw = round(float(w[0]), 2) if len(w) > 0 else 0.0
        u = tu[tu["Employee"] == e["Name"]]; uv, us = (u["Vacation_Used"].sum(), u["Sick_Used"].sum()) if not u.empty else (0.0, 0.0)
        sr = float(e.get("Sick_Rate", 40)) if float(e.get("Sick_Rate", 0)) > 0 else 40
        s_acc = min(round(tw / sr, 2), float(e.get("Sick_Max", 40)))
        summ.append({"Name": e["Name"], "Location": e["Location"], "Total Worked": tw, "Total Vacation": round(float(e["Vacation_Limit"]), 2), "Used Vacation": uv, "Retained Vacation": round(float(e["Vacation_Limit"]) - uv, 2), "Sick Leave Accrued": s_acc, "Used Sick Leave": us, "Retained Sick Leave": round(s_acc - us, 2)})
    if summ:
        df_summ = pd.DataFrame(summ); df_summ.insert(0, "Select", False)
        def apply_st(val): return "background-color: #FFC0CB; color: black; font-weight: bold" if isinstance(val, (int, float)) and val < 0 else "background-color: #FFFF00; color: black; font-weight: bold"
        edited_summ = st.data_editor(df_summ.style.applymap(apply_st, subset=["Retained Vacation", "Retained Sick Leave"]).format(precision=2), use_container_width=True, hide_index=True, key="bal_editor")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("💾 Save Balances Changes", use_container_width=True, disabled=not is_admin()):
                emp_data = load_employees()
                for _, row in edited_summ.iterrows():
                    emp_data.loc[emp_data["Name"] == row["Name"], "Vacation_Limit"] = row["Total Vacation"]
                if upsert_table("employees", emp_data):
                    st.success("✅ Saved!"); time.sleep(1); st.rerun()
        with col_b2:
            # ✅ Current Balances에서는 leave_usage 데이터만 삭제 — employees 절대 건드리지 않음
            if st.button("🗑️ Delete Leave Data (Selected)", use_container_width=True, disabled=not is_admin()):
                if edited_summ["Select"].any():
                    names = edited_summ[edited_summ["Select"] == True]["Name"].tolist()
                    cu_all = load_leave_usage()
                    remaining = cu_all[~cu_all["Employee"].isin(names)].reset_index(drop=True)
                    if upsert_table("leave_usage", remaining):
                        st.success("✅ Leave 데이터 삭제 완료! (직원 명단은 유지됩니다)"); time.sleep(1); st.rerun()

    with st.expander("📥 Download & Manage Usage/Plan Logs"):
        if not cu_filtered.empty:
            fl_emp = st.selectbox("Select Employee to Filter", ["All Employees"] + sorted(cu_filtered["Employee"].unique().tolist()), key="f_del")
            cu_to_edit = cu_filtered.copy() if fl_emp == "All Employees" else cu_filtered[cu_filtered["Employee"] == fl_emp]
            if not cu_to_edit.empty:
                st.download_button(label=f"📥 Download [{fl_emp}] Leave History", data=cu_to_edit.to_csv(index=False).encode("utf-8-sig"), file_name="LeaveHistory.csv", mime="text/csv", use_container_width=True)
            cu_disp = cu_to_edit.copy(); cu_disp.insert(0, "Select", False)
            ed_usage = st.data_editor(cu_disp, use_container_width=True, key="usage_final", hide_index=True)
            if st.button("Delete Selected Entries", disabled=not is_admin()):
                if ed_usage["Select"].any(): st.session_state.usage_del_conf = True
            if st.session_state.get("usage_del_conf", False):
                st.error("⚠️ Permanently delete selected entries?"); uy_c, un_c = st.columns([1, 4])
                with uy_c:
                    if st.button("Yes", key="uy_usage"):
                        sel_idx = ed_usage[ed_usage["Select"] == True].index.tolist()
                        cu_all = load_leave_usage()
                        updated = cu_all.drop(index=[i for i in sel_idx if i < len(cu_all)]).reset_index(drop=True)
                        if upsert_table("leave_usage", updated):
                            st.session_state.usage_del_conf = False; st.success("✅ Saved!"); time.sleep(1); st.rerun()
                with un_c:
                    if st.button("No", key="un_usage"): st.session_state.usage_del_conf = False; st.rerun()

# ═══════════════════════════════════════════════════════════════
# 4. Dashboard & Email
# ═══════════════════════════════════════════════════════════════
elif menu == "4. Dashboard & Email":
    st.title("📊 Dashboard & Email")
    rdf = load_resigned(); all_emp = load_employees()
    if not rdf.empty:
        st.subheader("👤 Resigned Employee List")
        def apply_res_style(df):
            style_df = pd.DataFrame("", index=df.index, columns=df.columns)
            mask = df.applymap(lambda x: isinstance(x, (int, float)) and x < 0)
            style_df[mask] = "background-color: #FFC0CB; color: black; font-weight: bold"
            return style_df
        edited_rdf = st.data_editor(rdf.assign(Select=False).style.apply(apply_res_style, axis=None).format(precision=2), use_container_width=True, height=300, key="res")
        c1, c2, c3 = st.columns([1, 1, 4])
        with c1:
            if st.button("🔄 Restore", disabled=not is_admin()):
                sel = edited_rdf[edited_rdf["Select"] == True]
                if not sel.empty:
                    m_emp = load_employees()
                    for _, r in sel.iterrows():
                        new = pd.DataFrame([{"Name": r["Name"], "Email": r["Email"], "Location": r["Location"], "Type": r["Type"], "Vacation_Limit": r["Vacation_Limit"], "Sick_Rate": r["Sick_Rate"], "Sick_Max": r["Sick_Max"]}])
                        m_emp = pd.concat([m_emp, new], ignore_index=True)
                    upsert_table("employees", m_emp)
                    upsert_table("resigned_employees", rdf.drop(sel.index).reset_index(drop=True))
                    st.success("✅ Saved!"); time.sleep(1); st.rerun()
        with c2:
            if st.button("🗑️ Delete", disabled=not is_admin()):
                upsert_table("resigned_employees", rdf.drop(edited_rdf[edited_rdf["Select"] == True].index).reset_index(drop=True))
                st.success("✅ Saved!"); time.sleep(1); st.rerun()
        with c3:
            if st.button("💾 Save", disabled=not is_admin()):
                upsert_table("resigned_employees", edited_rdf.drop(columns=["Select"]))
                st.success("✅ Saved!"); time.sleep(1); st.rerun()
        st.markdown("---")
        st.subheader("📑 Detailed History (Resigned Staff)")
        fl = load_work_logs(); usage_data = load_leave_usage()
        for name in rdf["Name"].tolist():
            with st.expander(f"🔍 Detail for: {name}"):
                r_info = rdf[rdf["Name"] == name].iloc[0]
                emp_logs  = fl[fl["Employee"] == name].sort_values(by="Start_Date")
                emp_usage = usage_data[(usage_data["Employee"] == name) & (usage_data["Status"] == "Used")].sort_values(by="Date")
                rate = float(r_info["Sick_Rate"]) if float(r_info.get("Sick_Rate", 0)) > 0 else 40
                tab1, tab2 = st.tabs(["🕒 Work Logs", "📅 Leave Usage"])
                with tab1:
                    rows = []
                    for _, log in emp_logs.iterrows():
                        acc = round(log["Hours_Worked"] / rate, 2)
                        rows.append({"From": log["Start_Date"], "To": log["End_Date"], "Hours": log["Hours_Worked"], "Accrued Sick": acc})
                    st.dataframe(pd.DataFrame(rows), use_container_width=True)
                with tab2:
                    if not emp_usage.empty: st.dataframe(emp_usage[["Date", "Vacation_Used", "Sick_Used"]], use_container_width=True)
                    else: st.info("No leave history.")
                st.write(f"**Final Balance:** Vacation {r_info['Retained_Vacation']}h | Sick {r_info['Retained_Sick']}h")

    st.markdown("---")
    st.subheader("📧 Send Official Leave Summary Email")
    full_list_names = []; combined_info = {}
    for _, r in all_emp.iterrows():
        label = f"{r['Name']} (Employed)"; full_list_names.append(label); combined_info[label] = r
    if not rdf.empty:
        for _, r in rdf.iterrows():
            label = f"{r['Name']} (Resigned)"; full_list_names.append(label); combined_info[label] = r
    if full_list_names:
        with st.container(border=True):
            e1, e2 = st.columns(2)
            with e1:
                sel_label = st.selectbox("Select Employee to Email", full_list_names)
                e_info = combined_info[sel_label]; sel_e_name = e_info["Name"]
                target_email = st.text_input("Recipient Email", value=e_info["Email"])
            with e2: dr = st.date_input("Report Period", value=(date(date.today().year, 1, 1), date.today()))
            if len(dr) == 2:
                logs = load_work_logs(); usage = load_leave_usage()
                p_logs  = logs[(logs["Employee"] == sel_e_name) & (logs["Start_Date"] >= dr[0]) & (logs["Start_Date"] <= dr[1])]
                p_usage = usage[(usage["Employee"] == sel_e_name) & (pd.to_datetime(usage["Date"]).dt.date >= dr[0]) & (pd.to_datetime(usage["Date"]).dt.date <= dr[1]) & (usage["Status"] == "Used")]
                tw_p = p_logs["Hours_Worked"].sum(); uv_p = p_usage["Vacation_Used"].sum(); us_p = p_usage["Sick_Used"].sum()
                total_w = logs[logs["Employee"] == sel_e_name]["Hours_Worked"].sum()
                rate  = float(e_info["Sick_Rate"]) if float(e_info.get("Sick_Rate", 0)) > 0 else 40
                ret_v = round(float(e_info["Vacation_Limit"]) - usage[(usage["Employee"] == sel_e_name) & (usage["Status"] == "Used")]["Vacation_Used"].sum(), 2)
                ret_s = round((total_w / rate) - usage[(usage["Employee"] == sel_e_name) & (usage["Status"] == "Used")]["Sick_Used"].sum(), 2)
                body_text  = f"Dear {sel_e_name},\n\nOfficial summary from {dr[0]} to {dr[1]}:\n\n"
                body_text += f"- Total Hours Worked: {tw_p:.2f} hrs\n"
                if uv_p > 0: body_text += f"- Used Vacation: {uv_p:.2f} hrs\n"
                if us_p > 0: body_text += f"- Used Sick leave: {us_p:.2f} hrs\n"
                if not p_usage.empty:
                    body_text += "\nLeave History Detail:\n"
                    for _, row in p_usage.iterrows():
                        h = row["Vacation_Used"] if float(row["Vacation_Used"]) > 0 else row["Sick_Used"]
                        body_text += f"- {row['Date']} ({h:.2f} hrs)\n"
                body_text += "\nCurrent Balance Status (As of today):\n"
                if ret_v != 0: body_text += f"- Retained Vacation: {ret_v:.2f} hrs\n"
                if ret_s != 0: body_text += f"- Retained Sick Leave: {ret_s:.2f} hrs\n"
                if "(Resigned)" in sel_label:
                    try:
                        paid_val = float(e_info.get("Paid_Amount", 0))
                        if paid_val > 0: body_text += f"\n[Settlement Info] ${paid_val:,.2f} was paid on {e_info.get('Paid_Date','N/A')}.\n"
                    except: pass
                body_text += "\nBest regards,\nAccounting Department\nAmlotus"
                edit_subject = st.text_input("Subject", value=f"Leave Summary - {sel_e_name}")
                edit_body    = st.text_area("Content", value=body_text, height=350)
                google_pw    = st.text_input("Enter Google App Password", type="password")
                if st.button("🚀 Send Email", type="primary", use_container_width=True, disabled=not is_admin()):
                    if google_pw:
                        html_c = edit_body.replace("\n", "<br>")
                        if ret_v < 0: html_c = html_c.replace(f"Retained Vacation: {ret_v:.2f} hrs", f"<span style='color:red; font-weight:bold;'>Retained Vacation: {ret_v:.2f} hrs</span>")
                        if ret_s < 0: html_c = html_c.replace(f"Retained Sick Leave: {ret_s:.2f} hrs", f"<span style='color:red; font-weight:bold;'>Retained Sick Leave: {ret_s:.2f} hrs</span>")
                        try:
                            msg = MIMEMultipart(); msg["From"] = "accounting@amlotus.edu"; msg["To"] = target_email
                            msg["Subject"] = edit_subject; msg.attach(MIMEText(f"<html><body>{html_c}</body></html>", "html"))
                            server = smtplib.SMTP("smtp.gmail.com", 587); server.starttls()
                            server.login("accounting@amlotus.edu", google_pw); server.send_message(msg); server.quit()
                            st.success("✅ Email sent!"); time.sleep(1)
                        except Exception as e: st.error(f"❌ Error: {e}")
