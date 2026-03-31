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
        st.markdown(f"**👤 Employee**")
        st.markdown(f"**📅 Date**")
        st.markdown(f"**🗂️ Leave Type**")
        st.markdown(f"**⏱️ Hours**")
        st.markdown(f"**📊 Status**")
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
            st.markdown("
