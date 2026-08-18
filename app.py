import io
import os
from datetime import date, time

import pandas as pd
import streamlit as st

try:
    from supabase import create_client
except ImportError:
    create_client = None

st.set_page_config(
    page_title="Electric Forklift Control",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Professional UI ----------
st.markdown(
    """
    <style>
    .stApp { background: #f5f7fa; }
    .block-container { max-width: 1380px; padding: 1.4rem 2rem 3rem; }
    [data-testid="stSidebar"] { background: #111827; }
    [data-testid="stSidebar"] * { color: #f9fafb !important; }
    .topbar {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        color: white; padding: 24px 28px; border-radius: 18px;
        margin-bottom: 18px; box-shadow: 0 8px 24px rgba(15,23,42,.12);
    }
    .topbar h1 { margin: 0; font-size: 30px; font-weight: 750; letter-spacing: -.5px; }
    .topbar p { margin: 7px 0 0; color: #cbd5e1; font-size: 14px; }
    .section-title { font-size: 20px; font-weight: 750; color: #0f172a; margin: 18px 0 10px; }
    .section-subtitle { color: #64748b; font-size: 13px; margin-top: -7px; margin-bottom: 14px; }
    .kpi-card {
        background: white; border: 1px solid #e2e8f0; border-radius: 14px;
        padding: 16px 18px; min-height: 92px; box-shadow: 0 2px 8px rgba(15,23,42,.04);
    }
    .kpi-label { color: #64748b; font-size: 12px; font-weight: 650; text-transform: uppercase; letter-spacing: .5px; }
    .kpi-value { color: #0f172a; font-size: 28px; font-weight: 800; margin-top: 4px; }
    .panel {
        background: white; border: 1px solid #e2e8f0; border-radius: 16px;
        padding: 18px; box-shadow: 0 2px 10px rgba(15,23,42,.04); margin-bottom: 14px;
    }
    .flt-badge {
        display: inline-flex; align-items: center; justify-content: center;
        width: 48px; height: 34px; border-radius: 9px; background: #e8f0ff;
        color: #1d4ed8; font-weight: 800; font-size: 15px;
    }
    .status-note { color: #64748b; font-size: 12px; }
    .footer-note { color: #94a3b8; text-align: center; font-size: 12px; padding-top: 20px; }
    div[data-testid="stMetric"] { background: white; border: 1px solid #e2e8f0; padding: 12px 15px; border-radius: 12px; }
    .stButton > button { border-radius: 10px; font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

FLT_CODES = ["10", "13", "14", "24", "29", "30", "31"]
STATUS_OPTIONS = ["Active", "Charging", "Maintenance", "Unavailable", "Other"]
SHIFT_OPTIONS = ["Morning", "Afternoon", "Night"]


def get_supabase():
    if create_client is None:
        return None
    url = st.secrets.get("SUPABASE_URL", os.getenv("SUPABASE_URL", ""))
    key = st.secrets.get("SUPABASE_KEY", os.getenv("SUPABASE_KEY", ""))
    if not url or not key:
        return None
    return create_client(url, key)


def save_records(df):
    client = get_supabase()
    if client is None:
        return False, "Supabase is not configured yet."
    rows = df.where(pd.notna(df), None).to_dict(orient="records")
    try:
        client.table("forklift_shift_records").insert(rows).execute()
        return True, "Saved permanently."
    except Exception as exc:
        return False, f"Could not save records: {exc}"


def load_history():
    client = get_supabase()
    if client is None:
        return pd.DataFrame()
    try:
        result = client.table("forklift_shift_records").select("*").order("created_at", desc=True).execute()
        return pd.DataFrame(result.data or [])
    except Exception:
        return pd.DataFrame()


def make_excel(history):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export = history.copy()
        if "id" in export.columns:
            export = export.drop(columns=["id"])
        if "created_at" in export.columns:
            export = export.drop(columns=["created_at"])
        if export.empty:
            export = pd.DataFrame(columns=["Date", "Time", "Shift", "FLT Code", "Status", "Operator", "Operation", "Charge %"])
        export.to_excel(writer, index=False, sheet_name="Shift Register")
        ws = writer.book["Shift Register"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = min(max(max(len(str(c.value or "")) for c in col) + 2, 10), 28)
            ws.column_dimensions[col[0].column_letter].width = width
    output.seek(0)
    return output


# ---------- Header ----------
st.markdown(
    """
    <div class="topbar">
        <h1>⚡ Electric Forklift Control</h1>
        <p>Shift-start status registration • operator allocation • deployment location • battery monitoring</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if "rows" not in st.session_state:
    st.session_state.rows = [
        {"code": c, "status": "Active", "operator": "", "operation": "", "charge": 0}
        for c in FLT_CODES
    ]

# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("### ⚡ Shift Control")
    st.caption("Enter the status of all electric forklifts at the beginning of the shift.")
    st.divider()
    shift_date = st.date_input("Shift date", value=date.today())
    shift_time = st.time_input("Start time", value=time(15, 0))
    shift = st.selectbox("Shift", SHIFT_OPTIONS)
    st.divider()
    st.markdown("**Status guide**")
    st.caption("🟢 Active  •  🔋 Charging  •  🔧 Maintenance  •  🔴 Unavailable")
    st.divider()
    st.caption("Each submission creates one record per forklift in the permanent register.")

# ---------- Entry ----------
st.markdown('<div class="section-title">Shift-start forklift status</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">Record who is operating each FLT, where it is deployed, and its current battery level.</div>', unsafe_allow_html=True)

with st.container(border=True):
    header = st.columns([0.75, 1.65, 1.05, 1.9, 1.7])
    for col, label in zip(header, ["FLT", "Status", "Battery", "Operator", "Deployment / Operation"]):
        col.markdown(f"**{label}**")

    for i, row in enumerate(st.session_state.rows):
        cols = st.columns([0.75, 1.65, 1.05, 1.9, 1.7])
        cols[0].markdown(f'<div class="flt-badge">FLT {row["code"]}</div>', unsafe_allow_html=True)
        row["status"] = cols[1].selectbox(
            "Status", STATUS_OPTIONS,
            index=STATUS_OPTIONS.index(row["status"]) if row["status"] in STATUS_OPTIONS else 0,
            key=f"status_{i}", label_visibility="collapsed",
        )
        row["charge"] = cols[2].number_input(
            "Battery", min_value=0, max_value=100, value=int(row["charge"]), step=5,
            key=f"charge_{i}", label_visibility="collapsed",
        )
        row["operator"] = cols[3].text_input(
            "Operator", value=row["operator"], key=f"operator_{i}",
            label_visibility="collapsed", placeholder="Assigned operator",
        )
        row["operation"] = cols[4].text_input(
            "Operation", value=row["operation"], key=f"operation_{i}",
            label_visibility="collapsed", placeholder="L1 / L2 / RM / Inside",
        )

st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)

# ---------- Save ----------
if st.button("💾  SAVE SHIFT REPORT", type="primary", use_container_width=True):
    records = []
    for row in st.session_state.rows:
        records.append({
            "Date": shift_date.isoformat(),
            "Time": shift_time.strftime("%H:%M"),
            "Shift": shift,
            "FLT Code": row["code"],
            "Status": row["status"],
            "Operator": row["operator"].strip(),
            "Operation": row["operation"].strip(),
            "Charge %": row["charge"],
        })
    df = pd.DataFrame(records)
    ok, msg = save_records(df)
    if ok:
        st.session_state.last_saved = df
        st.success("Shift report saved successfully.")
    else:
        st.error(msg)
        st.info("Configure SUPABASE_URL and SUPABASE_KEY in Streamlit secrets to enable permanent storage.")

# ---------- Summary ----------
if "last_saved" in st.session_state:
    df = st.session_state.last_saved
    total = len(df)
    active = int((df["Status"] == "Active").sum())
    charging = int((df["Status"] == "Charging").sum())
    assigned = int(df["Operator"].fillna("").str.strip().ne("").sum())
    missing = total - assigned
    avg = df["Charge %"].mean()

    st.markdown('<div class="section-title">Shift summary</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">A quick control-room view of forklift availability, manpower and battery condition.</div>', unsafe_allow_html=True)

    kpis = [
        ("TOTAL FLTs", total),
        ("ACTIVE", active),
        ("CHARGING", charging),
        ("AVG. BATTERY", f"{avg:.1f}%"),
        ("OPERATORS ASSIGNED", assigned),
        ("OPERATORS MISSING", missing),
    ]
    cols = st.columns(6)
    for col, (label, value) in zip(cols, kpis):
        with col:
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([1.15, 1])
    with c1:
        st.markdown('<div class="panel"><b>Shift status terminal</b>', unsafe_allow_html=True)
        st.code(
            f"ELECTRIC FORKLIFT SHIFT STATUS\n"
            f"DATE : {shift_date.strftime('%d/%m/%Y')}    TIME : {shift_time.strftime('%I:%M %p')}\n"
            f"SHIFT: {shift}\n"
            "----------------------------------------\n"
            f"TOTAL FLT        : {total}\n"
            f"ACTIVE           : {active}\n"
            f"CHARGING         : {charging}\n"
            f"AVG BATTERY      : {avg:.1f}%\n"
            f"OPERATOR ASSIGNED: {assigned}\n"
            f"OPERATOR MISSING : {missing}\n"
            "----------------------------------------"
        )
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        missing_codes = df.loc[df["Operator"].fillna("").str.strip().eq(""), "FLT Code"].tolist()
        low_battery = df.loc[df["Charge %"] < 30, "FLT Code"].tolist()
        st.markdown('<div class="panel"><b>Attention required</b>', unsafe_allow_html=True)
        if missing_codes:
            st.warning("Operator missing: " + ", ".join(missing_codes))
        else:
            st.success("All forklifts have an operator assigned.")
        if low_battery:
            st.warning("Low battery (<30%): " + ", ".join(low_battery))
        else:
            st.success("No forklift below 30% battery.")
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Submitted shift</div>', unsafe_allow_html=True)
    display = df.copy()
    display["FLT Code"] = "FLT " + display["FLT Code"].astype(str)
    display["Charge %"] = display["Charge %"].astype(str) + "%"
    st.dataframe(display, use_container_width=True, hide_index=True)

# ---------- History ----------
st.markdown('<div class="section-title">Shift history</div>', unsafe_allow_html=True)
st.markdown('<div class="section-subtitle">Review previously submitted records and export the complete register to Excel.</div>', unsafe_allow_html=True)
history = load_history()
if history.empty:
    st.info("No permanent records available yet, or the database connection is not configured.")
else:
    st.dataframe(history.drop(columns=["id", "created_at"], errors="ignore"), use_container_width=True, hide_index=True)
    st.download_button(
        "📥  Download Full Excel Register",
        data=make_excel(history),
        file_name="electric_forklift_shift_register.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown('<div class="footer-note">Electric Forklift Control • Shift-start registration • Operator • Deployment • Battery</div>', unsafe_allow_html=True)
