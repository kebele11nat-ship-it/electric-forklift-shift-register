import io
import os
from datetime import date, datetime, time

import pandas as pd
import streamlit as st

try:
    from supabase import create_client
except ImportError:
    create_client = None

st.set_page_config(page_title="Electric Forklift Shift Register", page_icon="🧧", layout="wide")

st.markdown("""
<style>
.block-container {max-width:1200px; padding-top:2rem;}
.hero {padding:1.4rem; border:1px solid rgba(128,128,128,.25); border-radius:16px; margin-bottom:1.2rem;}
</style>
""", unsafe_allow_html=True)

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
        if history.empty:
            pd.DataFrame(columns=["Date", "Time", "Shift", "FLT Code", "Status", "Operator", "Operation", "Charge %"]).to_excel(writer, index=False, sheet_name="Shift Register")
        else:
            export = history.copy()
            if "id" in export.columns:
                export = export.drop(columns=["id"])
            if "created_at" in export.columns:
                export = export.drop(columns=["created_at"])
            export.to_excel(writer, index=False, sheet_name="Shift Register")
        ws = writer.book["Shift Register"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = min(max(max(len(str(c.value or "")) for c in col) + 2, 10), 28)
            ws.column_dimensions[col[0].column_letter].width = width
    output.seek(0)
    return output


st.markdown("""
<div class="hero">
<h1>🧧 Electric Forklift Shift Register</h1>
<p>Enter the OC shift report directly in the app. Save it to the permanent register and view the shift summary immediately.</p>
</div>
""", unsafe_allow_html=True)

if "rows" not in st.session_state:
    st.session_state.rows = [{"code": c, "status": "Active", "operator": "", "operation": "", "charge": 0} for c in FLT_CODES]

with st.sidebar:
    st.header("Shift details")
    shift_date = st.date_input("Date", value=date.today())
    shift_time = st.time_input("Time", value=time(15, 0))
    shift = st.selectbox("Shift", SHIFT_OPTIONS)
    st.divider()
    st.caption("The register stores one row per forklift for every submitted shift.")

st.subheader("1. Enter Electric Forklift Status")

header = st.columns([1, 1.6, 1.2, 1.8, 1.5])
for col, label in zip(header, ["FLT Code", "Status", "Charge %", "Operator", "Operation"]):
    col.markdown(f"**{label}**")

for i, row in enumerate(st.session_state.rows):
    cols = st.columns([1, 1.6, 1.2, 1.8, 1.5])
    cols[0].markdown(f"### {row['code']}")
    row["status"] = cols[1].selectbox("Status", STATUS_OPTIONS, index=STATUS_OPTIONS.index(row["status"]) if row["status"] in STATUS_OPTIONS else 0, key=f"status_{i}", label_visibility="collapsed")
    row["charge"] = cols[2].number_input("Charge", min_value=0, max_value=100, value=int(row["charge"]), step=5, key=f"charge_{i}", label_visibility="collapsed")
    row["operator"] = cols[3].text_input("Operator", value=row["operator"], key=f"operator_{i}", label_visibility="collapsed", placeholder="Operator")
    row["operation"] = cols[4].text_input("Operation", value=row["operation"], key=f"operation_{i}", label_visibility="collapsed", placeholder="L1 / L2 / RM / Inside")

st.divider()
st.subheader("2. Save Shift Report")

if st.button("💾 SAVE SHIFT REPORT", type="primary", use_container_width=True):
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
        st.success("✅ Shift report saved permanently to the register.")
    else:
        st.error(msg)
        st.info("Configure SUPABASE_URL and SUPABASE_KEY in Streamlit secrets to enable permanent storage.")

if "last_saved" in st.session_state:
    df = st.session_state.last_saved
    st.subheader("📟 Terminal Summary")
    total = len(df)
    active = int((df["Status"] == "Active").sum())
    charging = int((df["Status"] == "Charging").sum())
    assigned = int(df["Operator"].fillna("").str.strip().ne("").sum())
    missing = total - assigned
    avg = df["Charge %"].mean()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total FLTs", total)
    c2.metric("Active", active)
    c3.metric("Charging", charging)
    c4.metric("Avg. Charge", f"{avg:.1f}%")
    c5.metric("Operators Missing", missing)

    st.code(
        f"=== ELECTRIC FORKLIFT SHIFT STATUS ===\n"
        f"DATE : {shift_date.strftime('%d/%m/%Y')}    TIME : {shift_time.strftime('%I:%M %p')}\n"
        f"SHIFT: {shift}\n"
        "----------------------------------------\n"
        f"TOTAL FLT        : {total}\n"
        f"ACTIVE           : {active}\n"
        f"CHARGING         : {charging}\n"
        f"AVG CHARGE       : {avg:.1f}%\n"
        f"OPERATOR ASSIGNED: {assigned}\n"
        f"OPERATOR MISSING : {missing}\n"
        "----------------------------------------"
    )

    missing_codes = df.loc[df["Operator"].fillna("").str.strip().eq(""), "FLT Code"].tolist()
    if missing_codes:
        st.warning("⚠️ Operator not assigned: " + ", ".join(missing_codes))

    st.subheader("📋 Saved Shift")
    st.dataframe(df, use_container_width=True, hide_index=True)

st.divider()
st.subheader("📚 Shift History")
history = load_history()
if history.empty:
    st.info("No permanent records available yet, or database connection is not configured.")
else:
    st.dataframe(history.drop(columns=["id", "created_at"], errors="ignore"), use_container_width=True, hide_index=True)
    st.download_button("📥 Download Full Excel Register", data=make_excel(history), file_name="electric_forklift_shift_register.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.caption("Free web app • Permanent storage uses the free Supabase tier • Excel export is generated from the permanent register")
