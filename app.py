import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Electric Forklift Shift Register", page_icon="🧧", layout="wide")

st.markdown("""
<style>
.block-container {max-width:1200px; padding-top:2rem;}
.hero {padding:1.4rem; border:1px solid rgba(128,128,128,.25); border-radius:16px; margin-bottom:1.2rem;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>🧧 Electric Forklift Shift Register</h1>
<p>Paste the OC shift status message → extract the forklift data → review → download Excel.</p>
</div>
""", unsafe_allow_html=True)

EXAMPLE = '''Electric Forklift Status Report
📅 Date: 17/8/2026
⏰ Time: 3:00 PM

🧧 FLT Code: 30
📊 Status: Charging
👷 Operator:
⚙️ Operation:
🔋 Charge: 50%

🧧 FLT Code: 29
📊 Status: Active
👷 Operator: Eyasu E
⚙️ Operation: L2
🔋 Charge: 90%'''


def parse_report(text):
    date_match = re.search(r'Date\s*:\s*(.+)', text, re.I)
    time_match = re.search(r'Time\s*:\s*(.+)', text, re.I)
    date_text = date_match.group(1).strip() if date_match else ''
    time_text = time_match.group(1).strip() if time_match else ''
    blocks = re.split(r'(?=FLT\s*Code\s*:)', text, flags=re.I)
    records = []

    for block in blocks:
        if not re.search(r'FLT\s*Code\s*:', block, re.I):
            continue

        def field(pattern):
            m = re.search(pattern, block, re.I)
            return m.group(1).strip() if m else ''

        code = field(r'FLT\s*Code\s*:\s*([^\r\n]+)')
        status = field(r'Status\s*:\s*([^\r\n]+)')
        operator = field(r'Operator\s*:\s*([^\r\n]*)')
        operation = field(r'Operation\s*:\s*([^\r\n]*)')
        charge_text = field(r'Charge\s*:\s*([^\r\n]+)')
        charge_match = re.search(r'(\d+(?:\.\d+)?)\s*%', charge_text)
        charge = float(charge_match.group(1)) if charge_match else None

        if code:
            records.append({
                'Date': date_text,
                'Time': time_text,
                'FLT Code': code,
                'Status': status,
                'Operator': operator,
                'Operation': operation,
                'Charge %': charge,
            })
    return records


def make_excel(df, summary):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Shift Register')
        summary.to_excel(writer, index=False, sheet_name='Summary')
        wb = writer.book
        ws = wb['Shift Register']
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            width = min(max(max(len(str(c.value or '')) for c in col) + 2, 10), 28)
            ws.column_dimensions[col[0].column_letter].width = width
        sws = wb['Summary']
        for col in sws.columns:
            width = min(max(max(len(str(c.value or '')) for c in col) + 2, 12), 30)
            sws.column_dimensions[col[0].column_letter].width = width
    output.seek(0)
    return output

st.sidebar.header('Shift input')
if st.sidebar.button('Load example report', use_container_width=True):
    st.session_state['report_text'] = EXAMPLE
    st.rerun()

report_text = st.text_area(
    'Paste the OC WhatsApp report here',
    value=st.session_state.get('report_text', ''),
    height=360,
    placeholder='Paste the full Electric Forklift Status Report exactly as received from OC...'
)
st.session_state['report_text'] = report_text

if st.button('⚙️ Process Shift Report', type='primary', use_container_width=True):
    records = parse_report(report_text)
    if not records:
        st.error('No forklift records were detected. Make sure the message contains FLT Code, Status, Operator, Operation and Charge fields.')
    else:
        df = pd.DataFrame(records)
        df['Charge %'] = pd.to_numeric(df['Charge %'], errors='coerce')
        st.session_state['df'] = df

if 'df' in st.session_state:
    df = st.session_state['df'].copy()
    total = len(df)
    active = int((df['Status'].str.strip().str.lower() == 'active').sum())
    charging = int((df['Status'].str.strip().str.lower() == 'charging').sum())
    assigned = int(df['Operator'].fillna('').str.strip().ne('').sum())
    missing = total - assigned
    avg_charge = df['Charge %'].mean()

    st.divider()
    st.subheader('📟 Terminal Summary')
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('Total FLTs', total)
    c2.metric('Active', active)
    c3.metric('Charging', charging)
    c4.metric('Avg. Charge', f'{avg_charge:.1f}%' if pd.notna(avg_charge) else '—')
    c5.metric('Operators Missing', missing)

    terminal = (
        "=== ELECTRIC FORKLIFT SHIFT STATUS ===\n"
        f"DATE : {df.iloc[0]['Date']}    TIME : {df.iloc[0]['Time']}\n"
        "----------------------------------------\n"
        f"TOTAL FLT        : {total}\n"
        f"ACTIVE           : {active}\n"
        f"CHARGING         : {charging}\n"
        f"AVG CHARGE       : {avg_charge:.1f}%\n"
        f"OPERATOR ASSIGNED: {assigned}\n"
        f"OPERATOR MISSING : {missing}\n"
        "----------------------------------------"
    )
    st.code(terminal)

    if missing:
        missing_codes = ', '.join(df.loc[df['Operator'].fillna('').str.strip().eq(''), 'FLT Code'].astype(str))
        st.warning(f'⚠️ Operator not assigned: {missing_codes}')

    st.subheader('📋 Extracted Shift Register')
    display_df = df.copy()
    display_df['Charge %'] = display_df['Charge %'].apply(lambda x: f'{x:.0f}%' if pd.notna(x) else '')
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    status_counts = df['Status'].fillna('').str.strip().replace('', 'Not specified').value_counts()
    summary_rows = [
        ['Shift Date', df.iloc[0]['Date']],
        ['Shift Time', df.iloc[0]['Time']],
        ['Total Electric FLTs', total],
        ['Active', active],
        ['Charging', charging],
        ['Average Charge %', round(avg_charge, 1) if pd.notna(avg_charge) else ''],
        ['Operators Assigned', assigned],
        ['Operators Missing', missing],
    ]
    for status, count in status_counts.items():
        summary_rows.append([f'Status - {status}', int(count)])
    summary_df = pd.DataFrame(summary_rows, columns=['Metric', 'Value'])

    excel = make_excel(df, summary_df)
    filename_date = re.sub(r'[^0-9A-Za-z_-]+', '-', str(df.iloc[0]['Date'])) or 'shift'
    st.download_button(
        '📥 Download Excel Register',
        data=excel,
        file_name=f'electric_forklift_shift_{filename_date}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )

    st.caption('The app only parses and summarizes the OC message; it does not alter the source report.')
