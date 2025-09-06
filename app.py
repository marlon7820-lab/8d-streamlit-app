import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import date

st.set_page_config(page_title="Step-by-Step 8D App", layout="wide")
st.title("📋 Step-by-Step Guided 8D Problem Solving App")
st.write("Follow each step carefully. Click on a step to expand and fill it out.")

# Auto-fill report info
st.subheader("Report Info")
report_date = st.date_input("Report Date", value=date.today())
prepared_by = st.text_input("Prepared By (your name)")

# 8D Steps with instructions and placeholders
steps = [
    ("D1 - Establish the Team", "List all team members involved in solving this problem. Include roles.", "e.g., John Doe – QA, Jane Smith – Production"),
    ("D2 - Describe the Problem", "Provide a clear description of the problem (what, where, when).", "e.g., PCB failed test on 08/01/2025 due to voltage drop."),
    ("D3 - Containment Actions", "Describe immediate actions to contain the problem.", "e.g., Isolated affected units, notified production."),
    ("D4 - Root Cause", "Identify the root cause using data and analysis.", "e.g., Faulty resistor R23 from supplier."),
    ("D5 - Corrective Actions", "List planned actions to fix the problem.", "e.g., Replace components, retrain team."),
    ("D6 - Permanent Corrective Actions", "Explain how this problem will be permanently solved.", "e.g., Change supplier, add inspection step."),
    ("D7 - Prevent Recurrence", "List steps to prevent similar issues in the future.", "e.g., Update SOP, monthly audits."),
    ("D8 - Congratulate the Team", "Add final remarks or acknowledgment for the team.", "e.g., Thanks to QA and Production teams for swift action.")
]

answers = {}

# Use collapsible sections for each step
with st.form("8D_Form"):
    for step, guide, placeholder in steps:
        with st.expander(step):
            st.write(guide)
            answers[step] = st.text_area("Your answer:", placeholder=placeholder, height=80)
    submitted = st.form_submit_button("Generate 8D Report")

if submitted:
    file_name = f"8D_Report_{report_date}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # Title formatting
    ws["A1"] = "Step-by-Step Guided 8D Problem Solving Report"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws.merge_cells("A1:C1")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    # Report info
    ws["A3"] = "Report Date"
    ws["B3"] = str(report_date)
    ws["A4"] = "Prepared By"
    ws["B4"] = prepared_by

    # Fill in 8D steps
    row = 6
    for step, guide, placeholder in steps:
        ws[f"A{row}"] = step
        ws[f"A{row}"].font = Font(bold=True)
        ws[f"B{row}"] = guide
        ws[f"B{row}"].font = Font(italic=True)
        ws[f"C{row}"] = answers.get(step, "")
        ws[f"C{row}"].alignment = Alignment(wrap_text=True)
        row += 3

    # Adjust column widths
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 80

    wb.save(file_name)
    st.success(f"✅ 8D Report generated: {file_name}")
    st.download_button("Download Excel Report", file_name, file_name)
        
