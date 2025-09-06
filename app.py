import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import date
import os

st.set_page_config(page_title="NPQP 8D App", layout="wide")
st.title("📋 Nissan NPQP 8D App")
st.write("Text-based NPQP 8D format with Summary sheet tracking.")

# Report info
report_date = st.date_input("Report Date", value=date.today())
prepared_by = st.text_input("Prepared By")

# --- NPQP Sections ---
concern_desc = st.text_area("Concern Details", height=100)
concern_image_name = st.text_input("Concern Image Filename or URL (optional)")

# Similar Part Consideration
st.subheader("Similar Part Consideration")
other_models = st.selectbox("Other Models Affected?", ["No","Yes"])
generic_parts = st.selectbox("Generic Parts Affected?", ["No","Yes"])
other_colors = st.selectbox("Other Colors?", ["No","Yes"])
opposite_hand = st.selectbox("Opposite Hand?", ["No","Yes"])
front_rear = st.selectbox("Front/Rear?", ["No","Yes"])

# Initial Analysis
st.subheader("Initial Analysis")
detect_process = st.selectbox("Detected during process?", ["No","Yes"])
detect_final = st.selectbox("Detected at final inspection?", ["No","Yes"])
detect_prior = st.selectbox("Detected prior to dispatch?", ["No","Yes"])
detect_other = st.text_input("Other detection points")

# Temporary Countermeasures
temp_actions = st.text_area("Temporary Countermeasures", height=80)
temp_date = st.date_input("Implementation Date", value=date.today())

# Root Cause Analysis
root_cause = st.text_area("Root Cause Analysis", height=80)

# Permanent Corrective Actions
perm_actions = st.text_area("Permanent Corrective Actions", height=80)

# Effectiveness Verification
effect_verif = st.text_area("Effectiveness Verification", height=80)

# Standardization
standardization = st.text_area("Standardization / Procedure Updates", height=80)

# --- Submit ---
submitted = st.button("Save 8D Report")

if submitted:
    file_name = "NPQP_8D_Reports.xlsx"

    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        if "8D Reports" not in wb.sheetnames:
            ws = wb.create_sheet("8D Reports")
        else:
            ws = wb["8D Reports"]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "8D Reports"

    row = ws.max_row + 2

    # Report Info
    ws[f"A{row}"] = "Report Date"
    ws[f"B{row}"] = str(report_date)
    row += 1
    ws[f"A{row}"] = "Prepared By"
    ws[f"B{row}"] = prepared_by
    row += 2

    npqp_steps = [
        ("Concern Details", concern_desc, concern_image_name),
        ("Similar Part Consideration",
         f"Other Models: {other_models}\nGeneric Parts: {generic_parts}\nOther Colors: {other_colors}\nOpposite Hand: {opposite_hand}\nFront/Rear: {front_rear}", ""),
        ("Initial Analysis",
         f"Detected during process: {detect_process}\nDetected final: {detect_final}\nDetected prior: {detect_prior}\nOther: {detect_other}", ""),
        ("Temporary Countermeasures", temp_actions, str(temp_date)),
        ("Root Cause Analysis", root_cause, ""),
        ("Permanent Corrective Actions", perm_actions, ""),
        ("Effectiveness Verification", effect_verif, ""),
        ("Standardization", standardization, "")
    ]

    colors = ["FFF2CC", "D9EAD3"]

    for idx, (step, answer, extra) in enumerate(npqp_steps):
        fill_color = colors[idx % len(colors)]
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        ws[f"A{row}"] = step
        ws[f"A{row}"].font = Font(bold=True)
        ws[f"A{row}"].fill = fill
        ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")

        ws[f"B{row}"] = answer
        ws[f"B{row}"].font = Font(italic=True)
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = Alignment(wrap_text=True, vertical="top")

        ws[f"C{row}"] = extra
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        row += 4

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 30

    # Summary sheet
    if "Summary" not in wb.sheetnames:
        summary_ws = wb.create_sheet("Summary")
        headers = ["Report Date","Prepared By"] + [step for step, _, _ in npqp_steps]
        summary_ws.append(headers)
        for col, header in enumerate(headers, 1):
            summary_ws.cell(row=1, column=col).font = Font(bold=True)
    else:
        summary_ws = wb["Summary"]

    summary_row = [str(report_date), prepared_by] + [answer[:20]+("..." if len(answer)>20 else "") for step, answer, extra in npqp_steps]
    summary_ws.append(summary_row)
    for col in range(1, len(summary_row)+1):
        summary_ws.column_dimensions[summary_ws.cell(row=1,column=col).column_letter].width=30

    wb.save(file_name)
    st.success(f"✅ NPQP 8D Report saved in {file_name}")
    st.download_button("Download Excel Workbook", file_name, file_name)
