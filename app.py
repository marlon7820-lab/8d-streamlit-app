import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from datetime import date
import os

st.set_page_config(page_title="NPQP 8D Training App", layout="wide")
st.title("📋 Nissan NPQP 8D Training App - iPhone-Friendly")
st.write("Step-by-step guided 8D form. The Excel file will always open correctly on iPhone.")

# Session state initialization
if "step_index" not in st.session_state:
    st.session_state.step_index = 0
if "answers" not in st.session_state:
    st.session_state.answers = {}
if "extra_info" not in st.session_state:
    st.session_state.extra_info = {}
if "report_date" not in st.session_state:
    st.session_state.report_date = date.today()
if "prepared_by" not in st.session_state:
    st.session_state.prepared_by = ""

# Report info
st.session_state.report_date = st.date_input("Report Date", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input("Prepared By", value=st.session_state.prepared_by)

# NPQP 8D steps
npqp_steps = [
    ("Concern Details",
     "Amplifier unit fails functional testing due to intermittent signal loss. Image: amp_fail_01.jpg",
     "Describe the issue, affected parts, symptoms. Optional: attach image filename."),
    ("Similar Part Consideration",
     "Other Models: No, Generic Parts: Yes, Other Colors: No, Opposite Hand: No, Front/Rear: No",
     "Indicate if other parts/models/colors/sides may be affected."),
    ("Initial Analysis",
     "Detected during process: Yes, Final inspection: No, Prior to dispatch: No, Other: Functional test not performed.",
     "Identify where defect could have been detected."),
    ("Temporary Countermeasures",
     "Segregated defective units, stopped shipment. Date: 08/01/2025",
     "List immediate containment actions and implementation date."),
    ("Root Cause Analysis",
     "Capacitor C23 on amplifier PCB defective due to supplier batch #A23",
     "Describe the fundamental cause clearly."),
    ("Permanent Corrective Actions",
     "Replaced defective capacitors, updated inspection process, revised SOP",
     "List long-term fixes to prevent recurrence."),
    ("Effectiveness Verification",
     "Functional test on 50 units passed, no customer complaints",
     "Describe verification tests/audits."),
    ("Standardization",
     "Updated assembly SOP, added capacitor check to QA checklist, trained staff",
     "Document procedure changes and training.")
]

# Current step
step_name, sample, instructions = npqp_steps[st.session_state.step_index]
st.markdown(f"### Step {st.session_state.step_index + 1}: {step_name}")

if st.checkbox("Show guidance for this step?"):
    st.write(f"**Instructions:** {instructions}")
    st.write(f"**Sample Answer:** {sample}")

# Input handling per step
if step_name == "Concern Details":
    answer = st.text_area("Your answer:", height=100, placeholder=sample)
    extra = st.text_input("Image filename or URL (optional)")
elif step_name == "Temporary Countermeasures":
    answer = st.text_area("Your answer:", height=80, placeholder=sample)
    extra = st.date_input("Implementation Date", value=date.today())
elif step_name == "Similar Part Consideration":
    other_models = st.selectbox("Other Models Affected?", ["No","Yes"], key="sim1")
    generic_parts = st.selectbox("Generic Parts Affected?", ["No","Yes"], key="sim2")
    other_colors = st.selectbox("Other Colors?", ["No","Yes"], key="sim3")
    opposite_hand = st.selectbox("Opposite Hand?", ["No","Yes"], key="sim4")
    front_rear = st.selectbox("Front/Rear?", ["No","Yes"], key="sim5")
    answer = f"Other Models: {other_models}\nGeneric Parts: {generic_parts}\nOther Colors: {other_colors}\nOpposite Hand: {opposite_hand}\nFront/Rear: {front_rear}"
    extra = ""
elif step_name == "Initial Analysis":
    detect_process = st.selectbox("Detected during process?", ["No","Yes"], key="init1")
    detect_final = st.selectbox("Detected at final inspection?", ["No","Yes"], key="init2")
    detect_prior = st.selectbox("Detected prior to dispatch?", ["No","Yes"], key="init3")
    detect_other = st.text_input("Other detection points", key="init4")
    answer = f"Detected during process: {detect_process}\nDetected final: {detect_final}\nDetected prior: {detect_prior}\nOther: {detect_other}"
    extra = ""
else:
    answer = st.text_area("Your answer:", height=80, placeholder=sample)
    extra = ""

# Store input
st.session_state.answers[step_name] = answer
st.session_state.extra_info[step_name] = extra

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button("Previous Step") and st.session_state.step_index > 0:
        st.session_state.step_index -= 1
with col2:
    if st.button("Next Step") and st.session_state.step_index < len(npqp_steps)-1:
        st.session_state.step_index += 1

# Save button
if st.button("Save Full 8D Report"):
    file_name = "NPQP_8D_Reports.xlsx"

    # Use default sheet for iPhone compatibility
    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        ws = wb.active  # Always use first sheet
    else:
        wb = Workbook()
        ws = wb.active  # Default sheet will be used

    # Start writing 8D data
    row = ws.max_row + 2
    ws[f"A{row}"] = "Report Date"
    ws[f"B{row}"] = str(st.session_state.report_date)
    row += 1
    ws[f"A{row}"] = "Prepared By"
    ws[f"B{row}"] = st.session_state.prepared_by
    row += 1

    for step, *_ in npqp_steps:
        ws[f"A{row}"] = step
        ws[f"B{row}"] = st.session_state.answers.get(step,"")
        ws[f"C{row}"] = str(st.session_state.extra_info.get(step,""))
        row += 1

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 30

    # Optional Summary sheet
    if "Summary" not in wb.sheetnames:
        summary_ws = wb.create_sheet("Summary")
        headers = ["Report Date","Prepared By"] + [step for step, *_ in npqp_steps]
        summary_ws.append(headers)
        summary_ws.freeze_panes = summary_ws["A2"]
        for col, header in enumerate(headers,1):
            summary_ws.cell(row=1,column=col).font = Font(bold=True)
    else:
        summary_ws = wb["Summary"]

    summary_row = [str(st.session_state.report_date), st.session_state.prepared_by] + \
        [ (st.session_state.answers.get(step,"")[:20]+("..." if len(st.session_state.answers.get(step,""))>20 else "")) for step,_ ,_ in npqp_steps ]

    summary_ws.append(summary_row)
    for col in range(1,len(summary_row)+1):
        summary_ws.column_dimensions[summary_ws.cell(row=1,column=col).column_letter].width=30

    wb.save(file_name)
    st.success(f"✅ NPQP 8D Report saved in {file_name}")
    st.download_button("Download Excel Workbook", file_name, file_name)
