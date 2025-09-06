import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import date
import os

st.set_page_config(page_title="NPQP 8D Training App", layout="wide")
st.title("📋 Nissan NPQP 8D Training App")
st.write("Guided 8D form with instructions, meaning, samples, and summary tracking for training beginners.")

# Report info
report_date = st.date_input("Report Date", value=date.today())
prepared_by = st.text_input("Prepared By")

# --- Define NPQP steps with training guidance ---
npqp_steps = [
    ("Concern Details",
     "💡 Meaning: Describe the issue in detail, include affected parts, symptoms, and context.",
     "Instructions: Write exactly what went wrong, where, and when. Attach image filename if available.",
     "Sample: Amplifier unit fails functional testing due to intermittent signal loss. Image: amp_fail_01.jpg"),
    ("Similar Part Consideration",
     "💡 Meaning: Determine if other parts may have the same problem.",
     "Instructions: Use the dropdowns to indicate if other models, generic parts, colors, or sides are affected.",
     "Sample: Other Models: No, Generic Parts: Yes, Other Colors: No, Opposite Hand: No, Front/Rear: No"),
    ("Initial Analysis",
     "💡 Meaning: Identify where the defect could have been detected during manufacturing or inspection.",
     "Instructions: Use dropdowns and input any other detection points.",
     "Sample: Detected during process: Yes, Final inspection: No, Prior to dispatch: No, Other: Functional test not performed."),
    ("Temporary Countermeasures",
     "💡 Meaning: Actions to stop the problem from spreading until permanent solution is applied.",
     "Instructions: List immediate containment actions and date implemented.",
     "Sample: Segregated defective amplifier units, stopped shipment, informed production team. Date: 08/01/2025"),
    ("Root Cause Analysis",
     "💡 Meaning: Determine the fundamental cause of the defect.",
     "Instructions: Describe the investigation and findings clearly.",
     "Sample: Capacitor C23 on amplifier PCB defective due to supplier batch #A23"),
    ("Permanent Corrective Actions",
     "💡 Meaning: Implement long-term fixes to prevent recurrence.",
     "Instructions: Explain what changes are made in design, process, or supplier.",
     "Sample: Replaced all defective capacitors, updated supplier inspection process, revised assembly SOP"),
    ("Effectiveness Verification",
     "💡 Meaning: Confirm corrective actions work and problem is resolved.",
     "Instructions: Describe tests or audits performed to verify effectiveness.",
     "Sample: Functional test on 50 units passed, no complaints reported in customer feedback."),
    ("Standardization",
     "💡 Meaning: Update procedures to ensure lessons learned are applied.",
     "Instructions: Document changes to SOPs, quality checks, and training.",
     "Sample: Updated assembly SOP, added capacitor check to QA checklist, trained staff")
]

# Collect answers
answers = {}
extra_info = {}  # for image filename or date
st.subheader("Fill NPQP 8D Sections")

with st.form("8D_Form"):
    for step, meaning, instructions, sample in npqp_steps:
        with st.expander(step):
            st.write(f"**Meaning:** {meaning}")
            st.write(f"**Instructions:** {instructions}")
            st.write(f"**Sample:** {sample}")
            # Special input for Concern Details and Temporary Countermeasures
            if step == "Concern Details":
                answers[step] = st.text_area("Your answer:", height=100, placeholder=sample)
                extra_info[step] = st.text_input("Image filename or URL (optional)")
            elif step == "Temporary Countermeasures":
                answers[step] = st.text_area("Your answer:", height=80, placeholder=sample)
                extra_info[step] = st.date_input("Implementation Date", value=date.today())
            elif step == "Similar Part Consideration":
                other_models = st.selectbox("Other Models Affected?", ["No","Yes"], key=step+"1")
                generic_parts = st.selectbox("Generic Parts Affected?", ["No","Yes"], key=step+"2")
                other_colors = st.selectbox("Other Colors?", ["No","Yes"], key=step+"3")
                opposite_hand = st.selectbox("Opposite Hand?", ["No","Yes"], key=step+"4")
                front_rear = st.selectbox("Front/Rear?", ["No","Yes"], key=step+"5")
                answers[step] = f"Other Models: {other_models}\nGeneric Parts: {generic_parts}\nOther Colors: {other_colors}\nOpposite Hand: {opposite_hand}\nFront/Rear: {front_rear}"
                extra_info[step] = ""
            elif step == "Initial Analysis":
                detect_process = st.selectbox("Detected during process?", ["No","Yes"], key=step+"1")
                detect_final = st.selectbox("Detected at final inspection?", ["No","Yes"], key=step+"2")
                detect_prior = st.selectbox("Detected prior to dispatch?", ["No","Yes"], key=step+"3")
                detect_other = st.text_input("Other detection points", key=step+"4")
                answers[step] = f"Detected during process: {detect_process}\nDetected at final inspection: {detect_final}\nDetected prior to dispatch: {detect_prior}\nOther: {detect_other}"
                extra_info[step] = ""
            else:
                answers[step] = st.text_area("Your answer:", height=80, placeholder=sample)
                extra_info[step] = ""

    submitted = st.form_submit_button("Save 8D Report")

if submitted:
    file_name = "NPQP_8D_Reports.xlsx"

    if os.path.exists(file_name):
        wb = load_workbook(file_name)
        ws = wb["8D Reports"] if "8D Reports" in wb.sheetnames else wb.create_sheet("8D Reports")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "8D Reports"

    row = ws.max_row + 2

    # Report info
    ws[f"A{row}"] = "Report Date"
    ws[f"B{row}"] = str(report_date)
    row += 1
    ws[f"A{row}"] = "Prepared By"
    ws[f"B{row}"] = prepared_by
    row += 2

    colors = ["FFF2CC", "D9EAD3"]

    # Write NPQP steps with guidance
    for idx, (step, *_ ) in enumerate(npqp_steps):
        fill_color = colors[idx % 2]
        fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        ws[f"A{row}"] = step
        ws[f"A{row}"].font = Font(bold=True)
        ws[f"A{row}"].fill = fill
        ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")

        ws[f"B{row}"] = answers.get(step,"")
        ws[f"B{row}"].font = Font(italic=True)
        ws[f"B{row}"].fill = fill
        ws[f"B{row}"].alignment = Alignment(wrap_text=True, vertical="top")

        ws[f"C{row}"] = str(extra_info.get(step,""))
        ws[f"C{row}"].fill = fill
        ws[f"C{row}"].alignment = Alignment(wrap_text=True, vertical="top")

        row += 4

    # Column widths
    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 30

    # Summary sheet
    if "Summary" not in wb.sheetnames:
        summary_ws = wb.create_sheet("Summary")
        headers = ["Report Date","Prepared By"] + [step for step, *_ in npqp_steps]
        summary_ws.append(headers)
        for col, header in enumerate(headers,1):
            summary_ws.cell(row=1,column=col).font=Font(bold=True)
    else:
        summary_ws = wb["Summary"]

    summary_row = [str(report_date), prepared_by] + [answers[step][:20]+("..." if len(answers[step])>20 else "") for step,_ ,_,_ in npqp_steps]
    summary_ws.append(summary_row)
    for col in range(1,len(summary_row)+1):
        summary_ws.column_dimensions[summary_ws.cell(row=1,column=col).column_letter].width=30

    wb.save(file_name)
    st.success(f"✅ NPQP 8D Report saved in {file_name}")
    st.download_button("Download Excel Workbook", file_name, file_name)
