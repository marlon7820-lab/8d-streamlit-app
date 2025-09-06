import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

st.set_page_config(page_title="8D Problem Solving App", layout="wide")
st.title("📋 8D Problem Solving App")
st.write("Fill out the form below to generate a professional 8D Excel report.")

# 8D Questions
questions = [
    ("D1 - Establish the Team", "Who is on the problem-solving team?"),
    ("D2 - Describe the Problem", "What is the problem? Provide details."),
    ("D3 - Containment Actions", "What actions were taken to contain the issue?"),
    ("D4 - Root Cause", "What is the root cause of the issue?"),
    ("D5 - Corrective Actions", "What corrective actions are planned?"),
    ("D6 - Permanent Corrective Actions", "How will you implement permanent corrective actions?"),
    ("D7 - Prevent Recurrence", "What steps will prevent this from happening again?"),
    ("D8 - Congratulate the Team", "Final remarks / congratulations.")
]

# Collect answers from user
answers = {}
with st.form("8D Form"):
    for step, question in questions:
        answers[step] = st.text_area(f"{step}: {question}", height=50)
    submitted = st.form_submit_button("Generate 8D Report")

# Generate Excel report if submitted
if submitted:
    file_name = "8D_Report_Professional.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # Title formatting
    ws["A1"] = "8D Problem Solving Report"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws.merge_cells("A1:C1")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws["A1"].fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    # Fill in questions and answers
    row = 3
    for step, question in questions:
        ws[f"A{row}"] = step
        ws[f"A{row}"].font = Font(bold=True)
        ws[f"B{row}"] = question
        ws[f"B{row}"].font = Font(italic=True)
        ws[f"C{row}"] = answers.get(step, "")
        ws[f"C{row}"].alignment = Alignment(wrap_text=True)
        row += 2

    # Adjust column widths
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 80

    wb.save(file_name)
    st.success(f"✅ 8D Report generated: {file_name}")
    st.download_button("Download Excel Report", file_name, file_name)
