import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("📑 Nissan NPQP 8D Report")

# -------------------------------------------------------------------
# NPQP 8D Steps
# -------------------------------------------------------------------
npqp_steps = [
    ("D1: Concern Details", ""),
    ("D2: Similar Part Considerations", ""),
    ("D3: Initial Analysis", ""),
    ("D4: Implement Containment", ""),
    ("D5: Root Cause", ""),  # Only step with extra info / 5-Why
    ("D6: Permanent Corrective Actions", ""),
    ("D7: Countermeasure Confirmation", ""),
    ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)", "")
]

# -------------------------------------------------------------------
# Initialize session state
# -------------------------------------------------------------------
for step, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_whys", [""] * 5)  # dynamic 5-Why for D5

# Color dictionary for Excel
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Root Cause": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)": "D3D3D3"
}

# -------------------------------------------------------------------
# Report info
# -------------------------------------------------------------------
st.subheader("Report Information")
st.session_state.report_date = st.text_input("📅 Report Date (YYYY-MM-DD)", st.session_state.report_date)
st.session_state.prepared_by = st.text_input("✍️ Prepared By", st.session_state.prepared_by)

# -------------------------------------------------------------------
# Tabs for each step
# -------------------------------------------------------------------
tabs = st.tabs([step for step, _ in npqp_steps])
for i, (step, _) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")

        # D5 Root Cause with dynamic 5-Why
        if step.startswith("D5"):
            st.markdown("#### Root Cause 5-Why Analysis")
            for idx, val in enumerate(st.session_state.d5_whys):
                st.session_state.d5_whys[idx] = st.text_input(f"Why {idx+1}", value=val, key=f"{step}_why_{idx}")
            if st.button("➕ Add another Why", key=f"add_why_{step}"):
                st.session_state.d5_whys.append("")

            # Combine 5-Why answers + extra info
            st.session_state[step]["answer"] = "\n".join([w for w in st.session_state.d5_whys if w.strip()])
            st.session_state[step]["extra"] = st.text_area("Additional Root Cause Details (optional)", value=st.session_state[step]["extra"], key="extra_rootcause")
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", value=st.session_state[step]["answer"], key=f"ans_{step}")

# -------------------------------------------------------------------
# Collect answers
# -------------------------------------------------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _ in npqp_steps]

# -------------------------------------------------------------------
# Save button with styled Excel
# -------------------------------------------------------------------
if st.button("💾 Save 8D Report"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
    else:
        # --- XLSX ---
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        # Title
        ws.merge_cells("A1:C1")
        ws["A1"] = "Nissan NPQP 8D Report"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        # Report info
        ws["A3"] = "Report Date"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By"
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Your Answer", "Extra Info"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        # Content
        row = 7
        for step, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=step)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)

            fill_color = step_colors.get(step, "FFFFFF")
            for col in range(1, 4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")

            row += 1

        # Adjust column widths
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)

        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX", f, file_name=xlsx_file)
