import streamlit as st
import csv
from openpyxl import Workbook

st.title("📑 Nissan NPQP 8D Report Trainer")

# -------------------------------------------------------------------
# Define NPQP 8D steps: (Step Title, Description, Sample Example)
# -------------------------------------------------------------------
npqp_steps = [
    ("D0: Prepare and Plan",
     "Define the problem clearly and plan resources. This is the stage where you set the scope and urgency.",
     "Example: Customer complaint from Nissan for static noise in amplifier during end-of-line testing."),
    
    ("D1: Establish Team",
     "Form a cross-functional team with the knowledge, time, and authority to solve the problem.",
     "Example: SQE, Design Engineer, Manufacturing Engineer, Supplier Representative."),
    
    ("D2: Describe the Problem",
     "Use 5W2H (What, Where, When, Why, How, How many). Be specific and measurable.",
     "Example: 200 radios failed in Plant A during functional test due to distorted audio."),
    
    ("D3: Implement Containment",
     "Protect the customer immediately while you investigate. Containment is temporary, not the final fix.",
     "Example: Implement 100% inspection of amplifier boards before shipment."),
    
    ("D4: Identify Root Cause",
     "Use root cause analysis tools (5-Why, Fishbone, data analysis). Identify both Occurrence and Detection causes.",
     "Example: Root cause traced to cold solder joint on DSP chip caused by insufficient heating profile."),
    
    ("D5: Choose Permanent Actions",
     "Define corrective actions that eliminate the root cause permanently.",
     "Example: Update soldering process parameters, retrain operators, and improve solder paste specification."),
    
    ("D6: Implement and Validate",
     "Put corrective actions in place and verify they solve the problem long-term.",
     "Example: Run accelerated life tests on corrected amplifiers to confirm no solder failures."),
    
    ("D7: Prevent Recurrence",
     "Update standards, procedures, training, and FMEAs to prevent the same issue in future.",
     "Example: Add automated solder inspection camera, update work instructions and PFMEA."),
    
    ("D8: Recognize the Team",
     "Celebrate success and acknowledge the team’s contribution.",
     "Example: Share results with management and recognize all engineers and operators involved.")
]

# -------------------------------------------------------------------
# Initialize session state
# -------------------------------------------------------------------
if "answers" not in st.session_state:
    st.session_state.answers = {}
if "extra_info" not in st.session_state:
    st.session_state.extra_info = {}
if "report_date" not in st.session_state:
    st.session_state.report_date = ""
if "prepared_by" not in st.session_state:
    st.session_state.prepared_by = ""

# -------------------------------------------------------------------
# Report header
# -------------------------------------------------------------------
st.subheader("Report Information")
st.session_state.report_date = st.text_input("📅 Report Date (YYYY-MM-DD)", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input("✍️ Prepared By", value=st.session_state.prepared_by)

# -------------------------------------------------------------------
# Form sections
# -------------------------------------------------------------------
st.subheader("NPQP 8D Steps")
for step, desc, example in npqp_steps:
    st.markdown(f"### {step}")
    st.info(f"**Training Note:** {desc}")
    st.write(f"💡 **Example:** {example}")

    st.session_state.answers[step] = st.text_area(
        f"Your Answer for {step}",
        value=st.session_state.answers.get(step, ""),
        key=f"ans_{step}"
    )

    st.session_state.extra_info[step] = st.text_area(
        f"Extra Information (optional) for {step}",
        value=st.session_state.extra_info.get(step, ""),
        key=f"extra_{step}"
    )

# -------------------------------------------------------------------
# Collect answers safely
# -------------------------------------------------------------------
data_rows = []
for step, _, _ in npqp_steps:
    ans = st.session_state.answers.get(step, "")
    extra = st.session_state.extra_info.get(step, "")
    data_rows.append((step, ans, extra))

# -------------------------------------------------------------------
# Save button
# -------------------------------------------------------------------
if st.button("💾 Save Full 8D Report"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
    else:
        # --- XLSX file ---
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        row = 1
        ws[f"A{row}"] = "NPQP 8D Report"
        row += 2
        ws[f"A{row}"] = "Report Date"
        ws[f"B{row}"] = str(st.session_state.report_date)
        row += 1
        ws[f"A{row}"] = "Prepared By"
        ws[f"B{row}"] = st.session_state.prepared_by
        row += 2

        ws[f"A{row}"] = "Step"
        ws[f"B{row}"] = "Answer"
        ws[f"C{row}"] = "Extra Info"
        row += 1

        for step, ans, extra in data_rows:
            ws[f"A{row}"] = step
            ws[f"B{row}"] = ans
            ws[f"C{row}"] = str(extra)
            row += 1

        wb.save(xlsx_file)

        # --- CSV file ---
        csv_file = "NPQP_8D_Report.csv"
        with open(csv_file, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["NPQP 8D Report"])
            writer.writerow(["Report Date", st.session_state.report_date])
            writer.writerow(["Prepared By", st.session_state.prepared_by])
            writer.writerow([])
            writer.writerow(["Step", "Answer", "Extra Info"])
            writer.writerows(data_rows)

        # --- Downloads ---
        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX (desktop)", f, file_name=xlsx_file)
        with open(csv_file, "rb") as f:
            st.download_button("📥 Download CSV (iPhone-friendly)", f, file_name=csv_file)
