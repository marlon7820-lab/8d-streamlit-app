import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide default Streamlit UI
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📑 8D Training App</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = [
    ("D1: Concern Details",
     "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
     "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
    ("D2: Similar Part Considerations",
     "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
     "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
    ("D3: Initial Analysis",
     "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
     "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4: Implement Containment",
     "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
     "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
    ("D5: Final Analysis",
     "Interactive 5-Why for root cause. Separate Occurrence and Detection.",
     ""),  # D5 guidance
    ("D6: Permanent Corrective Actions",
     "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
     "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
    ("D7: Countermeasure Confirmation",
     "Verify that corrective actions effectively resolve the issue long-term.",
     "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
    ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
     "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
     "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future.")
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("prev_lang", "en")

# Colors for Excel
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information")
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input("📅 Report Date", value=today_str)
st.session_state.prepared_by = st.text_input("✍️ Prepared By", st.session_state.prepared_by)

# ---------------------------
# Language selection
# ---------------------------
lang = st.radio("Language / Idioma", ["en", "es"], horizontal=True)

# ---------------------------
# Manual translation function (placeholder)
# ---------------------------
def translate_text(text, target_lang="es"):
    return text  # Placeholder for later AI integration

# ---------------------------
# Tabs for each D-step
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")
        st.info(f"**Guidance:** {note}\n\n💡 **Example:** {example}")

        if step == "D5: Final Analysis":
            st.markdown("#### Occurrence Analysis (Interactive)")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=val, key=f"{step}_occ_{idx}")
            if st.button("➕ Add Occurrence Why", key=f"add_occ_{step}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### Detection Analysis (Interactive)")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", value=val, key=f"{step}_det_{idx}")
            if st.button("➕ Add Detection Why", key=f"add_det_{step}"):
                st.session_state.d5_det_whys.append("")

            # Combine answers
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area("Root Cause (summary after 5-Whys)", value=st.session_state[step]["extra"], key="root_cause")

        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", value=st.session_state[step]["answer"], key=f"ans_{step}")
            st.session_state[step]["extra"] = st.text_area(f"Extra / Root Cause", value=st.session_state[step]["extra"], key=f"extra_{step}")

# ---------------------------
# Manual Translation Buttons
# ---------------------------
st.markdown("---")
st.subheader("Translate Answers / Traducir Respuestas")
col1, col2 = st.columns(2)
with col1:
    if st.button("Translate to English / Traducir a Inglés"):
        for step, _, _ in npqp_steps:
            st.session_state[step]["answer"] = translate_text(st.session_state[step]["answer"], target_lang="en")
            st.session_state[step]["extra"] = translate_text(st.session_state[step]["extra"], target_lang="en")
        st.success("✅ All answers translated to English")

with col2:
    if st.button("Translate to Spanish / Traducir a Español"):
        for step, _, _ in npqp_steps:
            st.session_state[step]["answer"] = translate_text(st.session_state[step]["answer"], target_lang="es")
            st.session_state[step]["extra"] = translate_text(st.session_state[step]["extra"], target_lang="es")
        st.success("✅ All answers translated to Spanish")

# ---------------------------
# Prepare Excel data
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button with Excel
# ---------------------------
if st.button("💾 Save 8D Report"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet.")
    else:
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        ws.merge_cells("A1:C1")
        ws["A1"] = "Nissan NPQP 8D Report"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws["A3"] = "Report Date"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By"
        ws["B4"] = st.session_state.prepared_by

        headers = ["Step", "Your Answer", "Root Cause"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

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

        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)
        st.success("✅ NPQP 8D Report saved successfully")
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX", f, file_name=xlsx_file)
