import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime

# ---------------------------
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide default Streamlit menu, header, footer
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Language toggle (fixed)
# ---------------------------
lang_choice = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
lang = "en" if lang_choice == "English" else "es"

# ---------------------------
# Language dictionary
# ---------------------------
texts = {
    "en": {
        "header": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_report": "💾 Save 8D Report",
        "download": "📥 Download XLSX",
        "no_answers": "⚠️ No answers filled in yet. Please complete some fields before saving.",
        "ai_helper": "💡 AI Helper Suggestions"
    },
    "es": {
        "header": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado Por",
        "save_report": "💾 Guardar Reporte 8D",
        "download": "📥 Descargar XLSX",
        "no_answers": "⚠️ No se han completado respuestas. Por favor complete algunos campos antes de guardar.",
        "ai_helper": "💡 Sugerencias del Asistente AI"
    }
}
t = texts[lang]

# ---------------------------
# Custom header
# ---------------------------
st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['header']}</h1>", unsafe_allow_html=True)

# ---------------------------
# Existing 8D steps
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
     "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected). Add more Whys if needed.",
     ""),
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
st.session_state.setdefault("interactive_whys", [""])
st.session_state.setdefault("interactive_root_cause", "")

# Color dictionary for Excel
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)": "D3D3D3",
    "Interactive 5-Why": "FFE4B5"
}

# ---------------------------
# Report Information
# ---------------------------
st.subheader("Report Information")
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input(t["report_date"], value=today_str)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# AI Helper functions
# ---------------------------
def ai_suggestion_for_next_why(prev_answer):
    if "solder" in prev_answer.lower():
        return "Was the soldering process performed correctly?"
    return "Why did this happen?"

def ai_root_cause_summary(whys_list):
    combined = " | ".join([w for w in whys_list if w.strip()])
    if combined.strip() == "":
        return ""
    return f"AI analysis of root cause based on: {combined}"

# ---------------------------
# Tabs
# ---------------------------
tab_current, tab_5why, tab_ai = st.tabs([
    "Current Features",
    "Interactive 5-Why",
    t["ai_helper"]
])

# ---------------------------
# Current Features Tab
# ---------------------------
with tab_current:
    st.info("✅ Current working 8D features remain unchanged.")
    st.markdown("### Existing 8D Inputs Here (D1-D8 tabs and fields)")

# ---------------------------
# Interactive 5-Why Tab
# ---------------------------
with tab_5why:
    st.header("Interactive 5-Why Analysis (AI-Powered)")

    for idx, val in enumerate(st.session_state.interactive_whys):
        placeholder = st.empty()
        st.session_state.interactive_whys[idx] = placeholder.text_input(
            f"Why {idx+1}", value=val, key=f"interactive_why_{idx}"
        )
        if val.strip() != "":
            suggestion = ai_suggestion_for_next_why(val)
            st.markdown(f"*Suggested next question: {suggestion}*")

    if st.button("➕ Add another Why", key="add_dynamic_why"):
        st.session_state.interactive_whys.append("")

    root_cause = ai_root_cause_summary(st.session_state.interactive_whys)
    st.session_state.interactive_root_cause = root_cause
    st.text_area("AI Suggested Root Cause", value=root_cause, height=150)

# ---------------------------
# AI Helper Tab
# ---------------------------
with tab_ai:
    st.header(t["ai_helper"])
    st.info("This tab can provide additional AI guidance or corrective action suggestions.")
    if st.button("Generate AI Suggestions"):
        whys_text = "\n".join([w for w in st.session_state.interactive_whys if w.strip()])
        st.text_area("AI Suggestions", f"AI would analyze the following 5-Whys:\n{whys_text}\n\nand provide recommendations here.")

# ---------------------------
# Save Button (Excel Export)
# ---------------------------
if st.button(t["save_report"]):
    data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]
    data_rows.append(("Interactive 5-Why", "\n".join([w for w in st.session_state.interactive_whys if w.strip()]),
                      st.session_state.interactive_root_cause))

    if not any(ans for _, ans, _ in data_rows):
        st.error(t["no_answers"])
    else:
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
        headers = ["Step", "Your Answer", "Root Cause"]
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
