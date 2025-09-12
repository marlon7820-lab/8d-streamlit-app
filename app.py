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

# Hide Streamlit default menu, header, and footer
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
is_en = lang == "English"

# ---------------------------
# Translation dictionary
# ---------------------------
t = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "Report Date",
        "prepared_by": "Prepared By",
        "save_report": "💾 Save 8D Report",
        "download_report": "📥 Download XLSX",
        "occurrence": "Occurrence Analysis",
        "detection": "Detection Analysis",
        "root_cause": "Root Cause (summary after 5-Whys)",
        "add_occ": "➕ Add another Occurrence Why",
        "add_det": "➕ Add another Detection Why"
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "Fecha del Reporte",
        "prepared_by": "Preparado Por",
        "save_report": "💾 Guardar Reporte 8D",
        "download_report": "📥 Descargar XLSX",
        "occurrence": "Análisis de Ocurrencia",
        "detection": "Análisis de Detección",
        "root_cause": "Causa Raíz (resumen después del 5-Whys)",
        "add_occ": "➕ Agregar otro Why de Ocurrencia",
        "add_det": "➕ Agregar otro Why de Detección"
    }
}[lang[:2]]

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
     "Use 5-Why analysis to determine the root cause. Occurrence and Detection separately.",
     ""),
    ("D6: Permanent Corrective Actions",
     "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
     "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
    ("D7: Countermeasure Confirmation",
     "Verify that corrective actions effectively resolve the issue long-term.",
     "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
    ("D8: Follow-up Activities",
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
st.session_state.setdefault("d5_occ_whys", [""])
st.session_state.setdefault("d5_det_whys", [""])

# Color dictionary for Excel
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input(t["report_date"], value=today_str)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")
        st.info(f"**Guidance:** {note}\n\n💡 **Example:** {example}" if note else "")

        # D5 interactive 5-Why
        if step.startswith("D5"):
            st.markdown(f"#### {t['occurrence']}")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"{t['occurrence']} Why {idx+1}", val)
                else:
                    options = ["Process issue", "Operator error", "Material defect"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"{t['occurrence']} Why {idx+1}", options, index=0)
            if st.button(t["add_occ"], key="add_occ_D5"):
                st.session_state.d5_occ_whys.append("")

            st.markdown(f"#### {t['detection']}")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"{t['detection']} Why {idx+1}", val)
                else:
                    options = ["Inspection missed it", "Checklist incomplete", "No automated test"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(f"{t['detection']} Why {idx+1}", options, index=0)
            if st.button(t["add_det"], key="add_det_D5"):
                st.session_state.d5_det_whys.append("")

            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area(t["root_cause"], st.session_state[step]["extra"], key="root_cause")
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", st.session_state[step]["answer"], key=f"ans_{step}")

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save Excel
# ---------------------------
if st.button(t["save_report"]):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
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
        ws.row_dimensions[1].height =
