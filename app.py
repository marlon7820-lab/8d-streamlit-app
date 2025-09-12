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
lang = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"], index=0)

# Translation dictionary for labels and tabs
t = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_report": "💾 Save 8D Report",
        "download_report": "📥 Download XLSX",
        "D1": "D1: Concern Details",
        "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis",
        "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis",
        "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation",
        "D8": "D8: Follow-up Activities",
        "occurrence_analysis": "#### Occurrence Analysis",
        "detection_analysis": "#### Detection Analysis",
        "root_cause": "Root Cause (summary after 5-Whys)"
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado por",
        "save_report": "💾 Guardar Reporte 8D",
        "download_report": "📥 Descargar XLSX",
        "D1": "D1: Detalles de la Preocupación",
        "D2": "D2: Consideraciones de Piezas Similares",
        "D3": "D3: Análisis Inicial",
        "D4": "D4: Implementar Contención",
        "D5": "D5: Análisis Final",
        "D6": "D6: Acciones Correctivas Permanentes",
        "D7": "D7: Confirmación de Contramedidas",
        "D8": "D8: Actividades de Seguimiento",
        "occurrence_analysis": "#### Análisis de Ocurrencia",
        "detection_analysis": "#### Análisis de Detección",
        "root_cause": "Causa Raíz (resumen después de 5-Whys)"
    }
}[lang[:2]]

st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['app_title']}</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = [
    ("D1", "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.", "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
    ("D2", "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.", "Example: Same speaker type used in another radio model; different amplifier colors."),
    ("D3", "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.", "Example: Visual inspection of solder joints, initial functional tests."),
    ("D4", "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.", "Example: 100% inspection of amplifiers before shipment; quarantine affected batches."),
    ("D5", "Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection.", ""),  # D5 interactive
    ("D6", "Define corrective actions that eliminate the root cause permanently and prevent recurrence.", "Example: Update soldering process, retrain operators."),
    ("D7", "Verify that corrective actions effectively resolve the issue long-term.", "Example: Functional tests on corrected amplifiers, monitoring first production runs."),
    ("D8", "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.", "Example: Update SOPs, PFMEA, work instructions, and employee training.")
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Color dictionary for Excel
# ---------------------------
step_colors = {
    "D1": "ADD8E6",
    "D2": "90EE90",
    "D3": "FFFF99",
    "D4": "FFD580",
    "D5": "FF9999",
    "D6": "D8BFD8",
    "D7": "E0FFFF",
    "D8": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information")
st.session_state.report_date = st.text_input(t["report_date"], value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([t[step] for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[step]}")
        st.info(f"**Training Guidance:** {note}\n\n💡 **Example:** {example}")

        # D5 interactive 5-Why
        if step == "D5":
            st.markdown(t["occurrence_analysis"])
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(
                        f"Occurrence Why {idx+1}",
                        value=st.session_state.d5_occ_whys[idx],
                        key=f"{step}_occ_{idx}"
                    )
                else:
                    prev = st.session_state.d5_occ_whys[idx - 1]
                    suggestions = [f"Follow-up based on '{prev}' #{n}" for n in range(1, 4)]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(
                        f"Occurrence Why {idx+1}",
                        options=[""] + suggestions,
                        index=0,
                        key=f"{step}_occ_{idx}"
                    )

            st.markdown(t["detection_analysis"])
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(
                        f"Detection Why {idx+1}",
                        value=st.session_state.d5_det_whys[idx],
                        key=f"{step}_det_{idx}"
                    )
                else:
                    prev = st.session_state.d5_det_whys[idx - 1]
                    suggestions = [f"Follow-up based on '{prev}' #{n}" for n in range(1, 4)]
                    st.session_state.d5_det_whys[idx] = st.selectbox(
                        f"Detection Why {idx+1}",
                        options=[""] + suggestions,
                        index=0,
                        key=f"{step}_det_{idx}"
                    )
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area(t["root_cause"], value=st.session_state[step]["extra"], key="root_cause")
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {t[step]}", value=st.session_state[step]["answer"], key=f"ans_{step}")

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button with styled Excel
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
        ws.row_dimensions[1].height = 25

        # Report info
        ws["A3"] = t["report_date"]
        ws["B3"] = st.session_state.report_date
        ws["A4"] = t["prepared_by"]
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Your Answer", "Root Cause"]
        header_fill = PatternFill(start
