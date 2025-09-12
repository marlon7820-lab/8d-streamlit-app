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
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("🌐 Language / Idioma", ["English", "Español"], index=0)

# Translation dictionary for labels & headings
t = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_btn": "💾 Save 8D Report",
        "download_btn": "📥 Download XLSX",
        "occurrence_analysis": "#### Occurrence Analysis",
        "detection_analysis": "#### Detection Analysis",
        "root_cause": "Root Cause (summary after 5-Whys)",
        "add_occ": "➕ Add another Occurrence Why",
        "add_det": "➕ Add another Detection Why"
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "📅 Fecha del Informe",
        "prepared_by": "✍️ Elaborado por",
        "save_btn": "💾 Guardar Reporte 8D",
        "download_btn": "📥 Descargar XLSX",
        "occurrence_analysis": "#### Análisis de Ocurrencia",
        "detection_analysis": "#### Análisis de Detección",
        "root_cause": "Causa Raíz (resumen después de 5-Porqués)",
        "add_occ": "➕ Agregar otro Porqué de Ocurrencia",
        "add_det": "➕ Agregar otro Porqué de Detección"
    }
}[lang[:2]]

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = [
    ("D1: Concern Details",
     {"en": "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
      "es": "Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de respaldo."},
     {"en": "Example: Customer reported static noise in amplifier during end-of-line test at Plant A.",
      "es": "Ejemplo: El cliente reportó ruido estático en el amplificador durante la prueba final en Planta A."}),
    ("D2: Similar Part Considerations",
     {"en": "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.",
      "es": "Verifique partes similares, modelos, piezas genéricas, otros colores, lado opuesto, delantero/trasero, etc."},
     {"en": "Example: Same speaker type used in another radio model.",
      "es": "Ejemplo: Mismo tipo de altavoz usado en otro modelo de radio."}),
    ("D3: Initial Analysis",
     {"en": "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
      "es": "Realice una investigación inicial para identificar problemas evidentes, recopilar datos y documentar hallazgos iniciales."},
     {"en": "Example: Visual inspection of solder joints, initial functional tests, checking connectors.",
      "es": "Ejemplo: Inspección visual de soldaduras, pruebas funcionales iniciales, verificación de conectores."}),
    ("D4: Implement Containment",
     {"en": "Define temporary containment actions to prevent customer exposure while permanent actions are developed.",
      "es": "Defina acciones de contención temporales para prevenir que el cliente vea el problema mientras se desarrollan acciones permanentes."},
     {"en": "Example: 100% inspection of amplifiers before shipment.",
      "es": "Ejemplo: Inspección del 100% de los amplificadores antes del envío."}),
    ("D5: Final Analysis",
     {"en": "Use interactive 5-Why analysis for root cause (Occurrence & Detection).",
      "es": "Use análisis interactivo de 5-Porqués para causa raíz (Ocurrencia y Detección)."},
     {"en": "", "es": ""}),
    ("D6: Permanent Corrective Actions",
     {"en": "Define corrective actions that eliminate the root cause permanently.",
      "es": "Defina acciones correctivas que eliminen permanentemente la causa raíz."},
     {"en": "Example: Update soldering process, retrain operators, update work instructions.",
      "es": "Ejemplo: Actualizar el proceso de soldadura, reentrenar operadores, actualizar instrucciones de trabajo."}),
    ("D7: Countermeasure Confirmation",
     {"en": "Verify that corrective actions effectively resolve the issue long-term.",
      "es": "Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo."},
     {"en": "Example: Functional tests on corrected units.",
      "es": "Ejemplo: Pruebas funcionales en unidades corregidas."}),
    ("D8: Follow-up Activities",
     {"en": "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
      "es": "Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y capacitación para prevenir recurrencia."},
     {"en": "Example: Update SOPs and employee training.",
      "es": "Ejemplo: Actualizar SOP y capacitación de empleados."})
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
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")

        # Training guidance & example
        st.info(f"**Guidance:** {note_dict[lang[:2]]}\n\n💡 **Example:** {example_dict[lang[:2]]}")

        # D5 interactive 5-Why
        if step.startswith("D5"):
            st.markdown(t["occurrence_analysis"])
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=st.session_state.d5_occ_whys[idx], key=f"{step}_occ_{idx}")
                else:
                    # simple suggestion based on previous why (can be improved later)
                    suggestions = [f"Follow up on: {st.session_state.d5_occ_whys[idx-1]}"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"Occurrence Why {idx+1}", options=suggestions, index=0, key=f"{step}_occ_{idx}")

            st.markdown(t["detection_analysis"])
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", value=st.session_state.d5_det_whys[idx], key=f"{step}_det_{idx}")
                else:
                    suggestions = [f"Follow up on: {st.session_state.d5_det_whys[idx-1]}"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(f"Detection Why {idx+1}", options=suggestions, index=0, key=f"{step}_det_{idx}")

            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area(t["root_cause"], value=st.session_state[step]["extra"])

        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", value=st.session_state[step]["answer"], key=f"ans_{step}")

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button with styled Excel
# ---------------------------
if st.button(t["save_btn"]):
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
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill
