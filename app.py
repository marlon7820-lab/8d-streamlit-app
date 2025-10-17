import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import json
import os

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

# ---------------------------
# App styles
# ---------------------------
st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Reset Session check
# ---------------------------
if st.session_state.get("_reset_8d_session", False):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    if "_reset_8d_session" in st.session_state:
        st.session_state["_reset_8d_session"] = False
    st.rerun()

# ---------------------------
# Main title
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Version info
# ---------------------------
version_number = "v1.1.0"
last_updated = "October 17, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar: Language selection & reset
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")
if st.sidebar.button("🔄 Reset 8D Session"):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys:
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.session_state["_reset_8d_session"] = True
    st.stop()

if st.session_state.get("_reset_8d_session", False):
    st.session_state["_reset_8d_session"] = False
    st.experimental_rerun()

# ---------------------------
# Language dictionary
# ---------------------------
t = {
    "en": {
        "D1": "D1: Concern Details",
        "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis",
        "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis",
        "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation",
        "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date",
        "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)",
        "Root_Cause_Det": "Root Cause (Detection)",
        "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Systemic_Why": "Systemic Why",
        "Save": "💾 Save 8D Report",
        "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance",
        "Example": "Example",
        "FMEA_Failure": "FMEA Failure Occurrence",
        "Location": "Material Location",
        "Status": "Activity Status",
        "Containment_Actions": "Containment Actions"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación",
        "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial",
        "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final",
        "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas",
        "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe",
        "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)",
        "Root_Cause_Det": "Causa raíz (Detección)",
        "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia",
        "Detection_Why": "Por qué Detección",
        "Systemic_Why": "Por qué Sistémico",
        "Save": "💾 Guardar Informe 8D",
        "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento",
        "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA",
        "Location": "Ubicación del material",
        "Status": "Estado de la actividad",
        "Containment_Actions": "Acciones de contención"
    }
}

# ---------------------------
# NPQP 8D steps with examples
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."}, {"en":"Customer reported static noise in amplifier during end-of-line test.", "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.", "es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."}, {"en":"Similar model radio, Front vs. rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.", "es":"Realice una investigación inicial para identificar problemas evidentes."}, {"en":"Visual inspection of solder joints, initial functional tests.", "es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions and material location.", "es":"Defina acciones de contención temporales y ubicación del material."}, {"en":"","es":""}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause.", "es":"Use el análisis de 5 Porqués para determinar la causa raíz."}, {"en":"","es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently.", "es":"Defina acciones correctivas que eliminen la causa raíz permanentemente."}, {"en":"Update soldering process, redesign fixture.", "es":"Actualizar proceso de soldadura, rediseñar herramienta."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue.", "es":"Verifique que las acciones correctivas resuelvan efectivamente el problema."}, {"en":"Functional tests on corrected amplifiers.", "es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned, update standards, FMEAs.", "es":"Documente lecciones aprendidas, actualice estándares, FMEAs."}, {"en":"Update SOPs, PFMEA, work instructions.", "es":"Actualizar SOPs, PFMEA, instrucciones de trabajo."})
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")

st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("d5_sys_whys", [""]*5)

st.session_state.setdefault("d4_location", "")
st.session_state.setdefault("d4_status", "")
st.session_state.setdefault("d4_containment", "")

# D6/D7 root cause notes
for step in ["D6", "D7"]:
    st.session_state.setdefault(f"{step.lower()}_occ_rc_note", "")
    st.session_state.setdefault(f"{step.lower()}_det_rc_note", "")
    st.session_state.setdefault(f"{step.lower()}_sys_rc_note", "")

# ---------------------------
# D5 helper functions
# ---------------------------
occurrence_categories = { ... } # keep your original categories
detection_categories = { ... }
systemic_categories = { ... }

def render_whys_no_repeat(whys_list, categories, label):
    for i in range(5):
        whys_list[i] = st.text_input(f"{label} Why {i+1}", value=whys_list[i])

# ---------------------------
# Tabs rendering
# ---------------------------
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
    [t[lang_key]["D1"], t[lang_key]["D2"], t[lang_key]["D3"], t[lang_key]["D4"], t[lang_key]["D5"],
     t[lang_key]["D6"], t[lang_key]["D7"], t[lang_key]["D8"]]
)

with tab5:
    st.subheader(t[lang_key]["D5"])
    st.text("Occurrence Root Cause Analysis (5-Why)")
    render_whys_no_repeat(st.session_state.d5_occ_whys, occurrence_categories, t[lang_key]["Occurrence_Why"])
    st.text("Detection Root Cause Analysis (5-Why)")
    render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories, t[lang_key]["Detection_Why"])
    st.text("Systemic Root Cause Analysis (5-Why)")
    render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories, t[lang_key]["Systemic_Why"])

with tab6:
    st.subheader(t[lang_key]["D6"])
    st.text("Occurrence Root Cause Note")
    st.session_state.d6_occ_rc_note = st.text_area("Occurrence RC Note", value=st.session_state.d6_occ_rc_note)
    st.text("Detection Root Cause Note")
    st.session_state.d6_det_rc_note = st.text_area("Detection RC Note", value=st.session_state.d6_det_rc_note)
    st.text("Systemic Root Cause Note")
    st.session_state.d6_sys_rc_note = st.text_area("Systemic RC Note", value=st.session_state.d6_sys_rc_note)

with tab7:
    st.subheader(t[lang_key]["D7"])
    st.text("Occurrence Root Cause Note")
    st.session_state.d7_occ_rc_note = st.text_area("Occurrence RC Note", value=st.session_state.d7_occ_rc_note)
    st.text("Detection Root Cause Note")
    st.session_state.d7_det_rc_note = st.text_area("Detection RC Note", value=st.session_state.d7_det_rc_note)
    st.text("Systemic Root Cause Note")
    st.session_state.d7_sys_rc_note = st.text_area("Systemic RC Note", value=st.session_state.d7_sys_rc_note)

# ---------------------------
# Excel export function
# ---------------------------
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"
    ws["A1"] = "8D Report"
    ws["A2"] = f"Report Date: {st.session_state.report_date}"
    ws["A3"] = f"Prepared By: {st.session_state.prepared_by}"
    # Insert other D1-D8 contents
    ws["A5"] = "D5 Occurrence Why"
    for i, why in enumerate(st.session_state.d5_occ_whys):
        ws[f"A{6+i}"] = why
    ws["B5"] = "D6 Occ RC Note"
    ws["B6"] = st.session_state.d6_occ_rc_note
    ws["C5"] = "D7 Occ RC Note"
    ws["C6"] = st.session_state.d7_occ_rc_note
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

st.download_button(label=t[lang_key]["Download"], data=export_excel(), file_name="8D_Report.xlsx")

st.markdown("<p style='text-align:center; font-size:12px; color:#555555;'>End of 8D Report Assistant</p>", unsafe_allow_html=True)
