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
# Initialize session state safely
# ---------------------------
if "_reset_8d_session" not in st.session_state:
    st.session_state["_reset_8d_session"] = False

if "lang" not in st.session_state:
    st.session_state["lang"] = "English"
if "lang_key" not in st.session_state:
    st.session_state["lang_key"] = "en"
if "current_tab" not in st.session_state:
    st.session_state["current_tab"] = 0
if "report_date" not in st.session_state:
    st.session_state["report_date"] = datetime.datetime.today().date()

# ---------------------------
# Reset Session check
# ---------------------------
if st.session_state["_reset_8d_session"]:
    preserve_keys = ["lang", "lang_key", "current_tab", "report_date"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.session_state["_reset_8d_session"] = False
    st.experimental_rerun()

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
# Sidebar: Language & date input
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"], index=0 if st.session_state["lang"]=="English" else 1)
st.session_state["lang"] = lang
st.session_state["lang_key"] = "en" if lang=="English" else "es"

st.sidebar.date_input(
    f"Report Date / Fecha del informe",
    value=st.session_state["report_date"],
    key="report_date"
)

# ---------------------------
# Sidebar: Reset
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")
if st.sidebar.button("🔄 Reset 8D Session"):
    st.session_state["_reset_8d_session"] = True
    st.stop()

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
        "D8": "D8: Follow-up Activities",
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
        "D8": "D8: Actividades de seguimiento",
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
# NPQP 8D steps
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
# Initialize session state for steps
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("d5_sys_whys", [""]*5)
st.session_state.setdefault("d4_location", "")
st.session_state.setdefault("d4_status", "")
st.session_state.setdefault("d4_containment", "")

# ---------------------------
# D5 categories
# ---------------------------
occurrence_categories = {
    "Machine / Equipment": [
        "Mechanical failure or breakdown",
        "Calibration issues or drift",
        "Tooling or fixture wear or damage",
        "Machine parameters not optimized",
        "Improper preventive maintenance schedule",
        "Sensor malfunction or misalignment",
        "Process automation fault not detected",
        "Unstable process due to poor machine setup"
    ],
    "Material / Component": [
        "Wrong material or component delivered",
        "Supplier provided off-spec component",
        "Material quality inconsistency",
        "Incorrect handling/storage of material",
        "Material contamination",
        "Defective incoming components",
        "Substituted or counterfeit components"
    ],
    "Method / Process": [
        "Incorrect process sequence",
        "Operator skipped steps",
        "Improper setup or changeover",
        "Process parameters not followed",
        "Inefficient or outdated procedures",
        "Unclear work instructions"
    ],
    "Man / Operator": [
        "Operator not trained adequately",
        "Operator fatigue or distraction",
        "Human error or oversight",
        "Failure to follow SOP"
    ],
    "Measurement / Inspection": [
        "Gauge calibration error",
        "Inspection missed defect",
        "Measurement method unsuitable",
        "Misinterpretation of results"
    ]
}

detection_categories = occurrence_categories.copy()
systemic_categories = occurrence_categories.copy()

# ---------------------------
# Function to render whys
# ---------------------------
def render_whys_no_repeat(session_list, categories, title):
    selected_whys = []
    for i in range(5):
        cat_key = f"{title}_{i}"
        st.selectbox(f"{title} Why {i+1}", [""] + [item for sublist in categories.values() for item in sublist], key=cat_key)
        selected_whys.append(st.session_state[cat_key])
    return selected_whys

# ---------------------------
# Function to generate Excel
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"
    ws["A1"] = "8D Report"
    ws["A2"] = f"Report Date: {st.session_state['report_date']}"
    ws["A3"] = f"Prepared By: {st.session_state['prepared_by']}"
    row = 5
    for step, desc, example in npqp_steps:
        ws[f"A{row}"] = t[st.session_state['lang_key']][step]
        ws[f"B{row}"] = st.session_state[step]["answer"]
        row += 2
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 80
    stream = io.BytesIO()
    wb.save(stream)
    return stream

# ---------------------------
# Tabs
# ---------------------------
tabs = [t[st.session_state['lang_key']][step] for step, _, _ in npqp_steps]
tab_objects = st.tabs(tabs)

for idx, (step, desc, example) in enumerate(npqp_steps):
    with tab_objects[idx]:
        st.markdown(f"**{desc[st.session_state['lang_key']]}**")
        st.session_state[step]["answer"] = st.text_area(f"{t[st.session_state['lang_key']][step]}", value=st.session_state[step]["answer"], height=150)

# ---------------------------
# Prepared by
# ---------------------------
st.session_state["prepared_by"] = st.text_input(t[st.session_state['lang_key']]["Prepared_By"], value=st.session_state["prepared_by"])

# ---------------------------
# Save & Download
# ---------------------------
st.markdown("---")
col1, col2 = st.columns(2)
with col1:
    if st.button(t[st.session_state['lang_key']]["Save"]):
        with open("8d_report.json", "w") as f:
            json.dump({k: st.session_state[k] for k in st.session_state.keys() if k not in ["_reset_8d_session"]}, f)
        st.success("✅ Report saved successfully!")

with col2:
    if st.button(t[st.session_state['lang_key']]["Download"]):
        excel_stream = generate_excel()
        st.download_button(label="📥 Download XLSX", data=excel_stream, file_name="8D_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown("<p style='text-align:center; font-size:12px; color:#555555;'>End of 8D Report Assistant</p>", unsafe_allow_html=True)
