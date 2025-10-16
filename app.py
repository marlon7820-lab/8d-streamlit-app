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
# Main title
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Version info
# ---------------------------
version_number = "v1.0.9"
last_updated = "October 10, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'> 
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'> Version {version_number} | Last updated: {last_updated} </p> 
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar: Language selection & reset
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

# Language selection
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# Sidebar: JSON Backup / Restore + Safe Reset
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")

# Safe Reset Button
if st.sidebar.button("🧹 Reset Session"):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}

    # Delete only user-input keys (avoid Streamlit internal keys starting with "_")
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and not key.startswith("_"):
            del st.session_state[key]

    # Restore preserved keys
    for k, v in preserved.items():
        st.session_state[k] = v

    st.experimental_rerun()

# JSON Backup
def generate_json():
    save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
    return json.dumps(save_data, indent=4)

st.download_button(
    label="💾 Save Progress (JSON)",
    data=generate_json(),
    file_name=f"8D_Report_Backup_{st.session_state.get('report_date', datetime.datetime.today().strftime('%B_%d_%Y'))}.json",
    mime="application/json"
)

# JSON Restore
uploaded_file = st.file_uploader("Upload JSON file to restore", type="json")
if uploaded_file:
    try:
        restore_data = json.load(uploaded_file)
        for k, v in restore_data.items():
            st.session_state[k] = v
        st.success("✅ Session restored from JSON!")
    except Exception as e:
        st.error(f"Error restoring JSON: {e}")

# ---------------------------
# Language dictionary
# ---------------------------
t = {
    "en": {
        "D1":"D1: Concern Details","D2":"D2: Similar Part Considerations",
        "D3":"D3: Initial Analysis","D4":"D4: Implement Containment",
        "D5":"D5: Final Analysis","D6":"D6: Permanent Corrective Actions",
        "D7":"D7: Countermeasure Confirmation","D8":"D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date":"Report Date","Prepared_By":"Prepared By",
        "Root_Cause_Occ":"Root Cause (Occurrence)","Root_Cause_Det":"Root Cause (Detection)","Root_Cause_Sys":"Root Cause (Systemic)",
        "Occurrence_Why":"Occurrence Why","Detection_Why":"Detection Why","Systemic_Why":"Systemic Why",
        "Save":"💾 Save 8D Report","Download":"📥 Download XLSX",
        "Training_Guidance":"Training Guidance","Example":"Example",
        "FMEA_Failure":"FMEA Failure Occurrence","Location":"Material Location","Status":"Activity Status","Containment_Actions":"Containment Actions"
    },
    "es": {
        "D1":"D1: Detalles de la preocupación","D2":"D2: Consideraciones de partes similares",
        "D3":"D3: Análisis inicial","D4":"D4: Implementar contención",
        "D5":"D5: Análisis final","D6":"D6: Acciones correctivas permanentes",
        "D7":"D7: Confirmación de contramedidas","D8":"D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date":"Fecha del informe","Prepared_By":"Preparado por",
        "Root_Cause_Occ":"Causa raíz (Ocurrencia)","Root_Cause_Det":"Causa raíz (Detección)","Root_Cause_Sys":"Causa raíz (Sistémica)",
        "Occurrence_Why":"Por qué Ocurrencia","Detection_Why":"Por qué Detección","Systemic_Why":"Por qué Sistémico",
        "Save":"💾 Guardar Informe 8D","Download":"📥 Descargar XLSX",
        "Training_Guidance":"Guía de Entrenamiento","Example":"Ejemplo",
        "FMEA_Failure":"Ocurrencia de falla FMEA","Location":"Ubicación del material","Status":"Estado de la actividad","Containment_Actions":"Acciones de contención"
    }
}

# ---------------------------
# NPQP 8D steps with examples
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."},
     {"en":"Customer reported static noise in amplifier during end-of-line test.", "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.", "es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."},
     {"en":"Similar model radio, Front vs. rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.", "es":"Realice una investigación inicial para identificar problemas evidentes."},
     {"en":"Visual inspection of solder joints, initial functional tests.", "es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions and material location.", "es":"Defina acciones de contención temporales y ubicación del material."}, {"en":"","es":""}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause.", "es":"Use el análisis de 5 Porqués para determinar la causa raíz."}, {"en":"","es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently.", "es":"Defina acciones correctivas que eliminen la causa raíz permanentemente."},
     {"en":"Update soldering process, redesign fixture.", "es":"Actualizar proceso de soldadura, rediseñar herramienta."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue.", "es":"Verifique que las acciones correctivas resuelvan efectivamente el problema."},
     {"en":"Functional tests on corrected amplifiers.", "es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned, update standards, FMEAs.", "es":"Documente lecciones aprendidas, actualice estándares, FMEAs."},
     {"en":"Update SOPs, PFMEA, work instructions.", "es":"Actualizar SOPs, PFMEA, instrucciones de trabajo."})
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

# ---------------------------
# Helper: Suggest root cause
# ---------------------------
def suggest_root_cause(whys):
    text = " ".join(whys).lower()
    if any(word in text for word in ["training", "knowledge", "human error"]):
        return "Lack of proper training / knowledge gap"
    if any(word in text for word in ["equipment", "tool", "machine", "fixture"]):
        return "Equipment, tooling, or maintenance issue"
    if any(word in text for word in ["procedure", "process", "standard"]):
        return "Procedure or process not followed or inadequate"
    if any(word in text for word in ["communication", "information", "handover"]):
        return "Poor communication or unclear information flow"
    if any(word in text for word in ["material", "supplier", "component", "part"]):
        return "Material, supplier, or logistics-related issue"
    if any(word in text for word in ["design", "specification", "drawing"]):
        return "Design or engineering issue"
    if any(word in text for word in ["management", "supervision", "resource"]):
        return "Management or resource-related issue"
    if any(word in text for word in ["temperature", "humidity", "contamination", "environment"]):
        return "Environmental or external factor"
    return "Systemic issue identified from analysis"

# ---------------------------
# The rest of your original tabs rendering, dynamic 5-Why, and Excel export
# should be pasted here exactly as in your original code
# ---------------------------

# ---------------------------
# Example: Render tabs (simplified)
# ---------------------------
tab_labels = [f"{t[lang_key][step]}" for step, _, _ in npqp_steps]
tabs = st.tabs(tab_labels)
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        st.text_area("Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")

# ---------------------------
# Example Excel download
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"
    ws.append(["Step", "Answer"])
    for step, _, _ in npqp_steps:
        ws.append([t[lang_key][step], st.session_state[step]["answer"]])
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
