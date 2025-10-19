import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import os
from PIL import Image as PILImage
from io import BytesIO

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

# ---------------------------
# App styles - Light mode by default, buttons always same color
# ---------------------------
st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"], .stDownloadButton button, .stSidebar button {
    background-color: #87AFC7 !important;
    color: #000000 !important;
    font-weight: bold !important;
}
div.stSelectbox, div.stTextInput, div.stTextArea {
    border: 2px solid #1E90FF !important;
    border-radius: 5px !important;
    padding: 5px !important;
    background-color: #ffffff !important;
    transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}
div.stSelectbox:hover, div.stTextInput:hover, div.stTextArea:hover {
    border: 2px solid #104E8B !important;
    box-shadow: 0 0 5px #1E90FF;
}
.image-thumbnail {width: 120px; height: 80px; object-fit: cover; margin:5px; border:1px solid #1E90FF; border-radius:4px;}
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
    st.session_state["_reset_8d_session"] = False
    st.rerun()

# ---------------------------
# Main title
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Version info
# ---------------------------
version_number = "v1.2.0"
last_updated = "October 18, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar: Language & Dark Mode
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"], key="lang_select")
lang_key = "en" if lang == "English" else "es"
dark_mode = st.sidebar.checkbox("🌙 Dark Mode", key="dark_mode")

# ---------------------------
# Dark Mode - form only, sidebar normal
# ---------------------------
if dark_mode:
    st.markdown("""
    <style>
    .stApp { background: #2b2b2b !important; color: #e0e0e0 !important; }
    div.stTextInput, div.stTextArea, div.stSelectbox {
        border: 2px solid #87AFC7 !important;
        border-radius: 5px !important;
        background-color: #3a3a3a !important;
        color: #e0e0e0 !important;
        padding: 5px !important;
        transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
    }
    div.stTextInput:hover, div.stTextArea:hover, div.stSelectbox:hover {
        border: 2px solid #1E90FF !important;
        box-shadow: 0 0 5px #1E90FF;
    }
    .stInfo { background-color: #444444 !important; border-left: 5px solid #87AFC7 !important; color: #e0e0e0 !important; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; color: #e0e0e0 !important; }
    .stTabs [data-baseweb="tab"]:hover { color: #87AFC7 !important; }
    button[kind="primary"], .stDownloadButton button, .stSidebar button { background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# ---------------------------
# Sidebar: App Controls
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")
if st.sidebar.button("🔄 Reset 8D Session", key="reset_button"):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys:
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
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
    ("D1", {"en":"Describe the customer concerns clearly.","es":"Describa claramente las preocupaciones del cliente."},
     {"en":"Customer reported static noise in amplifier during end-of-line test.","es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.","es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."},
     {"en":"Similar model radio, Front vs. rear speaker.","es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.","es":"Realice una investigación inicial para identificar problemas evidentes."},
     {"en":"Visual inspection of solder joints, initial functional tests.","es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions and material location.","es":"Defina acciones de contención temporales y ubicación del material."},
     {"en":"Post Quality Alert, Increase Inspection, Inventory Certification","es":"Implementar Ayuda Visual, Incrementar Inspeccion, Certificar Inventario"}),
    ("D5", {"en": "Use 5-Why analysis to determine the root cause.","es": "Use el análisis de 5 Porqués para determinar la causa raíz."},
     {"en": "Final 'Why' from the Analysis will give a good indication of the True Root Cause","es": "El último \"Por qué\" del análisis proporcionará una idea clara de la causa raíz del problema"}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently.","es":"Defina acciones correctivas que eliminen la causa raíz permanentemente."},
     {"en":"Update soldering process, redesign fixture.","es":"Actualizar proceso de soldadura, rediseñar herramienta."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue.","es":"Verifique que las acciones correctivas resuelvan efectivamente el problema."},
     {"en":"Functional tests on corrected amplifiers.","es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned, update standards, FMEAs.","es":"Documente lecciones aprendidas, actualice estándares, FMEAs."},
     {"en":"Update SOPs, PFMEA, work instructions.","es":"Actualizar SOPs, PFMEA, instrucciones de trabajo."})
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
    if step in ["D1","D3","D4","D7"]:
        st.session_state[step]["uploaded_files"] = []

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("d5_sys_whys", [""]*5)
st.session_state.setdefault("d5_occ_whys_free", [""]*0)
st.session_state.setdefault("d5_det_whys_free", [""]*0)
st.session_state.setdefault("d5_sys_whys_free", [""]*0)
st.session_state.setdefault("d4_location", "")
st.session_state.setdefault("d4_status", "")
st.session_state.setdefault("d4_containment", "")

for sub in ["occ_answer", "det_answer", "sys_answer"]:
    st.session_state.setdefault(("D6"), st.session_state.get("D6", {}))
    st.session_state["D6"].setdefault(sub, "")
    st.session_state.setdefault(("D7"), st.session_state.get("D7", {}))
    st.session_state["D7"].setdefault(sub, "")

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
        "Wrong material or component used",
        "Supplier provided off-spec component",
        "Material defect not visible during inspection",
        "Damage during storage, handling, or transport",
        "Incorrect labeling, Missing label or lot traceability error",
        "Material substitution without approval",
        "Incorrect specifications or revision mismatch"
    ],
    "Process / Method": [
        "Incorrect process step sequence",
        "Critical process parameters not controlled",
        "Work instructions unclear or missing details",
        "Process drift over time not detected",
        "Control plan not followed on production floor",
        "Incorrect torque, solder, or assembly process",
        "Outdated or missing process FMEA linkage",
        "Inadequate process capability (Cp/Cpk below target)"
    ],
    "Design / Engineering": [
        "Design not robust to real-use conditions",
        "Tolerance stack-up issue not evaluated",
        "Late design change not communicated to production",
        "Incorrect or unclear drawing specification",
        "Component placement design error (DFMEA gap)",
        "Lack of design verification or validation testing"
    ],
    "Environmental / External": [
        "Temperature or humidity out of control range",
        "Electrostatic discharge (ESD) not controlled",
        "Contamination or dust affecting product",
        "Power fluctuation or interruption",
        "External vibration or noise interference",
        "Unstable environmental monitoring process"
    ]
}

detection_categories = {
    "QA / Inspection": [
        "QA checklist incomplete or not updated",
        "No automated inspection system in place",
        "Manual inspection prone to human error",
        "Inspection frequency too low to detect issue",
        "Inspection criteria unclear or inconsistent",
        "Measurement system not capable (GR&R issues)",
        "Incoming inspection missed supplier issue",
        "Final inspection missed due to sampling plan"
    ],
    "Validation / Process": [
        "Process validation not updated after design/process change",
        "Insufficient verification of new parameters or components",
        "Design validation not complete or not representative of real conditions",
        "Inadequate control plan coverage for potential failure modes",
        "Lack of ongoing process monitoring (SPC / CpK tracking)",
        "Incorrect or outdated process limits not aligned with FMEA"
    ],
    "FMEA / Control Plan": [
        "Failure mode not captured in PFMEA",
        "Detection controls missing or ineffective in PFMEA",
        "Control plan not updated after corrective actions",
        "FMEA not reviewed after customer complaint",
        "Detection ranking not realistic to actual inspection capability",
        "PFMEA and control plan not properly linked"
    ],
    "Test / Equipment": [
        "Test equipment calibration overdue",
        "Testing software parameters incorrect",
        "Test setup does not detect this specific failure mode",
        "Detection threshold too wide to capture failure",
        "Test data not logged or reviewed regularly"
    ],
    "Systemic / Organizational": [
        "Feedback loop from quality incidents not implemented",
        "Lack of detection feedback in regular team meetings",
        "Training gaps in inspection or test personnel",
        "Quality alerts not properly communicated to operators"
    ]
}

systemic_categories = {
    "Management / Organization": [
        "Inadequate leadership or supervision structure",
        "Insufficient resource allocation to critical processes",
        "Delayed response to known production issues",
        "Lack of accountability or ownership of quality issues",
        "Ineffective escalation process for recurring problems",
        "Weak cross-functional communication between departments"
    ],
    "Process / Procedure": [
        "Standard Operating Procedures (SOPs) outdated or missing",
        "Process FMEA not reviewed regularly",
        "Control plan not aligned with PFMEA or actual process",
        "Lessons learned not integrated into similar processes",
        "Inefficient document control system",
        "Preventive maintenance procedures not standardized"
    ],
    "Training": [
        "No defined training matrix or certification tracking",
        "New hires not trained on critical control points",
        "Training effectiveness not evaluated",
        "Knowledge not shared between shifts or teams",
        "Competence requirements not clearly defined"
    ],
    "Supplier / External": [
        "Supplier not included in 8D or FMEA review process",
        "Supplier corrective actions not verified for effectiveness",
        "Inadequate incoming material audit process",
        "Supplier process changes not communicated to customer",
        "Long lead time for supplier quality issue closure",
        "Supplier violation of cleanpoint"
    ],
    "Quality System / Feedback": [
        "Internal audits ineffective or not completed",
        "Quality KPI tracking not linked to root cause analysis",
        "Ineffective use of 5-Why or other problem solving tools",
        "Customer complaints not feeding back into design reviews",
        "No systemic review after multiple 8Ds in same area"
    ]
}

# ---------------------------
# Root cause suggestion & helper functions
# ---------------------------
def suggest_root_cause(whys):
    text = " ".join(whys).lower()
    if any(word in text for word in ["training", "knowledge", "human error"]):
        return "The root cause may be attributed to insufficient training or a knowledge gap"
    if any(word in text for word in ["equipment", "tool", "machine", "fixture"]):
        return "The root cause may be attributed to equipment, tooling, or maintenance issue"
    if any(word in text for word in ["procedure", "process", "standard"]):
        return "The root cause may be attributed to procedure or process not followed or inadequate"
    if any(word in text for word in ["communication", "information", "handover"]):
        return "The root cause may be attributed to poor communication or unclear information flow"
    if any(word in text for word in ["material", "supplier", "component", "part"]):
        return "The root cause may be attributed to material, supplier, or logistics-related issue"
    if any(word in text for word in ["design", "specification", "drawing"]):
        return "The root cause may be attributed to design or engineering issue"
    if any(word in text for word in ["management", "supervision", "resource"]):
        return "The root cause may be attributed management or resource-related issue"
    if any(word in text for word in ["temperature", "humidity", "contamination", "environment"]):
        return "The root cause may be attributed to environmental or external factor"
    return "No clear root cause suggestion (provide more 5-Whys)"
# ---------------------------
# Helper: render 5-Whys section
# ---------------------------
def render_whys(whys_list, category_name, categories):
    st.markdown(f"**{category_name}**")
    for i in range(5):
        col1, col2 = st.columns([1,3])
        with col1:
            st.text_input(f"Why {i+1}", key=f"{category_name}_{i}", value=whys_list[i])
        with col2:
            st.selectbox("Select category", [""] + categories, key=f"{category_name}_cat_{i}", index=0)

# ---------------------------
# Tabs for D1-D8
# ---------------------------
tabs = st.tabs([t[lang_key] for t in ["D1","D2","D3","D4","D5","D6","D7","D8"]])

# ---------------------------
# D1
# ---------------------------
with tabs[0]:
    st.header("D1: Concern Details")
    st.text_input("Prepared By", key="prepared_by")
    st.date_input("Report Date", key="report_date")
    st.text_area("Customer Concern / Problem Description", key="D1_answer", height=120)
    uploaded_files = st.file_uploader("Upload evidence images", type=["png","jpg","jpeg"], accept_multiple_files=True, key="D1_upload")
    st.session_state["D1"]["uploaded_files"] = uploaded_files

# ---------------------------
# D2
# ---------------------------
with tabs[1]:
    st.header("D2: Similar Part Considerations")
    st.text_area("Similar Parts / Models / Generic Considerations", key="D2_answer", height=120)

# ---------------------------
# D3
# ---------------------------
with tabs[2]:
    st.header("D3: Initial Analysis")
    st.text_area("Initial Analysis / Findings", key="D3_answer", height=120)
    uploaded_files = st.file_uploader("Upload evidence images", type=["png","jpg","jpeg"], accept_multiple_files=True, key="D3_upload")
    st.session_state["D3"]["uploaded_files"] = uploaded_files

# ---------------------------
# D4
# ---------------------------
with tabs[3]:
    st.header("D4: Implement Containment")
    st.text_area("Temporary Containment Actions", key="D4_containment", height=80)
    st.text_input("Material Location", key="d4_location")
    st.selectbox("Activity Status", ["Open","Closed","In Progress"], key="d4_status")
    uploaded_files = st.file_uploader("Upload evidence images", type=["png","jpg","jpeg"], accept_multiple_files=True, key="D4_upload")
    st.session_state["D4"]["uploaded_files"] = uploaded_files

# ---------------------------
# D5
# ---------------------------
with tabs[4]:
    st.header("D5: Final Analysis (5-Whys)")
    st.markdown("**Occurrence Root Cause**")
    for i in range(5):
        st.text_input(f"Occurrence Why {i+1}", key=f"d5_occ_{i}", value=st.session_state["d5_occ_whys"][i])
    st.markdown("**Detection Root Cause**")
    for i in range(5):
        st.text_input(f"Detection Why {i+1}", key=f"d5_det_{i}", value=st.session_state["d5_det_whys"][i])
    st.markdown("**Systemic Root Cause**")
    for i in range(5):
        st.text_input(f"Systemic Why {i+1}", key=f"d5_sys_{i}", value=st.session_state["d5_sys_whys"][i])
    st.markdown("**Suggested Occurrence Root Cause**")
    st.info(suggest_root_cause([st.session_state[f"d5_occ_{i}"] for i in range(5)]))
    st.markdown("**Suggested Detection Root Cause**")
    st.info(suggest_root_cause([st.session_state[f"d5_det_{i}"] for i in range(5)]))
    st.markdown("**Suggested Systemic Root Cause**")
    st.info(suggest_root_cause([st.session_state[f"d5_sys_{i}"] for i in range(5)]))

# ---------------------------
# D6
# ---------------------------
with tabs[5]:
    st.header("D6: Permanent Corrective Actions")
    st.text_area("Corrective Actions (Occurrence)", key="D6_occ_answer", height=100)
    st.text_area("Corrective Actions (Detection)", key="D6_det_answer", height=100)
    st.text_area("Corrective Actions (Systemic)", key="D6_sys_answer", height=100)
    uploaded_files = st.file_uploader("Upload evidence images", type=["png","jpg","jpeg"], accept_multiple_files=True, key="D6_upload")
    st.session_state["D6"]["uploaded_files"] = uploaded_files

# ---------------------------
# D7
# ---------------------------
with tabs[6]:
    st.header("D7: Countermeasure Confirmation")
    st.text_area("Effectiveness Verification (Occurrence)", key="D7_occ_answer", height=80)
    st.text_area("Effectiveness Verification (Detection)", key="D7_det_answer", height=80)
    st.text_area("Effectiveness Verification (Systemic)", key="D7_sys_answer", height=80)
    uploaded_files = st.file_uploader("Upload verification evidence", type=["png","jpg","jpeg"], accept_multiple_files=True, key="D7_upload")
    st.session_state["D7"]["uploaded_files"] = uploaded_files

# ---------------------------
# D8
# ---------------------------
with tabs[7]:
    st.header("D8: Follow-up / Lessons Learned")
    st.text_area("Lessons Learned / Recurrence Prevention", key="D8_answer", height=120)
    st.text_area("Additional Notes / Observations", key="D8_extra", height=80)
    uploaded_files = st.file_uploader("Upload evidence images", type=["png","jpg","jpeg"], accept_multiple_files=True, key="D8_upload")

# ---------------------------
# Excel Export
# ---------------------------
def create_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # Set headers
    headers = ["Step", "Description", "Answer", "Extra Notes"]
    for col_num, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_num, value=header)
        ws.cell(row=1, column=col_num).font = Font(bold=True)
        ws.cell(row=1, column=col_num).alignment = Alignment(horizontal="center")

    # Fill data
    row_num = 2
    for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
        answer = st.session_state.get(step, {}).get("answer","") if isinstance(st.session_state.get(step,{}), dict) else st.session_state.get(f"{step}_answer","")
        extra = st.session_state.get(step, {}).get("extra","") if isinstance(st.session_state.get(step,{}), dict) else st.session_state.get(f"{step}_extra","")
        ws.cell(row=row_num, column=1, value=step)
        ws.cell(row=row_num, column=2, value=t[lang_key].get(step, step))
        ws.cell(row=row_num, column=3, value=answer)
        ws.cell(row=row_num, column=4, value=extra)
        row_num += 1

    # Add images for D1, D3, D4, D6, D7
    image_steps = ["D1","D3","D4","D6","D7"]
    for step in image_steps:
        files = st.session_state.get(step, {}).get("uploaded_files", [])
        if files:
            for f in files:
                image = XLImage(BytesIO(f.read()))
                ws.add_image(image, f"D{row_num}")

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    return output

# ---------------------------
# Download Button
# ---------------------------
if st.button("📥 Download XLSX"):
    excel_file = create_excel()
    st.download_button("Download 8D Report", data=excel_file.getvalue(), file_name=f"8D_Report_{st.session_state['report_date']}.xlsx")

# ---------------------------
# End of App
# ---------------------------
