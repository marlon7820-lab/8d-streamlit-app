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
# Sidebar: Language selection & Reset
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

# Language selection
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# Sidebar: Report Date & Prepared By
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.text_input("Report Date / Fecha del informe", key="report_date", value=st.session_state.get("report_date", datetime.datetime.today().strftime("%B %d, %Y")))
st.sidebar.text_input("Prepared By / Preparado por", key="prepared_by", value=st.session_state.get("prepared_by", ""))

# ---------------------------
# Sidebar: Reset 8D Report (merged button)
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")
if st.sidebar.button("🧹 Reset 8D Report"):
    preserve_keys = ["lang", "lang_key", "report_date", "prepared_by"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    
    # Clear everything except preserved values
    for key in list(st.session_state.keys()):
        if key not in preserve_keys:
            del st.session_state[key]
    
    # Restore preserved values
    for k, v in preserved.items():
        st.session_state[k] = v
    
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
        "Material defect not visible during inspection",
        "Damage during storage, handling, or transport",
        "Incorrect labeling or lot traceability error",
        "Material substitution without approval",
        "Incorrect specifications or revision mismatch"
    ],
    "Process / Method": [
        "Incorrect process step sequence",
        "Critical process parameters not controlled",
        "Work instructions unclear or missing detail",
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
    "Training / People": [
        "No defined training matrix or certification tracking",
        "New hires not trained on critical control points",
        "Training effectiveness not evaluated",
        "Knowledge not shared between shifts or teams",
        "Competence requirements not clearly defined"
    ],
    "Supplier / External": [
        "Supplier not included in 8D / corrective action process",
        "No supplier quality performance metrics tracked",
        "Suppliers not audited for process adherence",
        "Material or component selection not validated for reliability",
        "Supply chain change not communicated to manufacturing"
    ]
}

# ---------------------------
# Helper functions
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

def render_whys_no_repeat(why_list, categories, label_prefix):
    for idx in range(len(why_list)):
        selected_so_far = [w for i, w in enumerate(why_list) if w.strip() and i != idx]
        options = [""] + [f"{cat}: {item}" for cat, items in categories.items() for item in items if f"{cat}: {item}" not in selected_so_far]
        current_val = why_list[idx] if why_list[idx] in options else ""
        why_list[idx] = st.selectbox(
            f"{label_prefix} {idx+1}",
            options,
            index=options.index(current_val) if current_val in options else 0,
            key=f"{label_prefix}_{idx}"
        )
        free_text = st.text_input(f"Or enter your own {label_prefix} {idx+1}", value=why_list[idx], key=f"{label_prefix}_txt_{idx}")
        if free_text.strip():
            why_list[idx] = free_text

# ---------------------------
# Render Tabs D1–D8
# ---------------------------
tabs = st.tabs([t[lang_key]["D1"], t[lang_key]["D2"], t[lang_key]["D3"], t[lang_key]["D4"], t[lang_key]["D5"], t[lang_key]["D6"], t[lang_key]["D7"], t[lang_key]["D8"]])

# ---------------------------
# D1
# ---------------------------
with tabs[0]:
    st.markdown("**Customer Concerns / Detalles del cliente**")
    st.session_state["D1"]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state["D1"]["answer"], height=100)

# ---------------------------
# D2
# ---------------------------
with tabs[1]:
    st.markdown("**Similar Part Considerations / Consideraciones de partes similares**")
    st.session_state["D2"]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state["D2"]["answer"], height=100)

# ---------------------------
# D3
# ---------------------------
with tabs[2]:
    st.markdown("**Initial Analysis / Análisis inicial**")
    st.session_state["D3"]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state["D3"]["answer"], height=100)

# ---------------------------
# D4
# ---------------------------
with tabs[3]:
    st.markdown("**Containment Actions / Acciones de contención**")
    st.session_state["D4"]["containment"] = st.text_area("Containment Actions / Acciones de contención", value=st.session_state["D4"]["containment"], height=60)
    st.session_state["D4"]["location"] = st.text_input("Material Location / Ubicación del material", value=st.session_state["D4"]["location"])
    st.session_state["D4"]["status"] = st.text_input("Activity Status / Estado de la actividad", value=st.session_state["D4"]["status"])

# ---------------------------
# D5
# ---------------------------
with tabs[4]:
    st.markdown("**Final Analysis / Análisis Final**")
    st.write("Occurrence Why / Por qué Ocurrencia")
    render_whys_no_repeat(st.session_state["d5_occ_whys"], occurrence_categories, t[lang_key]["Occurrence_Why"])
    st.write("Detection Why / Por qué Detección")
    render_whys_no_repeat(st.session_state["d5_det_whys"], detection_categories, t[lang_key]["Detection_Why"])
    st.write("Systemic Why / Por qué Sistémico")
    render_whys_no_repeat(st.session_state["d5_sys_whys"], systemic_categories, t[lang_key]["Systemic_Why"])

# ---------------------------
# D6
# ---------------------------
with tabs[5]:
    st.markdown("**Permanent Corrective Actions / Acciones correctivas permanentes**")
    st.session_state["D6"]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state["D6"]["answer"], height=100)

# ---------------------------
# D7
# ---------------------------
with tabs[6]:
    st.markdown("**Countermeasure Confirmation / Confirmación de contramedidas**")
    st.session_state["D7"]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state["D7"]["answer"], height=100)

# ---------------------------
# D8
# ---------------------------
with tabs[7]:
    st.markdown("**Follow-up Activities / Actividades de seguimiento**")
    st.session_state["D8"]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state["D8"]["answer"], height=100)

# ---------------------------
# Excel Export
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"
    
    ws["A1"] = "Report Date"
    ws["B1"] = st.session_state["report_date"]
    ws["A2"] = "Prepared By"
    ws["B2"] = st.session_state["prepared_by"]
    
    row = 4
    for step, _, _ in npqp_steps:
        ws[f"A{row}"] = t[lang_key][step]
        ws[f"B{row}"] = st.session_state[step]["answer"]
        row += 2
    
    ws[f"A{row}"] = "D4 Location"
    ws[f"B{row}"] = st.session_state["D4"]["location"]
    row += 1
    ws[f"A{row}"] = "D4 Status"
    ws[f"B{row}"] = st.session_state["D4"]["status"]
    row += 1
    ws[f"A{row}"] = "D4 Containment"
    ws[f"B{row}"] = st.session_state["D4"]["containment"]
    row += 1
    
    # D5 Whys
    for cat, key in zip(["Occurrence", "Detection", "Systemic"], ["d5_occ_whys", "d5_det_whys", "d5_sys_whys"]):
        ws[f"A{row}"] = f"{cat} Why"
        ws[f"B{row}"] = ", ".join([w for w in st.session_state[key] if w.strip()])
        row += 1
    
    # Save to BytesIO
    stream = io.BytesIO()
    wb.save(stream)
    return stream

st.sidebar.markdown("---")
excel_stream = generate_excel()
st.sidebar.download_button(
    label=t[lang_key]["Download"],
    data=excel_stream,
    file_name=f"8D_Report_{datetime.datetime.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
