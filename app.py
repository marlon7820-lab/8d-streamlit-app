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

# Language selection
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# Safe reset functions
# ---------------------------
def reset_8d_session():
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    keys_to_delete = [k for k in st.session_state.keys() if k not in preserve_keys]
    for k in keys_to_delete:
        del st.session_state[k]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.experimental_rerun()

def reset_full_session():
    keys_to_delete = list(st.session_state.keys())
    for k in keys_to_delete:
        del st.session_state[k]
    st.experimental_rerun()

st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")

if st.sidebar.button("🔄 Reset 8D Session"):
    reset_8d_session()

if st.sidebar.button("🧹 Reset Full Session"):
    reset_full_session()

# ---------------------------
# Language dictionary
# ---------------------------
t = {
    "en": {
        "D1": "D1: Concern Details", "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis", "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis", "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)", "Root_Cause_Det": "Root Cause (Detection)", "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why", "Systemic_Why": "Systemic Why",
        "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example",
        "FMEA_Failure": "FMEA Failure Occurrence",
        "Location": "Material Location", "Status": "Activity Status", "Containment_Actions": "Containment Actions"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)", "Root_Cause_Det": "Causa raíz (Detección)", "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección", "Systemic_Why": "Por qué Sistémico",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA",
        "Location": "Ubicación del material", "Status": "Estado de la actividad", "Containment_Actions": "Acciones de contención"
    }
}

# ---------------------------
# NPQP 8D steps with examples
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."},
           {"en":"Customer reported static noise in amplifier during end-of-line test.",
            "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.", "es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."},
           {"en":"Similar model radio, Front vs. rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.", "es":"Realice una investigación inicial para identificar problemas evidentes."},
           {"en":"Visual inspection of solder joints, initial functional tests.", "es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions and material location.", "es":"Defina acciones de contención temporales y ubicación del material."},
           {"en":"","es":""}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause.", "es":"Use el análisis de 5 Porqués para determinar la causa raíz."},
           {"en":"","es":""}),
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
# D5 categories
# ---------------------------
occurrence_categories = {
    "Machine / Equipment": ["Mechanical failure", "Calibration issues", "Tooling wear", "Machine parameters not optimized", "Maintenance schedule issue", "Sensor misalignment", "Automation fault", "Unstable process setup"],
    "Material / Component": ["Wrong material delivered", "Supplier off-spec", "Material defect not visible", "Damage in storage/handling", "Labeling error", "Material substitution without approval", "Incorrect specs"],
    "Process / Method": ["Incorrect process step", "Critical parameter not controlled", "Work instructions unclear", "Process drift not detected", "Control plan not followed", "Incorrect assembly process", "Outdated FMEA linkage", "Inadequate process capability"],
    "Design / Engineering": ["Design not robust", "Tolerance stack-up issue", "Late design change not communicated", "Incorrect drawing spec", "Component placement error", "Lack of verification/testing"],
    "Environmental / External": ["Temperature/humidity out of range", "ESD not controlled", "Contamination/dust", "Power fluctuation", "Vibration/noise", "Unstable environment monitoring"]
}

detection_categories = {
    "QA / Inspection": ["QA checklist incomplete", "No automated inspection", "Manual inspection error", "Inspection too infrequent", "Criteria unclear", "Measurement system not capable", "Incoming inspection missed", "Final inspection missed"],
    "Validation / Process": ["Process validation outdated", "Insufficient verification", "Design validation incomplete", "Control plan coverage inadequate", "Lack of monitoring", "Process limits outdated"],
    "FMEA / Control Plan": ["Failure mode not captured", "Detection controls missing", "Control plan not updated", "FMEA not reviewed", "Detection ranking unrealistic", "PFMEA and control plan not linked"],
    "Test / Equipment": ["Test calibration overdue", "Testing software incorrect", "Test setup not detecting failure", "Detection threshold too wide", "Test data not reviewed"],
    "Systemic / Organizational": ["Feedback loop not implemented", "Lack of detection feedback", "Training gaps", "Quality alerts not communicated"]
}

systemic_categories = {
    "Management / Organization": ["Inadequate leadership", "Insufficient resources", "Delayed response", "Lack of accountability", "Ineffective escalation", "Weak cross-functional communication"],
    "Process / Procedure": ["SOPs outdated", "Process FMEA not reviewed", "Control plan misaligned", "Lessons learned not integrated", "Inefficient document control", "Maintenance procedures not standardized"],
    "Training / People": ["No training matrix", "New hires not trained", "Training effectiveness not evaluated", "Knowledge not shared", "Competence requirements unclear"],
    "Supplier / External": ["Supplier not included in 8D/FMEA", "Supplier corrective actions not verified", "Incoming audit inadequate", "Supplier changes not communicated", "Long lead time for supplier closure"],
    "Quality System / Feedback": ["Internal audits ineffective", "KPI not linked to root cause", "Ineffective 5-Why use", "Customer complaints not fed into design reviews", "No systemic review after multiple 8Ds"]
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

def render_whys_no_repeat(why_list, categories, label_prefix, key_prefix):
    for idx in range(len(why_list)):
        selected_so_far = [w for i, w in enumerate(why_list) if w.strip() and i != idx]
        options = [""] + [f"{cat}: {item}" for cat, items in categories.items() for item in items if f"{cat}: {item}" not in selected_so_far]
        current_val = why_list[idx] if why_list[idx] in options else ""
        why_list[idx] = st.selectbox(
            f"{label_prefix} {idx+1}",
            options,
            index=options.index(current_val) if current_val in options else 0,
            key=f"{key_prefix}_{idx}"
        )
        free_text = st.text_input(f"Or enter your own {label_prefix} {idx+1}", value=why_list[idx], key=f"{key_prefix}_txt_{idx}")
        if free_text.strip():
            why_list[idx] = free_text

# ---------------------------
# Render Tabs D1–D8
# ---------------------------
tab_labels = []
for step, _, _ in npqp_steps:
    if st.session_state[step]["answer"].strip() != "":
        tab_labels.append(f"🟢 {t[lang_key][step]}")
    else:
        tab_labels.append(f"🔴 {t[lang_key][step]}")

tabs = st.tabs(tab_labels)

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        st.markdown(f"""
        <div style="
            background-color:#b3e0ff; 
            color:black; 
            padding:12px; 
            border-left:5px solid #1E90FF; 
            border-radius:6px;
            width:100%;
            font-size:14px;
            line-height:1.5;
        ">
        <b>{t[lang_key]['Training_Guidance']}:</b> {note_dict[lang_key]}<br><br>
        💡 <b>{t[lang_key]['Example']}:</b> {example_dict[lang_key]}
        </div>
        """, unsafe_allow_html=True)

        if step == "D4":
            st.session_state[step]["location"] = st.selectbox("Location of Material", ["", "Work in Progress", "Stores Stock", "Warehouse Stock", "Service Parts", "Other"], index=0, key="d4_location")
            st.session_state[step]["status"] = st.selectbox("Status of Activities", ["", "Pending", "In Progress", "Completed", "Other"], index=0, key="d4_status")
            st.session_state[step]["answer"] = st.text_area("Containment Actions / Notes", value=st.session_state[step]["answer"], key=f"ans_{step}")
        elif step == "D5":
            st.markdown("#### Occurrence Analysis")
            render_whys_no_repeat(st.session_state.d5_occ_whys, occurrence_categories, t[lang_key]['Occurrence_Why'], key_prefix="d5_occ")
            if st.button("➕ Add another Occurrence Why", key="add_occ_why"):
                st.session_state.d5_occ_whys.append("")
            st.markdown("#### Detection Analysis")
            render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories, t[lang_key]['Detection_Why'], key_prefix="d5_det")
            if st.button("➕ Add another Detection Why", key="add_det_why"):
                st.session_state.d5_det_whys.append("")
            st.markdown("#### Systemic Analysis")
            render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories, t[lang_key]['Systemic_Why'], key_prefix="d5_sys")
            if st.button("➕ Add another Systemic Why", key="add_sys_why"):
                st.session_state.d5_sys_whys.append("")
            st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=suggest_root_cause([w for w in st.session_state.d5_occ_whys if w.strip()]), height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=suggest_root_cause([w for w in st.session_state.d5_det_whys if w.strip()]), height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=suggest_root_cause([w for w in st.session_state.d5_sys_whys if w.strip()]), height=80, disabled=True)
        else:
            st.session_state[step]["answer"] = st.text_area("Notes / Details", value=st.session_state[step]["answer"], key=f"ans_{step}")

# ---------------------------
# Save / Download Excel
# ---------------------------
def export_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"
    ws.append(["Step", "Notes / Details"])
    for step, _, _ in npqp_steps:
        answer = st.session_state[step]["answer"]
        ws.append([t[lang_key][step], answer])
    # D5 whys
    ws.append(["D5 Occurrence Why(s)"] + st.session_state.d5_occ_whys)
    ws.append(["D5 Detection Why(s)"] + st.session_state.d5_det_whys)
    ws.append(["D5 Systemic Why(s)"] + st.session_state.d5_sys_whys)
    # Containment actions
    ws.append(["D4 Material Location", st.session_state.d4_location])
    ws.append(["D4 Activity Status", st.session_state.d4_status])
    ws.append(["D4 Containment Actions", st.session_state.D4["answer"]])
    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()

if st.button(t[lang_key]["Download"]):
    excel_data = export_to_excel()
    st.download_button(
        label=f"{t[lang_key]['Download']}",
        data=excel_data,
        file_name=f"8D_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("<p style='text-align:center; font-size:12px; color:#555555;'>End of 8D Report Assistant</p>", unsafe_allow_html=True)
