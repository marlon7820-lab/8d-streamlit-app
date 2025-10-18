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

# App styles - updated for desktop selectbox outline
st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}

/* Outline all Streamlit widget containers (works on desktop) */
div.stSelectbox, div.stTextInput, div.stTextArea {
    border: 2px solid #1E90FF !important;
    border-radius: 5px !important;
    padding: 5px !important;
    background-color: #ffffff !important;
    transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}

/* Hover effect */
div.stSelectbox:hover, div.stTextInput:hover, div.stTextArea:hover {
    border: 2px solid #104E8B !important; /* slightly darker blue */
    box-shadow: 0 0 5px #1E90FF;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Reset Session check (safe, no KeyError)
# ---------------------------
if st.session_state.get("_reset_8d_session", False):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}

    # Clear everything except preserved values
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]

    # Restore preserved values
    for k, v in preserved.items():
        st.session_state[k] = v

    # Safely unset the flag only if it exists
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

# Language selection
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# Sidebar: Smart Session Reset Button
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")
# Reset 8D Session button
if st.sidebar.button("🔄 Reset 8D Session"):
    # Preserve essential keys
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}

    # Clear all other keys
    for key in list(st.session_state.keys()):
        if key not in preserve_keys:
            del st.session_state[key]

    # Restore preserved keys
    for k, v in preserved.items():
        st.session_state[k] = v

    # Set a dedicated reset flag
    st.session_state["_reset_8d_session"] = True

    # Stop further execution; the app will rerun safely
    st.stop()

# At the very top of your app (after imports), handle the reset flag safely:
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
    ("D4", {"en":"Define temporary containment actions and material location.", "es":"Defina acciones de contención temporales y ubicación del material."}, {"en":"Post Quality Alert, Increase Inspection, Inventory Certification","es":"Implementar Ayuda Visual, Incrementar Inspeccion, Certificar Inventario"}),
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

# Ensure D6/D7 subfields exist
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("d5_sys_whys", [""]*5)
st.session_state.setdefault("d4_location", "")
st.session_state.setdefault("d4_status", "")
st.session_state.setdefault("d4_containment", "")

# D6 fields
for sub in ["occ_answer", "det_answer", "sys_answer"]:
    st.session_state.setdefault(("D6"), st.session_state.get("D6", {}))
    st.session_state["D6"].setdefault(sub, "")

# D7 fields
for sub in ["occ_answer", "det_answer", "sys_answer"]:
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
# Helper: Suggest root cause based on whys
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
# Helper: Render 5-Why dropdowns without repeating selections
# ---------------------------
def render_whys_no_repeat(why_list, categories, label_prefix):
    for idx in range(len(why_list)):
        selected_so_far = [w for i, w in enumerate(why_list) if w.strip() and i != idx]
        options = [""] + [f"{cat}: {item}" for cat, items in categories.items() for item in items if f"{cat}: {item}" not in selected_so_far]
        current_val = why_list[idx] if why_list[idx] in options else ""
        why_list[idx] = st.selectbox(
            f"{label_prefix} {idx+1}",
            options,
            index=options.index(current_val) if current_val in options else 0,
            key=f"{label_prefix}_{idx}_{lang_key}"
        )
        free_text = st.text_input(f"Or enter your own {label_prefix} {idx+1}", value=why_list[idx], key=f"{label_prefix}_txt_{idx}_{lang_key}")
        if free_text.strip():
            why_list[idx] = free_text

# ---------------------------
# Render Tabs D1–D8
# ---------------------------
tab_labels = [
    f"🟢 {t[lang_key][step]}" if st.session_state[step]["answer"].strip() else f"🔴 {t[lang_key][step]}"
    for step, _, _ in npqp_steps
]
tabs = st.tabs(tab_labels)

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        note_text = note_dict[lang_key]
        example_text = example_dict[lang_key]
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
<b>{t[lang_key]['Training_Guidance']}:</b> {note_text}<br><br>
💡 <b>{t[lang_key]['Example']}:</b> {example_text}
</div>
""", unsafe_allow_html=True)

        # Default single-answer field (for steps that use it)
        # We'll override for D4, D5, D6, D7, D8 below as needed

        # D4 Nissan-style
        if step == "D4":
            st.session_state[step]["location"] = st.selectbox(
                "Location of Material",
                ["", "Work in Progress", "Stores Stock", "Warehouse Stock", "Service Parts", "Other"],
                index=0,
                key="d4_location"
            )
            st.session_state[step]["status"] = st.selectbox(
                "Status of Activities",
                ["", "Pending", "In Progress", "Completed", "Other"],
                index=0,
                key="d4_status"
            )
            st.session_state[step]["answer"] = st.text_area(
                "Containment Actions / Notes",
                value=st.session_state[step]["answer"],
                key=f"ans_{step}"
            )

        # D5 5-Why
        elif step == "D5":
            st.markdown("#### Occurrence Analysis")
            render_whys_no_repeat(st.session_state.d5_occ_whys, occurrence_categories, t[lang_key]['Occurrence_Why'])
            if st.button("➕ Add another Occurrence Why", key=f"add_occ_{i}"):
                st.session_state.d5_occ_whys.append("")
            st.markdown("#### Detection Analysis")
            render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories, t[lang_key]['Detection_Why'])
            if st.button("➕ Add another Detection Why", key=f"add_det_{i}"):
                st.session_state.d5_det_whys.append("")
            st.markdown("#### Systemic Analysis")
            render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories, t[lang_key]['Systemic_Why'])
            if st.button("➕ Add another Systemic Why", key=f"add_sys_{i}"):
                st.session_state.d5_sys_whys.append("")
            # Dynamic Root Causes
            occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
            det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
            sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]
            st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet", height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet", height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet", height=80, disabled=True)

        # D6: Permanent Corrective Actions (three text areas: Occ/Det/Sys)
        elif step == "D6":
            st.session_state[step].setdefault("occ_answer", st.session_state["D6"].get("occ_answer", ""))
            st.session_state[step].setdefault("det_answer", st.session_state["D6"].get("det_answer", ""))
            st.session_state[step].setdefault("sys_answer", st.session_state["D6"].get("sys_answer", ""))

            st.session_state[step]["occ_answer"] = st.text_area(
                "D6 - Corrective Actions for Occurrence Root Cause",
                value=st.session_state[step]["occ_answer"],
                key="d6_occ"
            )
            st.session_state[step]["det_answer"] = st.text_area(
                "D6 - Corrective Actions for Detection Root Cause",
                value=st.session_state[step]["det_answer"],
                key="d6_det"
            )
            st.session_state[step]["sys_answer"] = st.text_area(
                "D6 - Corrective Actions for Systemic Root Cause",
                value=st.session_state[step]["sys_answer"],
                key="d6_sys"
            )

            # Mirror into top-level D6 storage so export code can find them consistently
            st.session_state["D6"]["occ_answer"] = st.session_state[step]["occ_answer"]
            st.session_state["D6"]["det_answer"] = st.session_state[step]["det_answer"]
            st.session_state["D6"]["sys_answer"] = st.session_state[step]["sys_answer"]

        # D7: Countermeasure Confirmation (three text areas: verification for Occ/Det/Sys)
        elif step == "D7":
            st.session_state[step].setdefault("occ_answer", st.session_state["D7"].get("occ_answer", ""))
            st.session_state[step].setdefault("det_answer", st.session_state["D7"].get("det_answer", ""))
            st.session_state[step].setdefault("sys_answer", st.session_state["D7"].get("sys_answer", ""))

            st.session_state[step]["occ_answer"] = st.text_area(
                "D7 - Occurrence Countermeasure Verification",
                value=st.session_state[step]["occ_answer"],
                key="d7_occ"
            )
            st.session_state[step]["det_answer"] = st.text_area(
                "D7 - Detection Countermeasure Verification",
                value=st.session_state[step]["det_answer"],
                key="d7_det"
            )
            st.session_state[step]["sys_answer"] = st.text_area(
                "D7 - Systemic Countermeasure Verification",
                value=st.session_state[step]["sys_answer"],
                key="d7_sys"
            )

            # Mirror into top-level D7 storage so export code can find them consistently
            st.session_state["D7"]["occ_answer"] = st.session_state[step]["occ_answer"]
            st.session_state["D7"]["det_answer"] = st.session_state[step]["det_answer"]
            st.session_state["D7"]["sys_answer"] = st.session_state[step]["sys_answer"]

        # D8: Follow-up Activities / Lessons Learned (single text area)
        elif step == "D8":
            st.session_state[step]["answer"] = st.text_area(
                "Your Answer",
                value=st.session_state[step]["answer"],
                key=f"ans_{step}"
            )

        else:
            # Default for D1, D2, D3, or any other single-answer steps
            if step not in ["D4", "D5", "D6", "D7", "D8"]:
                st.session_state[step]["answer"] = st.text_area(
                    "Your Answer",
                    value=st.session_state[step]["answer"],
                    key=f"ans_{step}"
                )

# ---------------------------
# Collect all answers for Excel export
# ---------------------------
data_rows = []

occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]

occ_rc_text = suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet"
det_rc_text = suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet"
sys_rc_text = suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet"

for step, _, _ in npqp_steps:
    # D6 and D7 should export their 3 sub-answers as separate rows
    if step == "D6":
        data_rows.append(("D6 - Occurrence Countermeasure", st.session_state.get("D6", {}).get("occ_answer", ""), ""))
        data_rows.append(("D6 - Detection Countermeasure", st.session_state.get("D6", {}).get("det_answer", ""), ""))
        data_rows.append(("D6 - Systemic Countermeasure", st.session_state.get("D6", {}).get("sys_answer", ""), ""))
    elif step == "D7":
        data_rows.append(("D7 - Occurrence Countermeasure Verification", st.session_state.get("D7", {}).get("occ_answer", ""), ""))
        data_rows.append(("D7 - Detection Countermeasure Verification", st.session_state.get("D7", {}).get("det_answer", ""), ""))
        data_rows.append(("D7 - Systemic Countermeasure Verification", st.session_state.get("D7", {}).get("sys_answer", ""), ""))
    elif step == "D5":
        data_rows.append(("D5 - Root Cause (Occurrence)", occ_rc_text, " | ".join(occ_whys)))
        data_rows.append(("D5 - Root Cause (Detection)", det_rc_text, " | ".join(det_whys)))
        data_rows.append(("D5 - Root Cause (Systemic)", sys_rc_text, " | ".join(sys_whys)))
    elif step == "D4":
        loc = st.session_state[step].get("location", "")
        status = st.session_state[step].get("status", "")
        answer = st.session_state[step].get("answer", "")
        extra = f"Location: {loc} | Status: {status}"
        data_rows.append((step, answer, extra))
    else:
        answer = st.session_state[step].get("answer", "")
        extra = st.session_state[step].get("extra", "")
        data_rows.append((step, answer, extra))

# ---------------------------
# Excel generation (formatted)
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Add logo if exists
    if os.path.exists("logo.png"):
        try:
            img = XLImage("logo.png")
            img.width = 140
            img.height = 40
            ws.add_image(img, "A1")
        except:
            pass

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
    ws.cell(row=3, column=1, value="📋 8D Report Assistant").font = Font(bold=True, size=14)

    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    # Header row
    header_row = ws.max_row + 1
    headers = ["Step", "Answer", "Extra / Notes"]
    fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Append step answers
    for step_label, answer_text, extra_text in data_rows:
        ws.append([step_label, answer_text, extra_text])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            # Bold the Answer column content visually
            if c == 2:
                cell.font = Font(bold=True)
            cell.border = border

    # Set column widths
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
# ---------------------------
# Sidebar: Attach files / photos
# ---------------------------
with st.sidebar:
    st.markdown("## 📎 Attach Files / Photos")
    uploaded_files = st.file_uploader(
        "Upload files to include in 8D report",
        type=["png", "jpg", "jpeg", "pdf", "xlsx", "docx"],
        accept_multiple_files=True
    )

    # Store uploaded files in session state
    if uploaded_files:
        if "attached_files" not in st.session_state:
            st.session_state["attached_files"] = []
        st.session_state["attached_files"].extend(uploaded_files)

    # Display list of uploaded files
    if st.session_state.get("attached_files"):
        st.markdown("### Attached Files")
        for f in st.session_state["attached_files"]:
            st.write(f"- {f.name}")

# At the end of Excel generation
if st.session_state.get("attached_files"):
    # Add a blank row safely
ws.append([""])  # just one empty cell
ws.append(["Attached Files"])
    for f in st.session_state["attached_files"]:
    ws.append([f.name])


# ---------------------------
# (End)
# ---------------------------
