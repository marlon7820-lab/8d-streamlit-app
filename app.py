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
# App styles - updated for desktop selectbox outline + thumbnails
# ---------------------------
st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}
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

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

dark_mode = st.sidebar.checkbox("🌙 Dark Mode")
if dark_mode:
    st.markdown("""
    <style>
    /* Main app background & text */
    .stApp {
        background: linear-gradient(to right, #1e1e1e, #2c2c2c);
        color: #f5f5f5 !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab"] {
        font-weight: bold; 
        color: #f5f5f5 !important;
    }
    .stTabs [data-baseweb="tab"]:hover {
        color: #87AFC7 !important;
    }

    /* Text inputs, textareas, selectboxes */
    div.stTextInput, div.stTextArea, div.stSelectbox {
        border: 2px solid #87AFC7 !important;
        border-radius: 5px !important;
        background-color: #2c2c2c !important;
        color: #f5f5f5 !important;
        padding: 5px !important;
        transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
    }
    div.stTextInput:hover, div.stTextArea:hover, div.stSelectbox:hover {
        border: 2px solid #1E90FF !important;
        box-shadow: 0 0 5px #1E90FF;
    }

    /* Info boxes */
    .stInfo {
        background-color: #3a3a3a !important; 
        border-left: 5px solid #87AFC7 !important; 
        color: #f5f5f5 !important;
    }

    /* Sidebar background & text */
    .css-1d391kg {color: #87AFC7 !important; font-weight: bold !important;}
    .stSidebar {
        background-color: #1e1e1e !important;
        color: #f5f5f5 !important;
    }

    /* Sidebar buttons */
    .stSidebar button[kind="primary"] {
        background-color: #87AFC7 !important;
        color: #000000 !important;
        font-weight: bold;
    }
    .stSidebar button {
        background-color: #5a5a5a !important;
        color: #f5f5f5 !important;
    }

    /* Download button in sidebar */
    .stSidebar .stDownloadButton button {
        background-color: #87AFC7 !important;
        color: #000000 !important;
        font-weight: bold;
    }

    </style>
    """, unsafe_allow_html=True)
# ---------------------------
# Sidebar: App Controls
# ---------------------------
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
    ("D5", {"en": "Use 5-Why analysis to determine the root cause.", "es": "Use el análisis de 5 Porqués para determinar la causa raíz."}, {"en": "Final 'Why' from the Analysis will give a good indication of the True Root Cause", "es": "El último \"Por qué\" del análisis proporcionará una idea clara de la causa raíz del problema"}),
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

# ---------------------------
# Helper: Render 5-Why dropdowns without repeating selections
# ---------------------------
def render_whys_with_free_text(why_list, categories, label_prefix):
    for idx in range(len(why_list)):
        selected_so_far = [w for i, w in enumerate(why_list) if w.strip() and i != idx]
        options = [""] + [f"{cat}: {item}" for cat, items in categories.items() for item in items
                          if f"{cat}: {item}" not in selected_so_far]
        current_val = why_list[idx] if why_list[idx] in options else ""
        why_list[idx] = st.selectbox(
            f"{label_prefix} {idx+1}",
            options,
            index=options.index(current_val) if current_val in options else 0,
            key=f"{label_prefix}_{idx}_{lang_key}"
        )
        free_text = st.text_input(
            f"Or enter your own {label_prefix} {idx+1}",
            value=why_list[idx],
            key=f"{label_prefix}_txt_{idx}_{lang_key}"
        )
        if free_text.strip():
            why_list[idx] = free_text
    return why_list

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
    return why_list
# ---------------------------
# Render Tabs with Uploads
# ---------------------------
tab_labels = [
    f"🟢 {t[lang_key][step]}" if st.session_state[step]["answer"].strip() else f"🔴 {t[lang_key][step]}"
    for step, _, _ in npqp_steps
]
tabs = st.tabs(tab_labels)

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")

        # Training Guidance & Example
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

        # File uploads for D1, D3, D4, D7
        if step in ["D1","D3","D4","D7"]:
            uploaded_files = st.file_uploader(
                f"Upload files/photos for {step}",
                type=["png", "jpg", "jpeg", "pdf", "xlsx", "txt"],
                accept_multiple_files=True,
                key=f"upload_{step}"
            )
            if uploaded_files:
                for file in uploaded_files:
                    if file not in st.session_state[step]["uploaded_files"]:
                        st.session_state[step]["uploaded_files"].append(file)

        # Display uploaded files (aligned with file upload, not nested too deep)
        if step in ["D1","D3","D4","D7"] and st.session_state[step].get("uploaded_files"):
            st.markdown("**Uploaded Files / Photos:**")
            for f in st.session_state[step]["uploaded_files"]:
                st.write(f"{f.name}")
                if f.type.startswith("image/"):
                    st.image(f, width=192)  # roughly 2 inches wide, height auto-scaled
    
        # ---------------------------
# Step Rendering
# ---------------------------
if step == "D4":
    st.markdown("#### Root Cause by Category")
    st.session_state.d4_selection = render_whys_no_repeat(
        st.session_state.get("d4_selection", ["", "", ""]),
        categories=d4_categories,
        label_prefix=t[lang_key]["D4_Why"]
    )

elif step == "D5":
    st.markdown("#### Occurrence Analysis")
    st.session_state.d5_occ_whys = render_whys_with_free_text(
        st.session_state.get("d5_occ_whys", [""]),
        occurrence_categories,
        t[lang_key]['Occurrence_Why']
    )
    if st.button("➕ Add another Occurrence Why", key="add_occ"):
        st.session_state.d5_occ_whys.append("")

    st.markdown("#### Detection Analysis")
    st.session_state.d5_det_whys = render_whys_with_free_text(
        st.session_state.get("d5_det_whys", [""]),
        detection_categories,
        t[lang_key]['Detection_Why']
    )
    if st.button("➕ Add another Detection Why", key="add_det"):
        st.session_state.d5_det_whys.append("")

    st.markdown("#### Systemic Analysis")
    st.session_state.d5_sys_whys = render_whys_with_free_text(
        st.session_state.get("d5_sys_whys", [""]),
        systemic_categories,
        t[lang_key]['Systemic_Why']
    )
    if st.button("➕ Add another Systemic Why", key="add_sys"):
        st.session_state.d5_sys_whys.append("")

    # Dynamic Root Causes
    occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
    det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
    sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]

    st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet", height=80, disabled=True)
    st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet", height=80, disabled=True)
    st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet", height=80, disabled=True)

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

    # Mirror into top-level D6 storage for export
    st.session_state["D6"]["occ_answer"] = st.session_state[step]["occ_answer"]
    st.session_state["D6"]["det_answer"] = st.session_state[step]["det_answer"]
    st.session_state["D6"]["sys_answer"] = st.session_state[step]["sys_answer"]

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

    # Mirror into top-level D7 storage for export
    st.session_state["D7"]["occ_answer"] = st.session_state[step]["occ_answer"]
    st.session_state["D7"]["det_answer"] = st.session_state[step]["det_answer"]
    st.session_state["D7"]["sys_answer"] = st.session_state[step]["sys_answer"]

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
# Excel generation (formatted + images/files)
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
            if c == 2:
                cell.font = Font(bold=True)
            cell.border = border

    # Insert uploaded images below table
    from PIL import Image as PILImage
    from io import BytesIO

    last_row = ws.max_row + 2
    for step in ["D1","D3","D4","D7"]:
        uploaded_files = st.session_state[step].get("uploaded_files", [])
        if uploaded_files:
            ws.cell(row=last_row, column=1, value=f"{step} Uploaded Files / Photos").font = Font(bold=True)
            last_row += 1
            for f in uploaded_files:
                if f.type.startswith("image/"):
                    try:
                        img = PILImage.open(BytesIO(f.getvalue()))
                        max_width = 300
                        ratio = max_width / img.width
                        img = img.resize((int(img.width * ratio), int(img.height * ratio)))
                        temp_path = f"/tmp/{f.name}"
                        img.save(temp_path)
                        excel_img = XLImage(temp_path)
                        ws.add_image(excel_img, f"A{last_row}")
                        last_row += int(img.height / 15) + 2
                    except Exception as e:
                        ws.cell(row=last_row, column=1, value=f"Could not add image {f.name}: {e}")
                        last_row += 1
                else:
                    ws.cell(row=last_row, column=1, value=f.name)
                    last_row += 1

    # Set column widths
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    # ✅ The return must be inside the function
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Move download button to sidebar
with st.sidebar:
    st.download_button(
        label=t[lang_key]['Download'],  # no extra icon
        data=generate_excel(),  # function that returns BytesIO of XLSX
        file_name=f"8D_Report_{st.session_state['report_date']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ---------------------------
# (End)
# ---------------------------
