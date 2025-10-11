# ---------------------------
# 8D Report Assistant — PART 1 (Setup, imports, config, initialization, D1-D3)
# ---------------------------

import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.writer.excel import save_virtual_workbook
import datetime
import json
import os

# ---------------------------
# Safety / limits
# ---------------------------
MAX_WHYS = 10                # limit for number of 5-Why entries to avoid runaway session growth
LOGO_PATH = "logo.png"       # optional logo; handled safely (may be absent)

# ---------------------------
# Page config / styles
# ---------------------------
st.set_page_config(page_title="8D Report Assistant", layout="wide", page_icon=None)

st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.section-box {background-color:#b3e0ff; color:black; padding:12px; border-left:5px solid #1E90FF; border-radius:6px; width:100%; font-size:14px; line-height:1.5;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Title / optional safe logo display
# ---------------------------
col1, col2, col3 = st.columns([1,6,1])
with col1:
    pass
with col2:
    st.markdown("<h1 style='text-align:center; color:#1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='text-align:center; color:#666666; font-size:12px;'>Version v1.0.11 — Last updated: October 10, 2025</p>", unsafe_allow_html=True)
with col3:
    # show small logo if available but open safely
    if os.path.exists(LOGO_PATH):
        try:
            with open(LOGO_PATH, "rb") as _f:
                st.image(_f.read(), width=100)
        except Exception:
            st.warning("Logo could not be loaded (will not affect app functionality).")

st.sidebar.markdown("### ⚙️ Controls")

# ---------------------------
# Reset session state safely
# ---------------------------
if st.sidebar.button("🔄 Reset Session State"):
    # preserve nothing — clear everything and rerun
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.experimental_rerun()

# ---------------------------
# Language selection
# ---------------------------
if "language" not in st.session_state:
    st.session_state.language = "English"
lang_choice = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"], index=0 if st.session_state.language=="English" else 1)
st.session_state.language = lang_choice
lang = st.session_state.language

# ---------------------------
# Translation dictionary (kept consistent with previous UX)
# ---------------------------
t = {
    "English": {
        "D1": "D1: Concern Details", "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis", "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis", "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)", "Root_Cause_Det": "Root Cause (Detection)", "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why", "Systemic_Why": "Systemic Why",
        "Location": "Location of Material", "Status": "Activity Status",
        "Training_Guidance": "Training Guidance", "Example": "Example",
        "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX"
    },
    "Español": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)", "Root_Cause_Det": "Causa raíz (Detección)", "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección", "Systemic_Why": "Por qué Sistémico",
        "Location": "Ubicación del material", "Status": "Estado de la actividad",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX"
    }
}

# ---------------------------
# NPQP steps metadata (kept from your previous working app)
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."},
           {"en":"Customer reported static noise in amplifier during end-of-line test.", "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.", "es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."},
           {"en":"Similar model radio, Front vs. rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.", "es":"Realice una investigación inicial para identificar problemas evidentes."},
           {"en":"Visual inspection of solder joints, initial functional tests.", "es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions.", "es":"Defina acciones de contención temporales."},
           {"en":"100% inspection of amplifiers before shipment.", "es":"Inspección 100% de amplificadores antes del envío."}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause.", "es":"Use el análisis de 5 Porqués para determinar la causa raíz."},
           {"en":"", "es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently.", "es":"Defina acciones correctivas que eliminen la causa raíz permanentemente."},
           {"en":"Update soldering process, redesign fixture.", "es":"Actualizar proceso de soldadura, rediseñar herramienta."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue.", "es":"Verifique que las acciones correctivas resuelvan efectivamente el problema."},
           {"en":"Functional tests on corrected amplifiers.", "es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned, update standards, FMEAs.", "es":"Documente lecciones aprendidas, actualice estándares, FMEAs."},
           {"en":"Update SOPs, PFMEA, work instructions.", "es":"Actualizar SOPs, PFMEA, instrucciones de trabajo."})
]

# ---------------------------
# Safe session_state initialization for all keys we will use (prevents missing-key errors)
# ---------------------------
# Initialize step containers
for step, _, _ in npqp_steps:
    st.session_state.setdefault(step, {"answer": "", "extra": ""})

# Initialize D4 Nissan-style fields inside D4 container (kept from your previous app)
st.session_state.setdefault("D4_location", "")
st.session_state.setdefault("D4_status", "")
st.session_state.setdefault("D4_actions", "")

# D5 whys (start with 5 slots to preserve your original behavior, but user may add up to MAX_WHYS)
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("d5_sys_whys", [""]*5)

# D6 / D7 separate answers preserved
st.session_state.setdefault("D6_occ_answer", "")
st.session_state.setdefault("D6_det_answer", "")
st.session_state.setdefault("D6_sys_answer", "")
st.session_state.setdefault("D7_occ_answer", "")
st.session_state.setdefault("D7_det_answer", "")
st.session_state.setdefault("D7_sys_answer", "")

# Report metadata
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")

# ---------------------------
# Navigation: show tabs but keep session-safe rendering
# ---------------------------
tab_labels = []
for step, _, _ in npqp_steps:
    label = t[lang][step] if step in t[lang] else step
    # keep visual indicator of answered vs empty (like your previous app)
    answered = bool(st.session_state[step]["answer"] and str(st.session_state[step]["answer"]).strip())
    tab_labels.append(f"🟢 {label}" if answered else f"🔴 {label}")

tabs = st.tabs(tab_labels)

# ---------------------------
# Render D1 - D3 (the app will continue with D4..D8 in next parts)
# ---------------------------
for i, (step, note_dict, example_dict) in enumerate(npqp_steps[:3]):
    with tabs[i]:
        st.markdown(f"### {t[lang][step]}")
        # training guidance + example (keeps the same UX as before)
        guidance = note_dict["en"] if lang == "English" else note_dict["es"]
        example = example_dict["en"] if lang == "English" else example_dict["es"]
        st.markdown(f"<div class='section-box'><b>{t[lang]['Training_Guidance']}:</b> {guidance}<br><br>💡 <b>{t[lang]['Example']}:</b> {example}</div>", unsafe_allow_html=True)

        # preserve the text area keys used previously so JSON restore and Excel map cleanly
        st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}", height=180)

        # keep a tiny helper row with metadata inputs (consistent with earlier app)
        if step == "D2":
            st.text_input("Part / Model / Reference (optional)", key="d2_part_ref")
        if step == "D3":
            st.text_input("Containment owner", key="d3_containment_owner")

# End of Part 1
# ---------------------------
# PART 2: D4 (Nissan-style) + D5 (5-Why UI) + helpers
# ---------------------------

# ---------- Helpers ----------
def suggest_root_cause(whys):
    """Return a short suggested root cause based on presence of keywords in the whys list."""
    text = " ".join(whys).lower()
    if any(word in text for word in ["training", "knowledge", "human error"]):
        return "Lack of proper training / knowledge gap"
    if any(word in text for word in ["equipment", "tool", "machine", "fixture"]):
        return "Equipment, tooling, or maintenance issue"
    if any(word in text for word in ["procedure", "process", "standard", "sop"]):
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

def render_whys_no_repeat(key_prefix, why_list_name, categories, label_prefix):
    """
    Render a fixed-length list of selectboxes + free-text inputs for a why_list stored in session_state.
    - key_prefix used to make Streamlit keys unique across instances.
    - why_list_name is the name of the session_state list (e.g. 'd5_occ_whys').
    """
    # Ensure the list exists
    why_list = st.session_state.setdefault(why_list_name, [""]*5)
    # Cap the list length by MAX_WHYS
    if len(why_list) > MAX_WHYS:
        why_list = why_list[:MAX_WHYS]
        st.session_state[why_list_name] = why_list

    # Build flattened master options once for performance
    master_options = []
    for cat, items in categories.items():
        for item in items:
            master_options.append(f"{cat}: {item}")

    for idx in range(len(why_list)):
        # compute options excluding currently selected items in other indexes
        selected_others = {w for i, w in enumerate(why_list) if w.strip() and i != idx}
        options = [""] + [opt for opt in master_options if opt not in selected_others]

        current_val = why_list[idx] if why_list[idx] in options else ""
        select_key = f"{key_prefix}_select_{idx}"
        txt_key = f"{key_prefix}_txt_{idx}"

        # show selectbox
        new_val = st.selectbox(f"{label_prefix} {idx+1}", options, index=options.index(current_val) if current_val in options else 0, key=select_key)
        # free text override
        free_text = st.text_input(f"Or enter your own {label_prefix} {idx+1}", value=new_val if new_val and new_val not in [""] else why_list[idx], key=txt_key)
        # Decide final value: free_text has priority if non-empty and different
        final_val = free_text.strip() if free_text.strip() else (new_val if new_val else "")
        why_list[idx] = final_val

    # Save back
    st.session_state[why_list_name] = why_list

# ---------- D4: Nissan-style containment (render into correct tab) ----------
# Find D4 tab index
d4_index = None
for i, (step, _, _) in enumerate(npqp_steps):
    if step == "D4":
        d4_index = i
        break

if d4_index is not None:
    with tabs[d4_index]:
        st.markdown(f"### {t[lang]['D4']}")
        guidance = dict(npqp_steps)[ "D4" ][ "en" ] if lang == "English" else dict(npqp_steps)[ "D4" ][ "es" ]
        st.markdown(f"<div class='section-box'><b>{t[lang]['Training_Guidance']}:</b> {guidance}</div>", unsafe_allow_html=True)

        # Location dropdown (Nissan-style)
        location_opts = ["", "Work in Progress", "Stores Stock", "Warehouse Stock", "Service Parts", "Other"]
        # use top-level keys from Part1
        st.session_state["D4_location"] = st.selectbox(t[lang]["Location"], location_opts, index=location_opts.index(st.session_state.get("D4_location","")) if st.session_state.get("D4_location","") in location_opts else 0, key="D4_loc_select")

        # Activity status
        status_opts = ["", "Planned", "In Progress", "Completed", "On Hold"]
        st.session_state["D4_status"] = st.selectbox(t[lang]["Status"], status_opts, index=status_opts.index(st.session_state.get("D4_status","")) if st.session_state.get("D4_status","") in status_opts else 0, key="D4_status_select")

        # Containment actions text area
        st.session_state["D4_actions"] = st.text_area("Containment Actions (describe temporary measures)", value=st.session_state.get("D4_actions",""), key="D4_actions_text", height=150)

# ---------- D5: 5-Why UI (Occurrence / Detection / Systemic) ----------
# Find D5 tab index
d5_index = None
for i, (step, _, _) in enumerate(npqp_steps):
    if step == "D5":
        d5_index = i
        break

# categories (kept as improved set; "Human / Training" removed)
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
        "Supplier not included in 8D or FMEA review process",
        "Supplier corrective actions not verified for effectiveness",
        "Inadequate incoming material audit process",
        "Supplier process changes not communicated to customer",
        "Long lead time for supplier quality issue closure"
    ],
    "Quality System / Feedback": [
        "Internal audits ineffective or not completed",
        "Quality KPI tracking not linked to root cause analysis",
        "Ineffective use of 5-Why or fishbone tools",
        "Customer complaints not feeding back into design reviews",
        "No systemic review after multiple 8Ds in same area"
    ]
}

if d5_index is not None:
    with tabs[d5_index]:
        st.markdown(f"### {t[lang]['D5']}")
        guidance = dict(npqp_steps)["D5"]["en"] if lang == "English" else dict(npqp_steps)["D5"]["es"]
        st.markdown(f"<div class='section-box'><b>{t[lang]['Training_Guidance']}:</b> {guidance}</div>", unsafe_allow_html=True)

        # Occurrence Analysis
        st.markdown("#### Occurrence Analysis")
        render_whys_no_repeat("occ", "d5_occ_whys", occurrence_categories, t[lang]["Occurrence_Why"])
        # add button with unique key and enforce MAX_WHYS
        if st.button("➕ Add another Occurrence Why", key="add_occ_btn"):
            if len(st.session_state["d5_occ_whys"]) < MAX_WHYS:
                st.session_state["d5_occ_whys"].append("")
            else:
                st.warning(f"Maximum of {MAX_WHYS} whys reached for Occurrence.")

        # Detection Analysis
        st.markdown("#### Detection Analysis")
        render_whys_no_repeat("det", "d5_det_whys", detection_categories, t[lang]["Detection_Why"])
        if st.button("➕ Add another Detection Why", key="add_det_btn"):
            if len(st.session_state["d5_det_whys"]) < MAX_WHYS:
                st.session_state["d5_det_whys"].append("")
            else:
                st.warning(f"Maximum of {MAX_WHYS} whys reached for Detection.")

        # Systemic Analysis
        st.markdown("#### Systemic Analysis")
        render_whys_no_repeat("sys", "d5_sys_whys", systemic_categories, t[lang]["Systemic_Why"])
        if st.button("➕ Add another Systemic Why", key="add_sys_btn"):
            if len(st.session_state["d5_sys_whys"]) < MAX_WHYS:
                st.session_state["d5_sys_whys"].append("")
            else:
                st.warning(f"Maximum of {MAX_WHYS} whys reached for Systemic.")

        # Dynamic Root Cause suggestions (non-editable)
        occ_whys = [w for w in st.session_state.get("d5_occ_whys", []) if w.strip()]
        det_whys = [w for w in st.session_state.get("d5_det_whys", []) if w.strip()]
        sys_whys = [w for w in st.session_state.get("d5_sys_whys", []) if w.strip()]

        st.session_state["d5_occ_rc"] = suggest_root_cause(occ_whys) if occ_whys else ""
        st.session_state["d5_det_rc"] = suggest_root_cause(det_whys) if det_whys else ""
        st.session_state["d5_sys_rc"] = suggest_root_cause(sys_whys) if sys_whys else ""

        st.text_area(t[lang]["Root_Cause_Occ"], value=st.session_state["d5_occ_rc"], height=80, disabled=True)
        st.text_area(t[lang]["Root_Cause_Det"], value=st.session_state["d5_det_rc"], height=80, disabled=True)
        st.text_area(t[lang]["Root_Cause_Sys"], value=st.session_state["d5_sys_rc"], height=80, disabled=True)
        # ---------------------------
# PART 3: D6, D7, D8 rendering (separate Occ/Det/Sys answers) 
# ---------------------------

# Helper: find index of step in npqp_steps (should exist from Part 1)
step_index_map = {step: idx for idx, (step, _, _) in enumerate(npqp_steps)}

# D6 tab
if "D6" in step_index_map:
    with tabs[step_index_map["D6"]]:
        st.markdown(f"### {t[lang]['D6']}")
        guidance = dict(npqp_steps)["D6"]["en"] if lang == "English" else dict(npqp_steps)["D6"]["es"]
        example = dict(npqp_steps)["D6"][2]["en"] if lang == "English" else dict(npqp_steps)["D6"][2]["es"]
        st.markdown(f"<div class='section-box'><b>{t[lang]['Training_Guidance']}:</b> {guidance}<br><br>💡 <b>{t[lang]['Example']}:</b> {example}</div>", unsafe_allow_html=True)

        # Separate text areas for Occurrence / Detection / Systemic corrective actions
        st.session_state.setdefault("D6_occ_answer", st.session_state.get("D6_occ_answer", ""))
        st.session_state["D6_occ_answer"] = st.text_area("D6 - Corrective Actions for Occurrence Root Cause", value=st.session_state["D6_occ_answer"], key="D6_occ_text", height=120)

        st.session_state.setdefault("D6_det_answer", st.session_state.get("D6_det_answer", ""))
        st.session_state["D6_det_answer"] = st.text_area("D6 - Corrective Actions for Detection Root Cause", value=st.session_state["D6_det_answer"], key="D6_det_text", height=120)

        st.session_state.setdefault("D6_sys_answer", st.session_state.get("D6_sys_answer", ""))
        st.session_state["D6_sys_answer"] = st.text_area("D6 - Corrective Actions for Systemic Root Cause", value=st.session_state["D6_sys_answer"], key="D6_sys_text", height=120)

# D7 tab
if "D7" in step_index_map:
    with tabs[step_index_map["D7"]]:
        st.markdown(f"### {t[lang]['D7']}")
        guidance = dict(npqp_steps)["D7"]["en"] if lang == "English" else dict(npqp_steps)["D7"]["es"]
        example = dict(npqp_steps)["D7"][2]["en"] if lang == "English" else dict(npqp_steps)["D7"][2]["es"]
        st.markdown(f"<div class='section-box'><b>{t[lang]['Training_Guidance']}:</b> {guidance}<br><br>💡 <b>{t[lang]['Example']}:</b> {example}</div>", unsafe_allow_html=True)

        st.session_state.setdefault("D7_occ_answer", st.session_state.get("D7_occ_answer", ""))
        st.session_state["D7_occ_answer"] = st.text_area("D7 - Confirmation for Occurrence Root Cause", value=st.session_state["D7_occ_answer"], key="D7_occ_text", height=120)

        st.session_state.setdefault("D7_det_answer", st.session_state.get("D7_det_answer", ""))
        st.session_state["D7_det_answer"] = st.text_area("D7 - Confirmation for Detection Root Cause", value=st.session_state["D7_det_answer"], key="D7_det_text", height=120)

        st.session_state.setdefault("D7_sys_answer", st.session_state.get("D7_sys_answer", ""))
        st.session_state["D7_sys_answer"] = st.text_area("D7 - Confirmation for Systemic Root Cause", value=st.session_state["D7_sys_answer"], key="D7_sys_text", height=120)

# D8 tab
if "D8" in step_index_map:
    with tabs[step_index_map["D8"]]:
        st.markdown(f"### {t[lang]['D8']}")
        guidance = dict(npqp_steps)["D8"]["en"] if lang == "English" else dict(npqp_steps)["D8"]["es"]
        example = dict(npqp_steps)["D8"][2]["en"] if lang == "English" else dict(npqp_steps)["D8"][2]["es"]
        st.markdown(f"<div class='section-box'><b>{t[lang]['Training_Guidance']}:</b> {guidance}<br><br>💡 <b>{t[lang]['Example']}:</b> {example}</div>", unsafe_allow_html=True)

        st.session_state.setdefault("D8_answer", st.session_state.get("D8_answer", ""))
        st.session_state["D8_answer"] = st.text_area("D8 - Follow-up Activities / Lessons Learned", value=st.session_state["D8_answer"], key="D8_text", height=200)
