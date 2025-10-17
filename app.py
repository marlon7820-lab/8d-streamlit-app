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

# ---------------------------
# Sidebar: Smart Session Reset Button
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
# D6/D7 separate answers
st.session_state.setdefault("D6", {"occ_answer":"", "det_answer":"", "sys_answer":""})
st.session_state.setdefault("D7", {"occ_answer":"", "det_answer":"", "sys_answer":""})

# ---------------------------
# D5 categories
# ---------------------------
# (same as your original code: occurrence_categories, detection_categories, systemic_categories)
# ---------------------------

# ---------------------------
# Helper: Suggest root cause based on whys
# ---------------------------
# (same as original)

# ---------------------------
# Helper: Render 5-Why dropdowns without repeating selections
# ---------------------------
# (same as original)

# ---------------------------
# Render Tabs D1–D8
# ---------------------------
tab_labels = [f"🟢 {t[lang_key][step]}" if st.session_state[step].get("answer","").strip() else f"🔴 {t[lang_key][step]}" for step, _, _ in npqp_steps]
tabs = st.tabs(tab_labels)

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        st.markdown(f"""
        <div style=" background-color:#b3e0ff; color:black; padding:12px; border-left:5px solid #1E90FF; border-radius:6px; width:100%; font-size:14px; line-height:1.5; ">
        <b>{t[lang_key]['Training_Guidance']}:</b> {note_dict[lang_key]}<br><br>
        💡 <b>{t[lang_key]['Example']}:</b> {example_dict[lang_key]}
        </div>
        """, unsafe_allow_html=True)

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
            if st.button("➕ Add another Occurrence Why"):
                st.session_state.d5_occ_whys.append("")
            st.markdown("#### Detection Analysis")
            render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories, t[lang_key]['Detection_Why'])
            if st.button("➕ Add another Detection Why"):
                st.session_state.d5_det_whys.append("")
            st.markdown("#### Systemic Analysis")
            render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories, t[lang_key]['Systemic_Why'])
            if st.button("➕ Add another Systemic Why"):
                st.session_state.d5_sys_whys.append("")
            # Dynamic Root Causes
            occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
            det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
            sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]
            st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet", height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet", height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet", height=80, disabled=True)

        # D6
        elif step == "D6":
            st.session_state[step]["occ_answer"] = st.text_area(
                "Corrective Actions for Occurrence Root Cause",
                value=st.session_state[step]["occ_answer"],
                key="d6_occ"
            )
            st.session_state[step]["det_answer"] = st.text_area(
                "Corrective Actions for Detection Root Cause",
                value=st.session_state[step]["det_answer"],
                key="d6_det"
            )
            st.session_state[step]["sys_answer"] = st.text_area(
                "Corrective Actions for Systemic Root Cause",
                value=st.session_state[step]["sys_answer"],
                key="d6_sys"
            )

        # D7
        elif step == "D7":
            st.session_state[step]["occ_answer"] = st.text_area(
                "Countermeasure Confirmation for Occurrence Root Cause",
                value=st.session_state[step]["occ_answer"],
                key="d7_occ"
            )
            st.session_state[step]["det_answer"] = st.text_area(
                "Countermeasure Confirmation for Detection Root Cause",
                value=st.session_state[step]["det_answer"],
                key="d7_det"
            )
            st.session_state[step]["sys_answer"] = st.text_area(
                "Countermeasure Confirmation for Systemic Root Cause",
                value=st.session_state[step]["sys_answer"],
                key="d7_sys"
            )

        # D1-D3, D8
        else:
            st.session_state[step]["answer"] = st.text_area(
                "Your Answer",
                value=st.session_state[step]["answer"],
                key=f"ans_{step}"
            )

# ---------------------------
# Sidebar Backup/Restore/Reset
# ---------------------------
# (same as original)

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = []
occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]
occ_rc_text = suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet"
det_rc_text = suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet"
sys_rc_text = suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet"

for step, _, _ in npqp_steps:
    if step == "D4":
        location = st.session_state[step].get("location", "")
        status = st.session_state[step].get("status", "")
        data_rows.append((step, st.session_state[step]["answer"], f"Location: {location} | Status: {status}"))
    elif step == "D5":
        data_rows.append(("D5 - Root Cause (Occurrence)", occ_rc_text, " | ".join(occ_whys)))
        data_rows.append(("D5 - Root Cause (Detection)", det_rc_text, " | ".join(det_whys)))
        data_rows.append(("D5 - Root Cause (Systemic)", sys_rc_text, " | ".join(sys_whys)))
    elif step == "D6":
        data_rows.append((f"{step} - Occurrence", st.session_state[step]["occ_answer"], ""))
        data_rows.append((f"{step} - Detection", st.session_state[step]["det_answer"], ""))
        data_rows.append((f"{step} - Systemic", st.session_state[step]["sys_answer"], ""))
    elif step == "D7":
        data_rows.append((f"{step} - Occurrence", st.session_state[step]["occ_answer"], ""))
        data_rows.append((f"{step} - Detection", st.session_state[step]["det_answer"], ""))
        data_rows.append((f"{step} - Systemic", st.session_state[step]["sys_answer"], ""))
    else:
        data_rows.append((step, st.session_state[step]["answer"], st.session_state[step].get("extra", "")))

# ---------------------------
# Generate Excel
# ---------------------------
# (same as original)

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
