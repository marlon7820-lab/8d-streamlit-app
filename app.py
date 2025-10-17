import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font
import datetime
import io

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
    st.session_state["_reset_8d_session"] = False
    st.rerun()

# ---------------------------
# Main title and version
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)
version_number = "v1.1.0"
last_updated = "October 17, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar: Language & Reset
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
st.session_state.setdefault("occ_rc", "")
st.session_state.setdefault("det_rc", "")
st.session_state.setdefault("sys_rc", "")

# ---------------------------
# Helper: Suggest root cause
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
    return "Systemic issue identified from analysis"

# ---------------------------
# Render 5-Why dropdowns
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
            key=f"{label_prefix}_{idx}"
        )
        free_text = st.text_input(f"Or enter your own {label_prefix} {idx+1}", value=why_list[idx], key=f"{label_prefix}_txt_{idx}")
        if free_text.strip():
            why_list[idx] = free_text

# ---------------------------
# D5 categories
# ---------------------------
occurrence_categories = {
    "Machine / Equipment": ["Mechanical failure or breakdown","Calibration issues or drift"],
    "Material / Component": ["Wrong material or component used","Supplier provided off-spec component"],
    "Process / Method": ["Incorrect process step sequence","Critical process parameters not controlled"],
    "Design / Engineering": ["Design not robust to real-use conditions","Tolerance stack-up issue not evaluated"],
    "Environmental / External": ["Temperature or humidity out of control range","Contamination or dust affecting product"]
}

detection_categories = {
    "QA / Inspection": ["QA checklist incomplete or not updated","No automated inspection system in place"],
    "Validation / Process": ["Process validation not updated after design/process change"],
    "FMEA / Control Plan": ["Failure mode not captured in PFMEA"],
    "Test / Equipment": ["Test equipment calibration overdue"],
    "Systemic / Organizational": ["Feedback loop from quality incidents not implemented"]
}

systemic_categories = {
    "Management / Organization": ["Inadequate leadership or supervision structure","Insufficient resource allocation"],
    "Process / Procedure": ["Standard Operating Procedures (SOPs) outdated or missing","Process FMEA not reviewed regularly"],
    "Training": ["No defined training matrix or certification tracking"],
    "Supplier / External": ["Supplier not included in 8D or FMEA review process"],
    "Quality System / Feedback": ["Internal audits ineffective or not completed"]
}

# ---------------------------
# Render Tabs D1–D8
# ---------------------------
tab_labels = [f"🟢 {t[lang_key][step]}" if st.session_state[step]["answer"].strip() else f"🔴 {t[lang_key][step]}" for step, _, _ in npqp_steps]
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

        st.session_state[step]["answer"] = st.text_area(f"Your answer for {t[lang_key][step]}", st.session_state[step]["answer"], height=120, key=f"{step}_answer")
        st.session_state[step]["extra"] = st.text_area(f"Extra notes / observations", st.session_state[step]["extra"], height=80, key=f"{step}_extra")

        if step == "D4":
            st.session_state["d4_location"] = st.text_input(f"{t[lang_key]['Location']}", st.session_state["d4_location"])
            st.session_state["d4_status"] = st.text_input(f"{t[lang_key]['Status']}", st.session_state["d4_status"])
            st.session_state["d4_containment"] = st.text_area(f"{t[lang_key]['Containment_Actions']}", st.session_state["d4_containment"], height=80)

# ---------------------------
# Render D5 5-Why Analysis
# ---------------------------
with tabs[4]:
    st.markdown("### D5: 5-Why Analysis")
    st.subheader("Occurrence")
    render_whys_no_repeat(st.session_state.d5_occ_whys, occurrence_categories, "Occurrence Why")
    st.subheader("Detection")
    render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories, "Detection Why")
    st.subheader("Systemic")
    render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories, "Systemic Why")

    st.markdown("**Suggested Root Causes (editable)**")
    st.session_state["occ_rc"] = st.text_area("Occurrence Root Cause", value=suggest_root_cause(st.session_state.d5_occ_whys), height=40, key="occ_rc_txt")
    st.session_state["det_rc"] = st.text_area("Detection Root Cause", value=suggest_root_cause(st.session_state.d5_det_whys), height=40, key="det_rc_txt")
    st.session_state["sys_rc"] = st.text_area("Systemic Root Cause", value=suggest_root_cause(st.session_state.d5_sys_whys), height=40, key="sys_rc_txt")

# ---------------------------
# Top inputs: Report Date / Prepared By
# ---------------------------
st.session_state["report_date"] = st.text_input(t[lang_key]["Report_Date"], value=st.session_state["report_date"])
st.session_state["prepared_by"] = st.text_input(t[lang_key]["Prepared_By"], value=st.session_state["prepared_by"])

# ---------------------------
# Excel Export
# ---------------------------
def export_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    ws["A1"] = "8D Report Assistant"
    ws["A1"].font = Font(size=16, bold=True)
    ws.merge_cells("A1:D1")

    row = 3
    ws[f"A{row}"] = "Report Date"
    ws[f"B{row}"] = st.session_state["report_date"]
    row +=1
    ws[f"A{row}"] = "Prepared By"
    ws[f"B{row}"] = st.session_state["prepared_by"]
    row +=2

    for step, _, _ in npqp_steps:
        ws[f"A{row}"] = step
        ws[f"A{row}"].font = Font(bold=True)
        row +=1
        ws[f"A{row}"] = st.session_state[step]["answer"]
        row +=1
        ws[f"A{row}"] = st.session_state[step]["extra"]
        row +=2
        if step == "D4":
            ws[f"A{row}"] = f"{t[lang_key]['Location']}: {st.session_state['d4_location']}"
            row +=1
            ws[f"A{row}"] = f"{t[lang_key]['Status']}: {st.session_state['d4_status']}"
            row +=1
            ws[f"A{row}"] = f"{t[lang_key]['Containment_Actions']}: {st.session_state['d4_containment']}"
            row +=2
        if step == "D5":
            ws[f"A{row}"] = "Occurrence Why: " + ", ".join([w for w in st.session_state.d5_occ_whys if w.strip()])
            row +=1
            ws[f"A{row}"] = "Detection Why: " + ", ".join([w for w in st.session_state.d5_det_whys if w.strip()])
            row +=1
            ws[f"A{row}"] = "Systemic Why: " + ", ".join([w for w in st.session_state.d5_sys_whys if w.strip()])
            row +=1
            ws[f"A{row}"] = "Suggested Occurrence Root Cause: " + st.session_state["occ_rc"]
            row +=1
            ws[f"A{row}"] = "Suggested Detection Root Cause: " + st.session_state["det_rc"]
            row +=1
            ws[f"A{row}"] = "Suggested Systemic Root Cause: " + st.session_state["sys_rc"]
            row +=2

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    return file_stream

st.sidebar.markdown("---")
st.sidebar.header("💾 Download Report")
excel_file = export_to_excel()
st.sidebar.download_button("📥 Download Excel", excel_file, file_name=f"8D_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

st.markdown("<p style='text-align:center; font-size:12px; color:#555555;'>End of 8D Report Assistant</p>", unsafe_allow_html=True)
