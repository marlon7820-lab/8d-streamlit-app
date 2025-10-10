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
# Initialize session state for D5 whys
# ---------------------------
st.session_state.setdefault("d5_occ_whys", [""])
st.session_state.setdefault("d5_det_whys", [""])
st.session_state.setdefault("d5_sys_whys", [""])
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# Language dictionary
# ---------------------------
t = {
    "en": {
        "D1": "D1: Concern Details", "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis", "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis", "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)", "Root_Cause_Det": "Root Cause (Detection)", "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why", "Systemic_Why": "Systemic Why",
        "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)", "Root_Cause_Det": "Causa raíz (Detección)", "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección", "Systemic_Why": "Por qué Sistémico",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo"
    }
}

# ---------------------------
# NPQP steps
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."}, {"en":"Customer reported static noise", "es":"El cliente reportó ruido estático"}),
    ("D2", {"en":"Check for similar parts, models, generic parts.", "es":"Verifique partes similares, modelos, partes genéricas."}, {"en":"Similar model radio", "es":"Radio de modelo similar"}),
    ("D3", {"en":"Perform initial investigation to identify issues.", "es":"Realice investigación inicial para identificar problemas."}, {"en":"Visual inspection of solder joints", "es":"Inspección visual de soldaduras"}),
    ("D4", {"en":"Define temporary containment actions.", "es":"Defina acciones de contención temporales."}, {"en":"100% inspection before shipment", "es":"Inspección 100% antes del envío"}),
    ("D5", {"en":"Use 5-Why analysis to determine root cause.", "es":"Use análisis de 5 Porqués para determinar causa raíz."}, {"en":"", "es":""}),
    ("D6", {"en":"Define permanent corrective actions.", "es":"Defina acciones correctivas permanentes."}, {"en":"Update soldering process", "es":"Actualizar proceso de soldadura"}),
    ("D7", {"en":"Verify corrective actions effectiveness.", "es":"Verifique efectividad de acciones correctivas."}, {"en":"Functional tests on corrected amplifiers", "es":"Pruebas funcionales en amplificadores corregidos"}),
    ("D8", {"en":"Document lessons learned.", "es":"Documente lecciones aprendidas."}, {"en":"Update SOPs, PFMEA", "es":"Actualizar SOPs, PFMEA"})
]

# ---------------------------
# Initialize session state for each step
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

# ---------------------------
# Rule-based Root Cause Suggestion
# ---------------------------
def suggest_root_cause(whys):
    text = " ".join(whys).lower()
    if any(word in text for word in ["training", "knowledge", "operator", "human error"]):
        return "Lack of proper training / knowledge gap"
    if any(word in text for word in ["equipment", "tool", "machine", "maintenance", "fixture"]):
        return "Equipment, tooling, or maintenance issue"
    if any(word in text for word in ["procedure", "process", "standard", "sop"]):
        return "Procedure or process not followed"
    if any(word in text for word in ["communication", "handover", "reporting"]):
        return "Poor communication or unclear info flow"
    if any(word in text for word in ["material", "supplier", "component"]):
        return "Material, supplier, or logistics issue"
    if any(word in text for word in ["design", "drawing", "engineering"]):
        return "Design or engineering issue"
    if any(word in text for word in ["management", "oversight", "resource"]):
        return "Management or resource-related issue"
    if any(word in text for word in ["temperature", "humidity", "environment", "contamination"]):
        return "Environmental or external factor"
    return "Systemic issue identified from analysis"

# ---------------------------
# Render D1–D4 tabs
# ---------------------------
tab_labels = []
for step, _, _ in npqp_steps:
    tab_labels.append(f"🔴 {t[lang_key][step]}" if st.session_state[step]["answer"].strip()=="" else f"🟢 {t[lang_key][step]}")

tabs = st.tabs(tab_labels)

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    if step not in ["D5","D6","D7","D8"]:
        with tabs[i]:
            st.markdown(f"### {t[lang_key][step]}")
            st.markdown(f"<div style='background-color:#b3e0ff; padding:12px; border-left:5px solid #1E90FF;'>{note_dict[lang_key]}<br>💡 {example_dict[lang_key]}</div>", unsafe_allow_html=True)
            st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")
            # ---------------------------
# Initialize D5 Why lists
# ---------------------------
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("d5_sys_whys", [""] * 5)

# ---------------------------
# Expanded categories for dropdowns
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
    "Human / Training": [
        "Operator not properly trained for this task",
        "Human error due to fatigue or poor ergonomics",
        "Workload too high causing mistakes",
        "No standard method or checklist followed",
        "Incorrect rework or repair method used",
        "Training records not up to date"
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

# ---------------------------
# Render D5 tab
# ---------------------------
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    if step == "D5":
        with tabs[i]:
            st.markdown(f"### {t[lang_key][step]}")
            st.markdown(f"<div style='background-color:#b3e0ff; padding:12px; border-left:5px solid #1E90FF;'>{note_dict[lang_key]}</div>", unsafe_allow_html=True)
            # ---------------------------
# Function: Render 5-Why dropdowns without repeating selections
# ---------------------------
def render_whys_no_repeat(why_list, categories, label_prefix):
    """
    Render a list of Why dropdowns with options grouped by categories.
    Once an option is selected in one dropdown, it is removed from the remaining dropdowns.
    """
    # Flatten options with category prefix
    all_options = [f"{cat}: {item}" for cat, items in categories.items() for item in items]

    for idx in range(len(why_list)):
        # Compute available options for this dropdown (exclude selections already made)
        selected_values = [v for i, v in enumerate(why_list) if i != idx and v.strip()]
        available_options = [""] + [opt for opt in all_options if opt not in selected_values]

        # Ensure current value is in available_options
        current_val = why_list[idx]
        if current_val not in available_options and current_val.strip():
            available_options.append(current_val)

        # Render the selectbox
        why_list[idx] = st.selectbox(
            f"{label_prefix} {idx+1}",
            options=available_options,
            index=available_options.index(current_val) if current_val in available_options else 0,
            key=f"{label_prefix}_{idx}"
        )

        # Optional free text entry
        free_text = st.text_input(f"Or enter your own {label_prefix} {idx+1}", value=why_list[idx], key=f"{label_prefix}_txt_{idx}")
        if free_text.strip():
            why_list[idx] = free_text
            # ---------------------------
# Capture D5 root causes after user selections
# ---------------------------
occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]

occ_rc_text = suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet"
det_rc_text = suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet"
sys_rc_text = suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet"

# Display dynamic read-only root causes
st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=occ_rc_text, height=80, disabled=True)
st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=det_rc_text, height=80, disabled=True)
st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=sys_rc_text, height=80, disabled=True)

# ---------------------------
# Collect all answers for Excel export
# ---------------------------
data_rows = []
for step, _, _ in npqp_steps:
    answer = st.session_state[step]["answer"]
    extra = st.session_state[step].get("extra", "")

    if step == "D5":
        data_rows.append(("D5 - Root Cause (Occurrence)", occ_rc_text, " | ".join(occ_whys)))
        data_rows.append(("D5 - Root Cause (Detection)", det_rc_text, " | ".join(det_whys)))
        data_rows.append(("D5 - Root Cause (Systemic)", sys_rc_text, " | ".join(sys_whys)))
    else:
        data_rows.append((step, answer, extra))

# ---------------------------
# Function to generate Excel
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
    for step, answer, extra in data_rows:
        ws.append([t[lang_key].get(step, step), answer, extra])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font = Font(bold=True if c == 2 else False)
            cell.border = border

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------
# Download button
# ---------------------------
st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
# ---------------------------
# Sidebar: JSON Backup / Restore
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore")

    # Function to generate JSON backup
    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
        mime="application/json"
    )

    st.markdown("---")
    st.markdown("### Restore from JSON")

    # File uploader for restoring JSON backup
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
# Footer
# ---------------------------
st.markdown("<hr style='border:1px solid #1E90FF;'>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; font-size:12px; color:#555555;'>End of 8D Report Assistant</p>", unsafe_allow_html=True)}
