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
# persist the user's language in session_state so resets can preserve it
if "lang" not in st.session_state:
    st.session_state["lang"] = "English"
lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"], index=0 if st.session_state["lang"] == "English" else 1)
st.session_state["lang"] = lang
lang_key = "en" if lang == "English" else "es"
st.session_state["lang_key"] = lang_key

# ---------------------------
# Helper: clear/reset 8D session (preserve language optionally)
# ---------------------------
def clear_8d_session(preserve_keys=None):
    """Clear all 8D-related session state and reinitialize defaults.
       preserve_keys: list of keys to keep (e.g. ['lang', 'lang_key'])
    """
    if preserve_keys is None:
        preserve_keys = ["lang", "lang_key"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    # clear everything
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    # restore preserved keys
    for k, v in preserved.items():
        st.session_state[k] = v

    # Reinitialize core session keys (same defaults you had)
    for step, _, _ in npqp_steps:
        st.session_state[step] = {"answer": "", "extra": ""}
    st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
    st.session_state.setdefault("prepared_by", "")
    st.session_state.setdefault("d5_occ_whys", [""]*5)
    st.session_state.setdefault("d5_det_whys", [""]*5)
    st.session_state.setdefault("d5_sys_whys", [""]*5)
    st.session_state.setdefault("d4_location", "")
    st.session_state.setdefault("d4_status", "")
    st.session_state.setdefault("d4_containment", "")
    # keep language selections if preserved
    # done by restored preserved dict above

# ---------------------------
# Smart Session Reset Button (top of sidebar)
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")

if st.sidebar.button("🔄 Reset 8D Session"):
    # preserve language
    clear_8d_session(preserve_keys=["lang", "lang_key"])
    st.sidebar.success("Session data cleared. All 8D inputs reset, but language preserved.")
    st.experimental_rerun()

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
# Initialize session state (only if missing)
# ---------------------------
if "initialized_8d" not in st.session_state:
    # base step answers
    for step, _, _ in npqp_steps:
        st.session_state[step] = {"answer": "", "extra": ""}
    st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
    st.session_state.setdefault("prepared_by", "")
    st.session_state.setdefault("d5_occ_whys", [""]*5)
    st.session_state.setdefault("d5_det_whys", [""]*5)
    st.session_state.setdefault("d5_sys_whys", [""]*5)
    st.session_state.setdefault("d4_location", "")
    st.session_state.setdefault("d4_status", "")
    st.session_state.setdefault("d4_containment", "")
    st.session_state["initialized_8d"] = True

# ---------------------------
# Expanded categories for D5
# ---------------------------
occurrence_categories = { ... }  # identical to your original lists (omitted here for brevity)
detection_categories = { ... }
systemic_categories = { ... }

# (To keep the example concise here, in your real file keep the original lists exactly as you had them above.)

# ---------------------------
# Helper: Suggest root cause based on whys
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
# Helper: Render 5-Why dropdowns without repeating selections (unique keys)
# ---------------------------
def render_whys_no_repeat(why_list, categories, label_prefix, key_prefix):
    """
    why_list: list reference in session_state to be modified
    categories: dict of category -> list(items)
    label_prefix: human label for the selectboxes ("Occurrence Why", etc.)
    key_prefix: unique short string used to build widget keys (e.g. "d5_occ")
    """
    for idx in range(len(why_list)):
        # Gather all selected values except current index
        selected_so_far = [w for i, w in enumerate(why_list) if w.strip() and i != idx]

        # Build options excluding already selected
        options = [""] + [f"{cat}: {item}" for cat, items in categories.items() for item in items if f"{cat}: {item}" not in selected_so_far]

        current_val = why_list[idx] if why_list[idx] in options else ""
        why_list[idx] = st.selectbox(
            f"{label_prefix} {idx+1}",
            options,
            index=options.index(current_val) if current_val in options else 0,
            key=f"{key_prefix}_select_{idx}"
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

        # D4 special: Nissan-style fields
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
        # D5 special: 5-Why dropdowns + dynamic root causes
        elif step == "D5":
            st.markdown("#### Occurrence Analysis")
            render_whys_no_repeat(st.session_state.d5_occ_whys, occurrence_categories, t[lang_key]['Occurrence_Why'], key_prefix="d5_occ")
            if st.button("➕ Add another Occurrence Why", key="add_occ"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### Detection Analysis")
            render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories, t[lang_key]['Detection_Why'], key_prefix="d5_det")
            if st.button("➕ Add another Detection Why", key="add_det"):
                st.session_state.d5_det_whys.append("")

            st.markdown("#### Systemic Analysis")
            render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories, t[lang_key]['Systemic_Why'], key_prefix="d5_sys")
            if st.button("➕ Add another Systemic Why", key="add_sys"):
                st.session_state.d5_sys_whys.append("")

            # Dynamic Root Cause Suggestions (read-only)
            occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
            det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
            sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]

            st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet", height=80, disabled=True, key="rc_occ")
            st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet", height=80, disabled=True, key="rc_det")
            st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet", height=80, disabled=True, key="rc_sys")
        # D6–D8: text areas
        else:
            st.session_state[step]["answer"] = st.text_area(
                "Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}"
            )

# ---------------------------
# Sidebar: JSON Backup / Restore + Session Reset
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore / Reset")
    
    # JSON Backup
    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
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
            st.experimental_rerun()
        except Exception as e:
            st.error(f"Error restoring JSON: {e}")

    # Session Reset (full but preserve language)
    if st.button("🧹 Reset Session"):
        clear_8d_session(preserve_keys=["lang", "lang_key"])
        st.experimental_rerun()

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = []

# D5 whys
occ_whys = [w for w in st.session_state.d5_occ_whys if w.strip()]
det_whys = [w for w in st.session_state.d5_det_whys if w.strip()]
sys_whys = [w for w in st.session_state.d5_sys_whys if w.strip()]

occ_rc_text = suggest_root_cause(occ_whys) if occ_whys else "No occurrence whys provided yet"
det_rc_text = suggest_root_cause(det_whys) if det_whys else "No detection whys provided yet"
sys_rc_text = suggest_root_cause(sys_whys) if sys_whys else "No systemic whys provided yet"

for step, _, _ in npqp_steps:
    answer = st.session_state[step]["answer"]
    extra = st.session_state[step].get("extra", "")

    if step == "D4":
        # Include Location, Status, and Containment
        location = st.session_state[step].get("location", "")
        status = st.session_state[step].get("status", "")
        extra_text = f"Location: {location} | Status: {status}"
        data_rows.append((step, answer, extra_text))
    elif step == "D5":
        data_rows.append(("D5 - Root Cause (Occurrence)", occ_rc_text, " | ".join(occ_whys)))
        data_rows.append(("D5 - Root Cause (Detection)", det_rc_text, " | ".join(det_whys)))
        data_rows.append(("D5 - Root Cause (Systemic)", sys_rc_text, " | ".join(sys_whys)))
    else:
        data_rows.append((step, answer, extra))

# ---------------------------
# Generate Excel
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

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
