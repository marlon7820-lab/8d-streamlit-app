# --------------------------- Part 1 ---------------------------
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
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

# ---------------------------
# App colors and styles
# ---------------------------
st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(to right, #f0f8ff, #e6f2ff);
        color: #000000 !important;
    }
    .stTabs [data-baseweb="tab"] {
        font-weight: bold;
        color: #000000 !important;
    }
    textarea {
        background-color: #ffffff !important;
        border: 1px solid #1E90FF !important;
        border-radius: 5px;
        color: #000000 !important;
    }
    .stInfo {
        background-color: #e6f7ff !important;
        border-left: 5px solid #1E90FF !important;
        color: #000000 !important;
    }
    button[kind="primary"] {
        background-color: #87AFC7 !important;
        color: white !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# Main title
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Version info
# ---------------------------
version_number = "v1.0.8"
last_updated = "September 16, 2025"

st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

t = {
    "en": {
        "D1": "D1: Concern Details", "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis", "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis", "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)", "Root_Cause_Det": "Root Cause (Detection)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why",
        "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example",
        "FMEA_Failure": "FMEA Failure Occurrence"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)", "Root_Cause_Det": "Causa raíz (Detección)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA"
    }
}

# ---------------------------
# NPQP 8D steps
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly...", "es":"Describa claramente las preocupaciones..."},{"en":"Example text","es":"Texto de ejemplo"}),
    ("D2", {"en":"Check for similar parts...", "es":"Verifique partes similares..."},{"en":"Example text","es":"Texto de ejemplo"}),
    ("D3", {"en":"Perform an initial investigation...", "es":"Realice una investigación inicial..."},{"en":"Example text","es":"Texto de ejemplo"}),
    ("D4", {"en":"Define temporary containment actions...", "es":"Defina acciones de contención temporales..."},{"en":"Example text","es":"Texto de ejemplo"}),
    ("D5", {"en":"Use 5-Why analysis...", "es":"Use el análisis de 5 Porqués..."},{"en":"","es":""}),
    ("D6", {"en":"Define corrective actions...", "es":"Defina acciones correctivas..."},{"en":"Example text","es":"Texto de ejemplo"}),
    ("D7", {"en":"Verify that corrective actions...", "es":"Verifique que las acciones correctivas..."},{"en":"Example text","es":"Texto de ejemplo"}),
    ("D8", {"en":"Document lessons learned...", "es":"Documente lecciones aprendidas..."},{"en":"Example text","es":"Texto de ejemplo"})
]

# ---------------------------
# Initialize session state (safe)
# ---------------------------
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    for step, _, _ in npqp_steps:
        st.session_state[step] = {"answer": "", "extra": ""}
    st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
    st.session_state["prepared_by"] = ""
    st.session_state["d5_occ_whys"] = [""] * 5
    st.session_state["d5_det_whys"] = [""] * 5
    st.session_state["d5_occ_selected"] = []
    st.session_state["d5_det_selected"] = []
    # --------------------------- Part 2 ---------------------------
# Report info
st.subheader(f"{t[lang_key]['Report_Date']}")
st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state["report_date"], key="report_date")
st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state["prepared_by"], key="prepared_by")

# Tabs with status
tab_labels = []
for step, _, _ in npqp_steps:
    if st.session_state[step]["answer"].strip() != "":
        tab_labels.append(f"🟢 {t[lang_key][step]}")
    else:
        tab_labels.append(f"🔴 {t[lang_key][step]}")
tabs = st.tabs(tab_labels)

# Render D1–D8
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        if step not in ["D5","D6","D7","D8"]:
            st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")

        if step == "D5":
            with st.form(key="d5_form", clear_on_submit=False):
                st.markdown("#### Occurrence Analysis")
                occurrence_categories = {
                    "Machine / Equipment-related": ["Mechanical failure or breakdown","Calibration issues"],
                    "Material / Component-related": ["Wrong material delivered","Material defects"],
                }
                selected_occ = []
                for idx, val in enumerate(st.session_state.d5_occ_whys):
                    options = [""] + sorted([f"{cat}: {item}" for cat, items in occurrence_categories.items() for item in items])
                    st.session_state.d5_occ_whys[idx] = st.selectbox(
                        f"{t[lang_key]['Occurrence_Why']} {idx+1}",
                        options,
                        index=options.index(val) if val in options else 0,
                        key=f"occ_{idx}"
                    )
                    if st.session_state.d5_occ_whys[idx]:
                        selected_occ.append(st.session_state.d5_occ_whys[idx])
                if st.form_submit_button("➕ Add another Occurrence Why", on_click=lambda: st.session_state.d5_occ_whys.append("")):
                    pass
                st.session_state["d5_occ_selected"] = selected_occ

                st.markdown("#### Detection Analysis")
                detection_categories = {
                    "QA / Inspection-related": ["QA checklist incomplete","No automated test"],
                }
                selected_det = []
                for idx, val in enumerate(st.session_state.d5_det_whys):
                    options = [""] + sorted([f"{cat}: {item}" for cat, items in detection_categories.items() for item in items])
                    st.session_state.d5_det_whys[idx] = st.selectbox(
                        f"{t[lang_key]['Detection_Why']} {idx+1}",
                        options,
                        index=options.index(val) if val in options else 0,
                        key=f"det_{idx}"
                    )
                    if st.session_state.d5_det_whys[idx]:
                        selected_det.append(st.session_state.d5_det_whys[idx])
                if st.form_submit_button("➕ Add another Detection Why", on_click=lambda: st.session_state.d5_det_whys.append("")):
                    pass
                st.session_state["d5_det_selected"] = selected_det

                st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=", ".join(selected_occ), key="root_cause_occ")
                st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=", ".join(selected_det), key="root_cause_det")
                # --------------------------- Part 3 ---------------------------
# Collect answers for Excel
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# Generate Excel
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append([t[lang_key]['Report_Date'], st.session_state["report_date"]])
    ws.append([t[lang_key]['Prepared_By'], st.session_state["prepared_by"]])
    ws.append([])

    header_row = ws.max_row + 1
    headers = ["Step", "Answer", "Extra / Notes"]
    fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for step, answer, extra in data_rows:
        ws.append([t[lang_key][step], answer, extra])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state['report_date'].replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# Sidebar: JSON Backup / Restore + Reset
with st.sidebar:
    st.markdown("## Backup / Restore")
    def generate_json():
        return json.dumps({k: v for k, v in st.session_state.items() if not k.startswith("_")}, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state['report_date'].replace(' ', '_')}.json",
        mime="application/json"
    )

    st.markdown("---")
    uploaded_file = st.file_uploader("Upload JSON file to restore", type="json")
    if uploaded_file:
        try:
            restore_data = json.load(uploaded_file)
            for k, v in restore_data.items():
                st.session_state[k] = v
            st.success("✅ Session restored from JSON!")
        except Exception as e:
            st.error(f"Error restoring JSON: {e}")

    st.markdown("---")
    if st.button("🗑️ Clear All"):
        for step, _, _ in npqp_steps:
            st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["d5_occ_whys"] = [""]
        st.session_state["d5_det_whys"] = [""]
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.success("✅ All data has been reset!")
