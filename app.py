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
    .stAlert {
        background-color: #e6f7ff !important;
        border-left: 5px solid #1E90FF !important;
        color: #000000 !important;
    }
    button[kind="primary"] {
        background-color: #1E90FF !important;
        color: white !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

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
        "Root_Cause": "Root Cause (summary after 5-Whys)", "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why", "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause": "Causa raíz (resumen después de los 5 Porqués)", "Occurrence_Why": "Por qué Ocurrencia",
        "Detection_Why": "Por qué Detección", "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo"
    }
}

npqp_steps = [
    ("D1", "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.", "Customer reported static noise in amplifier during end-of-line test."),
    ("D2", "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.", "Same speaker type used in another radio model; different amplifier colors."),
    ("D3", "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.", "Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4", "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.", "100% inspection of amplifiers before shipment; temporary shielding."),
    ("D5", "Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection.", ""),
    ("D6", "Define corrective actions that eliminate the root cause permanently and prevent recurrence.", "Update soldering process, retrain operators, update work instructions."),
    ("D7", "Verify that corrective actions effectively resolve the issue long-term.", "Functional tests on corrected amplifiers, accelerated life testing."),
    ("D8", "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.", "Update SOPs, PFMEA, work instructions, and employee training.")
]

for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Report info inputs
# ---------------------------
st.subheader(f"{t[lang_key]['Report_Date']}")
st.session_state.report_date = st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state.prepared_by)

tab_labels = []
for step, _, _ in npqp_steps:
    if st.session_state[step]["answer"].strip() != "":
        tab_labels.append(f"🟢 {t[lang_key][step]}")
    else:
        tab_labels.append(f"🔴 {t[lang_key][step]}")
tabs = st.tabs(tab_labels)
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        if step != "D5":
            st.info(f"**{t[lang_key]['Training_Guidance']}:** {note}\n\n💡 **{t[lang_key]['Example']}:** {example}")
            st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.info(f"**{t[lang_key]['Training_Guidance']}:** {note}")
            occurrence_categories = {
                "Machine / Equipment-related": [
                    "Mechanical failure or breakdown",
                    "Calibration issues (incorrect settings)",
                    "Tooling or fixture failure",
                    "Machine wear and tear"
                ],
                "Material / Component-related": [
                    "Wrong material delivered",
                    "Material defects or impurities",
                    "Damage during storage or transport",
                    "Incorrect specifications or tolerance errors"
                ],
                "Process / Method-related": [
                    "Incorrect process steps due to poor process design",
                    "Inefficient workflow or bottlenecks",
                    "Lack of standardized procedures",
                    "Outdated or incomplete work instructions"
                ],
                "Environmental / External Factors": [
                    "Temperature, humidity, or other environmental conditions",
                    "Power fluctuations or outages",
                    "Contamination (dust, oil, chemicals)",
                    "Regulatory or compliance changes"
                ]
            }

            st.markdown("#### Occurrence Analysis")
            selected_occ = st.session_state.get("d5_occ_selected", [])
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"{t[lang_key]['Occurrence_Why']} {idx+1}", value=val, key=f"occ_{idx}")
                else:
                    remaining_options = []
                    for cat, items in occurrence_categories.items():
                        remaining_options.extend([f"{cat}: {item}" for item in items if f"{cat}: {item}" not in selected_occ])
                    if val and val not in remaining_options:
                        remaining_options.append(val)
                    choice = st.selectbox(f"{t[lang_key]['Occurrence_Why']} {idx+1}", [""] + remaining_options, key=f"occ_{idx}")
                    st.session_state.d5_occ_whys[idx] = choice
                    if choice and choice not in selected_occ:
                        selected_occ.append(choice)
            st.session_state["d5_occ_selected"] = selected_occ

            st.markdown("#### Detection Analysis")
            detection_categories = {
                "QA / Inspection-related": [
                    "QA checklist incomplete",
                    "No automated test",
                    "Missed inspection due to process gap",
                    "Tooling or equipment inspection not scheduled"
                ],
                "Validation / Process-related": [
                    "Insufficient validation steps",
                    "Design verification not complete",
                    "Inspection documentation missing or outdated"
                ]
            }
            selected_det = st.session_state.get("d5_det_selected", [])
            for idx, val in enumerate(st.session_state.d5_det_whys):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"{t[lang_key]['Detection_Why']} {idx+1}", value=val, key=f"det_{idx}")
                else:
                    remaining_det = []
                    for cat, items in detection_categories.items():
                        remaining_det.extend([f"{cat}: {item}" for item in items if f"{cat}: {item}" not in selected_det])
                    if val and val not in remaining_det:
                        remaining_det.append(val)
                    choice = st.selectbox(f"{t[lang_key]['Detection_Why']} {idx+1}", [""] + remaining_det, key=f"det_{idx}")
                    st.session_state.d5_det_whys[idx] = choice
                    if choice and choice not in selected_det:
                        selected_det.append(choice)
            st.session_state["d5_det_selected"] = selected_det

            st.session_state.D5["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state.D5["extra"] = st.text_area(f"{t[lang_key]['Root_Cause']}", value=st.session_state.D5["extra"], key="root_cause")

# ---------------------------
# Collect answers + Excel/JSON logic (unchanged)
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
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

with st.sidebar:
    st.markdown("## Backup / Restore")
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
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.success("✅ All data has been reset!")
