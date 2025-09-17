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
    .stApp { background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000 !important; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; color: #000 !important; }
    textarea { background-color: #fff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000 !important; }
    .stInfo { background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000 !important; }
    .css-1d391kg { color: #1E90FF !important; font-weight: bold !important; }
    button[kind="primary"] { background-color: #87AFC7 !important; color: white !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# Main title & version
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)
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
        "D1":"D1: Concern Details", "D2":"D2: Similar Part Considerations",
        "D3":"D3: Initial Analysis", "D4":"D4: Implement Containment",
        "D5":"D5: Final Analysis", "D6":"D6: Permanent Corrective Actions",
        "D7":"D7: Countermeasure Confirmation", "D8":"D8: Follow-up Activities",
        "Report_Date":"Report Date", "Prepared_By":"Prepared By",
        "Root_Cause_Occ":"Root Cause (Occurrence)", "Root_Cause_Det":"Root Cause (Detection)",
        "Occurrence_Why":"Occurrence Why", "Detection_Why":"Detection Why",
        "Save":"💾 Save 8D Report", "Download":"📥 Download XLSX",
        "Training_Guidance":"Training Guidance", "Example":"Example", "FMEA_Failure":"FMEA Failure Occurrence"
    },
    "es": {
        "D1":"D1: Detalles de la preocupación", "D2":"D2: Consideraciones de partes similares",
        "D3":"D3: Análisis inicial", "D4":"D4: Implementar contención",
        "D5":"D5: Análisis final", "D6":"D6: Acciones correctivas permanentes",
        "D7":"D7: Confirmación de contramedidas", "D8":"D8: Actividades de seguimiento",
        "Report_Date":"Fecha del informe", "Prepared_By":"Preparado por",
        "Root_Cause_Occ":"Causa raíz (Ocurrencia)", "Root_Cause_Det":"Causa raíz (Detección)",
        "Occurrence_Why":"Por qué Ocurrencia", "Detection_Why":"Por qué Detección",
        "Save":"💾 Guardar Informe 8D", "Download":"📥 Descargar XLSX",
        "Training_Guidance":"Guía de Entrenamiento", "Example":"Ejemplo", "FMEA_Failure":"Ocurrencia de falla FMEA"
    }
}

# ---------------------------
# NPQP 8D steps and examples
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."},
           {"en":"Customer reported static noise.", "es":"El cliente reportó ruido estático."}),
    ("D2", {"en":"Check similar parts or models.", "es":"Verifique partes similares o modelos."},
           {"en":"Similar model radio, front vs rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform initial investigation.", "es":"Realice una investigación inicial."},
           {"en":"Visual inspection of solder joints.", "es":"Inspección visual de soldaduras."}),
    ("D4", {"en":"Define temporary containment actions.", "es":"Defina acciones de contención temporales."},
           {"en":"100% inspection of amplifiers before shipment.", "es":"Inspección 100% de amplificadores antes del envío."}),
    ("D5", {"en":"Use 5-Why analysis for root cause.", "es":"Use el análisis de 5 Porqués para causa raíz."},
           {"en":"","es":""}),
    ("D6", {"en":"Define permanent corrective actions.", "es":"Defina acciones correctivas permanentes."},
           {"en":"Update soldering process.", "es":"Actualizar proceso de soldadura."}),
    ("D7", {"en":"Verify effectiveness of corrective actions.", "es":"Verifique la efectividad de las acciones correctivas."},
           {"en":"Functional tests on corrected amplifiers.", "es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned.", "es":"Documente lecciones aprendidas."},
           {"en":"Update SOPs and PFMEA.", "es":"Actualizar SOPs y PFMEA."})
]

# ---------------------------
# Initialize session_state safely
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("d5_occ_selected", [])
st.session_state.setdefault("d5_det_selected", [])
# --------------------------- Part 2 ---------------------------
# ---------------------------
# Restore from URL backup if available
# ---------------------------
if "backup" in st.query_params:
    try:
        data = json.loads(st.query_params["backup"][0])
        for k, v in data.items():
            st.session_state[k] = v
    except Exception:
        pass

# ---------------------------
# Report info
# ---------------------------
st.subheader(f"{t[lang_key]['Report_Date']}")
report_date_input = st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state["report_date"], key="report_date_input")
st.session_state["report_date"] = report_date_input

prepared_by_input = st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state["prepared_by"], key="prepared_by_input")
st.session_state["prepared_by"] = prepared_by_input

# ---------------------------
# Tabs with status indicators
# ---------------------------
tab_labels = []
for step, _, _ in npqp_steps:
    if st.session_state[step]["answer"].strip() != "":
        tab_labels.append(f"🟢 {t[lang_key][step]}")
    else:
        tab_labels.append(f"🔴 {t[lang_key][step]}")

tabs = st.tabs(tab_labels)

# ---------------------------
# Render D1–D5 Tabs
# ---------------------------
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

        # --------------------------- D1–D4 ---------------------------
        if step not in ["D5","D6","D7","D8"]:
            st.session_state[step]["answer"] = st.text_area(
                "Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}"
            )

        # --------------------------- D5 FIXED ---------------------------
        elif step == "D5":
            with st.form(key="d5_form", clear_on_submit=False):
                st.markdown("#### Occurrence Analysis")
                occurrence_categories = {
                    "Machine / Equipment-related": [
                        "Mechanical failure or breakdown",
                        "Calibration issues (incorrect settings)",
                        "Tooling or fixture failure",
                        "Machine wear and tear",
                        "Failure not identified in FMEA"
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

                selected_occ = []
                for idx, val in enumerate(st.session_state["d5_occ_whys"]):
                    remaining_options = []
                    for cat, items in occurrence_categories.items():
                        for item in items:
                            full_item = f"{cat}: {item}"
                            if full_item not in selected_occ:
                                remaining_options.append(full_item)
                    if val and val not in remaining_options:
                        remaining_options.append(val)
                    options = [""] + sorted(remaining_options)
                    current_value = st.session_state["d5_occ_whys"][idx]
                    st.session_state["d5_occ_whys"][idx] = st.selectbox(
                        f"{t[lang_key]['Occurrence_Why']} {idx+1}",
                        options,
                        index=options.index(current_value) if current_value in options else 0,
                        key=f"occ_{idx}"
                    )
                    free_text = st.text_input(f"Or enter your own Occurrence Why {idx+1}", value=st.session_state["d5_occ_whys"][idx], key=f"occ_txt_{idx}")
                    if free_text.strip():
                        st.session_state["d5_occ_whys"][idx] = free_text
                    if st.session_state["d5_occ_whys"][idx]:
                        selected_occ.append(st.session_state["d5_occ_whys"][idx])

                if st.form_submit_button("➕ Add another Occurrence Why", on_click=lambda: st.session_state["d5_occ_whys"].append("")):
                    pass

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

                selected_det = []
                for idx, val in enumerate(st.session_state["d5_det_whys"]):
                    remaining_options = []
                    for cat, items in detection_categories.items():
                        for item in items:
                            full_item = f"{cat}: {item}"
                            if full_item not in selected_det:
                                remaining_options.append(full_item)
                    if val and val not in remaining_options:
                        remaining_options.append(val)
                    options_det = [""] + sorted(remaining_options)
                    current_value = st.session_state["d5_det_whys"][idx]
                    st.session_state["d5_det_whys"][idx] = st.selectbox(
                        f"{t[lang_key]['Detection_Why']} {idx+1}",
                        options_det,
                        index=options_det.index(current_value) if current_value in options_det else 0,
                        key=f"det_{idx}"
                    )
                    free_text_det = st.text_input(f"Or enter your own Detection Why {idx+1}", value=st.session_state["d5_det_whys"][idx], key=f"det_txt_{idx}")
                    if free_text_det.strip():
                        st.session_state["d5_det_whys"][idx] = free_text_det
                    if st.session_state["d5_det_whys"][idx]:
                        selected_det.append(st.session_state["d5_det_whys"][idx])

                if st.form_submit_button("➕ Add another Detection Why", on_click=lambda: st.session_state["d5_det_whys"].append("")):
                    pass

                st.session_state["d5_det_selected"] = selected_det

                # Suggested Root Cause
                suggested_occ_rc = (
                    "The root cause that allowed this issue to occur may be related to: " + ", ".join(selected_occ)
                    if selected_occ else ""
                )
                suggested_det_rc = (
                    "The root cause that allowed this issue to escape detection may be related to: " + ", ".join(selected_det)
                    if selected_det else ""
                )

                st.session_state["D5"]["answer"] = st.text_area(
                    f"{t[lang_key]['Root_Cause_Occ']}",
                    value=suggested_occ_rc,
                    key="root_cause_occ"
                )
                st.text_area(
                    f"{t[lang_key]['Root_Cause_Det']}",
                    value=suggested_det_rc,
                    key="root_cause_det"
                )
                # --------------------------- Part 3 ---------------------------
# ---------------------------
# Render D6–D8 Tabs
# ---------------------------
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    if step in ["D6","D7","D8"]:
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

            st.session_state[step]["answer"] = st.text_area(
                "Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}"
            )

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save / Download Excel
# ---------------------------
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
    file_name=f"8D_Report_{st.session_state['report_date'].replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ---------------------------
# Sidebar: JSON Backup / Restore + Reset
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore")

    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state['report_date'].replace(' ', '_')}.json",
        mime="application/json"
    )

    st.markdown("---")
    st.markdown("### Restore from JSON")

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
    st.markdown("### Reset All Data")

    if st.button("🗑️ Clear All"):
        for step, _, _ in npqp_steps:
            st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["D5"] = {"answer": "", "extra": ""}
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.success("✅ All data has been reset!")
