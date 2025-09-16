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
from typing import List

# Page config
st.set_page_config(page_title="8D Report Assistant", page_icon="logo.png", layout="wide")

# Styles (kept from baseline)
st.markdown("""
    <style>
    .stApp { background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important; }
    .stTabs [data-baseweb="tab"] { font-weight: bold; color: #000000 !important; }
    textarea { background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important; }
    .stInfo { background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important; }
    .css-1d391kg { color: #1E90FF !important; font-weight: bold !important; }
    button[kind="primary"] { background-color: #87AFC7 !important; color: white !important; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# Title & version
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)
version_number = "v1.0.8-merged"
last_updated = "September 16, 2025"
st.markdown(f"<p style='text-align:center;color:#555'>{version_number} | Last updated: {last_updated}</p>", unsafe_allow_html=True)

# Language strings (kept from baseline)
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

t = {
    "en": {
        "D1":"D1: Concern Details","D2":"D2: Similar Part Considerations","D3":"D3: Initial Analysis","D4":"D4: Implement Containment",
        "D5":"D5: Final Analysis","D6":"D6: Permanent Corrective Actions","D7":"D7: Countermeasure Confirmation","D8":"D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date":"Report Date","Prepared_By":"Prepared By","Root_Cause_Occ":"Root Cause (Occurrence)","Root_Cause_Det":"Root Cause (Detection)",
        "Occurrence_Why":"Occurrence Why","Detection_Why":"Detection Why","Save":"💾 Save 8D Report","Download":"📥 Download XLSX",
        "Training_Guidance":"Training Guidance","Example":"Example"
    },
    "es": {
        "D1":"D1: Detalles de la preocupación","D2":"D2: Consideraciones de partes similares","D3":"D3: Análisis inicial","D4":"D4: Implementar contención",
        "D5":"D5: Análisis final","D6":"D6: Acciones correctivas permanentes","D7":"D7: Confirmación de contramedidas","D8":"D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date":"Fecha del informe","Prepared_By":"Preparado por","Root_Cause_Occ":"Causa raíz (Ocurrencia)","Root_Cause_Det":"Causa raíz (Detección)",
        "Occurrence_Why":"Por qué Ocurrencia","Detection_Why":"Por qué Detección","Save":"💾 Guardar Informe 8D","Download":"📥 Descargar XLSX",
        "Training_Guidance":"Guía de Entrenamiento","Example":"Ejemplo"
    }
}

# NPQP steps (exact items kept)
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.","es":"Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de soporte."}, {"en":"Customer reported static noise in amplifier during end-of-line test.","es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.","es":"Verifique partes similares, modelos, partes genéricas, otros colores, mano opuesta, frente/trasero, etc."}, {"en":"Similar model radio, Front vs. rear speaker; for amplifiers consider 8, 12, or 24 channels.","es":"Radio de modelo similar, altavoz delantero vs trasero; para amplificadores considere 8, 12 o 24 canales."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues, collect data, and document initial findings.","es":"Realice una investigación inicial para identificar problemas evidentes, recopile datos y documente hallazgos iniciales."}, {"en":"Visual inspection of solder joints, initial functional tests, checking connectors, etc.","es":"Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores, etc."}),
    ("D4", {"en":"Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.","es":"Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes."}, {"en":"100% inspection of amplifiers before shipment; temporary shielding.","es":"Inspección 100% de amplificadores antes del envío; blindaje temporal."}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection. Include FMEA failure occurrence if applicable.","es":"Use el análisis de 5 Porqués para determinar la causa raíz. Separe Ocurrencia y Detección. Incluya la ocurrencia de falla FMEA si aplica."}, {"en":"","es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently and prevent recurrence.","es":"Defina acciones correctivas que eliminen la causa raíz permanentemente y eviten recurrencia."}, {"en":"Update soldering process, redesign fixture, improve component handling.","es":"Actualizar proceso de soldadura, rediseñar herramienta, mejorar manejo de componentes."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue long-term.","es":"Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo."}, {"en":"Functional tests on corrected amplifiers, accelerated life testing.","es":"Pruebas funcionales en amplificadores corregidos, pruebas de vida aceleradas."}),
    ("D8", {"en":"Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.","es":"Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y capacitación para prevenir recurrencia."}, {"en":"Update SOPs, PFMEA, work instructions, and maintenance procedures.","es":"Actualizar SOPs, PFMEA, instrucciones de trabajo y procedimientos de mantenimiento."})
]

# Safe initialization pattern — use a single flag so we can fully reinitialize on Clear All
if not st.session_state.get("initialized", False):
    # initialize main containers and defaults (only when not initialized)
    for step, _, _ in npqp_steps:
        st.session_state[step] = {"answer": "", "extra": ""}
    st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
    st.session_state["prepared_by"] = ""
    # default 5 Why rows (matches your baseline)
    st.session_state["d5_occ_whys"] = [""] * 5
    st.session_state["d5_det_whys"] = [""] * 5
    st.session_state["d5_occ_selected"] = []
    st.session_state["d5_det_selected"] = []
    st.session_state["initialized"] = True

# small helper
def flatten_categories(cat_dict: dict) -> List[str]:
    out = []
    for cat, items in cat_dict.items():
        for item in items:
            out.append(f"{cat}: {item}")
    return out
    # --------------------------- Part 2 ---------------------------
# Report info inputs — widget keys are "report_date" and "prepared_by"
st.subheader(f"{t[lang_key]['Report_Date']}")
# Use widget keys; do NOT assign the widget's return to st.session_state here (avoid key overwrite)
st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state.get("report_date", ""), key="report_date")
st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state.get("prepared_by", ""), key="prepared_by")

# Build tab labels
def step_answer_nonempty(step_key: str) -> bool:
    widget_key = f"ans_{step_key}"
    # check widget value first (if user typed), else stored answer
    wval = st.session_state.get(widget_key, None)
    if wval is not None:
        return str(wval).strip() != ""
    return str(st.session_state.get(step_key, {}).get("answer", "")).strip() != ""

tab_labels = []
for step, _, _ in npqp_steps:
    tab_labels.append(f"🟢 {t[lang_key][step]}" if step_answer_nonempty(step) else f"🔴 {t[lang_key][step]}")

tabs = st.tabs(tab_labels)

# D5 categories (kept full lists from baseline)
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

all_occ_options = sorted(flatten_categories(occurrence_categories))
all_det_options = sorted(flatten_categories(detection_categories))

# Render tabs
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")

        # D1-D4 simple textareas (unique widget keys)
        if step not in ["D5", "D6", "D7", "D8"]:
            widget_key = f"ans_{step}"
            st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"], key=widget_key)

        # ---------------- D5 ----------------
        if step == "D5":
            st.markdown(f"""
            <div style='background-color:#b3e0ff; padding:12px; border-left:5px solid #1E90FF; border-radius:6px;'>
              <b>{t[lang_key]['Training_Guidance']}:</b> {note_dict[lang_key]}
            </div>
            """, unsafe_allow_html=True)

            # Use a form to avoid widget-triggered reruns while selecting
            with st.form(key="d5_form", clear_on_submit=False):
                st.markdown("#### Occurrence Analysis")
                # render each occurrence row (selectbox + free text)
                for idx in range(len(st.session_state["d5_occ_whys"])):
                    occ_sel_key = f"occ_select_{idx}"
                    occ_txt_key = f"occ_text_{idx}"
                    current_saved = st.session_state["d5_occ_whys"][idx] if idx < len(st.session_state["d5_occ_whys"]) else ""
                    # if current_saved is a known option, preselect it; otherwise show it in text
                    default_text = "" if current_saved in all_occ_options else current_saved
                    default_select = current_saved if current_saved in all_occ_options else ""
                    # build options (include custom if present so user can re-select it)
                    options_for_widget = [""] + all_occ_options
                    if default_text:
                        options_for_widget = [""] + sorted(set(options_for_widget + [default_text]))
                    try:
                        default_index = options_for_widget.index(default_select) if default_select else 0
                    except ValueError:
                        default_index = 0
                    st.selectbox(f"{t[lang_key]['Occurrence_Why']} {idx+1}", options_for_widget, index=default_index, key=occ_sel_key)
                    st.text_input(f"Or enter your own Occurrence Why {idx+1}", value=default_text, key=occ_txt_key)

                if st.form_submit_button("➕ Add another Occurrence Why", on_click=lambda: st.session_state["d5_occ_whys"].append("")):
                    pass

                st.markdown("#### Detection Analysis")
                for idx in range(len(st.session_state["d5_det_whys"])):
                    det_sel_key = f"det_select_{idx}"
                    det_txt_key = f"det_text_{idx}"
                    current_saved = st.session_state["d5_det_whys"][idx] if idx < len(st.session_state["d5_det_whys"]) else ""
                    default_text = "" if current_saved in all_det_options else current_saved
                    default_select = current_saved if current_saved in all_det_options else ""
                    options_for_widget = [""] + all_det_options
                    if default_text:
                        options_for_widget = [""] + sorted(set(options_for_widget + [default_text]))
                    try:
                        default_index = options_for_widget.index(default_select) if default_select else 0
                    except ValueError:
                        default_index = 0
                    st.selectbox(f"{t[lang_key]['Detection_Why']} {idx+1}", options_for_widget, index=default_index, key=det_sel_key)
                    st.text_input(f"Or enter your own Detection Why {idx+1}", value=default_text, key=det_txt_key)

                if st.form_submit_button("➕ Add another Detection Why", on_click=lambda: st.session_state["d5_det_whys"].append("")):
                    pass

                # Save D5 form action
                save_clicked = st.form_submit_button("💾 Save D5")
                if save_clicked:
                    # Commit occurrences
                    final_occ = []
                    for idx in range(len(st.session_state["d5_occ_whys"])):
                        sel = st.session_state.get(f"occ_select_{idx}", "") or ""
                        txt = st.session_state.get(f"occ_text_{idx}", "") or ""
                        chosen = txt.strip() if txt.strip() else sel
                        final_occ.append(chosen)
                    st.session_state["d5_occ_whys"] = final_occ

                    # Commit detections
                    final_det = []
                    for idx in range(len(st.session_state["d5_det_whys"])):
                        sel = st.session_state.get(f"det_select_{idx}", "") or ""
                        txt = st.session_state.get(f"det_text_{idx}", "") or ""
                        chosen = txt.strip() if txt.strip() else sel
                        final_det.append(chosen)
                    st.session_state["d5_det_whys"] = final_det

                    # Build suggested root cause strings
                    selected_occ = [x for x in final_occ if x and x.strip()]
                    selected_det = [x for x in final_det if x and x.strip()]

                    suggested_occ_rc = "The root cause that allowed this issue to occur may be related to: " + ", ".join(selected_occ) if selected_occ else ""
                    suggested_det_rc = "The root cause that allowed this issue to escape detection may be related to: " + ", ".join(selected_det) if selected_det else ""

                    # Save suggestions into D5 slot and root cause fields
                    st.session_state["D5"] = {"answer": suggested_occ_rc, "extra": suggested_det_rc}
                    st.session_state["root_cause_occ"] = suggested_occ_rc
                    st.session_state["root_cause_det"] = suggested_det_rc

                    st.success("✅ D5 saved")

            # show/edit root cause outside form (so editing doesn't trigger D5 form re-submission)
            if "root_cause_occ" not in st.session_state:
                st.session_state["root_cause_occ"] = st.session_state.get("D5", {}).get("answer", "")
            if "root_cause_det" not in st.session_state:
                st.session_state["root_cause_det"] = st.session_state.get("D5", {}).get("extra", "")

            # Allow manual editing of the suggested root causes
            st.session_state["D5"]["answer"] = st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=st.session_state.get("root_cause_occ", ""), key="root_cause_occ")
            st.session_state["D5"]["extra"] = st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=st.session_state.get("root_cause_det", ""), key="root_cause_det")
                    # --------------------------- D6–D8 ---------------------------
        elif step in ["D6","D7","D8"]:
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
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
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
            if step != "D5":
                st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["D5"] = {"answer": "", "extra": ""}
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
            st.session_state.setdefault(step, {"answer":"", "extra":""})
        st.success("✅ All data has been reset!")
