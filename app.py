# --------------------------- Part 1/3 ---------------------------
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
# App colors and styles (neutral buttons)
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
    /* neutral / light buttons so they are readable on mobile */
    div.stButton > button, button[kind="primary"] {
        background-color: #e6e9ec !important;
        color: #111 !important;
        border: 1px solid #cfcfcf !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# Title & Version
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

version_number = "v1.0.8"
last_updated = "September 15, 2025"

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
        "Training_Guidance": "Training Guidance", "Example": "Example"
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
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo"
    }
}

# ---------------------------
# NPQP 8D steps (keep examples updated as you requested)
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
            "es":"Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de soporte."},
     {"en":"Customer reported static noise in amplifier during end-of-line test.",
      "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.",
            "es":"Verifique partes similares, modelos, partes genéricas, otros colores, mano opuesta, frente/trasero, etc."},
     {"en":"Similar model radio — front vs. rear speaker; for amplifiers consider 8, 12, or 24 channels as applicable.",
      "es":"Radio de modelo similar — altavoz delantero vs trasero; para amplificadores considere 8, 12 o 24 canales según corresponda."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
            "es":"Realice una investigación inicial para identificar problemas evidentes, recopile datos y documente hallazgos iniciales."},
     {"en":"Visual inspection of solder joints, initial functional tests, checking connectors, etc.",
      "es":"Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores, etc."}),
    ("D4", {"en":"Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
            "es":"Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes."},
     {"en":"100% inspection of amplifiers before shipment; temporary shielding.",
      "es":"Inspección 100% de amplificadores antes del envío; blindaje temporal."}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection. Each why helps drill down to the underlying process, equipment, material, or FMEA gaps.",
            "es":"Use el análisis de 5 Porqués para determinar la causa raíz. Separe Ocurrencia y Detección. Cada porqué ayuda a identificar problemas en el proceso, equipo, material o brechas en FMEA."},
     {"en":"","es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
            "es":"Defina acciones correctivas que eliminen la causa raíz permanentemente y eviten recurrencia."},
     {"en":"Update soldering process, redesign fixture, improve component handling.",
      "es":"Actualizar proceso de soldadura, rediseñar herramienta, mejorar manejo de componentes."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue long-term.",
            "es":"Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo."},
     {"en":"Functional tests on corrected amplifiers, accelerated life testing.",
      "es":"Pruebas funcionales en amplificadores corregidos, pruebas de vida aceleradas."}),
    ("D8", {"en":"Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
            "es":"Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y capacitación para prevenir recurrencia."},
     {"en":"Update SOPs, PFMEA, work instructions, and maintenance procedures.",
      "es":"Actualizar SOPs, PFMEA, instrucciones de trabajo y procedimientos de mantenimiento."})
]

# ---------------------------
# Initialize session state (preserve past behavior)
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
# default D5 arrays (start with 5 whys)
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
# keep track of selected suggestions (to avoid duplicates in dropdown)
st.session_state.setdefault("d5_occ_selected", [""] * len(st.session_state.d5_occ_whys))
st.session_state.setdefault("d5_det_selected", [""] * len(st.session_state.d5_det_whys))
# place to persist root cause text (editable)
st.session_state.setdefault("root_cause_occ", "")
st.session_state.setdefault("root_cause_det", "")

# ---------------------------
# Restore from URL (st.query_params) - unchanged
# ---------------------------
if "backup" in st.query_params:
    try:
        data = json.loads(st.query_params["backup"][0])
        for k, v in data.items():
            st.session_state[k] = v
    except Exception:
        pass

# ---------------------------
# Report input fields (keeps same keys so values persist)
# ---------------------------
st.subheader(f"{t[lang_key]['Report_Date']}")
st.session_state.report_date = st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state.report_date, key="report_date_input")
st.session_state.prepared_by = st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state.prepared_by, key="prepared_by_input")
# --------------------------- Part 2/3 ---------------------------
# Build tab labels with status indicators (stable ordering)
tab_labels = []
for step, _, _ in npqp_steps:
    indicator = "🟢" if st.session_state[step]["answer"].strip() else "🔴"
    tab_labels.append(f"{indicator} {t[lang_key][step]}")

# create tabs (no key argument — avoids TypeError in some Streamlit versions)
tabs = st.tabs(tab_labels)

# small helper to infer a single professional root cause phrase from list of whys
def infer_occurrence_root(occ_list):
    combined = " | ".join([s.lower() for s in occ_list if s])
    if not combined:
        return "Review occurrence analysis for final root cause."
    # heuristics (ordered)
    if any(k in combined for k in ["spec", "specification", "tolerance", "dimension", "drawing", "out of spec"]):
        return "The root cause that allowed this issue to occur appears to be inadequate or incorrect specifications (specification/drawing/tolerance issue)."
    if any(k in combined for k in ["fmea", "risk priority", "rpn", "fmea failure", "fmea"]):
        return "The root cause that allowed this issue to occur appears to be a gap in the FMEA / risk assessment (controls or failure modes not identified)."
    if any(k in combined for k in ["calibration", "setting", "incorrect settings"]):
        return "The root cause that allowed this issue to occur appears to be incorrect calibration or machine setup."
    if any(k in combined for k in ["material", "component", "wrong material", "impurity", "damage during storage"]):
        return "The root cause that allowed this issue to occur appears to be a material or component quality issue (wrong or defective material)."
    if any(k in combined for k in ["process", "work instruction", "procedure", "process design", "outdated"]):
        return "The root cause that allowed this issue to occur appears to be an inadequate or unclear process / work instruction."
    # fallback - base on last why if available
    last = next((s for s in reversed(occ_list) if s), "")
    return f"The root cause that allowed this issue to occur appears to be: {last}" if last else "Review occurrence analysis for final root cause."

def infer_detection_root(det_list):
    combined = " | ".join([s.lower() for s in det_list if s])
    if not combined:
        return "Review detection analysis for final root cause."
    if any(k in combined for k in ["qa checklist incomplete", "missed inspection", "no automated test", "inspection documentation missing", "tooling or equipment inspection not scheduled"]):
        return "The root cause that allowed this issue to escape detection appears to be insufficient inspection/detection controls (checklist/tests not defined or executed)."
    if any(k in combined for k in ["validation", "design verification", "insufficient validation"]):
        return "The root cause that allowed this issue to escape detection appears to be incomplete validation or design verification activities."
    if any(k in combined for k in ["documentation", "outdated", "missing"]):
        return "The root cause that allowed this issue to escape detection appears to be inadequate or outdated inspection documentation and records."
    last = next((s for s in reversed(det_list) if s), "")
    return f"The root cause that allowed this issue to escape detection appears to be: {last}" if last else "Review detection analysis for final root cause."

# Single tab rendering loop (keeps everything together and prevents tab-jump bugs)
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        # non-D5 steps: show guidance + fields (unchanged behavior)
        if step != "D5":
            note_text = note_dict[lang_key]
            example_text = example_dict[lang_key]
            st.markdown(f"""
            <div style="
                background-color:#b3effb;
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
            # answer and extra text areas (persist via session_state keys)
            st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")
            st.session_state[step]["extra"] = st.text_area("Extra / Notes", value=st.session_state[step]["extra"], key=f"extra_{step}")

        # ---------------------------
        # D5: Final Analysis (Occurrence + Detection 5-Whys)
        # ---------------------------
        else:
            st.markdown(f"""
            <div style="
                background-color:#b3effb;
                color:black;
                padding:12px;
                border-left:5px solid #1E90FF;
                border-radius:6px;
                width:100%;
                font-size:14px;
                line-height:1.5;
            ">
            <b>{t[lang_key]['Training_GuidANCE'] if False else t[lang_key]['Training_Guidance']}:</b> {note_dict[lang_key]}
            </div>
            """, unsafe_allow_html=True)

            # Occurrence categories (includes FMEA / Risk-related)
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
                ],
                "FMEA / Risk-related": [
                    "FMEA failure not detected",
                    "Process controls not defined",
                    "Risk priority number underestimated"
                ]
            }

            st.markdown("#### Occurrence Analysis")
            # dynamic number of occurrence why boxes (start with current length)
            occ_len = len(st.session_state.d5_occ_whys)
            selected_occ_local = []
            # build a flattened suggestions list for removal logic
            all_occ_suggestions = []
            for cat_items in occurrence_categories.values():
                all_occ_suggestions.extend(cat_items)

            for idx in range(occ_len):
                # compute remaining suggestions excluding already selected ones
                already = [x for x in selected_occ_local if x]
                remaining_options = []
                for cat, items in occurrence_categories.items():
                    for item in items:
                        full_item = f"{cat}: {item}"
                        # avoid adding suggestion already chosen in previous boxes
                        if item not in already and (st.session_state.d5_occ_whys[idx] != full_item):
                            remaining_options.append(full_item)
                # ensure current val remains selectable
                current_val = st.session_state.d5_occ_whys[idx] if idx < len(st.session_state.d5_occ_whys) else ""
                if current_val and current_val not in remaining_options:
                    remaining_options.append(current_val)
                remaining_options = [""] + sorted(set(remaining_options))

                col1, col2 = st.columns([0.75, 0.25])
                with col1:
                    # dropdown (suggestions)
                    sel = st.selectbox(f"{t[lang_key]['Occurrence_Why']} {idx+1}", remaining_options, index=remaining_options.index(current_val) if current_val in remaining_options else 0, key=f"occ_select_{idx}")
                with col2:
                    # free text (overrides)
                    free = st.text_input(f"Free text {idx+1}", value=(current_val if (current_val not in remaining_options or current_val=="" ) else ""), key=f"occ_free_{idx}")
                # decide final value: free text takes precedence if non-empty, else selection
                final = free.strip() if free.strip() else (sel.strip() if sel and sel.strip() else "")
                # store back to session state list
                st.session_state.d5_occ_whys[idx] = final
                selected_occ_local.append(final)

            # Add / remove buttons for Occurrence Why (only add as requested)
            add_occ = st.button("➕ Add another Occurrence Why", key="add_occ_why")
            if add_occ:
                st.session_state.d5_occ_whys.append("")
                # extend selected placeholder list as well
                st.session_state.d5_occ_selected.append("")

            # Detection categories
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

            st.markdown("#### Detection Analysis")
            det_len = len(st.session_state.d5_det_whys)
            selected_det_local = []
            for idx in range(det_len):
                remaining_options = []
                for cat, items in detection_categories.items():
                    for item in items:
                        full_item = f"{cat}: {item}"
                        if item not in selected_det_local and (st.session_state.d5_det_whys[idx] != full_item):
                            remaining_options.append(full_item)
                current_val = st.session_state.d5_det_whys[idx] if idx < len(st.session_state.d5_det_whys) else ""
                if current_val and current_val not in remaining_options:
                    remaining_options.append(current_val)
                options_det = [""] + sorted(set(remaining_options))

                col1, col2 = st.columns([0.75, 0.25])
                with col1:
                    sel = st.selectbox(f"{t[lang_key]['Detection_Why']} {idx+1}", options_det, index=options_det.index(current_val) if current_val in options_det else 0, key=f"det_select_{idx}")
                with col2:
                    free = st.text_input(f"Free text {idx+1}", value=(current_val if (current_val not in options_det or current_val=="") else ""), key=f"det_free_{idx}")
                final = free.strip() if free.strip() else (sel.strip() if sel and sel.strip() else "")
                st.session_state.d5_det_whys[idx] = final
                selected_det_local.append(final)

            add_det = st.button("➕ Add another Detection Why", key="add_det_why")
            if add_det:
                st.session_state.d5_det_whys.append("")
                st.session_state.d5_det_selected.append("")

            # Compute suggested root causes using heuristics
            suggested_occ = infer_occurrence_root([w for w in st.session_state.d5_occ_whys if w.strip()])
            suggested_det = infer_detection_root([w for w in st.session_state.d5_det_whys if w.strip()])

            # Provide editable root-cause areas; any edits become the final D5["answer"]
            st.markdown("#### Suggested Root Cause (you can edit this)")
            occ_edit = st.text_area("Occurrence root cause (editable)", value=(st.session_state.get("root_cause_occ") or suggested_occ), key="root_cause_occ_text", height=120)
            det_edit = st.text_area("Detection root cause (editable)", value=(st.session_state.get("root_cause_det") or suggested_det), key="root_cause_det_text", height=120)

            # persist edited text in session for future runs
            st.session_state["root_cause_occ"] = occ_edit
            st.session_state["root_cause_det"] = det_edit

            # Save combined root causes into D5 answer (so Excel will include it under Answer)
            st.session_state["D5"]["answer"] = f"{occ_edit}\n\n{det_edit}"
            # Keep any additional notes in D5["extra"] (not required, but preserved)
            st.session_state["D5"]["extra"] = "\n".join([f"Occ Why {i+1}: {w}" for i,w in enumerate(st.session_state.d5_occ_whys) if w.strip()] + [""] + [f"Det Why {i+1}: {w}" for i,w in enumerate(st.session_state.d5_det_whys) if w.strip()])
            # --------------------------- Part 3/3 ---------------------------
# ---------------------------
# Collect answers for Excel (use D5["answer"] which contains root cause)
# ---------------------------
def generate_excel_bytes():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # logo if available
    if os.path.exists("logo.png"):
        try:
            img = XLImage("logo.png")
            img.width = 140
            img.height = 40
            ws.add_image(img, "A1")
        except Exception:
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

    # append each step, ensuring D5 answer contains the suggested root cause saved above
    for step, _, _ in npqp_steps:
        answer = st.session_state.get(step, {}).get("answer", "")
        extra = st.session_state.get(step, {}).get("extra", "")
        ws.append([t[lang_key][step], answer, extra])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            # make the Answer column bold for readability
            cell.font = Font(bold=True if c == 2 else False)
            cell.border = border

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Download button (visible in main page)
st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel_bytes(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ---------------------------
# Sidebar: JSON Backup / Restore + Reset
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore")

    def generate_json_backup():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        # but keep only the keys we care about (prevent serializing non-json-serializable objects)
        return json.dumps(save_data, indent=4, default=str)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json_backup(),
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
            st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = [""] * 5
        st.session_state["d5_det_selected"] = [""] * 5
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.session_state["root_cause_occ"] = ""
        st.session_state["root_cause_det"] = ""
        st.success("✅ All data has been reset!")
