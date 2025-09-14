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
    .css-1d391kg {
        color: #1E90FF !important;
        font-weight: bold !important;
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
# Version info (kept as original)
# ---------------------------
version_number = "v1.0.6"
last_updated = "September 13, 2025"

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
# NPQP 8D steps with updated examples (kept as original)
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
            "es":"Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de soporte."},
     {"en":"Customer reported static noise in amplifier during end-of-line test.",
      "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.",
            "es":"Verifique partes similares, modelos, partes genéricas, otros colores, mano opuesta, frente/trasero, etc."},
     {"en":"Similar model radio, Front vs. rear speaker; for amplifiers consider 8, 12, or 24 channels.",
      "es":"Radio de modelo similar, altavoz delantero vs trasero; para amplificadores considere 8, 12 o 24 canales."}),
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
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
# keep existing d5 lists and selection storage
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("d5_occ_selected", [])
st.session_state.setdefault("d5_det_selected", [])

# Ensure per-widget text storage exists (so free text persists)
for i in range(len(st.session_state["d5_occ_whys"])):
    st.session_state.setdefault(f"occ_txt_{i}", st.session_state["d5_occ_whys"][i] if st.session_state["d5_occ_whys"][i] not in (None, "") else "")
    st.session_state.setdefault(f"occ_sel_{i}", st.session_state["d5_occ_whys"][i] if st.session_state["d5_occ_whys"][i] else "")

for i in range(len(st.session_state["d5_det_whys"])):
    st.session_state.setdefault(f"det_txt_{i}", st.session_state["d5_det_whys"][i] if st.session_state["d5_det_whys"][i] not in (None, "") else "")
    st.session_state.setdefault(f"det_sel_{i}", st.session_state["d5_det_whys"][i] if st.session_state["d5_det_whys"][i] else "")

# ---------------------------
# Restore from URL (st.query_params)
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
st.session_state.report_date = st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state.prepared_by)

# ---------------------------
# Tabs with ✅ / 🔴 status indicators
# ---------------------------
tab_labels = []
for step, _, _ in npqp_steps:
    if st.session_state[step]["answer"].strip() != "":
        tab_labels.append(f"🟢 {t[lang_key][step]}")
    else:
        tab_labels.append(f"🔴 {t[lang_key][step]}")

tabs = st.tabs(tab_labels)
# --------------------------- Part 2 (continue) ---------------------------

# ---------------------------
# helper mapping functions (map selected whys -> professional root cause phrasing)
# ---------------------------
def map_occurrence_root(selected_list):
    # selected_list: list of strings (occurrence why entries)
    joined = " ".join(selected_list).lower()
    # priority checks (more specific items first)
    if any(k in joined for k in ["incorrect specifications", "incorrect spec", "tolerance", "drawing", "dimensions", "out of spec"]):
        return "The root cause that allowed this issue to occur is related to inadequate specifications or drawing errors."
    if "failure not identified in fmea" in joined or "fmea" in joined:
        return "The root cause that allowed this issue to occur is related to an incomplete or insufficient FMEA / risk assessment."
    if any(k in joined for k in ["wrong material", "material defect", "material defects", "supplier"]):
        return "The root cause that allowed this issue to occur is related to material/component quality or incorrect component specification from the supplier."
    if any(k in joined for k in ["calibration", "calibrat", "tooling", "fixture", "setup"]):
        return "The root cause that allowed this issue to occur is related to equipment setup, calibration or tooling issues."
    if any(k in joined for k in ["mechanical failure", "breakdown", "wear and tear", "wear"]):
        return "The root cause that allowed this issue to occur is related to equipment design, maintenance or reliability."
    if any(k in joined for k in ["process", "work instruction", "procedure", "lack of standardized", "outdated", "steps"]):
        return "The root cause that allowed this issue to occur is related to inadequate process control or incomplete/unclear work instructions."
    if any(k in joined for k in ["temperature", "humidity", "contamination", "power fluctuation", "power fluctuations"]):
        return "The root cause that allowed this issue to occur is related to environmental conditions or contamination affecting the process."
    # fallback:
    return "The root cause that allowed this issue to occur is related to a combination of the items identified in the analysis; further detailed investigation is recommended."

def map_detection_root(selected_list):
    joined = " ".join(selected_list).lower()
    if any(k in joined for k in ["qa checklist incomplete", "missed inspection", "inspection not scheduled", "inspection documentation missing", "tooling or equipment inspection not scheduled"]):
        return "The root cause that allowed this issue to escape detection is related to insufficient inspection planning, incomplete QA checks, or gaps in inspection scheduling."
    if any(k in joined for k in ["no automated test", "insufficient validation", "design verification not complete", "validation steps"]):
        return "The root cause that allowed this issue to escape detection is related to inadequate validation/verification and lack of appropriate automated testing."
    if "inspection documentation missing" in joined or "documentation" in joined or "outdated" in joined:
        return "The root cause that allowed this issue to escape detection is related to outdated or missing inspection/validation documentation."
    # fallback:
    return "The root cause that allowed this issue to escape detection is related to gaps in the detection system (inspection/validation) identified during analysis."

# ---------------------------
# Render all tabs (single loop to prevent tab reset issues)
# ---------------------------
for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")

        # Non-D5 standard steps
        if step not in ["D5", "D6", "D7", "D8"]:
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

        # D5: 5-Why Occurrence & Detection with free text + dropdown + add-more + suggested root cause
        elif step == "D5":
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
            <b>{t[lang_key]['Training_Guidance']}:</b> {note_dict[lang_key]}
            </div>
            """, unsafe_allow_html=True)

            # Occurrence choices
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

            # ensure d5 lists exist
            if "d5_occ_whys" not in st.session_state:
                st.session_state["d5_occ_whys"] = [""] * 5
            if "d5_det_whys" not in st.session_state:
                st.session_state["d5_det_whys"] = [""] * 5

            selected_occ = []
            # iterate existing occurrence why slots
            for idx in range(len(st.session_state["d5_occ_whys"])):
                # build available options excluding already-chosen items
                remaining_options = []
                for cat, items in occurrence_categories.items():
                    for item in items:
                        full_item = f"{cat}: {item}"
                        # exclude items already selected in earlier slots
                        if full_item not in selected_occ and full_item not in st.session_state["d5_occ_whys"]:
                            remaining_options.append(full_item)
                # keep current value even if it would otherwise be excluded
                current_val = st.session_state["d5_occ_whys"][idx]
                if current_val and current_val not in remaining_options:
                    remaining_options.append(current_val)

                options = [""] + sorted(remaining_options)

                # decide initial index (use existing selection if it exists and is in options)
                default_choice = st.session_state.get(f"occ_sel_{idx}", current_val if current_val in options else "")
                try:
                    default_index = options.index(default_choice) if default_choice in options else 0
                except Exception:
                    default_index = 0

                # create selectbox and free-text input (values persist via keys)
                choice = st.selectbox(
                    f"{t[lang_key]['Occurrence_Why']} {idx+1}",
                    options,
                    index=default_index,
                    key=f"occ_sel_{idx}"
                )
                txt_val = st.text_input(f"Or enter your own Occurrence Why {idx+1}", value=st.session_state.get(f"occ_txt_{idx}", ""), key=f"occ_txt_{idx}")

                # precedence: free-text overrides dropdown if non-empty
                final_val = txt_val.strip() if txt_val.strip() else choice
                st.session_state["d5_occ_whys"][idx] = final_val

                if final_val:
                    selected_occ.append(final_val)

            # add another occurrence why
            if st.button("➕ Add another Occurrence Why", key="add_occ_why"):
                st.session_state["d5_occ_whys"].append("")
                # initialize new text/select keys
                new_idx = len(st.session_state["d5_occ_whys"]) - 1
                st.session_state.setdefault(f"occ_txt_{new_idx}", "")
                st.session_state.setdefault(f"occ_sel_{new_idx}", "")

            st.session_state["d5_occ_selected"] = selected_occ

            # Detection Section
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
            for idx in range(len(st.session_state["d5_det_whys"])):
                remaining_options = []
                for cat, items in detection_categories.items():
                    for item in items:
                        full_item = f"{cat}: {item}"
                        if full_item not in selected_det and full_item not in st.session_state["d5_det_whys"]:
                            remaining_options.append(full_item)
                current_val_det = st.session_state["d5_det_whys"][idx]
                if current_val_det and current_val_det not in remaining_options:
                    remaining_options.append(current_val_det)

                options_det = [""] + sorted(remaining_options)
                default_choice_det = st.session_state.get(f"det_sel_{idx}", current_val_det if current_val_det in options_det else "")
                try:
                    default_index_det = options_det.index(default_choice_det) if default_choice_det in options_det else 0
                except Exception:
                    default_index_det = 0

                choice_det = st.selectbox(
                    f"{t[lang_key]['Detection_Why']} {idx+1}",
                    options_det,
                    index=default_index_det,
                    key=f"det_sel_{idx}"
                )
                txt_val_det = st.text_input(f"Or enter your own Detection Why {idx+1}", value=st.session_state.get(f"det_txt_{idx}", ""), key=f"det_txt_{idx}")

                final_det_val = txt_val_det.strip() if txt_val_det.strip() else choice_det
                st.session_state["d5_det_whys"][idx] = final_det_val

                if final_det_val:
                    selected_det.append(final_det_val)

            if st.button("➕ Add another Detection Why", key="add_det_why"):
                st.session_state["d5_det_whys"].append("")
                new_idx = len(st.session_state["d5_det_whys"]) - 1
                st.session_state.setdefault(f"det_txt_{new_idx}", "")
                st.session_state.setdefault(f"det_sel_{new_idx}", "")

            st.session_state["d5_det_selected"] = selected_det

            # ---------------------------
            # Suggested professional Root Causes based on mapped logic
            # ---------------------------
            suggested_occ = map_occurrence_root(selected_occ) if selected_occ else ""
            suggested_det = map_detection_root(selected_det) if selected_det else ""

            # show suggestions in editable text areas (user can tweak)
            occ_edit = st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=st.session_state.get("root_cause_occ_text", suggested_occ), key="root_cause_occ_text", height=120)
            det_edit = st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=st.session_state.get("root_cause_det_text", suggested_det), key="root_cause_det_text", height=120)

            # update session (persist edits)
            st.session_state["root_cause_occ_text"] = occ_edit
            st.session_state["root_cause_det_text"] = det_edit

            # Build final D5 Answer (so it appears under "Answer" in Excel)
            occ_section = "Occurrence Analysis:\n" + ("\n".join([f"- {s}" for s in selected_occ]) if selected_occ else "- (none)")
            det_section = "Detection Analysis:\n" + ("\n".join([f"- {s}" for s in selected_det]) if selected_det else "- (none)")
            rc_section = "\nSuggested Root Causes (editable):\n" + (occ_edit + ("\n" + det_edit if det_edit else "") if occ_edit or det_edit else "(none)")

            final_d5_answer = occ_section + "\n\n" + det_section + "\n\n" + rc_section

            st.session_state["D5"]["answer"] = final_d5_answer

        # D6-D8 rendering (unchanged behavior)
        elif step in ["D6", "D7", "D8"]:
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
