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
        background-color: #90a4ae !important;
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

# ---------------------------
# NPQP 8D steps
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
            "es":"Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de soporte."},
     {"en":"Customer reported static noise in amplifier during end-of-line test.",
      "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.",
            "es":"Verifique partes similares, modelos, partes genéricas, otros colores, mano opuesta, frente/trasero, etc."},
     {"en":"Similar model radio, front vs. rear speaker, amplifier with 8, 12, or 24 channels.",
      "es":"Radio de modelo similar, altavoz delantero vs trasero, amplificador con 8, 12 o 24 canales."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
            "es":"Realice una investigación inicial para identificar problemas evidentes, recopile datos y documente hallazgos iniciales."},
     {"en":"Visual inspection of solder joints, initial functional tests, checking connectors, etc.",
      "es":"Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores, etc."}),
    ("D4", {"en":"Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
            "es":"Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes."},
     {"en":"100% inspection of amplifiers before shipment; temporary shielding.",
      "es":"Inspección 100% de amplificadores antes del envío; blindaje temporal."}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause. The 5-Why analysis is a structured method to drill down into cause-and-effect by repeatedly asking 'why' until the true root cause is identified. Separate Occurrence and Detection.",
            "es":"Use el análisis de 5 Porqués para determinar la causa raíz. El análisis de 5 Porqués es un método estructurado para profundizar en la causa-efecto preguntando repetidamente 'por qué' hasta identificar la verdadera causa raíz. Separe Ocurrencia y Detección."},
     {"en":"","es":""}),
    # ---------------------------
# Tabs with ✅ / 🔴 status indicators (continued)
# ---------------------------

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        if step == "D5":
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
            <br>Use the 5-Why analysis to dig deeper into the root cause. Start from the observed issue, ask 'Why?' for each answer, and continue until you reach a true root cause.
            </div>
            """, unsafe_allow_html=True)

            # Occurrence Section
            st.markdown("#### Occurrence Analysis")
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
                    "Outdated or incomplete work instructions",
                    "Failure not identified in FMEA"
                ],
                "Environmental / External Factors": [
                    "Temperature, humidity, or other environmental conditions",
                    "Power fluctuations or outages",
                    "Contamination (dust, oil, chemicals)",
                    "Regulatory or compliance changes"
                ]
            }

            # Initialize dynamic number of why boxes
            if "d5_occ_count" not in st.session_state:
                st.session_state.d5_occ_count = len(st.session_state.d5_occ_whys)

            for idx in range(st.session_state.d5_occ_count):
                # Merge free-text and dropdown
                options = [""] + sorted(
                    [f"{cat}: {item}" for cat, items in occurrence_categories.items() for item in items
                     if f"{cat}: {item}" not in st.session_state.d5_occ_selected]
                )
                val = st.session_state.d5_occ_whys[idx] if idx < len(st.session_state.d5_occ_whys) else ""
                val_input = st.text_input(f"{t[lang_key]['Occurrence_Why']} {idx+1} (or choose from dropdown)",
                                          value=val, key=f"occ_input_{idx}")
                val_dropdown = st.selectbox(f"{t[lang_key]['Occurrence_Why']} {idx+1} dropdown", options, index=0, key=f"occ_dropdown_{idx}")
                final_val = val_input if val_input.strip() != "" else val_dropdown
                st.session_state.d5_occ_whys[idx] = final_val
                if final_val:
                    st.session_state.d5_occ_selected.append(final_val)

            if st.button("Add another Occurrence Why", key="add_occ"):
                st.session_state.d5_occ_whys.append("")
                st.session_state.d5_occ_count += 1

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

            if "d5_det_count" not in st.session_state:
                st.session_state.d5_det_count = len(st.session_state.d5_det_whys)

            for idx in range(st.session_state.d5_det_count):
                options_det = [""] + sorted(
                    [f"{cat}: {item}" for cat, items in detection_categories.items() for item in items
                     if f"{cat}: {item}" not in st.session_state.d5_det_selected]
                )
                val_det = st.session_state.d5_det_whys[idx] if idx < len(st.session_state.d5_det_whys) else ""
                val_det_input = st.text_input(f"{t[lang_key]['Detection_Why']} {idx+1} (or choose from dropdown)",
                                              value=val_det, key=f"det_input_{idx}")
                val_det_dropdown = st.selectbox(f"{t[lang_key]['Detection_Why']} {idx+1} dropdown", options_det, index=0, key=f"det_dropdown_{idx}")
                final_det_val = val_det_input if val_det_input.strip() != "" else val_det_dropdown
                st.session_state.d5_det_whys[idx] = final_det_val
                if final_det_val:
                    st.session_state.d5_det_selected.append(final_det_val)

            if st.button("Add another Detection Why", key="add_det"):
                st.session_state.d5_det_whys.append("")
                st.session_state.d5_det_count += 1

            # Combine answers into D5 answer field
            st.session_state.D5["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )

            # Root cause suggestions
            occurrence_root_suggestion = "The root cause that allowed this issue to occur: " + (
                st.session_state.d5_occ_whys[-1] if st.session_state.d5_occ_whys else ""
            )
            detection_root_suggestion = "The root cause that allowed this issue to escape detection: " + (
                st.session_state.d5_det_whys[-1] if st.session_state.d5_det_whys else ""
            )

            st.session_state.D5["extra"] = st.text_area(
                f"{t[lang_key]['Root_Cause']} - Occurrence", value=occurrence_root_suggestion, key="root_occ"
            )
            st.session_state.D5["extra_det"] = st.text_area(
                f"{t[lang_key]['Root_Cause']} - Detection", value=detection_root_suggestion, key="root_det"
            )

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"],
              st.session_state[step].get("extra", "") + "\n" + st.session_state[step].get("extra_det", "")) for step, _, _ in npqp_steps]

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
        st.session_state["D5"] = {"answer": "", "extra": "", "extra_det": ""}
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["d5_occ_count"] = 5
        st.session_state["d5_det_count"] = 5
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
            st.session_state.setdefault(step, {"answer":"", "extra":""})
        st.success("✅ All data has been reset!")
