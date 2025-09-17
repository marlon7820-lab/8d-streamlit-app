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
.stApp { background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important; }
.stTabs [data-baseweb="tab"] { font-weight: bold; color: #000000 !important; }
textarea { background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important; }
.stInfo { background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important; }
.css-1d391kg { color: #1E90FF !important; font-weight: bold !important; }
button[kind="primary"] { background-color: #87AFC7 !important; color: white !important; font-weight: bold; }
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
last_updated = "September 17, 2025"
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
        "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why", "Systemic_Why": "Systemic Why",
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
        "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección", "Systemic_Why": "Por qué Sistémica",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA"
    }
}

# ---------------------------
# NPQP 8D steps with examples
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
    ("D5", {"en":"Use 5-Why analysis to determine the root cause. Separate Occurrence, Detection, and Systemic. Include FMEA failure occurrence if applicable.",
            "es":"Use el análisis de 5 Porqués para determinar la causa raíz. Separe Ocurrencia, Detección y Sistémica. Incluya la ocurrencia de falla FMEA si aplica."},
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

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("d5_sys_whys", [""] * 5)
st.session_state.setdefault("d5_occ_selected", [])
st.session_state.setdefault("d5_det_selected", [])
st.session_state.setdefault("d5_sys_selected", [])
# --------------------------- Part 2 ---------------------------
# Main Tabs for 8D
tab_names = [t[lang_key][step] for step, _, _ in npqp_steps]
tabs = st.tabs(tab_names)

for idx, (step, guidance_dict, example_dict) in enumerate(npqp_steps):
    with tabs[idx]:
        st.markdown(f"### {t[lang_key][step]}")
        st.markdown(f"**{t[lang_key]['Training_Guidance']}**: {guidance_dict[lang_key]}")
        st.info(f"{t[lang_key]['Example']}: {example_dict[lang_key]}")

        # Text input / area for answers
        st.session_state[step]["answer"] = st.text_area(
            f"{t[lang_key][step]} - Your Answer",
            value=st.session_state[step]["answer"],
            key=f"{step}_answer",
            height=120
        )

        # ---------------------------
        # D5 Special Section: 5-Why Analysis
        # ---------------------------
        if step == "D5":
            st.markdown("---")
            st.markdown("#### Root Cause Analysis (5-Why)")

            # Occurrence
            st.markdown("**Occurrence Root Cause**")
            for i in range(5):
                st.session_state["d5_occ_whys"][i] = st.text_input(
                    f"{t[lang_key]['Occurrence_Why']} {i+1}",
                    value=st.session_state["d5_occ_whys"][i],
                    key=f"d5_occ_{i}"
                )

            # Detection
            st.markdown("**Detection Root Cause**")
            for i in range(5):
                st.session_state["d5_det_whys"][i] = st.text_input(
                    f"{t[lang_key]['Detection_Why']} {i+1}",
                    value=st.session_state["d5_det_whys"][i],
                    key=f"d5_det_{i}"
                )

            # Systemic
            st.markdown("**Systemic Root Cause**")
            for i in range(5):
                st.session_state["d5_sys_whys"][i] = st.text_input(
                    f"{t[lang_key]['Systemic_Why']} {i+1}",
                    value=st.session_state["d5_sys_whys"][i],
                    key=f"d5_sys_{i}"
                )

            st.markdown("---")
            # --------------------------- Part 3 ---------------------------
# ---------------------------
# Sidebar: Backup / Restore / Clear
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore / Reset")

    # Generate JSON backup
    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
        mime="application/json"
    )

    # Restore JSON
    uploaded_file = st.file_uploader("Upload JSON to restore", type="json")
    if uploaded_file:
        try:
            restore_data = json.load(uploaded_file)
            for k, v in restore_data.items():
                st.session_state[k] = v
            st.success("✅ Session restored from JSON!")
        except Exception as e:
            st.error(f"Error restoring JSON: {e}")

    # Clear all
    if st.button("🗑️ Clear All"):
        for step, _, _ in npqp_steps:
            st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_sys_whys"] = [""] * 5
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.success("✅ All fields cleared!")

# ---------------------------
# Generate Excel file
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    # Add logo
    if os.path.exists("logo.png"):
        try:
            img = XLImage("logo.png")
            img.width = 140
            img.height = 40
            ws.add_image(img, "A1")
        except:
            pass

    # Report info
    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    # Headers
    headers = ["Step", "Answer", "Extra / Notes"]
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c_idx, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Step data
    row = 5
    for step, _, _ in npqp_steps:
        answer = st.session_state[step]["answer"]
        extra = st.session_state[step].get("extra", "")
        ws.append([t[lang_key][step], answer, extra])
        row += 1

    # D5 Whys
    ws.append([])
    ws.append(["D5 Occurrence Whys"] + st.session_state["d5_occ_whys"])
    ws.append(["D5 Detection Whys"] + st.session_state["d5_det_whys"])
    ws.append(["D5 Systemic Whys"] + st.session_state["d5_sys_whys"])

    # Adjust column width
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    # Save workbook to BytesIO
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------
# Download Excel
# ---------------------------
st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
