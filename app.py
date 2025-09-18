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
# Version info
# ---------------------------
version_number = "v1.0.7"
last_updated = "September 14, 2025"

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
    ("D5", {"en":"Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection. Include FMEA failure occurrence if applicable.",
            "es":"Use el análisis de 5 Porqués para determinar la causa raíz. Separe Ocurrencia y Detección. Incluya la ocurrencia de falla FMEA si aplica."},
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
st.session_state.setdefault("d5_occ_selected", [])
st.session_state.setdefault("d5_det_selected", [])
# ---------------------------
# Helper functions
# ---------------------------
def add_d5_input_fields():
    """Generate 5-Why input fields for Occurrence and Detection."""
    st.markdown(f"### {t[lang_key]['Root_Cause_Occ']}")
    for i in range(5):
        st.session_state["d5_occ_whys"][i] = st.text_input(
            f"{t[lang_key]['Occurrence_Why']} {i+1}", 
            value=st.session_state["d5_occ_whys"][i],
            key=f"d5_occ_why_{i}"
        )

    st.markdown(f"### {t[lang_key]['Root_Cause_Det']}")
    for i in range(5):
        st.session_state["d5_det_whys"][i] = st.text_input(
            f"{t[lang_key]['Detection_Why']} {i+1}", 
            value=st.session_state["d5_det_whys"][i],
            key=f"d5_det_why_{i}"
        )

# ---------------------------
# UI Tabs
# ---------------------------
tabs = st.tabs([t[lang_key][step] for step, _, _ in npqp_steps])

for i, (step, instructions, examples) in enumerate(npqp_steps):
    with tabs[i]:
        st.info(instructions[lang_key])
        if examples[lang_key]:
            st.caption(f"💡 {t[lang_key]['Example']}: {examples[lang_key]}")

        if step == "D5":
            add_d5_input_fields()
            st.text_area(
                t[lang_key]['FMEA_Failure'],
                key=f"{step}_extra",
                value=st.session_state[step]["extra"],
                height=100,
                on_change=lambda s=step: st.session_state.update({s: {"extra": st.session_state[f"{s}_extra"], "answer": st.session_state[s]["answer"]}})
            )
        else:
            st.text_area(
                "Enter details here:",
                key=f"{step}_answer",
                value=st.session_state[step]["answer"],
                height=150,
                on_change=lambda s=step: st.session_state.update({s: {"answer": st.session_state[f"{s}_answer"], "extra": st.session_state[s]["extra"]}})
            )

# ---------------------------
# D1 Meta Info Section
# ---------------------------
with st.sidebar:
    st.markdown("### Report Info")
    st.session_state["report_date"] = st.date_input(
        t[lang_key]["Report_Date"],
        value=datetime.datetime.today()
    ).strftime("%B %d, %Y")

    st.session_state["prepared_by"] = st.text_input(
        t[lang_key]["Prepared_By"],
        value=st.session_state["prepared_by"]
    )

# ---------------------------
# Save and Download Buttons
# ---------------------------
col1, col2 = st.columns([1, 1])

with col1:
    save_btn = st.button(t[lang_key]["Save"])

with col2:
    download_btn = st.button(t[lang_key]["Download"])
    # ---------------------------
# Save Data to Excel
# ---------------------------
def save_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    ws["A1"] = "8D Report"
    ws["A2"] = f"Date: {st.session_state['report_date']}"
    ws["A3"] = f"Prepared By: {st.session_state['prepared_by']}"

    row = 5
    for step, _, _ in npqp_steps:
        ws[f"A{row}"] = step
        ws[f"B{row}"] = st.session_state[step]["answer"]
        row += 1

    for i, why in enumerate(st.session_state["d5_occ_whys"], start=1):
        ws[f"A{row}"] = f"D5 - Occurrence Why {i}"
        ws[f"B{row}"] = why
        row += 1

    for i, why in enumerate(st.session_state["d5_det_whys"], start=1):
        ws[f"A{row}"] = f"D5 - Detection Why {i}"
        ws[f"B{row}"] = why
        row += 1

    ws[f"A{row}"] = "D5 - Extra"
    ws[f"B{row}"] = st.session_state["D5"]["extra"]

    # Save Excel file in memory
    file_buffer = io.BytesIO()
    wb.save(file_buffer)
    file_buffer.seek(0)
    return file_buffer

# ---------------------------
# Handle Save & Download
# ---------------------------
if save_btn:
    st.session_state["saved_file"] = save_to_excel()
    st.success(t[lang_key]["Saved_Message"])

if download_btn:
    file_buffer = save_to_excel()
    st.download_button(
        label=t[lang_key]["Download_Label"],
        data=file_buffer,
        file_name=f"8D_Report_{datetime.datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    # ---------------------------
# Translations & Session Initialization
# ---------------------------
t = {
    "en": {
        "Saved_Message": "Report saved successfully!",
        "Download_Label": "Download 8D Report"
    },
    "es": {
        "Saved_Message": "¡Informe guardado con éxito!",
        "Download_Label": "Descargar Informe 8D"
    }
}

# ---------------------------
# Initialize Session State
# ---------------------------
if "report_date" not in st.session_state:
    st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
if "prepared_by" not in st.session_state:
    st.session_state["prepared_by"] = ""
if "saved_file" not in st.session_state:
    st.session_state["saved_file"] = None

for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": ""}

if "d5_occ_whys" not in st.session_state:
    st.session_state["d5_occ_whys"] = [""] * 5
if "d5_det_whys" not in st.session_state:
    st.session_state["d5_det_whys"] = [""] * 5
if "D5" not in st.session_state:
    st.session_state["D5"] = {"extra": ""}

# ---------------------------
# Build App UI
# ---------------------------
st.title("8D Report Assistant")
tabs = st.tabs([step for step, _, _ in npqp_steps])

for i, (step, desc, color) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"<h3 style='color:{color}'>{step}</h3>", unsafe_allow_html=True)
        st.text_area(
            f"{step} - Your Answer",
            key=f"{step}_answer",
            value=st.session_state[step]["answer"],
            on_change=lambda s=step: st.session_state[step].update({"answer": st.session_state[f"{s}_answer"]})
        )

# ---------------------------
# Save / Download Buttons
# ---------------------------
save_btn = st.button("💾 Save Report")
download_btn = st.button("⬇️ Download Excel Report")
