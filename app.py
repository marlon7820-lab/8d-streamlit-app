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
      "es":"Inspección 100% de amplificadores antes del envío; blindaje temporal."})
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    st.session_state.setdefault(step, {"answer": "", "extra": ""})

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
# --------------------------- Part 2 ---------------------------
# D5 Tab: Final Analysis
st.markdown(f"## {t[lang_key]['D5']}")

# Buttons for clearing D5
col1, col2 = st.columns([1,4])
with col1:
    if st.button("🧹 Clear D5 Section"):
        for i in range(5):
            st.session_state[f"occurrence_{i}"] = ""
            st.session_state[f"detection_{i}"] = ""
            st.session_state[f"systemic_{i}"] = ""

# Occurrence section
st.markdown("### Occurrence Root Cause")
for idx in range(5):
    key_occ = f"occurrence_{idx}"
    st.session_state.setdefault(key_occ, "")
    st.session_state[key_occ] = st.text_area(
        f"{t[lang_key]['Occurrence_Why']} {idx+1}",
        value=st.session_state[key_occ],
        key=f"{key_occ}_text"
    )

# Detection section
st.markdown("### Detection Root Cause")
for idx in range(5):
    key_det = f"detection_{idx}"
    st.session_state.setdefault(key_det, "")
    st.session_state[key_det] = st.text_area(
        f"{t[lang_key]['Detection_Why']} {idx+1}",
        value=st.session_state[key_det],
        key=f"{key_det}_text"
    )

# Systemic section
st.markdown("### Systemic Root Cause")
for idx in range(5):
    key_sys = f"systemic_{idx}"
    st.session_state.setdefault(key_sys, "")
    st.session_state[key_sys] = st.text_area(
        f"{t[lang_key]['Systemic_Why']} {idx+1}",
        value=st.session_state[key_sys],
        key=f"{key_sys}_text"
    )

# Optional guidance/note
st.markdown(f"**{t[lang_key]['Training_Guidance']}:** Use this section to guide the team through systemic improvements and lessons learned.")
# --------------------------- Part 3 ---------------------------
# D6–D8 Tabs
for step in ["D6", "D7", "D8"]:
    st.markdown(f"## {t[lang_key][step]}")
    st.session_state.setdefault(step, {"answer": "", "extra": ""})
    st.session_state[step]["answer"] = st.text_area(
        f"{t[lang_key][step]} - Your Answer",
        value=st.session_state[step]["answer"],
        key=f"{step}_answer"
    )
    st.session_state[step]["extra"] = st.text_area(
        "Extra / Notes",
        value=st.session_state[step]["extra"],
        key=f"{step}_extra"
    )

# ---------------------------
# Collect all answers for Excel
# ---------------------------
data_rows = []

# D1–D8 answers
for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
    if step == "D5":
        occ_answers = [st.session_state.get(f"occurrence_{i}", "") for i in range(5)]
        det_answers = [st.session_state.get(f"detection_{i}", "") for i in range(5)]
        sys_answers = [st.session_state.get(f"systemic_{i}", "") for i in range(5)]
        answer_text = f"Occurrence:\n" + "\n".join(occ_answers) + \
                      f"\n\nDetection:\n" + "\n".join(det_answers) + \
                      f"\n\nSystemic:\n" + "\n".join(sys_answers)
        extra_text = ""
        data_rows.append((t[lang_key][step], answer_text, extra_text))
    else:
        answer_text = st.session_state[step]["answer"]
        extra_text = st.session_state[step]["extra"]
        data_rows.append((t[lang_key][step], answer_text, extra_text))

# ---------------------------
# Generate Excel
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # Header row
    headers = ["Step", "Answer", "Extra / Notes"]
    ws.append(headers)

    for step, answer, extra in data_rows:
        ws.append([step, answer, extra])

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 50

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{datetime.datetime.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ---------------------------
# Sidebar: JSON Backup / Restore
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore")
    def generate_json():
        return json.dumps({k:v for k,v in st.session_state.items() if not k.startswith("_")}, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Backup_{datetime.datetime.today().strftime('%Y%m%d')}.json",
        mime="application/json"
    )

    uploaded_file = st.file_uploader("Upload JSON to Restore", type="json")
    if uploaded_file:
        try:
            restore_data = json.load(uploaded_file)
            for k, v in restore_data.items():
                st.session_state[k] = v
            st.success("✅ Session restored from JSON!")
        except Exception as e:
            st.error(f"Error restoring JSON: {e}")

    if st.button("🗑️ Clear All Data"):
        for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
            st.session_state[step] = {"answer": "", "extra": ""}
        for i in range(5):
            st.session_state[f"occurrence_{i}"] = ""
            st.session_state[f"detection_{i}"] = ""
            st.session_state[f"systemic_{i}"] = ""
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.success("✅ All data cleared!")
