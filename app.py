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
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)", "Root_Cause_Det": "Root Cause (Detection)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why",
        "Systematic_Why": "Systematic Why",
        "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example",
        "FMEA_Failure": "FMEA Failure Occurrence"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)", "Root_Cause_Det": "Causa raíz (Detección)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección",
        "Systematic_Why": "Por qué Sistémico",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA"
    }
}

# ---------------------------
# NPQP 8D steps with examples
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."},
           {"en":"Customer reported static noise.", "es":"El cliente reportó ruido estático."}),
    ("D2", {"en":"Check for similar parts or models.", "es":"Verifique partes similares o modelos."},
           {"en":"Similar model radio, Front vs. rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify issues.", "es":"Realice una investigación inicial para identificar problemas."},
           {"en":"Visual inspection of solder joints.", "es":"Inspección visual de soldaduras."}),
    ("D4", {"en":"Define temporary containment actions.", "es":"Defina acciones de contención temporales."},
           {"en":"100% inspection before shipment.", "es":"Inspección 100% antes del envío."}),
    ("D5", {"en":"Use 5-Why analysis for root cause, separate Occurrence, Detection, and Systematic.", 
            "es":"Use análisis de 5 Porqués para la causa raíz, separando Ocurrencia, Detección y Sistémico."},
           {"en":"", "es":""}),
    ("D6", {"en":"Define permanent corrective actions.", "es":"Defina acciones correctivas permanentes."},
           {"en":"Update soldering process.", "es":"Actualizar proceso de soldadura."}),
    ("D7", {"en":"Verify that corrective actions resolve the issue long-term.", "es":"Verifique que las acciones correctivas resuelvan el problema a largo plazo."},
           {"en":"Functional tests on corrected units.", "es":"Pruebas funcionales en unidades corregidas."}),
    ("D8", {"en":"Document lessons learned, update standards and training.", "es":"Documente lecciones aprendidas y actualice estándares y capacitación."},
           {"en":"Update SOPs, PFMEA.", "es":"Actualizar SOPs, PFMEA."})
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

# ---------------------------
# Sidebar controls
# ---------------------------
st.sidebar.header("Controls")
if st.sidebar.button("Clear D5 Section"):
    st.session_state.d5_occ_whys = [""] * 5
    st.session_state.d5_det_whys = [""] * 5
    st.session_state.d5_sys_whys = [""] * 5
    st.session_state.d5_occ_selected = []
    st.session_state.d5_det_selected = []
    st.session_state.d5_sys_selected = []

# ---------------------------
# Report Date and Prepared By
# ---------------------------
st.session_state.report_date = st.text_input(
    f"{t[lang_key]['Report_Date']}",
    value=st.session_state.report_date,
    key="report_date_input"
)
st.session_state.prepared_by = st.text_input(
    f"{t[lang_key]['Prepared_By']}",
    value=st.session_state.prepared_by,
    key="prepared_by_input"
)

# ---------------------------
# Render 8D Tabs
# ---------------------------
tabs = st.tabs([t[lang_key][step] for step, _, _ in npqp_steps])

for idx, (step, guidance, example) in enumerate(npqp_steps):
    with tabs[idx]:
        st.markdown(f"### {t[lang_key][step]}")
        st.info(guidance[lang_key])

        # Normal answer field
        st.session_state[step]["answer"] = st.text_area(
            f"{t[lang_key]['Example']}:", 
            value=st.session_state[step]["answer"], 
            height=120, key=f"{step}_answer"
        )

        # Extra field (if needed for internal notes)
        st.session_state[step]["extra"] = st.text_area(
            "Internal Notes / Notas Internas:", 
            value=st.session_state[step]["extra"], 
            height=80, key=f"{step}_extra"
        )

        # D5 specific: Occurrence, Detection, Systematic analysis
        if step == "D5":
            st.subheader("Occurrence Analysis")
            for i in range(5):
                st.session_state.d5_occ_whys[i] = st.text_input(
                    f"{t[lang_key]['Occurrence_Why']} {i+1}",
                    value=st.session_state.d5_occ_whys[i],
                    key=f"d5_occ_{i}"
                )

            st.subheader("Detection Analysis")
            for i in range(5):
                st.session_state.d5_det_whys[i] = st.text_input(
                    f"{t[lang_key]['Detection_Why']} {i+1}",
                    value=st.session_state.d5_det_whys[i],
                    key=f"d5_det_{i}"
                )

            st.subheader("Systematic Analysis")
            for i in range(5):
                st.session_state.d5_sys_whys[i] = st.text_input(
                    f"{t[lang_key]['Systematic_Why']} {i+1}",
                    value=st.session_state.d5_sys_whys[i],
                    key=f"d5_sys_{i}"
                )

            st.markdown("---")
            st.info(t[lang_key]["Training_Guidance"])
            # --------------------------- Part 3 ---------------------------

# ---------------------------
# D6–D8 Tabs
# ---------------------------
for step in ["D6", "D7", "D8"]:
    st.session_state[step]["answer"] = st.text_area(
        f"{t[lang_key][step]} - Your Answer",
        value=st.session_state[step]["answer"],
        height=120,
        key=f"{step}_answer"
    )
    st.session_state[step]["extra"] = st.text_area(
        "Internal Notes / Notas Internas:",
        value=st.session_state[step]["extra"],
        height=80,
        key=f"{step}_extra"
    )

# ---------------------------
# Prepare data for Excel
# ---------------------------
data_rows = []
for step, _, _ in npqp_steps:
    data_rows.append(
        [t[lang_key][step], st.session_state[step]["answer"], st.session_state[step]["extra"]]
    )

# ---------------------------
# Excel Export Function
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # Add logo if exists
    if os.path.exists("logo.png"):
        img = XLImage("logo.png")
        img.width = 140
        img.height = 40
        ws.add_image(img, "A1")

    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    # Headers
    headers = ["Step", "Answer", "Extra / Notes"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=ws.max_row + 1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")

    # Data
    for row in data_rows:
        ws.append(row)

    # Wrap text
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=3):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

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
# Sidebar: Backup / Restore / Reset
# ---------------------------
with st.sidebar:
    st.header("Backup / Restore / Reset")

    # JSON Backup
    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
        mime="application/json"
    )

    # Restore
    uploaded_file = st.file_uploader("Upload JSON to restore", type="json")
    if uploaded_file:
        restore_data = json.load(uploaded_file)
        for k, v in restore_data.items():
            st.session_state[k] = v
        st.success("✅ Session restored from JSON!")

    # Reset All
    if st.button("🗑️ Clear All Data"):
        for step, _, _ in npqp_steps:
            st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_sys_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["d5_sys_selected"] = []
        st.success("✅ All data has been reset!")
