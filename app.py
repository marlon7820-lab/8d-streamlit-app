import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import os

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

# App styles - updated for desktop selectbox outline
st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}

/* Outline all Streamlit widget containers (works on desktop) */
div.stSelectbox, div.stTextInput, div.stTextArea {
    border: 2px solid #1E90FF !important;
    border-radius: 5px !important;
    padding: 5px !important;
    background-color: #ffffff !important;
    transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}

/* Hover effect */
div.stSelectbox:hover, div.stTextInput:hover, div.stTextArea:hover {
    border: 2px solid #104E8B !important; /* slightly darker blue */
    box-shadow: 0 0 5px #1E90FF;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Reset Session check (safe, no KeyError)
# ---------------------------
if st.session_state.get("_reset_8d_session", False):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    if "_reset_8d_session" in st.session_state:
        st.session_state["_reset_8d_session"] = False
    st.rerun()

# ---------------------------
# Main title
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Version info
# ---------------------------
version_number = "v1.1.0"
last_updated = "October 17, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar: Language selection & reset
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

st.sidebar.markdown("---")
st.sidebar.header("⚙️ App Controls")
if st.sidebar.button("🔄 Reset 8D Session"):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys:
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.session_state["_reset_8d_session"] = True
    st.stop()

if st.session_state.get("_reset_8d_session", False):
    st.session_state["_reset_8d_session"] = False
    st.experimental_rerun()

# ---------------------------
# Language dictionary
# ---------------------------
t = {
    "en": {
        "D1": "D1: Concern Details",
        "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis",
        "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis",
        "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation",
        "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date",
        "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)",
        "Root_Cause_Det": "Root Cause (Detection)",
        "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Systemic_Why": "Systemic Why",
        "Save": "💾 Save 8D Report",
        "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance",
        "Example": "Example",
        "FMEA_Failure": "FMEA Failure Occurrence",
        "Location": "Material Location",
        "Status": "Activity Status",
        "Containment_Actions": "Containment Actions"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación",
        "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial",
        "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final",
        "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas",
        "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe",
        "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)",
        "Root_Cause_Det": "Causa raíz (Detección)",
        "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia",
        "Detection_Why": "Por qué Detección",
        "Systemic_Why": "Por qué Sistémico",
        "Save": "💾 Guardar Informe 8D",
        "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento",
        "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA",
        "Location": "Ubicación del material",
        "Status": "Estado de la actividad",
        "Containment_Actions": "Acciones de contención"
    }
}

# ---------------------------
# NPQP 8D steps with examples
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."}, {"en":"Customer reported static noise in amplifier during end-of-line test.", "es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.", "es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."}, {"en":"Similar model radio, Front vs. rear speaker.", "es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.", "es":"Realice una investigación inicial para identificar problemas evidentes."}, {"en":"Visual inspection of solder joints, initial functional tests.", "es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions and material location.", "es":"Defina acciones de contención temporales y ubicación del material."}, {"en":"Post Quality Alert, Increase Inspection, Inventory Certification","es":"Implementar Ayuda Visual, Incrementar Inspeccion, Certificar Inventario"}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause.", "es":"Use el análisis de 5 Porqués para determinar la causa raíz."}, {"en":"","es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently.", "es":"Defina acciones correctivas que eliminen la causa raíz permanentemente."}, {"en":"Update soldering process, redesign fixture.", "es":"Actualizar proceso de soldadura, rediseñar herramienta."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue.", "es":"Verifique que las acciones correctivas resuelvan efectivamente el problema."}, {"en":"Functional tests on corrected amplifiers.", "es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned, update standards, FMEAs.", "es":"Documente lecciones aprendidas, actualice estándares, FMEAs."}, {"en":"Update SOPs, PFMEA, work instructions.", "es":"Actualizar SOPs, PFMEA, instrucciones de trabajo."})
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

# Ensure D6/D7 subfields exist
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("d5_sys_whys", [""]*5)
st.session_state.setdefault("d4_location", "")
st.session_state.setdefault("d4_status", "")
st.session_state.setdefault("d4_containment", "")

# D6 fields
for sub in ["occ_answer", "det_answer", "sys_answer"]:
    st.session_state.setdefault(("D6"), st.session_state.get("D6", {}))
    st.session_state["D6"].setdefault(sub, "")

# D7 fields
for sub in ["occ_answer", "det_answer", "sys_answer"]:
    st.session_state.setdefault(("D7"), st.session_state.get("D7", {}))
    st.session_state["D7"].setdefault(sub, "")

# ---------------------------
# D5 categories
# ---------------------------
occurrence_categories = {
    "en":["Process","Material","Design","Machine","Human Error"], 
    "es":["Proceso","Material","Diseño","Máquina","Error Humano"]
}
detection_categories = {
    "en":["Inspection","Testing","Customer"], 
    "es":["Inspección","Prueba","Cliente"]
}
systemic_categories = {
    "en":["Process System","Supplier System","Quality System"], 
    "es":["Sistema de Proceso","Sistema de Proveedor","Sistema de Calidad"]
}

# ---------------------------
# Function to render 5-Whys for each category
# ---------------------------
def render_whys_no_repeat(whys_list, categories, text_prefix="Why"):
    cols = st.columns(len(whys_list))
    for i in range(len(whys_list)):
        whys_list[i] = cols[i].text_area(f"{text_prefix} #{i+1}", whys_list[i])

# ---------------------------
# Layout: Tabs D1-D8
# ---------------------------
tabs_labels = [t[lang_key][0] for t in npqp_steps]
tabs = st.tabs(tabs_labels)

for idx, (step, guidance, example) in enumerate(npqp_steps):
    with tabs[idx]:
        st.markdown(f"### {t[lang_key][step]}")
        st.info(guidance[lang_key])
        st.text_area("Your Answer", key=f"{step}_answer", height=120)
        st.markdown(f"**Example:** {example[lang_key]}")
        if step == "D4":
            st.text_input("Material Location", key="d4_location")
            st.text_input("Activity Status", key="d4_status")
            st.text_area("Containment Actions", key="d4_containment")
        if step == "D5":
            st.subheader("Occurrence Why")
            render_whys_no_repeat(st.session_state.d5_occ_whys, occurrence_categories[lang_key])
            st.subheader("Detection Why")
            render_whys_no_repeat(st.session_state.d5_det_whys, detection_categories[lang_key])
            st.subheader("Systemic Why")
            render_whys_no_repeat(st.session_state.d5_sys_whys, systemic_categories[lang_key])
        if step == "D6":
            st.text_area("Permanent Corrective Action - Occurrence", key="D6_occ_answer")
            st.text_area("Permanent Corrective Action - Detection", key="D6_det_answer")
            st.text_area("Permanent Corrective Action - Systemic", key="D6_sys_answer")
        if step == "D7":
            st.text_area("Countermeasure Confirmation - Occurrence", key="D7_occ_answer")
            st.text_area("Countermeasure Confirmation - Detection", key="D7_det_answer")
            st.text_area("Countermeasure Confirmation - Systemic", key="D7_sys_answer")

# ---------------------------
# Excel Export Function
# ---------------------------
def export_to_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    row = 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    ws.cell(row=row, column=1, value="8D Report").font = Font(bold=True, size=14)
    row += 2

    # Header info
    ws.cell(row=row, column=1, value="Report Date")
    ws.cell(row=row, column=2, value=st.session_state.get("report_date", ""))
    row += 1
    ws.cell(row=row, column=1, value="Prepared By")
    ws.cell(row=row, column=2, value=st.session_state.get("prepared_by", ""))
    row += 2

    # Loop through D1-D8
    for step, guidance, example in npqp_steps:
        ws.cell(row=row, column=1, value=t[lang_key][step])
        ws.cell(row=row, column=2, value=st.session_state.get(f"{step}_answer", ""))
        row += 2
        if step == "D4":
            ws.cell(row=row, column=1, value="Material Location")
            ws.cell(row=row, column=2, value=st.session_state.get("d4_location",""))
            row += 1
            ws.cell(row=row, column=1, value="Activity Status")
            ws.cell(row=row, column=2, value=st.session_state.get("d4_status",""))
            row += 1
            ws.cell(row=row, column=1, value="Containment Actions")
            ws.cell(row=row, column=2, value=st.session_state.get("d4_containment",""))
            row += 2
        if step == "D5":
            for i, why in enumerate(st.session_state.d5_occ_whys):
                ws.cell(row=row, column=1, value=f"Occurrence Why #{i+1}")
                ws.cell(row=row, column=2, value=why)
                row += 1
            for i, why in enumerate(st.session_state.d5_det_whys):
                ws.cell(row=row, column=1, value=f"Detection Why #{i+1}")
                ws.cell(row=row, column=2, value=why)
                row += 1
            for i, why in enumerate(st.session_state.d5_sys_whys):
                ws.cell(row=row, column=1, value=f"Systemic Why #{i+1}")
                ws.cell(row=row, column=2, value=why)
                row += 1
        if step in ["D6","D7"]:
            for sub in ["occ_answer","det_answer","sys_answer"]:
                ws.cell(row=row, column=1, value=sub)
                ws.cell(row=row, column=2, value=st.session_state[step].get(sub,""))
                row += 1
        row += 2

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------
# Download Button
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("💾 Export 8D Report")
if st.sidebar.button("Download XLSX"):
    excel_data = export_to_excel()
    st.sidebar.download_button(
        label="📥 Download 8D Report",
        data=excel_data,
        file_name=f"8D_Report_{datetime.datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------------------------
# File/Photo Upload Capability
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("📎 Upload Files / Photos")
uploaded_files = st.sidebar.file_uploader(
    "Upload files or images from computer/phone",
    type=None,
    accept_multiple_files=True
)
if uploaded_files:
    st.sidebar.markdown("**Uploaded Files:**")
    for f in uploaded_files:
        st.sidebar.write(f.name)
        if f.type.startswith("image/"):
            st.image(f, width=200)
