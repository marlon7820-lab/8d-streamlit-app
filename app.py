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
# Main title & version
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
# Session state initialization
# ---------------------------
if "initialized" not in st.session_state:
    for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
        st.session_state[step] = {"answer": ""}
    st.session_state["report_date"] = datetime.date.today()
    st.session_state["prepared_by"] = ""
    st.session_state["d5_occ_whys"] = [""] * 5
    st.session_state["d5_det_whys"] = [""] * 5
    st.session_state["d5_occ_selected"] = ""
    st.session_state["d5_det_selected"] = ""
    st.session_state["initialized"] = True

# ---------------------------
# Restore from URL param
# ---------------------------
if "saved_data" in st.query_params:
    try:
        saved_data = json.loads(st.query_params["saved_data"])
        for k, v in saved_data.items():
            if k in st.session_state:
                st.session_state[k] = v
    except Exception:
        pass

# ---------------------------
# Report Info
# ---------------------------
col1, col2 = st.columns([1, 1])
with col1:
    st.session_state["report_date"] = st.date_input(
        t[lang_key]["Report_Date"], value=st.session_state["report_date"]
    )
with col2:
    st.session_state["prepared_by"] = st.text_input(
        t[lang_key]["Prepared_By"], value=st.session_state["prepared_by"]
    )

# ---------------------------
# Tab labels with status
# ---------------------------
npqp_steps = [
    ("D1", "🔴", "🟢"), ("D2", "🔴", "🟢"), ("D3", "🔴", "🟢"),
    ("D4", "🔴", "🟢"), ("D5", "🔴", "🟢"), ("D6", "🔴", "🟢"),
    ("D7", "🔴", "🟢"), ("D8", "🔴", "🟢")
]

tab_labels = []
for step, red, green in npqp_steps:
    if st.session_state[step]["answer"]:
        tab_labels.append(f"{green} {t[lang_key][step]}")
    else:
        tab_labels.append(f"{red} {t[lang_key][step]}")

tabs = st.tabs(tab_labels)

# ---------------------------
# Render Tabs D1–D4
# ---------------------------
with tabs[0]:
    st.markdown(f"### {t[lang_key]['D1']}")
    st.info("Enter the concern description with as much detail as possible.")
    st.session_state["D1"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D1"]["answer"], key="ans_D1"
    )

with tabs[1]:
    st.markdown(f"### {t[lang_key]['D2']}")
    st.info("List other parts that could be impacted by the same issue.")
    st.session_state["D2"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D2"]["answer"], key="ans_D2"
    )

with tabs[2]:
    st.markdown(f"### {t[lang_key]['D3']}")
    st.info("Describe the initial analysis steps taken to verify the issue.")
    st.session_state["D3"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D3"]["answer"], key="ans_D3"
    )

with tabs[3]:
    st.markdown(f"### {t[lang_key]['D4']}")
    st.info("Describe containment actions to protect the customer.")
    st.session_state["D4"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D4"]["answer"], key="ans_D4"
    )
    # ---------------------------
# Render Tab D5 (Final Analysis)
# ---------------------------
with tabs[4]:
    st.markdown(f"### {t[lang_key]['D5']}")
    st.info("Perform root cause analysis using 5-Why for Occurrence and Detection.")

    # --- Root Cause (Occurrence) Free-text ---
    st.markdown(f"**{t[lang_key]['Root_Cause_Occ']}**")
    st.session_state["d5_occ_selected"] = st.text_input(
        "Occurrence Root Cause",
        value=st.session_state.get("d5_occ_selected", ""),
        key="occ_free_text"
    )

    # --- 5 Whys for Occurrence ---
    for i in range(5):
        st.session_state["d5_occ_whys"][i] = st.text_area(
            f"{t[lang_key]['Occurrence_Why']} {i+1}",
            value=st.session_state["d5_occ_whys"][i],
            key=f"occ_why_{i}"
        )

    # --- Root Cause (Detection) Free-text ---
    st.markdown(f"**{t[lang_key]['Root_Cause_Det']}**")
    st.session_state["d5_det_selected"] = st.text_input(
        "Detection Root Cause",
        value=st.session_state.get("d5_det_selected", ""),
        key="det_free_text"
    )

    # --- 5 Whys for Detection ---
    for i in range(5):
        st.session_state["d5_det_whys"][i] = st.text_area(
            f"{t[lang_key]['Detection_Why']} {i+1}",
            value=st.session_state["d5_det_whys"][i],
            key=f"det_why_{i}"
        )

    # --- Clear All Button (Fixed) ---
    if st.button("🧹 Clear All D5 Fields"):
        st.session_state["d5_occ_selected"] = ""
        st.session_state["d5_det_selected"] = ""
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["D5"]["answer"] = ""
        st.experimental_rerun()

    # --- D5 Notes / Final Analysis Text Area ---
    st.session_state["D5"]["answer"] = st.text_area(
        "Additional Notes (D5)",
        value=st.session_state["D5"]["answer"],
        key="ans_D5"
    )

# ---------------------------
# Render Tabs D6–D8
# ---------------------------
with tabs[5]:
    st.markdown(f"### {t[lang_key]['D6']}")
    st.info("Document the permanent corrective actions to eliminate root cause.")
    st.session_state["D6"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D6"]["answer"], key="ans_D6"
    )

with tabs[6]:
    st.markdown(f"### {t[lang_key]['D7']}")
    st.info("Provide verification that the corrective actions were effective.")
    st.session_state["D7"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D7"]["answer"], key="ans_D7"
    )

with tabs[7]:
    st.markdown(f"### {t[lang_key]['D8']}")
    st.info("Document lessons learned and actions to prevent recurrence.")
    st.session_state["D8"]["answer"] = st.text_area(
        "Your Answer", value=st.session_state["D8"]["answer"], key="ans_D8"
    )
    # ---------------------------
# Save Progress Function
# ---------------------------
def get_save_data():
    return {
        "D1": st.session_state["D1"],
        "D2": st.session_state["D2"],
        "D3": st.session_state["D3"],
        "D4": st.session_state["D4"],
        "D5": st.session_state["D5"],
        "D6": st.session_state["D6"],
        "D7": st.session_state["D7"],
        "D8": st.session_state["D8"],
        "report_date": str(st.session_state["report_date"]),
        "prepared_by": st.session_state["prepared_by"],
        "d5_occ_selected": st.session_state["d5_occ_selected"],
        "d5_det_selected": st.session_state["d5_det_selected"],
        "d5_occ_whys": st.session_state["d5_occ_whys"],
        "d5_det_whys": st.session_state["d5_det_whys"],
    }

def restore_from_json(uploaded_file):
    try:
        data = json.load(uploaded_file)
        for key, value in data.items():
            if key in st.session_state:
                st.session_state[key] = value
        st.experimental_rerun()
    except Exception as e:
        st.error(f"Error restoring file: {e}")

# ---------------------------
# Excel Generation Function
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # Title Row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    cell = ws.cell(row=1, column=1)
    cell.value = "8D Report"
    cell.font = Font(size=16, bold=True)
    cell.alignment = Alignment(horizontal="center")

    # Report Info
    ws["A3"] = "Report Date"
    ws["B3"] = st.session_state["report_date"].strftime("%Y-%m-%d")
    ws["A4"] = "Prepared By"
    ws["B4"] = st.session_state["prepared_by"]

    row = 6
    for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
        ws[f"A{row}"] = step
        ws[f"A{row}"].font = Font(bold=True)
        ws[f"B{row}"] = st.session_state[step]["answer"]
        row += 2

    # D5 Root Causes & Whys
    ws["A16"] = "Root Cause (Occurrence)"
    ws["B16"] = st.session_state["d5_occ_selected"]
    ws["A17"] = "Root Cause (Detection)"
    ws["B17"] = st.session_state["d5_det_selected"]

    for i, why in enumerate(st.session_state["d5_occ_whys"]):
        ws[f"A{19+i}"] = f"Occurrence Why {i+1}"
        ws[f"B{19+i}"] = why

    for i, why in enumerate(st.session_state["d5_det_whys"]):
        ws[f"A{25+i}"] = f"Detection Why {i+1}"
        ws[f"B{25+i}"] = why

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ---------------------------
# Sidebar: Save, Restore, Download
# ---------------------------
with st.sidebar:
    st.header("📂 Report Options")

    # Save Progress
    save_data = json.dumps(get_save_data(), indent=2)
    st.download_button(
        label="💾 Save Progress (JSON)",
        data=save_data,
        file_name="8D_report_progress.json",
        mime="application/json"
    )

    # Restore Progress
    uploaded_file = st.file_uploader("Restore Saved Progress", type=["json"])
    if uploaded_file:
        restore_from_json(uploaded_file)

    # Download Excel
    if st.button("📥 Generate Excel Report"):
        bio = generate_excel()
        st.download_button(
            label="Download Excel",
            data=bio,
            file_name="8D_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
