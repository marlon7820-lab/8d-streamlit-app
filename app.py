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
    .stMarkdown, .stText, .stTextArea, .stTextInput, .stButton {
        color: #000000 !important;
    }
    textarea {
        background-color: #ffffff !important;
        border: 1px solid #1E90FF !important;
        border-radius: 5px;
        color: #000000 !important;
    }
    .css-1d391kg {
        color: #1E90FF !important;
        font-weight: bold !important;
    }
    button[kind="primary"] {
        background-color: #1E90FF !important;
        color: white !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

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
    ("D1", "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.", "Customer reported static noise in amplifier during end-of-line test."),
    ("D2", "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.", "Same speaker type used in another radio model; different amplifier colors."),
    ("D3", "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.", "Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4", "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.", "100% inspection of amplifiers before shipment; temporary shielding."),
    ("D5", "Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection.", ""),
    ("D6", "Define corrective actions that eliminate the root cause permanently and prevent recurrence.", "Update soldering process, retrain operators, update work instructions."),
    ("D7", "Verify that corrective actions effectively resolve the issue long-term.", "Functional tests on corrected amplifiers, accelerated life testing."),
    ("D8", "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.", "Update SOPs, PFMEA, work instructions, and employee training.")
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
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        if step != "D5":
            # MOBILE-SAFE guidance/example box
            st.markdown(f"""
            <div style="background-color:#b3e0ff; padding:15px; border-left:5px solid #1E90FF; color:black; border-radius:5px;">
            <b>{t[lang_key]['Training_Guidance']}:</b> {note}<br><br>
            💡 <b>{t[lang_key]['Example']}:</b> {example}
            </div>
            """, unsafe_allow_html=True)

            st.session_state[step]["answer"] = st.text_area(f"Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.markdown(f"""
            <div style="background-color:#b3e0ff; padding:15px; border-left:5px solid #1E90FF; color:black; border-radius:5px;">
            <b>{t[lang_key]['Training_Guidance']}:</b> {note}
            </div>
            """, unsafe_allow_html=True)

            st.markdown("#### Occurrence Analysis")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(
                        f"{t[lang_key]['Occurrence_Why']} {idx+1}", value=val, key=f"occ_{idx}")
                else:
                    suggestions = ["Operator error", "Process not followed", "Equipment malfunction"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(
                        f"{t[lang_key]['Occurrence_Why']} {idx+1}", [""] + suggestions + [st.session_state.d5_occ_whys[idx]], key=f"occ_{idx}"
                    )

            st.markdown("#### Detection Analysis")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(
                        f"{t[lang_key]['Detection_Why']} {idx+1}", value=val, key=f"det_{idx}")
                else:
                    suggestions = ["QA checklist incomplete", "No automated test", "Missed inspection"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(
                        f"{t[lang_key]['
