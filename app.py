import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import os
from PIL import Image as PILImage
from io import BytesIO

st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}
div.stSelectbox, div.stTextInput, div.stTextArea {
    border: 2px solid #1E90FF !important;
    border-radius: 5px !important;
    padding: 5px !important;
    background-color: #ffffff !important;
    transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}
div.stSelectbox:hover, div.stTextInput:hover, div.stTextArea:hover {
    border: 2px solid #104E8B !important;
    box-shadow: 0 0 5px #1E90FF;
}
.image-thumbnail {width: 120px; height: 80px; object-fit: cover; margin:5px; border:1px solid #1E90FF; border-radius:4px;}
</style>
""", unsafe_allow_html=True)

if st.session_state.get("_reset_8d_session", False):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.session_state["_reset_8d_session"] = False
    st.rerun()

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

version_number = "v1.2.0"
last_updated = "October 18, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

dark_mode = st.sidebar.checkbox("🌙 Dark Mode")
if dark_mode:
    st.markdown("""
    <style>
    .stApp {background: linear-gradient(to right, #1e1e1e, #2c2c2c); color: #f5f5f5 !important;}
    .stTabs [data-baseweb="tab"] {font-weight: bold; color: #f5f5f5 !important;}
    .stTabs [data-baseweb="tab"]:hover {color: #87AFC7 !important;}
    div.stTextInput, div.stTextArea, div.stSelectbox {
        border: 2px solid #87AFC7 !important;
        border-radius: 5px !important;
        background-color: #2c2c2c !important;
        color: #f5f5f5 !important;
        padding: 5px !important;
        transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
    }
    div.stTextInput:hover, div.stTextArea:hover, div.stSelectbox:hover {
        border: 2px solid #1E90FF !important;
        box-shadow: 0 0 5px #1E90FF;
    }
    .stInfo {background-color: #3a3a3a !important; border-left: 5px solid #87AFC7 !important; color: #f5f5f5 !important;}
    .css-1d391kg {color: #87AFC7 !important; font-weight: bold !important;}
    .stSidebar {background-color: #1e1e1e !important; color: #f5f5f5 !important;}
    .stSidebar button[kind="primary"] {background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold;}
    .stSidebar button {background-color: #5a5a5a !important; color: #f5f5f5 !important;}
    .stSidebar .stDownloadButton button {background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold;}
    </style>
    """, unsafe_allow_html=True)

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
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import os
from PIL import Image as PILImage
from io import BytesIO

st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}
div.stSelectbox, div.stTextInput, div.stTextArea {
    border: 2px solid #1E90FF !important;
    border-radius: 5px !important;
    padding: 5px !important;
    background-color: #ffffff !important;
    transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}
div.stSelectbox:hover, div.stTextInput:hover, div.stTextArea:hover {
    border: 2px solid #104E8B !important;
    box-shadow: 0 0 5px #1E90FF;
}
.image-thumbnail {width: 120px; height: 80px; object-fit: cover; margin:5px; border:1px solid #1E90FF; border-radius:4px;}
</style>
""", unsafe_allow_html=True)

if st.session_state.get("_reset_8d_session", False):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.session_state["_reset_8d_session"] = False
    st.rerun()

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

version_number = "v1.2.0"
last_updated = "October 18, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

dark_mode = st.sidebar.checkbox("🌙 Dark Mode")
if dark_mode:
    st.markdown("""
    <style>
    .stApp {background: linear-gradient(to right, #1e1e1e, #2c2c2c); color: #f5f5f5 !important;}
    .stTabs [data-baseweb="tab"] {font-weight: bold; color: #f5f5f5 !important;}
    .stTabs [data-baseweb="tab"]:hover {color: #87AFC7 !important;}
    div.stTextInput, div.stTextArea, div.stSelectbox {
        border: 2px solid #87AFC7 !important;
        border-radius: 5px !important;
        background-color: #2c2c2c !important;
        color: #f5f5f5 !important;
        padding: 5px !important;
        transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
    }
    div.stTextInput:hover, div.stTextArea:hover, div.stSelectbox:hover {
        border: 2px solid #1E90FF !important;
        box-shadow: 0 0 5px #1E90FF;
    }
    .stInfo {background-color: #3a3a3a !important; border-left: 5px solid #87AFC7 !important; color: #f5f5f5 !important;}
    .css-1d391kg {color: #87AFC7 !important; font-weight: bold !important;}
    .stSidebar {background-color: #1e1e1e !important; color: #f5f5f5 !important;}
    .stSidebar button[kind="primary"] {background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold;}
    .stSidebar button {background-color: #5a5a5a !important; color: #f5f5f5 !important;}
    .stSidebar .stDownloadButton button {background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold;}
    </style>
    """, unsafe_allow_html=True)

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
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import os
from PIL import Image as PILImage
from io import BytesIO

st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
.css-1d391kg {color: #1E90FF !important; font-weight: bold !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}
div.stSelectbox, div.stTextInput, div.stTextArea {
    border: 2px solid #1E90FF !important;
    border-radius: 5px !important;
    padding: 5px !important;
    background-color: #ffffff !important;
    transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
}
div.stSelectbox:hover, div.stTextInput:hover, div.stTextArea:hover {
    border: 2px solid #104E8B !important;
    box-shadow: 0 0 5px #1E90FF;
}
.image-thumbnail {width: 120px; height: 80px; object-fit: cover; margin:5px; border:1px solid #1E90FF; border-radius:4px;}
</style>
""", unsafe_allow_html=True)

if st.session_state.get("_reset_8d_session", False):
    preserve_keys = ["lang", "lang_key", "current_tab"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and key != "_reset_8d_session":
            del st.session_state[key]
    for k, v in preserved.items():
        st.session_state[k] = v
    st.session_state["_reset_8d_session"] = False
    st.rerun()

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

version_number = "v1.2.0"
last_updated = "October 18, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'>
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

dark_mode = st.sidebar.checkbox("🌙 Dark Mode")
if dark_mode:
    st.markdown("""
    <style>
    .stApp {background: linear-gradient(to right, #1e1e1e, #2c2c2c); color: #f5f5f5 !important;}
    .stTabs [data-baseweb="tab"] {font-weight: bold; color: #f5f5f5 !important;}
    .stTabs [data-baseweb="tab"]:hover {color: #87AFC7 !important;}
    div.stTextInput, div.stTextArea, div.stSelectbox {
        border: 2px solid #87AFC7 !important;
        border-radius: 5px !important;
        background-color: #2c2c2c !important;
        color: #f5f5f5 !important;
        padding: 5px !important;
        transition: border 0.2s ease-in-out, box-shadow 0.2s ease-in-out;
    }
    div.stTextInput:hover, div.stTextArea:hover, div.stSelectbox:hover {
        border: 2px solid #1E90FF !important;
        box-shadow: 0 0 5px #1E90FF;
    }
    .stInfo {background-color: #3a3a3a !important; border-left: 5px solid #87AFC7 !important; color: #f5f5f5 !important;}
    .css-1d391kg {color: #87AFC7 !important; font-weight: bold !important;}
    .stSidebar {background-color: #1e1e1e !important; color: #f5f5f5 !important;}
    .stSidebar button[kind="primary"] {background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold;}
    .stSidebar button {background-color: #5a5a5a !important; color: #f5f5f5 !important;}
    .stSidebar .stDownloadButton button {background-color: #87AFC7 !important; color: #000000 !important; font-weight: bold;}
    </style>
    """, unsafe_allow_html=True)

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
# D1-D8 Tabs
# ---------------------------
tabs = st.tabs([t[lang_key]["D1"], t[lang_key]["D2"], t[lang_key]["D3"],
                t[lang_key]["D4"], t[lang_key]["D5"], t[lang_key]["D6"],
                t[lang_key]["D7"], t[lang_key]["D8"]])

# Initialize session state for dynamic fields if not exists
if "d5_occ_whys" not in st.session_state: st.session_state.d5_occ_whys = [""]
if "d5_det_whys" not in st.session_state: st.session_state.d5_det_whys = [""]
if "d5_sys_whys" not in st.session_state: st.session_state.d5_sys_whys = [""]

# ---------------------------
# D1: Concern Details
# ---------------------------
with tabs[0]:
    st.header(t[lang_key]["D1"])
    st.session_state.d1_report_date = st.date_input(t[lang_key]["Report_Date"], 
                                                    st.session_state.get("d1_report_date", datetime.date.today()))
    st.session_state.d1_prepared_by = st.text_input(t[lang_key]["Prepared_By"], 
                                                    st.session_state.get("d1_prepared_by", ""))

# ---------------------------
# D2: Similar Part Considerations
# ---------------------------
with tabs[1]:
    st.header(t[lang_key]["D2"])
    st.session_state.d2_similar_parts = st.text_area("Similar Part Considerations", 
                                                     st.session_state.get("d2_similar_parts", ""))

# ---------------------------
# D3: Initial Analysis
# ---------------------------
with tabs[2]:
    st.header(t[lang_key]["D3"])
    st.session_state.d3_initial_analysis = st.text_area("Initial Analysis", 
                                                       st.session_state.get("d3_initial_analysis", ""))

# ---------------------------
# D4: Implement Containment
# ---------------------------
with tabs[3]:
    st.header(t[lang_key]["D4"])
    st.session_state.d4_containment_actions = st.text_area(t[lang_key]["Containment_Actions"], 
                                                           st.session_state.get("d4_containment_actions", ""))
    st.session_state.d4_material_location = st.text_input(t[lang_key]["Location"], 
                                                          st.session_state.get("d4_material_location", ""))
    st.session_state.d4_status = st.selectbox(t[lang_key]["Status"], 
                                              ["Open", "In Progress", "Closed"], 
                                              index=0 if "d4_status" not in st.session_state else ["Open","In Progress","Closed"].index(st.session_state.d4_status))
    st.session_state.d4_fmea_failure = st.text_input(t[lang_key]["FMEA_Failure"], 
                                                     st.session_state.get("d4_fmea_failure", ""))

# ---------------------------
# D5: Final Analysis with dynamic Why boxes
# ---------------------------
with tabs[4]:
    st.header(t[lang_key]["D5"])
    
    st.subheader(t[lang_key]["Root_Cause_Occ"])
    for i, val in enumerate(st.session_state.d5_occ_whys):
        st.session_state.d5_occ_whys[i] = st.text_area(f"{t[lang_key]['Occurrence_Why']} {i+1}", val)
    if st.button("➕ Add Occurrence Why"):
        st.session_state.d5_occ_whys.append("")
        st.experimental_rerun()

    st.subheader(t[lang_key]["Root_Cause_Det"])
    for i, val in enumerate(st.session_state.d5_det_whys):
        st.session_state.d5_det_whys[i] = st.text_area(f"{t[lang_key]['Detection_Why']} {i+1}", val)
    if st.button("➕ Add Detection Why"):
        st.session_state.d5_det_whys.append("")
        st.experimental_rerun()

    st.subheader(t[lang_key]["Root_Cause_Sys"])
    for i, val in enumerate(st.session_state.d5_sys_whys):
        st.session_state.d5_sys_whys[i] = st.text_area(f"{t[lang_key]['Systemic_Why']} {i+1}", val)
    if st.button("➕ Add Systemic Why"):
        st.session_state.d5_sys_whys.append("")
        st.experimental_rerun()

# ---------------------------
# D6: Permanent Corrective Actions
# ---------------------------
with tabs[5]:
    st.header(t[lang_key]["D6"])
    st.session_state.d6_action1 = st.text_area("Action 1", st.session_state.get("d6_action1", ""))
    st.session_state.d6_action2 = st.text_area("Action 2", st.session_state.get("d6_action2", ""))
    st.session_state.d6_action3 = st.text_area("Action 3", st.session_state.get("d6_action3", ""))

# ---------------------------
# D7: Countermeasure Confirmation
# ---------------------------
with tabs[6]:
    st.header(t[lang_key]["D7"])
    st.session_state.d7_action1 = st.text_area("Action 1", st.session_state.get("d7_action1", ""))
    st.session_state.d7_action2 = st.text_area("Action 2", st.session_state.get("d7_action2", ""))
    st.session_state.d7_action3 = st.text_area("Action 3", st.session_state.get("d7_action3", ""))

# ---------------------------
# D8: Follow-up Activities
# ---------------------------
with tabs[7]:
    st.header(t[lang_key]["D8"])
    st.session_state.d8_lessons_learned = st.text_area("Lessons Learned", st.session_state.get("d8_lessons_learned", ""))
    st.session_state.d8_recurrence_prevention = st.text_area("Recurrence Prevention", st.session_state.get("d8_recurrence_prevention", ""))
# ---------------------------
# Excel Export
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    # ---------------------------
    # Styling helpers
    # ---------------------------
    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill_gray = PatternFill("solid", fgColor="DDDDDD")
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    row = 1

    # ---------------------------
    # D1-D8 Sections
    # ---------------------------
    def write_section(title, content_dict):
        nonlocal row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.cell(row=row, column=1, value=title).font = bold_font
        ws.cell(row=row, column=1).alignment = center_alignment
        ws.cell(row=row, column=1).fill = fill_gray
        row += 1
        for key, value in content_dict.items():
            ws.cell(row=row, column=1, value=key).font = bold_font
            ws.cell(row=row, column=2, value=value)
            row += 1
        row += 1

    # ---------------------------
    # D1: Concern Details
    # ---------------------------
    write_section("D1: Concern Details", {
        "Report Date": str(st.session_state.d1_report_date),
        "Prepared By": st.session_state.d1_prepared_by
    })

    # D2: Similar Parts
    write_section("D2: Similar Part Considerations", {
        "Similar Parts": st.session_state.d2_similar_parts
    })

    # D3: Initial Analysis
    write_section("D3: Initial Analysis", {
        "Analysis": st.session_state.d3_initial_analysis
    })

    # D4: Containment
    write_section("D4: Containment", {
        "Containment Actions": st.session_state.d4_containment_actions,
        "Material Location": st.session_state.d4_material_location,
        "Status": st.session_state.d4_status,
        "FMEA Failure": st.session_state.d4_fmea_failure
    })

    # D5: Final Analysis with dynamic Why boxes
    def whys_to_text(whys_list):
        return "\n".join([f"{i+1}. {w}" for i, w in enumerate(whys_list) if w.strip() != ""])

    write_section("D5: Root Cause Analysis - Occurrence", {
        "Whys": whys_to_text(st.session_state.d5_occ_whys)
    })
    write_section("D5: Root Cause Analysis - Detection", {
        "Whys": whys_to_text(st.session_state.d5_det_whys)
    })
    write_section("D5: Root Cause Analysis - Systemic", {
        "Whys": whys_to_text(st.session_state.d5_sys_whys)
    })

    # D6: Permanent Corrective Actions
    write_section("D6: Permanent Corrective Actions", {
        "Action 1": st.session_state.d6_action1,
        "Action 2": st.session_state.d6_action2,
        "Action 3": st.session_state.d6_action3
    })

    # D7: Countermeasure Confirmation
    write_section("D7: Countermeasure Confirmation", {
        "Action 1": st.session_state.d7_action1,
        "Action 2": st.session_state.d7_action2,
        "Action 3": st.session_state.d7_action3
    })

    # D8: Follow-up Activities
    write_section("D8: Follow-up Activities", {
        "Lessons Learned": st.session_state.d8_lessons_learned,
        "Recurrence Prevention": st.session_state.d8_recurrence_prevention
    })

    # ---------------------------
    # Column width adjustment
    # ---------------------------
    for col in range(1, 3):
        ws.column_dimensions[get_column_letter(col)].width = 40

    # ---------------------------
    # Return Excel Bytes
    # ---------------------------
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# ---------------------------
# Download Button
# ---------------------------
st.download_button(
    label="📥 Download 8D Excel",
    data=generate_excel(),
    file_name=f"8D_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
