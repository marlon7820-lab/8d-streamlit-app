import streamlit as st
from datetime import datetime

# Initialize session state
if "lang" not in st.session_state:
    st.session_state.lang = "en"

languages = {"English": "en", "Español": "es"}
lang_key = st.selectbox("Select language / Seleccione idioma", options=list(languages.keys()))
lang = languages[lang_key]

# Define translations
t = {
    "en": {
        "D1": "D1 – Problem Description",
        "D2": "D2 – Containment Action",
        "D3": "D3 – Interim Action",
        "D4": "D4 – Root Cause Analysis",
        "D5": "D5 – 5-Why Analysis",
        "D6": "D6 – Permanent Corrective Action",
        "D7": "D7 – Preventive Measures",
        "D8": "D8 – Closure & Verification",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Date": "Date",
        "Prepared_By": "Prepared by",
        "Download": "📥 Download XLS"
    },
    "es": {
        "D1": "D1 – Descripción del Problema",
        "D2": "D2 – Acción de Contención",
        "D3": "D3 – Acción Interina",
        "D4": "D4 – Análisis de Causa Raíz",
        "D5": "D5 – Análisis de 5 Porqués",
        "D6": "D6 – Acción Correctiva Permanente",
        "D7": "D7 – Medidas Preventivas",
        "D8": "D8 – Cierre y Verificación",
        "Occurrence_Why": "Porqué de la ocurrencia",
        "Detection_Why": "Porqué de la detección",
        "Date": "Fecha",
        "Prepared_By": "Elaborado por",
        "Download": "📥 Descargar XLS"
    }
}

# Initialize answers
for d in ["D1","D2","D3","D4","D6","D7","D8"]:
    if d not in st.session_state:
        st.session_state[d] = ""

# Initialize D5
if "d5_occ_whys" not in st.session_state:
    st.session_state.d5_occ_whys = [""]*5
if "d5_det_whys" not in st.session_state:
    st.session_state.d5_det_whys = [""]*5

# Sidebar metadata
st.sidebar.write(f"{t[lang]['Date']}: {datetime.now().strftime('%Y-%m-%d')}")
if "prepared_by" not in st.session_state:
    st.session_state.prepared_by = ""
st.sidebar.text_input(f"{t[lang]['Prepared_By']}", key="prepared_by")

# Tabs
tabs = st.tabs([t[lang][f"D{i}"] for i in range(1,9)])

# Free-text entries for D1–D4, D6–D8
for idx, d in enumerate(["D1","D2","D3","D4","D6","D7","D8"]):
    with tabs[idx]:
        st.text_area(t[lang][d], value=st.session_state[d], key=f"{d}_text")

# Function to generate suggestions based on prior answers
def generate_suggestions(prev_answers, domain_hints):
    suggestions = []
    for hint in domain_hints:
        if all(hint.lower() not in a.lower() for a in prev_answers):
            suggestions.append(hint)
    # Include previous answers for continuity
    suggestions += [a for a in prev_answers if a.strip()]
    return suggestions

# D5 – interactive 5-Why
with tabs[4]:
    st.write("### Occurrence Analysis")
    occurrence_hints = [
        "Missing protection", "Cold solder joint", "Inspection gap",
        "Testing method inadequate", "Process step skipped"
    ]
    
    for idx in range(5):
        prompt = f"{t[lang]['Occurrence_Why']} {idx+1}"
        if idx == 0:
            st.session_state.d5_occ_whys[idx] = st.text_input(prompt, value=st.session_state.d5_occ_whys[idx], key=f"d5_occ_text_{idx}")
        else:
            options = generate_suggestions(st.session_state.d5_occ_whys[:idx], occurrence_hints)
            choice = st.selectbox(prompt, options=options, key=f"d5_occ_select_{idx}")
            free_text = st.text_input(f"{prompt} (free text)", value="", key=f"d5_occ_free_{idx}")
            st.session_state.d5_occ_whys[idx] = free_text if free_text.strip() else choice

    st.write("### Detection Analysis")
    detection_hints = [
        "Detection process missing", "Insufficient test coverage", "Customer complaint overlooked",
        "Inspection not performed", "Measurement method inadequate"
    ]
    
    for idx in range(5):
        prompt = f"{t[lang]['Detection_Why']} {idx+1}"
        if idx == 0:
            st.session_state.d5_det_whys[idx] = st.text_input(prompt, value=st.session_state.d5_det_whys[idx], key=f"d5_det_text_{idx}")
        else:
            options = generate_suggestions(st.session_state.d5_det_whys[:idx], detection_hints)
            choice = st.selectbox(prompt, options=options, key=f"d5_det_select_{idx}")
            free_text = st.text_input(f"{prompt} (free text)", value="", key=f"d5_det_free_{idx}")
            st.session_state.d5_det_whys[idx] = free_text if free_text.strip() else choice

# Download stub
st.download_button(t[lang]["Download"], data="Excel export placeholder", file_name="8D_report.xlsx")
