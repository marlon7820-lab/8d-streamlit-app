import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from datetime import datetime

# ---------------------------
# Session state initialization
# ---------------------------
for d in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
    if d not in st.session_state:
        st.session_state[d] = {"answer":"", "extra":""}
if "d5_occ_whys" not in st.session_state:
    st.session_state.d5_occ_whys = [""]*5
if "d5_det_whys" not in st.session_state:
    st.session_state.d5_det_whys = [""]*5
if "prev_lang" not in st.session_state:
    st.session_state.prev_lang = "en"

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccione Idioma", ["English","Español"])
lang_key = lang[:2].lower()

# ---------------------------
# Translations dictionary
# ---------------------------
t = {
    "en": {
        "D1":"D1: Problem Description",
        "D2":"D2: Containment Actions",
        "D3":"D3: Interim Actions",
        "D4":"D4: Root Cause Analysis",
        "D5":"D5: 5-Why Analysis",
        "D6":"D6: Corrective Actions",
        "D7":"D7: Preventive Actions",
        "D8":"D8: Verification",
        "Training_Guidance":"Training Guidance",
        "Occurrence_Why":"Occurrence Why",
        "Detection_Why":"Detection Why",
        "Root_Cause":"Root Cause (summary after 5-Whys)",
        "Prepared_By":"Prepared By",
        "Report_Date":"Report Date",
        "Download_XLS":"Download XLS"
    },
    "es": {
        "D1":"D1: Descripción del Problema",
        "D2":"D2: Acciones de Contención",
        "D3":"D3: Acciones Interinas",
        "D4":"D4: Análisis de Causa Raíz",
        "D5":"D5: Análisis de 5 Porqués",
        "D6":"D6: Acciones Correctivas",
        "D7":"D7: Acciones Preventivas",
        "D8":"D8: Verificación",
        "Training_Guidance":"Guía de Entrenamiento",
        "Occurrence_Why":"Porqué de la Ocurrencia",
        "Detection_Why":"Porqué de la Detección",
        "Root_Cause":"Causa Raíz (resumen después de 5-Porqués)",
        "Prepared_By":"Preparado Por",
        "Report_Date":"Fecha del Reporte",
        "Download_XLS":"Descargar XLS"
    }
}

# ---------------------------
# Sample training guidance per D
# ---------------------------
npqp_steps = {
    0:["D1 sample guidance...","D1 sample text..."],
    1:["D2 guidance...","Sample containment actions..."],
    2:["D3 guidance...","Sample interim actions..."],
    3:["D4 guidance...","Sample root cause methods..."],
    4:["D5 guidance...","Use 5-Why to find root causes"],
    5:["D6 guidance...","Corrective actions examples"],
    6:["D7 guidance...","Preventive actions examples"],
    7:["D8 guidance...","Verification examples"]
}

# ---------------------------
# Tabs for D1-D8
# ---------------------------
tabs = st.tabs([t[lang_key][f"D{i}"] for i in range(1,9)])

# ---------------------------
# D1-D4, D6-D8: Free text + guidance
# ---------------------------
for idx, d in enumerate(["D1","D2","D3","D4","D6","D7","D8"]):
    with tabs[idx if idx<4 else idx+1]:
        st.markdown(f"### {t[lang_key][d]}")
        st.info(f"**{t[lang_key]['Training_Guidance']}:** {npqp_steps[idx][1]}")
        st.session_state[d]["answer"] = st.text_area("Answer / Respuesta", value=st.session_state[d]["answer"], height=150)
        if d=="D1":
            st.session_state[d]["extra"] = st.text_area("Extra notes", value=st.session_state[d]["extra"], height=100)

# ---------------------------
# D5: Interactive 5-Why
# ---------------------------
with tabs[4]:
    st.markdown(f"### {t[lang_key]['D5']}")
    st.info(f"**{t[lang_key]['Training_Guidance']}:** {npqp_steps[4][1]}")

    st.markdown("#### Occurrence Analysis")
    for idx in range(5):
        prompt = f"{t[lang_key]['Occurrence_Why']} {idx+1}"
        if idx==0:
            st.session_state.d5_occ_whys[idx] = st.text_input(prompt, value=st.session_state.d5_occ_whys[idx], key=f"occ_{idx}")
        else:
            suggestions = [w for w in st.session_state.d5_occ_whys[:idx] if w.strip()]
            context_hints = ["Equipment","Process","Materials","Methods","Environment"]
            options = suggestions + context_hints
            choice = st.selectbox(prompt, options=options, index=0, key=f"occ_{idx}_select")
            free_text = st.text_input(f"{prompt} (free text)", value="", key=f"occ_{idx}_free")
            st.session_state.d5_occ_whys[idx] = free_text if free_text.strip() else choice

    st.markdown("#### Detection Analysis")
    for idx in range(5):
        prompt = f"{t[lang_key]['Detection_Why']} {idx+1}"
        if idx==0:
            st.session_state.d5_det_whys[idx] = st.text_input(prompt, value=st.session_state.d5_det_whys[idx], key=f"det_{idx}")
        else:
            suggestions = [w for w in st.session_state.d5_det_whys[:idx] if w.strip()]
            context_hints = ["Inspection methods","Testing gaps","Detection process"]
            options = suggestions + context_hints
            choice = st.selectbox(prompt, options=options, index=0, key=f"det_{idx}_select")
            free_text = st.text_input(f"{prompt} (free text)", value="", key=f"det_{idx}_free")
            st.session_state.d5_det_whys[idx] = free_text if free_text.strip() else choice

    st.session_state.D5["answer"] = (
        "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
        "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
    )
    st.session_state.D5["extra"] = st.text_area(f"{t[lang_key]['Root_Cause']}", value=st.session_state.D5["extra"], key="root_cause")

# ---------------------------
# Report metadata
# ---------------------------
report_date = datetime.today().strftime("%Y-%m-%d")
prepared_by = st.text_input(t[lang_key]['Prepared_By'], value="Marlon Ordonez")
st.markdown(f"**{t[lang_key]['Report_Date']}:** {report_date}")

# ---------------------------
# Download XLS
# ---------------------------
wb = Workbook()
ws = wb.active
ws.title = "8D Report"
row = 1
for d in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
    ws.cell(row=row, column=1, value=d)
    ws.cell(row=row, column=2, value=st.session_state[d]["answer"])
    row+=1

from io import BytesIO
output = BytesIO()
wb.save(output)
st.download_button(t[lang_key]['Download_XLS'], data=output.getvalue(), file_name="8D_Report.xlsx")
