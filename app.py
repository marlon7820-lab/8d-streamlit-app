import streamlit as st
from googletrans import Translator
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ----------------------------
# PAGE CONFIG
# ----------------------------
st.set_page_config(page_title="8D Report Tool", layout="wide")
st.markdown("<style>footer{visibility:hidden}</style>", unsafe_allow_html=True)

# ----------------------------
# TRANSLATION SETUP
# ----------------------------
translator = Translator()

def translate_text(text, direction="en_to_es"):
    if not text.strip():
        return text
    try:
        if direction == "en_to_es":
            return translator.translate(text, src="en", dest="es").text
        else:
            return translator.translate(text, src="es", dest="en").text
    except Exception:
        return text

# ----------------------------
# UI TEXTS
# ----------------------------
texts = {
    "en": {
        "title": "8D Corrective Action Report (NPQP Format)",
        "prepared_by": "Prepared By",
        "report_date": "Report Date",
        "save": "Download Excel Report",
        "d1": "D1: Team",
        "d2": "D2: Problem Description",
        "d3": "D3: Interim Containment",
        "d4": "D4: Root Cause Analysis",
        "d5": "D5: Permanent Corrective Action",
        "d6": "D6: Implement & Verify",
        "d7": "D7: Prevent Recurrence",
        "d8": "D8: Congratulate the Team",
        "why_occ": "5-Why for Occurrence",
        "why_det": "5-Why for Detection",
        "root_summary": "Root Cause Summary"
    },
    "es": {
        "title": "Reporte 8D de Acción Correctiva (Formato NPQP)",
        "prepared_by": "Preparado Por",
        "report_date": "Fecha de Reporte",
        "save": "Descargar Reporte en Excel",
        "d1": "D1: Equipo",
        "d2": "D2: Descripción del Problema",
        "d3": "D3: Contención Temporal",
        "d4": "D4: Análisis de Causa Raíz",
        "d5": "D5: Acción Correctiva Permanente",
        "d6": "D6: Implementar y Verificar",
        "d7": "D7: Prevenir Recurrencia",
        "d8": "D8: Felicitar al Equipo",
        "why_occ": "5-Why para Ocurrencia",
        "why_det": "5-Why para Detección",
        "root_summary": "Resumen de Causa Raíz"
    }
}

# ----------------------------
# SESSION STATE INITIALIZATION
# ----------------------------
if "lang" not in st.session_state:
    st.session_state.lang = "en"
if "answers" not in st.session_state:
    st.session_state.answers = {f"D{i}": "" for i in range(1, 9)}
if "d5_occ" not in st.session_state:
    st.session_state.d5_occ = ["" for _ in range(5)]
if "d5_det" not in st.session_state:
    st.session_state.d5_det = ["" for _ in range(5)]
if "root_summary" not in st.session_state:
    st.session_state.root_summary = ""

# ----------------------------
# LANGUAGE SWITCHER
# ----------------------------
col1, col2 = st.columns([3, 1])
with col1:
    st.title(texts[st.session_state.lang]["title"])
with col2:
    lang = st.selectbox("🌐", ["English", "Español"], index=0 if st.session_state.lang=="en" else 1)
    st.session_state.lang = "en" if lang == "English" else "es"

# ----------------------------
# REPORT INFO
# ----------------------------
col1, col2 = st.columns(2)
with col1:
    prepared_by = st.text_input(texts[st.session_state.lang]["prepared_by"])
with col2:
    report_date = st.date_input(texts[st.session_state.lang]["report_date"], datetime.today())

# ----------------------------
# 8D FORM TABS
# ----------------------------
tabs = st.tabs([texts[st.session_state.lang][f"d{i}"] for i in range(1, 9)])

# D1: TEAM
with tabs[0]:
    st.session_state.answers["D1"] = st.text_area(texts[st.session_state.lang]["d1"], st.session_state.answers["D1"], height=150)
    extra_comment = st.text_area("Additional Notes", "")

# D2 - D4, D6 - D8: Simple input
for i, t in enumerate([tabs[1], tabs[2], tabs[3], tabs[5], tabs[6], tabs[7]], start=2):
    with t:
        st.session_state.answers[f"D{i}"] = st.text_area(texts[st.session_state.lang][f"d{i}"], st.session_state.answers[f"D{i}"], height=150)

# D5: Interactive 5-Why
with tabs[4]:
    st.subheader(texts[st.session_state.lang]["why_occ"])
    for idx in range(5):
        st.session_state.d5_occ[idx] = st.text_input(f"Why {idx+1}", st.session_state.d5_occ[idx])

    st.subheader(texts[st.session_state.lang]["why_det"])
    for idx in range(5):
        st.session_state.d5_det[idx] = st.text_input(f"Why {idx+1}", st.session_state.d5_det[idx])

    st.session_state.root_summary = st.text_area(texts[st.session_state.lang]["root_summary"], st.session_state.root_summary, height=100)

# ----------------------------
# EXPORT TO EXCEL
# ----------------------------
if st.button(texts[st.session_state.lang]["save"]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "8D Report"

    ws.merge_cells("A1:B1")
    ws["A1"] = texts[st.session_state.lang]["title"]
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center")

    row = 3
    ws.append(["Prepared By", prepared_by])
    ws.append(["Report Date", report_date.strftime("%Y-%m-%d")])
    row += 2

    for i in range(1, 9):
        ws.append([texts[st.session_state.lang][f"d{i}"], st.session_state.answers[f"D{i}"]])
        row += 1

    ws.append(["5-Why Occurrence", "\n".join(st.session_state.d5_occ)])
    ws.append(["5-Why Detection", "\n".join(st.session_state.d5_det)])
    ws.append([texts[st.session_state.lang]["root_summary"], st.session_state.root_summary])

    filename = "8D_Report.xlsx"
    wb.save(filename)
    with open(filename, "rb") as f:
        st.download_button(texts[st.session_state.lang]["save"], f, file_name=filename)
