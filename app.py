import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide Streamlit default menu/footer/header
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
t = {
    "en": {
        "title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save": "💾 Save 8D Report",
        "download": "📥 Download XLSX",
        "d1": "D1: Concern Details",
        "d2": "D2: Similar Part Considerations",
        "d3": "D3: Initial Analysis",
        "d4": "D4: Implement Containment",
        "d5": "D5: Final Analysis",
        "d6": "D6: Permanent Corrective Actions",
        "d7": "D7: Countermeasure Confirmation",
        "d8": "D8: Follow-up Activities",
        "d5_occ": "Occurrence Analysis",
        "d5_det": "Detection Analysis",
        "root_cause": "Root Cause (summary after 5-Whys)",
        "training_note": "Training Guidance"
    },
    "es": {
        "title": "📑 App de Entrenamiento 8D",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado Por",
        "save": "💾 Guardar Reporte 8D",
        "download": "📥 Descargar XLSX",
        "d1": "D1: Detalles del Problema",
        "d2": "D2: Consideraciones de Partes Similares",
        "d3": "D3: Análisis Inicial",
        "d4": "D4: Implementar Contención",
        "d5": "D5: Análisis Final",
        "d6": "D6: Acciones Correctivas Permanentes",
        "d7": "D7: Confirmación de Contramedidas",
        "d8": "D8: Seguimiento / Lecciones Aprendidas",
        "d5_occ": "Análisis de Ocurrencia",
        "d5_det": "Análisis de Detección",
        "root_cause": "Causa Raíz (resumen después de 5-Whys)",
        "training_note": "Guía de Entrenamiento"
    }
}[lang[:2]]

# App header
st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['title']}</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP Steps + examples
# ---------------------------
npqp_steps = [
    {"id":"D1", "title": t["d1"], "note": "Describe the customer concerns clearly. Include what, where, when, and data.", "example":"Customer reported static noise in amplifier during end-of-line test."},
    {"id":"D2", "title": t["d2"], "note": "Check similar parts, models, etc.", "example":"Same speaker type used in another radio model; front vs rear."},
    {"id":"D3", "title": t["d3"], "note": "Initial investigation, collect data.", "example":"Visual inspection of solder joints; initial functional tests."},
    {"id":"D4", "title": t["d4"], "note": "Temporary containment actions.", "example":"100% inspection of amplifiers; quarantine affected batches."},
    {"id":"D5", "title": t["d5"], "note": "Use 5-Why analysis to determine root cause.", "example":""},  # interactive 5-Why
    {"id":"D6", "title": t["d6"], "note": "Define corrective actions to eliminate root cause.", "example":"Update soldering process, retrain operators."},
    {"id":"D7", "title": t["d7"], "note": "Verify corrective actions are effective.", "example":"Functional tests, accelerated life testing."},
    {"id":"D8", "title": t["d8"], "note": "Document lessons learned.", "example":"Update SOPs, FMEA, training."}
]

# ---------------------------
# Session state initialization
# ---------------------------
for step in npqp_steps:
    sid = step["id"]
    st.session_state.setdefault(sid, {"answer": "", "extra": ""})

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")

# D5 interactive 5-Why
st.session_state.setdefault("d5_occ_whys", [""])
st.session_state.setdefault("d5_det_whys", [""])

# ---------------------------
# Report info
# ---------------------------
st.subheader(t["report_date"])
st.session_state.report_date = st.text_input(t["report_date"], value=st.session_state.report_date)
st.subheader(t["prepared_by"])
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([step["title"] for step in npqp_steps])

for i, step in enumerate(npqp_steps):
    sid = step["id"]
    with tabs[i]:
        st.markdown(f"### {step['title']}")
        st.info(f"**{t['training_note']}:** {step['note']}\n\n💡 **Example:** {step['example']}")

        if sid == "D5":
            st.markdown(f"#### {t['d5_occ']}")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=val, key=f"{sid}_occ_{idx}")
            if st.button("➕ Add Occurrence Why", key="add_occ_d5"):
                st.session_state.d5_occ_whys.append("")

            st.markdown(f"#### {t['d5_det']}")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", value=val, key=f"{sid}_det_{idx}")
            if st.button("➕ Add Detection Why", key="add_det_d5"):
                st.session_state.d5_det_whys.append("")

            st.session_state[sid]["answer"] = "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) + \
                                              "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            st.session_state[sid]["extra"] = st.text_area(t["root_cause"], value=st.session_state[sid]["extra"])

        else:
            st.session_state[sid]["answer"] = st.text_area(f"Your Answer", value=st.session_state[sid]["answer"], key=f"ans_{sid}")

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step["title"], st.session_state[step["id"]]["answer"], st.session_state[step["id"]]["extra"]) for step in npqp_steps]

# ---------------------------
# Save button with Excel
# ---------------------------
if st.button(t["save"]):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet.")
    else:
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        # Title
        ws.merge_cells("A1:C1")
        ws["A1"] = "Nissan NPQP 8D Report"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        # Report info
        ws["A3"] = t["report_date"]
        ws["B3"] = st.session_state.report_date
        ws["A4"] = t["prepared_by"]
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Your Answer", "Root Cause"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        # Content
        row = 7
        for step_title, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=step_title)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)

            for col in range(1, 4):
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")

            row += 1

        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)
        st.success("✅ NPQP 8D Report saved successfully")

        with open(xlsx_file, "rb") as f:
            st.download_button(
                label=t["download"],
                data=f,
                file_name=xlsx_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
