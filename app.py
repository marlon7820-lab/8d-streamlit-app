import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
import datetime

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📑 8D Training App</h1>", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
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
        "Root_Cause": "Root Cause (summary after 5-Whys)",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Save": "💾 Save 8D Report",
        "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance",
        "Example": "Example"
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
        "Root_Cause": "Causa raíz (resumen después de los 5 Porqués)",
        "Occurrence_Why": "Por qué Ocurrencia",
        "Detection_Why": "Por qué Detección",
        "Save": "💾 Guardar Informe 8D",
        "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento",
        "Example": "Ejemplo"
    }
}
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# 8D Steps
# ---------------------------
npqp_steps = [
    ("D1", "Describe the customer concerns clearly.", "Example: Customer reported static noise."),
    ("D2", "Check for similar parts, models, generic parts, etc.", "Example: Same speaker type in other radio."),
    ("D3", "Perform an initial investigation.", "Example: Visual inspection of solder joints."),
    ("D4", "Define temporary containment actions.", "Example: 100% inspection before shipment."),
    ("D5", "Use 5-Why analysis. Separate Occurrence and Detection.", ""),
    ("D6", "Define corrective actions.", "Example: Update process and retrain operators."),
    ("D7", "Verify corrective actions effectiveness.", "Example: Functional tests, accelerated life testing."),
    ("D8", "Document lessons learned.", "Example: Update SOPs, PFMEA, training.")
]

# ---------------------------
# Session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Report info
# ---------------------------
st.subheader(t[lang_key]["Report_Date"])
st.session_state.report_date = st.text_input(t[lang_key]["Report_Date"], value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t[lang_key]["Prepared_By"], value=st.session_state.prepared_by)

# ---------------------------
# Tabs for steps
# ---------------------------
tabs = st.tabs([t[lang_key][step] for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        if step != "D5":
            st.session_state[step]["answer"] = st.text_area(f"Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.markdown("#### Occurrence Analysis")
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"{t[lang_key]['Occurrence_Why']} {idx+1}", value=st.session_state.d5_occ_whys[idx], key=f"occ_{idx}")
                else:
                    suggestions = ["Operator error", "Process not followed", "Equipment malfunction"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"{t[lang_key]['Occurrence_Why']} {idx+1}", [""] + suggestions, key=f"occ_{idx}")
            st.markdown("#### Detection Analysis")
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"{t[lang_key]['Detection_Why']} {idx+1}", value=st.session_state.d5_det_whys[idx], key=f"det_{idx}")
                else:
                    suggestions = ["QA checklist incomplete", "No automated test", "Missed inspection"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(f"{t[lang_key]['Detection_Why']} {idx+1}", [""] + suggestions, key=f"det_{idx}")
            st.session_state.D5["answer"] = "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) + "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            st.session_state.D5["extra"] = st.text_area(t[lang_key]["Root_Cause"], value=st.session_state.D5["extra"], key="root_cause")

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save Excel
# ---------------------------
if st.button(t[lang_key]["Save"]):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet.")
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        # Title
        ws.merge_cells("A1:C1")
        ws["A1"] = "NPQP 8D Report"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")

        # Report info
        ws.append([t[lang_key]["Report_Date"], st.session_state.report_date])
        ws.append([t[lang_key]["Prepared_By"], st.session_state.prepared_by])
        ws.append([])

        # Header
        ws.append(["Step", "Answer", "Extra Notes"])
        for col in range(1, 4):
            ws.cell(row=ws.max_row, column=col).font = Font(bold=True)
            ws.cell(row=ws.max_row, column=col).alignment = Alignment(horizontal="center")

        # Write data
        for step, answer, extra in data_rows:
            ws.append([step, answer, extra])

        # Adjust column widths
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws.column_dimensions[column].width = max(15, max_length + 2)

        # Save file
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb.save(xlsx_file)
        st.success("✅ 8D report saved!")
        st.download_button(t[lang_key]["Download"], xlsx_file, file_name=xlsx_file)
