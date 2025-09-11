import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime

# ---------------------------
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide Streamlit default menu, header, and footer
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])

# ---------------------------
# Text dictionary
# ---------------------------
texts = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "Report Date",
        "prepared_by": "Prepared By",
        "save_btn": "💾 Save 8D Report",
        "download": "📥 Download XLSX",
        "occurrence": "Occurrence Analysis",
        "detection": "Detection Analysis",
        "root_cause": "Root Cause (summary after 5-Whys)",
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "Fecha del Reporte",
        "prepared_by": "Preparado Por",
        "save_btn": "💾 Guardar Reporte 8D",
        "download": "📥 Descargar XLSX",
        "occurrence": "Análisis de Ocurrencia",
        "detection": "Análisis de Detección",
        "root_cause": "Causa Raíz (resumen después del 5-Whys)",
    }
}
t = texts["en"] if lang == "English" else texts["es"]

# ---------------------------
# NPQP 8D Steps and training notes
# ---------------------------
npqp_steps = [
    ("D1: Concern Details", "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.", "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
    ("D2: Similar Part Considerations", "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.", "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
    ("D3: Initial Analysis", "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.", "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4: Implement Containment", "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.", "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
    ("D5: Final Analysis", "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected).", ""),
    ("D6: Permanent Corrective Actions", "Define corrective actions that eliminate the root cause permanently and prevent recurrence.", "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
    ("D7: Countermeasure Confirmation", "Verify that corrective actions effectively resolve the issue long-term.", "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
    ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)", "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.", "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future.")
]

# ---------------------------
# Session state initialization
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Excel color mapping
# ---------------------------
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.markdown(f"<h1 style='text-align:center;color:#1E90FF;'>{t['app_title']}</h1>", unsafe_allow_html=True)
st.session_state.report_date = st.text_input(t["report_date"], st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# 5-Why suggestion engine
# ---------------------------
context_suggestions_en = {
    "solder": ["Cold solder joint", "Solder bridge", "Insufficient heating"],
    "component": ["Wrong component", "Defective batch", "Misplaced polarity"],
    "process": ["Incorrect assembly step", "Step skipped", "Operator not trained"],
    "inspection": ["Visual inspection missed", "Test step skipped", "Checklists incomplete"]
}
context_suggestions_es = {
    "soldadura": ["Unión en frío", "Puente de soldadura", "Calentamiento insuficiente"],
    "componente": ["Componente incorrecto", "Lote defectuoso", "Polaridad mal colocada"],
    "proceso": ["Paso de ensamblaje incorrecto", "Paso omitido", "Operador no capacitado"],
    "inspeccion": ["Inspección visual omitida", "Prueba omitida", "Lista de verificación incompleta"]
}

def get_suggestions(prev_whys, lang):
    text = " ".join(prev_whys).lower()
    suggestions = []
    keywords = context_suggestions_en.keys() if lang=="English" else context_suggestions_es.keys()
    context_map = context_suggestions_en if lang=="English" else context_suggestions_es
    for k in keywords:
        if k in text:
            suggestions.extend(context_map[k])
    if not suggestions:
        # fallback to all general suggestions
        suggestions = sum(context_map.values(), [])
    # remove duplicates and already used
    suggestions = [s for s in suggestions if s not in prev_whys]
    return suggestions[:5]

# ---------------------------
# Tabs
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")

        # Show training note
        st.info(f"**Training Guidance:** {note}\n\n💡 **Example:** {example}" if example else f"**Training Guidance:** {note}")

        # Inputs
        if step.startswith("D5"):
            st.markdown(f"#### {t['occurrence']}")
            for idx in range(len(st.session_state.d5_occ_whys)):
                prev_whys = st.session_state.d5_occ_whys[:idx]
                suggestions = get_suggestions(prev_whys, lang)
                st.session_state.d5_occ_whys[idx] = st.selectbox(f"Occurrence Why {idx+1}", options=[""] + suggestions + [st.session_state.d5_occ_whys[idx]], index=0 if not st.session_state.d5_occ_whys[idx] else 1)
            st.markdown(f"#### {t['detection']}")
            for idx in range(len(st.session_state.d5_det_whys)):
                prev_whys = st.session_state.d5_det_whys[:idx]
                suggestions = get_suggestions(prev_whys, lang)
                st.session_state.d5_det_whys[idx] = st.selectbox(f"Detection Why {idx+1}", options=[""] + suggestions + [st.session_state.d5_det_whys[idx]], index=0 if not st.session_state.d5_det_whys[idx] else 1)
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area(t["root_cause"], value=st.session_state[step]["extra"])
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", value=st.session_state[step]["answer"])

# ---------------------------
# Save button
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]
if st.button(t["save_btn"]):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet.")
    else:
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"

        ws.merge_cells("A1:C1")
        ws["A1"] = "Nissan NPQP 8D Report"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        ws["A3"] = "Report Date"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By"
        ws["B4"] = st.session_state.prepared_by

        headers = ["Step", "Your Answer", "Root Cause"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        row = 7
        for step, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=step)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)
            fill_color = step_colors.get(step, "FFFFFF")
            for col in range(1, 4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            row += 1

        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)
        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button(t["download"], f, file_name=xlsx_file)
