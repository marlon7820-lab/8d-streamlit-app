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

# Hide Streamlit menu and footer
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# Translation dictionary
# ---------------------------
TEXTS = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "occurrence_analysis": "Occurrence Analysis",
        "detection_analysis": "Detection Analysis",
        "root_cause": "Root Cause (summary after 5-Whys)",
        "save_button": "💾 Save 8D Report",
        "success": "✅ NPQP 8D Report saved successfully."
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Elaborado Por",
        "occurrence_analysis": "Análisis de Ocurrencia",
        "detection_analysis": "Análisis de Detección",
        "root_cause": "Causa Raíz (resumen después de 5-Whys)",
        "save_button": "💾 Guardar Reporte 8D",
        "success": "✅ Reporte 8D NPQP guardado correctamente."
    }
}

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Language / Idioma", ["English", "Español"])
lang = "en" if lang == "English" else "es"

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = [
    ("D1: Concern Details",
     "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
     "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
    ("D2: Similar Part Considerations",
     "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
     "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
    ("D3: Initial Analysis",
     "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
     "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4: Implement Containment",
     "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
     "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
    ("D5: Final Analysis",
     "Use 5-Why analysis to determine the root cause. Occurrence & Detection separate.",
     ""),
    ("D6: Permanent Corrective Actions",
     "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
     "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
    ("D7: Countermeasure Confirmation",
     "Verify that corrective actions effectively resolve the issue long-term.",
     "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
    ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
     "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
     "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future.")
]

# ---------------------------
# Step colors for Excel
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
# Session state initialization
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)

# ---------------------------
# Offline 5-Why suggestions
# ---------------------------
OCCURRENCE_SUGGESTIONS = [
    "Cold solder joint on DSP chip",
    "Loose connector",
    "Incorrect assembly",
    "Component misalignment",
    "Design tolerance issue"
]
DETECTION_SUGGESTIONS = [
    "No inspection at this step",
    "Inspection procedure unclear",
    "Test not sensitive enough",
    "Human error during check",
    "Missing test step"
]

def suggest_next_why(previous_whys, mode="occurrence"):
    if mode=="occurrence":
        suggestions = [s for s in OCCURRENCE_SUGGESTIONS if s not in previous_whys]
    else:
        suggestions = [s for s in DETECTION_SUGGESTIONS if s not in previous_whys]
    return suggestions[:3]

# ---------------------------
# Report info
# ---------------------------
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input(TEXTS[lang]["report_date"], value=today_str)
st.session_state.prepared_by = st.text_input(TEXTS[lang]["prepared_by"], value=st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])

for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")
        if note:
            st.info(f"**Training Guidance:** {note}\n\n💡 **Example:** {example}")

        if step.startswith("D5"):
            st.markdown(f"#### {TEXTS[lang]['occurrence_analysis']}")
            for idx in range(5):
                suggestions = suggest_next_why(st.session_state.d5_occ_whys[:idx], mode="occurrence")
                options = [""] + suggestions
                st.session_state.d5_occ_whys[idx] = st.selectbox(
                    f"Occurrence Why {idx+1}",
                    options=options,
                    index=options.index(st.session_state.d5_occ_whys[idx]) if st.session_state.d5_occ_whys[idx] in options else 0,
                    key=f"d5_occ_{idx}"
                )
            
            st.markdown(f"#### {TEXTS[lang]['detection_analysis']}")
            for idx in range(5):
                suggestions = suggest_next_why(st.session_state.d5_det_whys[:idx], mode="detection")
                options = [""] + suggestions
                st.session_state.d5_det_whys[idx] = st.selectbox(
                    f"Detection Why {idx+1}",
                    options=options,
                    index=options.index(st.session_state.d5_det_whys[idx]) if st.session_state.d5_det_whys[idx] in options else 0,
                    key=f"d5_det_{idx}"
                )
            
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", value=st.session_state[step]["answer"])

        if step.startswith("D1"):
            st.session_state[step]["extra"] = st.text_area(TEXTS[lang]['root_cause'], value=st.session_state[step]["extra"])

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button
# ---------------------------
if st.button(TEXTS[lang]["save_button"]):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
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
        ws.row_dimensions[1].height = 25

        # Report info
        ws["A3"] = "Report Date"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By"
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Your Answer", "Extra"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        # Content
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
        st.success(TEXTS[lang]["success"])
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX", f, file_name=xlsx_file)
