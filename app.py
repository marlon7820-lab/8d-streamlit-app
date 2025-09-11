import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
from googletrans import Translator

# ---------------------------
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide Streamlit default menu, header, and footer
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Translator
translator = Translator()

# -------------------------------------------------------------------
# Language selection
# -------------------------------------------------------------------
lang = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
prev_lang = st.session_state.get("prev_lang", "English")
st.session_state["prev_lang"] = lang

def translate_text(text, target_lang):
    if target_lang.startswith("Esp"):
        return translator.translate(text, src="en", dest="es").text
    else:
        return translator.translate(text, src="es", dest="en").text

# -------------------------------------------------------------------
# NPQP 8D Steps with training notes/examples
# -------------------------------------------------------------------
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
     "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected).",
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

# -------------------------------------------------------------------
# Initialize session state
# -------------------------------------------------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("interactive_whys", [""] * 5)

# Color dictionary for Excel
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

# -------------------------------------------------------------------
# Report info with formatted date
# -------------------------------------------------------------------
st.subheader("Report Information / Información del Informe")
today_str = datetime.datetime.today().strftime("%B %d, %Y")
if not st.session_state.report_date:
    st.session_state.report_date = today_str
st.session_state.report_date = st.text_input(
    translate_text("📅 Report Date", lang) if lang=="Español" else "📅 Report Date",
    value=st.session_state.report_date
)
st.session_state.prepared_by = st.text_input(
    translate_text("✍️ Prepared By", lang) if lang=="Español" else "✍️ Prepared By",
    value=st.session_state.prepared_by
)

# -------------------------------------------------------------------
# 5-Why interactive suggestion logic
# -------------------------------------------------------------------
def suggest_occurrence(previous_whys):
    suggestions = []
    for why in previous_whys:
        w = why.lower()
        if "solder" in w:
            suggestions.append("Check soldering process and training.")
        elif "connection" in w:
            suggestions.append("Verify connector type and assembly procedure.")
        elif "material" in w:
            suggestions.append("Review material specs and batch.")
        else:
            suggestions.append("Review previous failures or standard processes.")
    return suggestions

def suggest_detection(previous_whys):
    suggestions = []
    for why in previous_whys:
        w = why.lower()
        if "inspection" in w:
            suggestions.append("Check QA inspection checklist.")
        elif "test" in w:
            suggestions.append("Review automated testing procedure.")
        else:
            suggestions.append("Verify process controls and monitoring.")
    return suggestions

# -------------------------------------------------------------------
# Tabs for each step
# -------------------------------------------------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {translate_text(step, lang) if lang=='Español' else step}")

        if step.startswith("D5"):
            st.markdown("#### Occurrence Analysis / Análisis de Ocurrencia")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(
                    f"Occurrence Why {idx+1}" if lang=="English" else f"Por qué de Ocurrencia {idx+1}",
                    value=val,
                    key=f"occ_{idx}"
                )
                if st.session_state.d5_occ_whys[idx].strip():
                    sugg = suggest_occurrence(st.session_state.d5_occ_whys[:idx+1])
                    st.text(f"💡 Suggestions: {', '.join(sugg[-1:])}" if lang=="English" else f"💡 Sugerencias: {', '.join(sugg[-1:])}")

            st.markdown("#### Detection Analysis / Análisis de Detección")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(
                    f"Detection Why {idx+1}" if lang=="English" else f"Por qué de Detección {idx+1}",
                    value=val,
                    key=f"det_{idx}"
                )
                if st.session_state.d5_det_whys[idx].strip():
                    sugg = suggest_detection(st.session_state.d5_det_whys[:idx+1])
                    st.text(f"💡 Suggestions: {', '.join(sugg[-1:])}" if lang=="English" else f"💡 Sugerencias: {', '.join(sugg[-1:])}")

            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )

        else:
            st.session_state[step]["answer"] = st.text_area(
                translate_text(f"Your Answer for {step}", lang) if lang=="Español" else f"Your Answer for {step}",
                value=st.session_state[step]["answer"],
                key=f"ans_{step}"
            )

# -------------------------------------------------------------------
# Collect answers
# -------------------------------------------------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step].get("extra","")) for step, _, _ in npqp_steps]

# -------------------------------------------------------------------
# Save button with styled Excel
# -------------------------------------------------------------------
if st.button(translate_text("💾 Save 8D Report", lang) if lang=="Español" else "💾 Save 8D Report"):
    if not any(ans for _, ans, _ in data_rows):
        st.error(translate_text("⚠️ No answers filled in yet. Please complete some fields before saving.", lang) if lang=="Español" else "⚠️ No answers filled in yet. Please complete some fields before saving.")
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
        for step, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=step)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)
            fill_color = step_colors.get(step, "FFFFFF")
            for col in range(1, 4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            row += 1

        # Adjust column widths
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)
        st.success(translate_text("✅ NPQP 8D Report saved successfully.", lang) if lang=="Español" else "✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button(translate_text("📥 Download XLSX", lang) if lang=="Español" else "📥 Download XLSX", f, file_name=xlsx_file)
