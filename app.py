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

# Hide default Streamlit UI
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# App title
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📑 8D Training App</h1>", unsafe_allow_html=True)

# ---------------------------
# Bilingual labels / guidance dictionary
# ---------------------------
texts = {
    "en": {
        "report_info": "Report Information",
        "report_date": "Report Date",
        "prepared_by": "Prepared By",
        "save_button": "💾 Save 8D Report",
        "download_button": "📥 Download XLSX",
        "d5_occ": "Occurrence Analysis (Interactive)",
        "d5_det": "Detection Analysis (Interactive)",
        "root_cause": "Root Cause (summary after 5-Whys)",
        "translate_en": "Translate to English",
        "translate_es": "Translate to Spanish",
        "add_occ": "➕ Add Occurrence Why",
        "add_det": "➕ Add Detection Why",
        "translate_section": "Translate Answers / Traducir Respuestas"
    },
    "es": {
        "report_info": "Información del Reporte",
        "report_date": "Fecha del Reporte",
        "prepared_by": "Preparado Por",
        "save_button": "💾 Guardar Reporte 8D",
        "download_button": "📥 Descargar XLSX",
        "d5_occ": "Análisis de Ocurrencia (Interactivo)",
        "d5_det": "Análisis de Detección (Interactivo)",
        "root_cause": "Causa Raíz (resumen después del 5-Whys)",
        "translate_en": "Traducir a Inglés",
        "translate_es": "Traducir a Español",
        "add_occ": "➕ Agregar Ocurrencia Why",
        "add_det": "➕ Agregar Detección Why",
        "translate_section": "Traducir Respuestas / Translate Answers"
    }
}

# ---------------------------
# NPQP 8D Steps with guidance/examples
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
     "Interactive 5-Why for root cause. Separate Occurrence and Detection.",
     ""),  # D5 guidance placeholder
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
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)   # start with 5 empty whys
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("prev_lang", "en")

# Excel color mapping
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
# Report info (uses dynamic labels)
# ---------------------------
st.subheader(texts["en"]["report_info"])  # header stays in English by default here
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input(texts["en"]["report_date"], value=today_str)
st.session_state.prepared_by = st.text_input(texts["en"]["prepared_by"], value=st.session_state.prepared_by)

# ---------------------------
# Language selection for labels (not answers)
# ---------------------------
lang = st.radio("Language / Idioma", ["en", "es"], horizontal=True)

# Update report info labels to selected lang (redraw)
# We must re-render input labels in the chosen language
# Recreate report info inputs with correct labels:
st.session_state.report_date = st.text_input(texts[lang]["report_date"], value=st.session_state.report_date, key="__report_date")
st.session_state.prepared_by = st.text_input(texts[lang]["prepared_by"], value=st.session_state.prepared_by, key="__prepared_by")

# ---------------------------
# Placeholder translation function for manual buttons
# ---------------------------
def translate_text_placeholder(text, target_lang):
    # Placeholder: returns same text (no external API)
    # If you later add an API, replace this function.
    return text

# ---------------------------
# Tabs for each D-step with bilingual labels
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")
        # show guidance in selected language? guidance texts are static English examples;
        # keep them in English for now but you could provide translated guidance too.
        st.info(f"**Guidance:** {note}\n\n💡 **Example:** {example}")

        # D5: interactive only (Occurrence & Detection)
        if step == "D5: Final Analysis":
            st.markdown(f"#### {texts[lang]['d5_occ']}")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(
                    f"{texts[lang]['d5_occ']} - {idx+1}",
                    value=val,
                    key=f"{step}_occ_{idx}"
                )
            if st.button(texts[lang]["add_occ"], key=f"add_occ_{step}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown(f"#### {texts[lang]['d5_det']}")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(
                    f"{texts[lang]['d5_det']} - {idx+1}",
                    value=val,
                    key=f"{step}_det_{idx}"
                )
            if st.button(texts[lang]["add_det"], key=f"add_det_{step}"):
                st.session_state.d5_det_whys.append("")

            # Save aggregated D5 answer and show root cause summary field
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area(
                texts[lang]["root_cause"],
                value=st.session_state[step]["extra"],
                key="d5_root_cause_text"
            )

        else:
            # Other steps: standard answer + extra fields
            st.session_state[step]["answer"] = st.text_area(
                f"{step} - Answer",
                value=st.session_state[step]["answer"],
                key=f"ans_{i}"
            )
            st.session_state[step]["extra"] = st.text_area(
                f"{step} - Root Cause / Extra (if applicable)",
                value=st.session_state[step]["extra"],
                key=f"extra_{i}"
            )

# ---------------------------
# Manual Translation Buttons (placeholders)
# ---------------------------
st.markdown("---")
st.subheader(texts[lang]["translate_section"])
col1, col2 = st.columns(2)
with col1:
    if st.button(texts[lang]["translate_en"]):
        # Translate each stored answer & extra using placeholder
        for step, _, _ in npqp_steps:
            st.session_state[step]["answer"] = translate_text_placeholder(st.session_state[step]["answer"], target_lang="en")
            st.session_state[step]["extra"] = translate_text_placeholder(st.session_state[step]["extra"], target_lang="en")
        st.success("✅ All answers converted to English (placeholder)")

with col2:
    if st.button(texts[lang]["translate_es"]):
        for step, _, _ in npqp_steps:
            st.session_state[step]["answer"] = translate_text_placeholder(st.session_state[step]["answer"], target_lang="es")
            st.session_state[step]["extra"] = translate_text_placeholder(st.session_state[step]["extra"], target_lang="es")
        st.success("✅ All answers converted to Spanish (placeholder)")

# ---------------------------
# Prepare Excel data rows
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save to Excel and download
# ---------------------------
if st.button(texts[lang]["save_button"]):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet.")
    else:
        xlsx_file = f"NPQP_8D_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
        ws["A3"] = texts[lang]["report_date"]
        ws["B3"] = st.session_state.report_date
        ws["A4"] = texts[lang]["prepared_by"]
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Answer", "Root Cause / Extra"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        # Content rows
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

        # Column widths
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        # Save and provide download
        wb.save(xlsx_file)
        st.success("✅ NPQP 8D Report saved successfully")
        with open(xlsx_file, "rb") as f:
            st.download_button(texts[lang]["download_button"], f, file_name=xlsx_file)

# ---------------------------
# Notes / Next steps hint
# ---------------------------
st.markdown("---")
st.info("Notes: Labels and guidance switch with the language selector. The manual translate buttons are placeholders — they do not perform actual translations until you integrate an API (OpenAI or Hugging Face). If you'd like, I can add a free-model Hugging Face option for D5 AI suggestions or wire up OpenAI when you have a standard API key.")
