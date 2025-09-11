import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime

# ---------------------------
# Optional: Google Translate
# ---------------------------
try:
    from googletrans import Translator
    translator = Translator()
except:
    translator = None
    st.warning("Translation library not installed; bilingual answers won't auto-translate.")

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide default menu and footer
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# App header
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📑 8D Training App</h1>", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar idioma", ["English", "Español"])
if "prev_lang" not in st.session_state:
    st.session_state.prev_lang = lang

# ---------------------------
# NPQP steps
# ---------------------------
npqp_steps = [
    ("D1: Concern Details", 
     {"en": "Describe the customer concerns clearly...", "es": "Describa claramente las preocupaciones del cliente..."},
     {"en": "Example: Customer reported static noise...", "es": "Ejemplo: El cliente reportó ruido estático..."}),
    ("D2: Similar Part Considerations",
     {"en": "Check for similar parts...", "es": "Verifique piezas similares..."},
     {"en": "Example: Same speaker type...", "es": "Ejemplo: mismo tipo de altavoz..."}),
    ("D3: Initial Analysis",
     {"en": "Perform an initial investigation...", "es": "Realice una investigación inicial..."},
     {"en": "Example: Visual inspection...", "es": "Ejemplo: inspección visual..."}),
    ("D4: Implement Containment",
     {"en": "Define temporary containment actions...", "es": "Defina acciones de contención temporal..."},
     {"en": "Example: 100% inspection...", "es": "Ejemplo: inspección 100%..."}),
    ("D5: Final Analysis",
     {"en": "Use 5-Why analysis to determine the root cause.", "es": "Use análisis de 5-porqués para determinar la causa raíz."},
     {"en": "", "es": ""}),
    ("D6: Permanent Corrective Actions",
     {"en": "Define corrective actions to eliminate the root cause permanently.", "es": "Defina acciones correctivas para eliminar permanentemente la causa raíz."},
     {"en": "Example: Update soldering process...", "es": "Ejemplo: Actualizar el proceso de soldadura..."}),
    ("D7: Countermeasure Confirmation",
     {"en": "Verify corrective actions resolve the issue long-term.", "es": "Verifique que las acciones correctivas resuelvan el problema a largo plazo."},
     {"en": "Example: Functional tests on corrected units...", "es": "Ejemplo: pruebas funcionales en unidades corregidas..."}),
    ("D8: Follow-up Activities",
     {"en": "Document lessons learned, update standards, procedures...", "es": "Documente lecciones aprendidas, actualice estándares, procedimientos..."},
     {"en": "Example: Update SOPs, PFMEA...", "es": "Ejemplo: Actualizar SOP, PFMEA..."}),
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
st.session_state.setdefault("interactive_whys", [""] * 5)

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
    "D8: Follow-up Activities": "D3D3D3"
}

# ---------------------------
# Translation function
# ---------------------------
def translate_answers(to_lang_code):
    if translator:
        for step, _, _ in npqp_steps:
            ans = st.session_state[step]["answer"]
            if ans.strip():
                st.session_state[step]["answer"] = translator.translate(ans, dest=to_lang_code).text
            extra = st.session_state[step]["extra"]
            if extra.strip():
                st.session_state[step]["extra"] = translator.translate(extra, dest=to_lang_code).text

# Detect language switch
if st.session_state.prev_lang != lang:
    to_lang_code = "es" if lang == "Español" else "en"
    translate_answers(to_lang_code)
    st.session_state.prev_lang = lang

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information" if lang=="English" else "Información del reporte")
st.session_state.report_date = st.text_input("📅 Report Date" if lang=="English" else "📅 Fecha del reporte", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input("✍️ Prepared By" if lang=="English" else "✍️ Preparado por", value=st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, notes, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")

        # Training guidance
        st.info(f"**Training Guidance:** {notes[lang[:2].lower()]}\n\n💡 **Example:** {example[lang[:2].lower()]}")

        # Input fields
        if step.startswith("D5"):
            st.markdown("#### Occurrence Analysis" if lang=="English" else "#### Análisis de ocurrencia")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}" if lang=="English" else f"Por qué de ocurrencia {idx+1}", value=val, key=f"{step}_occ_{idx}")

            st.markdown("#### Detection Analysis" if lang=="English" else "#### Análisis de detección")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}" if lang=="English" else f"Por qué de detección {idx+1}", value=val, key=f"{step}_det_{idx}")

            # Interactive AI suggestions (rule-based)
            for idx in range(5):
                if st.session_state.d5_occ_whys[idx]:
                    # Simple suggestion logic
                    st.text(f"Suggestion for Occurrence Why {idx+1}: Check process controls and previous similar failures.")

            # Combine answers
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )

            # Extra field for Root Cause
            st.session_state[step]["extra"] = st.text_area("Root Cause (summary after 5-Whys)" if lang=="English" else "Causa raíz (resumen después de 5-porqués)", value=st.session_state[step]["extra"])

        else:
            # Only D1 has extra field
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}" if lang=="English" else f"Su respuesta para {step}", value=st.session_state[step]["answer"])
            if step.startswith("D1"):
                st.session_state[step]["extra"] = st.text_area("Additional notes / Notas adicionales", value=st.session_state[step]["extra"])

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button with Excel styling
# ---------------------------
if st.button("💾 Save 8D Report" if lang=="English" else "💾 Guardar reporte 8D"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving." if lang=="English" else "⚠️ No hay respuestas. Por favor complete los campos antes de guardar.")
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
        ws["A3"] = "Report Date" if lang=="English" else "Fecha del reporte"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By" if lang=="English" else "Preparado por"
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
        st.success("✅ NPQP 8D Report saved successfully." if lang=="English" else "✅ Reporte 8D guardado correctamente.")
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX" if lang=="English" else "📥 Descargar XLSX", f, file_name=xlsx_file)
