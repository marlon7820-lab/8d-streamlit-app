import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(page_title="8D Training App", page_icon="📑", layout="wide")

# Hide default menu/footer/header
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
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Spanish"])
is_en = lang == "English"

# ---------------------------
# Translations
# ---------------------------
texts = {
    "English": {
        "D1": "D1: Concern Details",
        "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis",
        "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis",
        "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation",
        "D8": "D8: Follow-up Activities",
        "report_date": "Report Date",
        "prepared_by": "Prepared By",
        "your_answer": "Your Answer",
        "root_cause": "Root Cause (summary after 5-Whys)",
        "occurrence_analysis": "Occurrence Analysis",
        "detection_analysis": "Detection Analysis",
        "save_button": "💾 Save 8D Report"
    },
    "Spanish": {
        "D1": "D1: Detalles de la preocupación",
        "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial",
        "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final",
        "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas",
        "D8": "D8: Actividades de seguimiento",
        "report_date": "Fecha",
        "prepared_by": "Preparado Por",
        "your_answer": "Tu Respuesta",
        "root_cause": "Causa Raíz (resumen después del 5-Whys)",
        "occurrence_analysis": "Análisis de Ocurrencia",
        "detection_analysis": "Análisis de Detección",
        "save_button": "💾 Guardar Reporte 8D"
    }
}
t = texts[lang]

# ---------------------------
# Step info and training guidance (example)
# ---------------------------
npqp_steps = [
    ("D1", "Describe the customer concerns clearly...", "Example: Customer reported static noise in amplifier..."),
    ("D2", "Check for similar parts, models, generic parts...", "Example: Same speaker type used in another radio model..."),
    ("D3", "Perform an initial investigation...", "Example: Visual inspection of solder joints..."),
    ("D4", "Define temporary containment actions...", "Example: 100% inspection before shipment..."),
    ("D5", "Use 5-Why analysis to determine the root cause.", ""),  # D5 interactive
    ("D6", "Define corrective actions to eliminate the root cause permanently.", "Example: Update soldering process..."),
    ("D7", "Verify that corrective actions effectively resolve the issue long-term.", "Example: Functional tests on corrected amplifiers..."),
    ("D8", "Document lessons learned, update standards to prevent recurrence.", "Example: Update SOPs, PFMEA, training...")
]

# ---------------------------
# Session State Initialization
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
st.subheader(t["report_date"])
st.session_state.report_date = st.text_input(t["report_date"], st.session_state.report_date)
st.subheader(t["prepared_by"])
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([t[step] for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[step]}")
        if step != "D5":
            st.info(f"{note}\n💡 {example}")
            st.session_state[step]["answer"] = st.text_area(t["your_answer"], st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.info(note)
            # Interactive 5-Why suggestions
            def get_5why_suggestions(prev_whys):
                # Simple context-aware rules
                context = {
                    "solder": ["Cold solder joint", "Solder bridge", "Insufficient heating"],
                    "component": ["Wrong component", "Defective batch", "Misplaced polarity"],
                    "process": ["Step skipped", "Operator not trained", "Incorrect assembly"],
                    "inspection": ["Visual inspection missed", "Test step skipped"]
                }
                prev_text = " ".join(prev_whys).lower()
                suggestions = []
                for k, v in context.items():
                    if k in prev_text:
                        suggestions.extend(v)
                if not suggestions:
                    suggestions = sum(context.values(), [])
                # Remove duplicates & already used
                suggestions = [s for s in suggestions if s not in prev_whys]
                return suggestions[:5]

            st.markdown(f"#### {t['occurrence_analysis']}")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                prev = st.session_state.d5_occ_whys[:idx]
                options = [""] + get_5why_suggestions(prev) + [st.session_state.d5_occ_whys[idx]]
                st.session_state.d5_occ_whys[idx] = st.selectbox(f"Occurrence Why {idx+1}", options, index=0 if not st.session_state.d5_occ_whys[idx] else 1, key=f"occ_{idx}")

            st.markdown(f"#### {t['detection_analysis']}")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                prev = st.session_state.d5_det_whys[:idx]
                options = [""] + get_5why_suggestions(prev) + [st.session_state.d5_det_whys[idx]]
                st.session_state.d5_det_whys[idx] = st.selectbox(f"Detection Why {idx+1}", options, index=0 if not st.session_state.d5_det_whys[idx] else 1, key=f"det_{idx}")

            st.session_state.D5["answer"] = "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) + "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            st.session_state.D5["extra"] = st.text_area(t["root_cause"], st.session_state.D5["extra"])

# ---------------------------
# Save button / Excel export
# ---------------------------
if st.button(t["save_button"]):
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    ws.merge_cells("A1:C1")
    ws["A1"] = "Nissan NPQP 8D Report"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws["A3"] = t["report_date"]
    ws["B3"] = st.session_state.report_date
    ws["A4"] = t["prepared_by"]
    ws["B4"] = st.session_state.prepared_by

    headers = ["Step", "Answer", "Root Cause"]
    header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=6, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    row = 7
    step_colors = {
        "D1": "ADD8E6", "D2": "90EE90", "D3": "FFFF99", "D4": "FFD580",
        "D5": "FF9999", "D6": "D8BFD8", "D7": "E0FFFF", "D8": "D3D3D3"
    }
    for step, _, _ in npqp_steps:
        ans = st.session_state[step]["answer"]
        extra = st.session_state[step]["extra"]
        ws.cell(row=row, column=1, value=t[step])
        ws.cell(row=row, column=2, value=ans)
        ws.cell(row=row, column=3, value=extra)
        fill_color = step_colors.get(step, "FFFFFF")
        for col in range(1, 4):
            ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    xlsx_file = "NPQP_8D_Report.xlsx"
    wb.save(xlsx_file)
    st.success("✅ Report saved successfully")
    with open(xlsx_file, "rb") as f:
        st.download_button("📥 Download XLS
