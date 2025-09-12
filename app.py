import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import datetime

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="📄",
    layout="wide"
)

# Hide menu/footer/header
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
lang = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
is_spanish = lang == "Español"

# Translation dictionary
t = {
    "en": {
        "D1": "D1: Concern Details",
        "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis",
        "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis",
        "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation",
        "D8": "D8: Follow-up Activities",
        "Report Date": "Report Date",
        "Prepared By": "Prepared By",
        "Save Report": "💾 Save 8D Report",
        "Download Report": "📥 Download XLSX",
        "Occurrence Why": "Occurrence Why",
        "Detection Why": "Detection Why",
        "Root Cause": "Root Cause (summary after 5-Whys)"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación",
        "D2": "D2: Consideraciones de piezas similares",
        "D3": "D3: Análisis inicial",
        "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final",
        "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas",
        "D8": "D8: Actividades de seguimiento",
        "Report Date": "Fecha del reporte",
        "Prepared By": "Preparado por",
        "Save Report": "💾 Guardar reporte 8D",
        "Download Report": "📥 Descargar XLSX",
        "Occurrence Why": "Por qué (Ocurrencia)",
        "Detection Why": "Por qué (Detección)",
        "Root Cause": "Causa raíz (resumen después del 5-Whys)"
    }
}[lang[:2].lower()]

# ---------------------------
# NPQP 8D steps and examples
# ---------------------------
npqp_steps = [
    ("D1", "Describe the customer concerns clearly.", "Example: Customer reported static noise in amplifier."),
    ("D2", "Check for similar parts/models.", "Example: Same speaker type used in other radio models."),
    ("D3", "Perform initial investigation and document findings.", "Example: Visual inspection of solder joints."),
    ("D4", "Define temporary containment actions.", "Example: Quarantine affected batches."),
    ("D5", "Use 5-Why to determine root cause (Occurrence/Detection).", ""),
    ("D6", "Define permanent corrective actions.", "Example: Update soldering process."),
    ("D7", "Verify corrective actions work long-term.", "Example: Functional testing."),
    ("D8", "Document lessons learned and update standards.", "Example: Update SOPs and training.")
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Report info
# ---------------------------
st.subheader(t["Report Date"])
st.session_state.report_date = st.text_input(t["Report Date"], st.session_state.report_date)
st.subheader(t["Prepared By"])
st.session_state.prepared_by = st.text_input(t["Prepared By"], st.session_state.prepared_by)

# ---------------------------
# Tabs for D1-D8
# ---------------------------
tabs = st.tabs([t[step] for step, _, _ in npqp_steps])
for i, (step, guidance, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[step]}")
        if step != "D5":
            st.info(f"**Guidance:** {guidance}\n\n💡 **Example:** {example}")
            st.session_state[step]["answer"] = st.text_area(f"Your answer for {t[step]}", st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.info("**Occurrence Analysis:**")
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"{t['Occurrence Why']} {idx+1}", value=st.session_state.d5_occ_whys[idx], key=f"occ_{idx}")
                else:
                    options = ["Incorrect process", "Insufficient training", "Equipment issue", "Material defect"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"{t['Occurrence Why']} {idx+1}", options, index=0, key=f"occ_{idx}")

            st.info("**Detection Analysis:**")
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"{t['Detection Why']} {idx+1}", value=st.session_state.d5_det_whys[idx], key=f"det_{idx}")
                else:
                    options = ["Inspection missed defect", "Checklist incomplete", "No automated test", "Batch testing not performed"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(f"{t['Detection Why']} {idx+1}", options, index=0, key=f"det_{idx}")

            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["root_cause"] = st.text_area(t["Root Cause"], value=st.session_state[step].get("root_cause", ""), key="root_cause_d5")

# ---------------------------
# Save button and Excel export
# ---------------------------
if st.button(t["Save Report"]):
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
    ws["A3"] = t["Report Date"]
    ws["B3"] = st.session_state.report_date
    ws["A4"] = t["Prepared By"]
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
    step_colors = {
        "D1": "ADD8E6", "D2": "90EE90", "D3": "FFFF99", "D4": "FFD580",
        "D5": "FF9999", "D6": "D8BFD8", "D7": "E0FFFF", "D8": "D3D3D3"
    }
    for step, _, _ in npqp_steps:
        ans = st.session_state[step]["answer"]
        extra = st.session_state[step].get("root_cause", "")
        ws.cell(row=row, column=1, value=t[step])
        ws.cell(row=row, column=2, value=ans)
        ws.cell(row=row, column=3, value=extra)
        for col in range(1, 4):
            ws.cell(row=row, column=col).fill = PatternFill(start_color=step_colors.get(step, "FFFFFF"), end_color=step_colors.get(step, "FFFFFF"), fill_type="solid")
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    # Adjust column widths
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    # Save workbook in-memory and provide download
    from io import BytesIO
    output = BytesIO()
    wb.save(output)
    st.download_button(
        label=t["Download Report"],
        data=output.getvalue(),
        file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
