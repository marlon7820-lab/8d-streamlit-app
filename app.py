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

hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📑 8D Training App</h1>", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Language / Idioma", ["English", "Español"])
en = lang == "English"

# ---------------------------
# 8D Steps with guidance/examples
# ---------------------------
npqp_steps = [
    ("D1: Concern Details", 
     "Describe the customer concerns clearly.", 
     "Example: Customer reported static noise in amplifier."),
    ("D2: Similar Part Considerations",
     "Check for similar parts/models to see if issue is recurring or isolated.",
     "Example: Same speaker type used in another radio model."),
    ("D3: Initial Analysis",
     "Perform initial investigation and document findings.",
     "Example: Visual inspection of solder joints."),
    ("D4: Implement Containment",
     "Define temporary containment actions to prevent customer impact.",
     "Example: 100% inspection of amplifiers before shipment."),
    ("D5: Final Analysis (Root Cause / 5-Why)", 
     "Use interactive 5-Why for Occurrence and Detection.",
     ""),
    ("D6: Permanent Corrective Actions",
     "Define corrective actions that eliminate the root cause permanently.",
     "Example: Update soldering process, retrain operators."),
    ("D7: Countermeasure Confirmation",
     "Verify that corrective actions effectively resolve the issue long-term.",
     "Example: Functional tests on corrected amplifiers."),
    ("D8: Follow-up Activities",
     "Document lessons learned, update standards/procedures to prevent recurrence.",
     "Example: Update SOPs, PFMEA, work instructions, training.")
]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")

# D5 interactive state
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# Example suggestion mapping for interactive 5-Why
why_suggestions = {
    "Occurrence": {
        "": ["Cold solder joint", "Incorrect part used", "Loose connector"],
        "Cold solder joint": ["Soldering temperature too low", "Operator error"],
        "Incorrect part used": ["Supplier sent wrong component", "Inventory labeling issue"],
        "Loose connector": ["Assembly not secured", "Design tolerance too tight"]
    },
    "Detection": {
        "": ["Inspection missed defect", "No automated test", "Checklist incomplete"],
        "Inspection missed defect": ["Inspector not trained", "Defect subtle"],
        "No automated test": ["Test step missing", "Equipment unavailable"],
        "Checklist incomplete": ["Step missing in SOP", "QA overlooked step"]
    }
}

# ---------------------------
# Excel colors
# ---------------------------
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis (Root Cause / 5-Why)": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information" if en else "Información del Reporte")
st.session_state.report_date = st.text_input("Report Date / Fecha", st.session_state.report_date)
st.session_state.prepared_by = st.text_input("Prepared By / Preparado por", st.session_state.prepared_by)

# ---------------------------
# Tabs
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}" if en else f"### {step.replace('D1','D1').replace('D2','D2')}")  # You can expand for translation

        # Show guidance and examples except D5
        if not step.startswith("D5"):
            st.info(f"**Training Guidance / Guía:** {note}\n\n💡 **Example / Ejemplo:** {example}")
            st.session_state[step]["answer"] = st.text_area("Your Answer / Su respuesta", value=st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.info("**Interactive 5-Why for Occurrence & Detection / 5-Why Interactivo**")
            st.markdown("#### Occurrence / Ocurrencia")
            for idx in range(5):
                prev = st.session_state.d5_occ_whys[idx-1] if idx > 0 else ""
                options = why_suggestions["Occurrence"].get(prev, ["Other"])
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"Why {idx+1}", value=st.session_state.d5_occ_whys[idx], key=f"d5_occ_{idx}")
                else:
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"Why {idx+1}", options, index=0, key=f"d5_occ_{idx}")

            st.markdown("#### Detection / Detección")
            for idx in range(5):
                prev = st.session_state.d5_det_whys[idx-1] if idx > 0 else ""
                options = why_suggestions["Detection"].get(prev, ["Other"])
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"Why {idx+1}", value=st.session_state.d5_det_whys[idx], key=f"d5_det_{idx}")
                else:
                    st.session_state.d5_det_whys[idx] = st.selectbox(f"Why {idx+1}", options, index=0, key=f"d5_det_{idx}")

            # Combine into answer
            st.session_state[step]["answer"] = (
                "Occurrence:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area("Root Cause / Causa Raíz", value=st.session_state[step]["extra"], key="d5_root")

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button
# ---------------------------
if st.button("💾 Save 8D Report / Guardar Reporte"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet / No hay respuestas")
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
        ws["A3"] = "Report Date / Fecha"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By / Preparado por"
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Your Answer / Respuesta", "Root Cause / Causa Raíz"]
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
            for col in range(1,4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=
                                                               )
