import streamlit as st
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import os

# ---------------------------
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
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

# App title
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# App-wide custom CSS
# ---------------------------
st.markdown("""
    <style>
    .stApp {
        background-color: #F5F7FA;
    }
    .stTabs [role="tab"] {
        font-weight: bold;
        background-color: #E6F0FA;
        color: #1E1E1E;
        border-radius: 5px;
    }
    .stTabs [role="tab"]:hover {
        background-color: #CDE0F7;
    }
    .stSidebar h2, .stSidebar h3 {
        color: #1E90FF;
    }
    .stSidebar button {
        background-color: #1E90FF;
        color: white;
        font-weight: bold;
        border-radius: 5px;
    }
    .stSidebar button:hover {
        background-color: #1C86EE;
        color: white;
    }
    .stTextArea textarea {
        border-radius: 5px;
        border: 1px solid #C3D7F0;
        padding: 8px;
        background-color: #FFFFFF;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

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

# ---------------------------
# NPQP 8D steps
# ---------------------------
npqp_steps = [
    ("D1", "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.", "Customer reported static noise in amplifier during end-of-line test."),
    ("D2", "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc.", "Same speaker type used in another radio model; different amplifier colors."),
    ("D3", "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.", "Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4", "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.", "100% inspection of amplifiers before shipment; temporary shielding."),
    ("D5", "Use 5-Why analysis to determine the root cause. Separate Occurrence and Detection.", ""),
    ("D6", "Define corrective actions that eliminate the root cause permanently and prevent recurrence.", "Update soldering process, retrain operators, update work instructions."),
    ("D7", "Verify that corrective actions effectively resolve the issue long-term.", "Functional tests on corrected amplifiers, accelerated life testing."),
    ("D8", "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.", "Update SOPs, PFMEA, work instructions, and employee training.")
]

# ---------------------------
# Initialize session state
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
st.subheader(f"{t[lang_key]['Report_Date']}")
st.session_state.report_date = st.text_input(f"{t[lang_key]['Report_Date']}", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(f"{t[lang_key]['Prepared_By']}", value=st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([t[lang_key][step] for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        if step != "D5":
            st.info(f"**{t[lang_key]['Training_Guidance']}:** {note}\n\n💡 **{t[lang_key]['Example']}:** {example}")
            st.session_state[step]["answer"] = st.text_area(f"Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}")
        else:
            st.info(f"**{t[lang_key]['Training_Guidance']}:** {note}")
            st.markdown("#### Occurrence Analysis")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"{t[lang_key]['Occurrence_Why']} {idx+1}", value=val, key=f"occ_{idx}")
                else:
                    suggestions = ["Operator error", "Process not followed", "Equipment malfunction"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"{t[lang_key]['Occurrence_Why']} {idx+1}", [""] + suggestions + [st.session_state.d5_occ_whys[idx]], key=f"occ_{idx}")

            st.markdown("#### Detection Analysis")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(f"{t[lang_key]['Detection_Why']} {idx+1}", value=val, key=f"det_{idx}")
                else:
                    suggestions = ["QA checklist incomplete", "No automated test", "Missed inspection"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(f"{t[lang_key]['Detection_Why']} {idx+1}", [""] + suggestions + [st.session_state.d5_det_whys[idx]], key=f"det_{idx}")

            st.session_state.D5["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state.D5["extra"] = st.text_area(f"{t[lang_key]['Root_Cause']}", value=st.session_state.D5["extra"], key="root_cause")

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save / Download Excel with formatting
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Add logo
    if os.path.exists("logo.png"):
        try:
            img = XLImage("logo.png")
            img.width = 140
            img.height = 40
            ws.add_image(img, "A1")
        except:
            pass

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
    ws.cell(row=3, column=1, value="📋 8D Report Assistant").font = Font(bold=True, size=14)

    # Report info
    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    # Header row
    header_row = ws.max_row + 1
    headers = ["Step", "Answer", "Extra / Notes"]
    header_fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Row fills
    answer_color1 = PatternFill(start_color="E6F0FA", end_color="E6F0FA", fill_type="solid")
    answer_color2 = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    step_color = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    # Add data rows
    for idx, (step, answer, extra) in enumerate(data_rows):
        ws.append([t[lang_key][step], answer, extra])
        r = ws.max_row

        # Step cell
        cell_step = ws.cell(row=r, column=1)
        cell_step.font = Font(bold=True)
        cell_step.fill = step_color
        cell_step.alignment = Alignment(horizontal="center", vertical="top")
        cell_step.border = border

        # Answer & Extra
        row_fill = answer_color1 if idx % 2 == 0 else answer_color2
        for c in [2, 3]:
            cell = ws.cell(row=r, column=c)
            cell.font = Font(bold=True)
            cell.fill = row_fill
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border

    # Adjust column widths
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# Download button
st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
