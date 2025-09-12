import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import json
import os

# ---------------------------
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

# ---------------------------
# Force Light Mode for consistent colors across devices
# ---------------------------
st.markdown(
    """
    <style>
    body {
        color-scheme: light !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

t = {
    "en": {
        "D1": "D1: Concern Details", "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis", "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis", "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause": "Root Cause (summary after 5-Whys)", "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why", "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause": "Causa raíz (resumen después de los 5 Porqués)", "Occurrence_Why": "Por qué Ocurrencia",
        "Detection_Why": "Por qué Detección", "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo"
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
# Restore from URL (st.query_params)
# ---------------------------
if "backup" in st.query_params:
    try:
        data = json.loads(st.query_params["backup"][0])
        for k, v in data.items():
            st.session_state[k] = v
    except Exception:
        pass

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
                    st.session_state.d5_occ_whys[idx] = st.text_input(
                        f"{t[lang_key]['Occurrence_Why']} {idx+1}", value=val, key=f"occ_{idx}")
                else:
                    suggestions = ["Operator error", "Process not followed", "Equipment malfunction"]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(
                        f"{t[lang_key]['Occurrence_Why']} {idx+1}", [""] + suggestions + [st.session_state.d5_occ_whys[idx]], key=f"occ_{idx}"
                    )

            st.markdown("#### Detection Analysis")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                if idx == 0:
                    st.session_state.d5_det_whys[idx] = st.text_input(
                        f"{t[lang_key]['Detection_Why']} {idx+1}", value=val, key=f"det_{idx}")
                else:
                    suggestions = ["QA checklist incomplete", "No automated test", "Missed inspection"]
                    st.session_state.d5_det_whys[idx] = st.selectbox(
                        f"{t[lang_key]['Detection_Why']} {idx+1}", [""] + suggestions + [st.session_state.d5_det_whys[idx]], key=f"det_{idx}"
                    )

            st.session_state.D5["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state.D5["extra"] = st.text_area(
                f"{t[lang_key]['Root_Cause']}", value=st.session_state.D5["extra"], key="root_cause"
            )

# ---------------------------
# Collect answers
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save / Download Excel
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

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

    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    header_row = ws.max_row + 1
    headers = ["Step", "Answer", "Extra / Notes"]
    fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for step, answer, extra in data_rows:
        ws.append([t[lang_key][step], answer, extra])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ---------------------------
# Sidebar: JSON Backup / Restore + Reset
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore")

    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
        mime="application/json"
    )

    st.markdown("---")
    st.markdown("### Restore from JSON")

    uploaded_file = st.file_uploader("Upload JSON file to restore", type="json")
    if uploaded_file:
        try:
            restore_data = json.load(uploaded_file)
            for k, v in restore_data.items():
                st.session_state[k] = v
            st.success("✅ Session restored from JSON!")
        except Exception as e:
            st.error(f"Error restoring JSON: {e}")

    st.markdown("---")
    st.markdown("### Reset All Data")

    if st.button("🗑️ Clear All"):
        for step, _, _ in npqp_steps:
            if step != "D5":
                st.session_state[step] = {"answer": "", "extra": ""}
        st.session_state["D5"] = {"answer": "", "extra": ""}
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
            st.session_state.setdefault(step, {"answer":"", "extra":""})
        st.success("✅ All data has been reset!")
