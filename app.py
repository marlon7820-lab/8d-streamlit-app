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

# Hide Streamlit default menu, header, and footer
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
lang = st.selectbox("Language / Idioma", ["English", "Español"])

# Text labels
t = {}
if lang=="English":
    t = {
        "report_date":"Report Date",
        "prepared_by":"Prepared By",
        "occurrence":"Occurrence Analysis",
        "detection":"Detection Analysis",
        "why":"Why",
        "suggestions":"Suggestions",
        "answer":"Your Answer",
        "root_cause":"Root Cause",
        "save":"💾 Save 8D Report",
        "download":"📥 Download XLSX"
    }
else:
    t = {
        "report_date":"Fecha",
        "prepared_by":"Preparado Por",
        "occurrence":"Análisis de Ocurrencia",
        "detection":"Análisis de Detección",
        "why":"Por Qué",
        "suggestions":"Sugerencias",
        "answer":"Su Respuesta",
        "root_cause":"Causa Raíz",
        "save":"💾 Guardar Reporte 8D",
        "download":"📥 Descargar XLSX"
    }

# ---------------------------
# 8D Steps
# ---------------------------
npqp_steps = [
    ("D1", "Concern Details", "Describe the customer concerns clearly.", ["Customer reported static noise in amplifier","Customer received wrong part","Customer reported overheating"]),
    ("D2", "Similar Part Considerations", "Check similar parts or models.", ["Same speaker used in another radio model","Check same batch of PCB","Similar connector in other assembly"]),
    ("D3", "Initial Analysis", "Perform initial investigation.", ["Visual inspection","Functional tests","Check logs for error codes"]),
    ("D4", "Implement Containment", "Define temporary containment actions.", ["Quarantine affected batches","Immediate replacement","Customer notification"]),
    ("D5", "Final Analysis", "Perform root cause analysis with 5-Why.", ["Check occurrence root cause","Check detection gap","Review process documentation"]),
    ("D6", "Permanent Corrective Actions", "Define corrective actions.", ["Update process","Retrain operators","Modify inspection procedure"]),
    ("D7", "Countermeasure Confirmation", "Verify effectiveness of corrective actions.", ["Functional tests","Monitor next batch","Customer feedback verification"]),
    ("D8", "Follow-up Activities", "Document lessons learned.", ["Update SOPs","Update FMEA","Team training"])
]

# ---------------------------
# Language mapping for step titles
# ---------------------------
step_titles = {
    "English": {
        "D1":"D1: Concern Details",
        "D2":"D2: Similar Part Considerations",
        "D3":"D3: Initial Analysis",
        "D4":"D4: Implement Containment",
        "D5":"D5: Final Analysis",
        "D6":"D6: Permanent Corrective Actions",
        "D7":"D7: Countermeasure Confirmation",
        "D8":"D8: Follow-up Activities"
    },
    "Español": {
        "D1":"D1: Detalles de la Preocupación",
        "D2":"D2: Consideraciones de Partes Similares",
        "D3":"D3: Análisis Inicial",
        "D4":"D4: Implementar Contención",
        "D5":"D5: Análisis Final",
        "D6":"D6: Acciones Correctivas Permanentes",
        "D7":"D7: Confirmación de Contramedidas",
        "D8":"D8: Actividades de Seguimiento"
    }
}

# ---------------------------
# Session state initialization
# ---------------------------
for step, _, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")

# D5 interactive lists
st.session_state.setdefault("d5_occ", [""]*5)
st.session_state.setdefault("d5_det", [""]*5)

# Excel colors
step_colors = {
    "D1": "ADD8E6",
    "D2": "90EE90",
    "D3": "FFFF99",
    "D4": "FFD580",
    "D5": "FF9999",
    "D6": "D8BFD8",
    "D7": "E0FFFF",
    "D8": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information")
st.session_state.report_date = st.text_input(t["report_date"], value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], value=st.session_state.prepared_by)

# ---------------------------
# Tabs for steps
# ---------------------------
tabs = st.tabs([step_titles[lang][step] for step,_,_,_ in npqp_steps])
for i, (step, name, note, examples) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step_titles[lang][step]}")
        if step!="D5":
            st.info(f"**Training Guidance:** {note}\n\n💡 **Examples:** {', '.join(examples)}")
            st.session_state[step]["answer"] = st.text_area(t["answer"], value=st.session_state[step]["answer"], key=f"ans_{step}")
            if step=="D1":
                st.session_state[step]["extra"] = st.text_area(t["root_cause"], value=st.session_state[step]["extra"], key=f"extra_{step}")
        else:
            st.info(f"**Training Guidance:** {note}\n\n💡 **Examples:** {', '.join(examples)}")
            # Interactive 5-Why
            def get_suggestions(prev):
                if not prev: return []
                k = prev.lower()
                suggestions=[]
                if "operator" in k or "operador" in k:
                    suggestions = ["Operator skipped step","Operator misread instructions","Operator not trained properly"]
                elif "process" in k or "proceso" in k:
                    suggestions = ["Process not standardized","Process not monitored","Equipment settings incorrect"]
                elif "inspection" in k or "inspección" in k:
                    suggestions = ["Inspection step missing","Checklist incomplete","Test not performed"]
                return suggestions[:3]

            st.markdown(f"#### {t['occurrence']}")
            for idx in range(5):
                prev = st.session_state.d5_occ[idx-1] if idx>0 else ""
                sug = get_suggestions(prev)
                if sug:
                    st.markdown(f"💡 {t['suggestions']}: {', '.join(sug)}")
                st.session_state.d5_occ[idx] = st.text_input(f"{t['why']} {idx+1}", value=st.session_state.d5_occ[idx], key=f"occ_{idx}")

            st.markdown(f"#### {t['detection']}")
            for idx in range(5):
                prev = st.session_state.d5_det[idx-1] if idx>0 else ""
                sug = get_suggestions(prev)
                if sug:
                    st.markdown(f"💡 {t['suggestions']}: {', '.join(sug)}")
                st.session_state.d5_det[idx] = st.text_input(f"{t['why']} {idx+1}", value=st.session_state.d5_det[idx], key=f"det_{idx}")

            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det if w.strip()])
            )
            st.session_state[step]["extra"] = st.text_area(t["root_cause"], value=st.session_state[step]["extra"], key="root_cause_d5")

# ---------------------------
# Collect data rows
# ---------------------------
data_rows = []
for step,_ ,_,_ in npqp_steps:
    ans = st.session_state[step]["answer"]
    extra = st.session_state[step]["extra"]
    data_rows.append((step, ans, extra))

# ---------------------------
# Save button and Excel export
# ---------------------------
if st.button(t["save"]):
    if not any(ans for _,ans,_ in data_rows):
        st.error("⚠️ No answers filled yet.")
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
        ws["A3"] = t["report_date"]
        ws["B3"] = st.session_state.report_date
        ws["A4"] = t["prepared_by"]
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
            ws.cell(row=row, column=1, value=step_titles[lang][step])
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)
            fill_color = step_colors.get(step, "FFFFFF")
            for col in range(1,4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            row += 1

        # Adjust column widths
        for col in range(1,4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)
        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file,"rb") as f:
            st.download_button(t["download"], f, file_name=xlsx_file)
