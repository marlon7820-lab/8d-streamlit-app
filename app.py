import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import openai

# ---------------------------
# Page config and branding
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide Streamlit default menu, header, and footer
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
# Initialize session state
# ---------------------------
steps = ["D1","D2","D3","D4","D5","D6","D7","D8"]
for sid in steps:
    if sid not in st.session_state:
        st.session_state[sid] = {"answer": "", "extra": ""}
    if f"ans_{sid}" not in st.session_state:
        st.session_state[f"ans_{sid}"] = ""

if "interactive_whys" not in st.session_state:
    st.session_state.interactive_whys = [""]

if "prev_lang" not in st.session_state:
    st.session_state.prev_lang = "en"

st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")

# ---------------------------
# Guidelines for each step
# ---------------------------
guidelines = {
    "D1": {"en":"Describe customer concerns clearly.", "es":"Describa claramente las preocupaciones del cliente."},
    "D2": {"en":"Check similar parts and occurrences.", "es":"Verifique partes similares y ocurrencias."},
    "D3": {"en":"Initial analysis and data collection.", "es":"Análisis inicial y recopilación de datos."},
    "D4": {"en":"Define temporary containment actions.", "es":"Defina acciones de contención temporales."},
    "D5": {"en":"Perform 5-Why analysis (Occurrence & Detection).", "es":"Realice análisis de 5 porqués (Ocurrencia y Detección)."},
    "D6": {"en":"Define permanent corrective actions.", "es":"Defina acciones correctivas permanentes."},
    "D7": {"en":"Verify corrective actions effectiveness.", "es":"Verifique la efectividad de las acciones correctivas."},
    "D8": {"en":"Document lessons learned / prevent recurrence.", "es":"Documente lecciones aprendidas / prevenir recurrencia."}
}

# ---------------------------
# Language selection
# ---------------------------
lang = st.radio("Language / Idioma", ["en", "es"], horizontal=True)

# Translate entered answers if language changed
if lang != st.session_state.prev_lang:
    for sid in steps:
        # Translate main answer
        if st.session_state[f"ans_{sid}"]:
            try:
                response = openai.ChatCompletion.create(
                    model="gpt-4",
                    messages=[
                        {"role":"system","content":f"Translate text to {'English' if lang=='en' else 'Spanish'}."},
                        {"role":"user","content":st.session_state[f"ans_{sid}"]}
                    ]
                )
                st.session_state[f"ans_{sid}"] = response.choices[0].message["content"].strip()
            except Exception:
                pass
    # Translate 5-Whys
    st.session_state.interactive_whys = [w for w in st.session_state.interactive_whys]
    st.session_state.prev_lang = lang

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information / Información del Reporte")
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input("📅 Report Date / Fecha del Reporte", value=today_str)
st.session_state.prepared_by = st.text_input("✍️ Prepared By / Preparado por", st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs(steps)
for i, sid in enumerate(steps):
    with tabs[i]:
        st.markdown(f"### {sid}")
        st.info(guidelines[sid][lang])
        st.session_state[f"ans_{sid}"] = st.text_area(
            f"Your Answer / Su Respuesta ({sid})",
            value=st.session_state[f"ans_{sid}"],
            key=f"input_{sid}"
        )
        st.session_state[sid]["answer"] = st.session_state[f"ans_{sid}"]

        # Only D5: interactive 5-Why
        if sid=="D5":
            st.markdown("### Interactive 5-Why / 5-Porqués Interactivo")
            for idx in range(len(st.session_state.interactive_whys)):
                st.session_state.interactive_whys[idx] = st.text_input(
                    f"Why {idx+1} / Por qué {idx+1}?",
                    value=st.session_state.interactive_whys[idx],
                    key=f"why_{idx}"
                )
                # show next input if current filled
                if idx == len(st.session_state.interactive_whys)-1 and st.session_state.interactive_whys[idx].strip():
                    st.session_state.interactive_whys.append("")
            
            if st.button("Generate Root Cause Suggestion / Sugerencia de Causa Raíz"):
                chain = " -> ".join([w for w in st.session_state.interactive_whys if w.strip()])
                st.success(f"Suggested Root Cause / Causa Raíz Sugerida: {chain}")

# ---------------------------
# Save to Excel
# ---------------------------
if st.button("💾 Save 8D Report / Guardar Reporte 8D"):
    xlsx_file = f"NPQP_8D_Report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    # Title
    ws.merge_cells("A1:C1")
    ws["A1"] = "Nissan NPQP 8D Report / Reporte 8D Nissan"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25

    # Report info
    ws["A3"] = "Report Date / Fecha del Reporte"
    ws["B3"] = st.session_state.report_date
    ws["A4"] = "Prepared By / Preparado por"
    ws["B4"] = st.session_state.prepared_by

    # Headers
    headers = ["Step / Paso", "Answer / Respuesta", "Root Cause / Causa Raíz"]
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
        "D1":"ADD8E6","D2":"90EE90","D3":"FFFF99","D4":"FFD580","D5":"FF9999",
        "D6":"D8BFD8","D7":"E0FFFF","D8":"D3D3D3"
    }

    for sid in steps:
        ans = st.session_state[sid]["answer"]
        extra = ""  # Optional root cause column, can include interactive 5why
        if sid=="D5":
            extra = "\n".join([w for w in st.session_state.interactive_whys if w.strip()])

        ws.cell(row=row, column=1, value=sid)
        ws.cell(row=row, column=2, value=ans)
        ws.cell(row=row, column=3, value=extra)
        fill_color = step_colors.get(sid,"FFFFFF")
        for col in range(1,4):
            ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    # Adjust column widths
    for col in range(1,4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    wb.save(xlsx_file)
    st.success("✅ NPQP 8D Report saved successfully / Guardado correctamente.")
    with open(xlsx_file,"rb") as f:
        st.download_button("📥 Download XLSX / Descargar XLSX", f, file_name=xlsx_file)
