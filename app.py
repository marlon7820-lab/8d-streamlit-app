import streamlit as st
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import openai

# ---------------------------
# Page Config & Branding
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

# ---------------------------
# Language Selector
# ---------------------------
lang = st.selectbox("🌐 Select Language / Seleccione Idioma", ["English", "Español"])
lang_code = "en" if lang == "English" else "es"
st.session_state.setdefault("prev_lang", lang_code)

# ---------------------------
# Translations for UI
# ---------------------------
ui_texts = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_info": "Report Information",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_btn": "💾 Save 8D Report",
        "download_btn": "📥 Download XLSX",
        "ai_helper": "🤖 AI Helper (Optional)",
        "add_occ": "➕ Add another Occurrence Why",
        "add_det": "➕ Add another Detection Why",
        "ai_btn": "Generate AI Suggestions",
        "ai_output": "AI Suggestions (copy/edit into fields as needed)",
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_info": "Información del Reporte",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado Por",
        "save_btn": "💾 Guardar Reporte 8D",
        "download_btn": "📥 Descargar XLSX",
        "ai_helper": "🤖 Asistente de IA (Opcional)",
        "add_occ": "➕ Agregar otro Porqué de Ocurrencia",
        "add_det": "➕ Agregar otro Porqué de Detección",
        "ai_btn": "Generar Sugerencias con IA",
        "ai_output": "Sugerencias de IA (copiar/editar según sea necesario)",
    }
}
t = ui_texts[lang_code]

st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['app_title']}</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = [
    ("D1", "D1: Concern Details", "Describe the customer concerns clearly.", "Describa claramente las preocupaciones del cliente."),
    ("D2", "D2: Similar Part Considerations", "Check for similar parts, models, etc.", "Verifique piezas, modelos o colores similares."),
    ("D3", "D3: Initial Analysis", "Perform initial investigation and document findings.", "Realice la investigación inicial y documente hallazgos."),
    ("D4", "D4: Implement Containment", "Define temporary containment actions.", "Defina acciones de contención temporales."),
    ("D5", "D5: Final Analysis (Root Cause)", "Use 5-Why analysis to determine root cause.", "Use el análisis de 5-porqués para determinar la causa raíz."),
    ("D6", "D6: Permanent Corrective Actions", "Define corrective actions to eliminate root cause.", "Defina acciones correctivas permanentes."),
    ("D7", "D7: Countermeasure Confirmation", "Verify corrective actions are effective.", "Verifique que las acciones correctivas sean efectivas."),
    ("D8", "D8: Follow-up Activities", "Document lessons learned and prevent recurrence.", "Documente lecciones aprendidas y prevenga recurrencias."),
]

# ---------------------------
# Session State Initialization
# ---------------------------
for sid, _, _, _ in npqp_steps:
    if sid not in st.session_state:
        st.session_state[sid] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Translate Answers on Language Switch
# ---------------------------
if st.session_state.prev_lang != lang_code:
    openai.api_key = st.secrets.get("OPENAI_API_KEY", "")
    if openai.api_key:
        for sid, _, _, _ in npqp_steps:
            for field in ["answer", "extra"]:
                text = st.session_state[sid][field]
                if text.strip():
                    try:
                        prompt = f"Translate this 8D report text from {'English' if st.session_state.prev_lang == 'en' else 'Spanish'} to {'English' if lang_code == 'en' else 'Spanish'}:\n{text}"
                        response = openai.ChatCompletion.create(
                            model="gpt-4",
                            messages=[{"role": "user", "content": prompt}],
                        )
                        st.session_state[sid][field] = response.choices[0].message.content
                    except:
                        pass
    st.session_state.prev_lang = lang_code

# ---------------------------
# Report Info
# ---------------------------
st.subheader(t["report_info"])
st.session_state.report_date = st.text_input(t["report_date"], value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], value=st.session_state.prepared_by)

# ---------------------------
# Tabs for Each D Step
# ---------------------------
tabs = st.tabs([title for _, title, _, _ in npqp_steps])
for i, (sid, title, note_en, note_es) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {title}")
        st.info(note_en if lang_code == "en" else note_es)

        if sid == "D5":
            st.markdown("#### Occurrence Analysis")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=val, key=f"{sid}_occ_{idx}")
            if st.button(t["add_occ"], key=f"add_occ_{sid}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### Detection Analysis")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", value=val, key=f"{sid}_det_{idx}")
            if st.button(t["add_det"], key=f"add_det_{sid}"):
                st.session_state.d5_det_whys.append("")

            st.session_state[sid]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state[sid]["extra"] = st.text_area("Root Cause (summary after 5-Whys)", value=st.session_state[sid]["extra"], key="root_cause")

            # AI Helper
            st.markdown(f"### {t['ai_helper']}")
            if st.button(t["ai_btn"], key="ai_suggest"):
                openai.api_key = st.secrets.get("OPENAI_API_KEY", "")
                if not openai.api_key:
                    st.warning("⚠️ Please set your OpenAI API key in Streamlit secrets.")
                else:
                    occ_text = "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()])
                    det_text = "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
                    prompt = f"""
You are an expert NPQP 8D coach. Language: {"English" if lang_code == "en" else "Spanish"}.
Current inputs:

Occurrence Whys:
{occ_text}

Detection Whys:
{det_text}

Suggest 1-3 additional Whys for each (Occurrence and Detection) if missing, 
and summarize a Root Cause statement.
"""
                    try:
                        response = openai.ChatCompletion.create(
                            model="gpt-4",
                            messages=[{"role": "user", "content": prompt}],
                        )
                        ai_text = response.choices[0].message.content
                        st.text_area(t["ai_output"], value=ai_text, height=250)
                    except Exception as e:
                        st.error(f"AI generation failed: {e}")
        else:
            st.session_state[sid]["answer"] = st.text_area(f"Your Answer for {title}", value=st.session_state[sid]["answer"], key=f"ans_{sid}")

# ---------------------------
# Save Button & Excel Export
# ---------------------------
if st.button(t["save_btn"]):
    data_rows = [(title, st.session_state[sid]["answer"], st.session_state[sid]["extra"]) for sid, title, _, _ in npqp_steps]
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet.")
    else:
        xlsx_file = "NPQP_8D_Report.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "NPQP 8D Report"
        ws.merge_cells("A1:C1")
        ws["A1"] = "Nissan NPQP 8D Report"
        ws["A1"].font = Font(size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A3"] = "Report Date"; ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By"; ws["B4"] = st.session_state.prepared_by

        headers = ["Step", "Your Answer", "Root Cause"]
        header_fill = PatternFill(start_color="C0C0C0", fill_type="solid")
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=6, column=col, value=header)
            cell.font = Font(bold=True); cell.alignment = Alignment(horizontal="center"); cell.fill = header_fill

        row = 7
        for step, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=step)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)
            row += 1

        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)
        st.success("✅ Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button(t["download_btn"], f, file_name=xlsx_file)
