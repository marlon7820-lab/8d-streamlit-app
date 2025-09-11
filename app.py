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

# Hide Streamlit default menu, header, and footer
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Language selector (visible label), internal key = 'en' or 'es'
# ---------------------------
lang_choice = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
lang = "en" if lang_choice == "English" else "es"

# ---------------------------
# UI strings (common)
# ---------------------------
ui_texts = {
    "en": {
        "header": "📑 8D Training App",
        "report_info": "Report Information",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_report": "💾 Save 8D Report",
        "download": "📥 Download XLSX",
        "no_answers": "⚠️ No answers filled in yet. Please complete some fields before saving.",
        "ai_helper": "💡 AI Helper Suggestions",
        "current_features_tab": "Current Features",
        "interactive_5why_tab": "Interactive 5-Why",
        "ai_tab": "AI Helper"
    },
    "es": {
        "header": "📑 Aplicación de Entrenamiento 8D",
        "report_info": "Información del Reporte",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado Por",
        "save_report": "💾 Guardar Reporte 8D",
        "download": "📥 Descargar XLSX",
        "no_answers": "⚠️ No se han completado respuestas. Por favor complete algunos campos antes de guardar.",
        "ai_helper": "💡 Sugerencias del Asistente AI",
        "current_features_tab": "Funciones Actuales",
        "interactive_5why_tab": "5-Why Interactivo",
        "ai_tab": "Asistente AI"
    }
}
t = ui_texts[lang]

# Header
st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['header']}</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP content for both languages
# Use language-independent step IDs so session state isn't lost when switching language.
# ---------------------------
step_ids = ["D1","D2","D3","D4","D5","D6","D7","D8"]

npqp_steps_texts = {
    "en": {
        "D1": ("D1: Concern Details",
               "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
               "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
        "D2": ("D2: Similar Part Considerations",
               "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
               "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
        "D3": ("D3: Initial Analysis",
               "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
               "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
        "D4": ("D4: Implement Containment",
               "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
               "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
        "D5": ("D5: Final Analysis",
               "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected). Add more Whys if needed.",
               ""),  # D5 example/guidance shown dynamically
        "D6": ("D6: Permanent Corrective Actions",
               "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
               "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
        "D7": ("D7: Countermeasure Confirmation",
               "Verify that corrective actions effectively resolve the issue long-term.",
               "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
        "D8": ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
               "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
               "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future.")
    },
    "es": {
        "D1": ("D1: Detalles de la Preocupación",
               "Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de soporte.",
               "Ejemplo: El cliente reportó ruido estático en el amplificador durante la prueba final en Planta A."),
        "D2": ("D2: Consideraciones de Piezas Similares",
               "Revise piezas, modelos, colores, manos opuestas, frontales/traseras, etc. para ver si el problema es recurrente o aislado.",
               "Ejemplo: Mismo tipo de altavoz usado en otro modelo de radio; diferentes colores de amplificador; unidades de audio frontales vs traseras."),
        "D3": ("D3: Análisis Inicial",
               "Realice una investigación inicial para identificar problemas evidentes, recopilar datos y documentar hallazgos.",
               "Ejemplo: Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores."),
        "D4": ("D4: Implementar Contención",
               "Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes.",
               "Ejemplo: Inspección 100% de amplificadores antes del envío; uso de blindaje temporal; cuarentena de lotes afectados."),
        "D5": ("D5: Análisis Final",
               "Use análisis de 5 porqués para determinar la causa raíz. Separe por Ocurrencia (por qué sucedió) y Detección (por qué no se detectó). Agregue más porqués si es necesario.",
               ""),
        "D6": ("D6: Acciones Correctivas Permanentes",
               "Defina acciones correctivas que eliminen la causa raíz permanentemente y prevengan recurrencias.",
               "Ejemplo: Actualizar proceso de soldadura, reentrenar operadores, actualizar instrucciones de trabajo y agregar inspección automatizada."),
        "D7": ("D7: Confirmación de Contramedidas",
               "Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo.",
               "Ejemplo: Pruebas funcionales en amplificadores corregidos, pruebas aceleradas de vida útil, y monitoreo de primeras unidades de producción."),
        "D8": ("D8: Actividades de Seguimiento (Lecciones Aprendidas / Prevención de Recurrencias)",
               "Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y entrenamiento para prevenir recurrencia.",
               "Ejemplo: Actualizar SOPs, PFMEA, instrucciones de trabajo y capacitación de empleados para prevenir el mismo problema en el futuro.")
    }
}

# Build localized list preserving canonical IDs order
npqp_steps = [ (sid, ) + npqp_steps_texts[lang][sid] for sid in step_ids ]
# Each item: (step_id, localized_title, guidance, example)

# ---------------------------
# Initialize session state (use canonical IDs)
# ---------------------------
for sid, localized_title, guidance, example in npqp_steps:
    if sid not in st.session_state:
        st.session_state[sid] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
# D5 specific whys
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
# Interactive 5-Why tab state
st.session_state.setdefault("interactive_whys", [""])
st.session_state.setdefault("interactive_root_cause", "")

# Color by canonical ID (used in Excel)
step_colors_by_id = {
    "D1": "ADD8E6",
    "D2": "90EE90",
    "D3": "FFFF99",
    "D4": "FFD580",
    "D5": "FF9999",
    "D6": "D8BFD8",
    "D7": "E0FFFF",
    "D8": "D3D3D3",
    "Interactive 5-Why": "FFE4B5"
}

# ---------------------------
# Report info inputs
# ---------------------------
st.subheader(t["report_info"])
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input(t["report_date"], value=today_str)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs: one tab per NPQP step (preserve original look), plus two extra tabs
# ---------------------------
# Create titles for each NPQP tab from localized titles
localized_titles = [localized_title for (_, localized_title, _, _) in npqp_steps]
# Create the tabs in the same order as before
all_tabs = st.tabs(localized_titles + [t["interactive_5why_tab"] if "interactive_5why_tab" in t else "Interactive 5-Why", t["ai_tab"]])

# Map NPQP step tabs to indices
num_npqp = len(npqp_steps)
npqp_tab_objs = all_tabs[:num_npqp]
interactive_tab_obj = all_tabs[num_npqp]
ai_tab_obj = all_tabs[num_npqp + 1]

# ---------------------------
# Render NPQP tabs exactly like original app (preserve D1-D8 layout)
# ---------------------------
for i, (sid, localized_title, guidance, example) in enumerate(npqp_steps):
    with npqp_tab_objs[i]:
        st.markdown(f"### {localized_title}")

        # D5: show enhanced guidance example block similar to your original
        if sid == "D5":
            # show bilingual guidance block — we include both languages examples inside so user sees training guidance clearly
            if lang == "en":
                full_training_note = (
                    "**Training Guidance:** Use 5-Why analysis to determine the root cause.\n\n"
                    "**Occurrence Example (5-Whys):**\n"
                    "1. Cold solder joint on DSP chip\n"
                    "2. Soldering temperature too low\n"
                    "3. Operator didn’t follow profile\n"
                    "4. Work instructions were unclear\n"
                    "5. No visual confirmation step\n\n"
                    "**Detection Example (5-Whys):**\n"
                    "1. QA inspection missed cold joint\n"
                    "2. Inspection checklist incomplete\n"
                    "3. No automated test step\n"
                    "4. Batch testing not performed\n"
                    "5. Early warning signal not tracked\n\n"
                    "**Root Cause Example:**\n"
                    "Insufficient process control on soldering operation, combined with inadequate QA checklist, "
                    "allowed defective DSP soldering to pass undetected."
                )
            else:
                full_training_note = (
                    "**Guía de Entrenamiento:** Use análisis de 5 porqués para determinar la causa raíz.\n\n"
                    "**Ejemplo de Ocurrencia (5-Whys):**\n"
                    "1. Junta de soldadura fría en el chip DSP\n"
                    "2. Temperatura de soldadura demasiado baja\n"
                    "3. El operador no siguió el perfil\n"
                    "4. Las instrucciones de trabajo no eran claras\n"
                    "5. No hay paso de confirmación visual\n\n"
                    "**Ejemplo de Detección (5-Whys):**\n"
                    "1. Inspección de QA no detectó junta fría\n"
                    "2. Lista de verificación de inspección incompleta\n"
                    "3. No hay paso de prueba automatizado\n"
                    "4. No se realizaron pruebas por lote\n"
                    "5. Señal de advertencia temprana no registrada\n\n"
                    "**Ejemplo de Causa Raíz:**\n"
                    "Control de proceso insuficiente en la operación de soldadura, combinado con una lista de verificación de QA inadecuada, "
                    "permitió que la soldadura defectuosa del DSP pasara sin ser detectada."
                )
            st.info(full_training_note)
        else:
            # Show guidance + example in the selected language
            st.info(f"**Training Guidance / Guía de Entrenamiento:** {guidance}\n\n💡 **Example / Ejemplo:** {example}")

        # Input fields: D5 has special structure (occurrence/detection whys), others simple text_area like original
        if sid == "D5":
            st.markdown("#### " + ("Occurrence Analysis" if lang=="en" else "Análisis de Ocurrencia"))
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(
                    f"{('Occurrence Why' if lang=='en' else 'Porqué de Ocurrencia')} {idx+1}",
                    value=val,
                    key=f"{sid}_occ_{idx}"
                )
            if st.button(("➕ Add another Occurrence Why" if lang=="en" else "➕ Agregar otro Porqué de Ocurrencia"), key=f"add_occ_{sid}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### " + ("Detection Analysis" if lang=="en" else "Análisis de Detección"))
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(
                    f"{('Detection Why' if lang=='en' else 'Porqué de Detección')} {idx+1}",
                    value=val,
                    key=f"{sid}_det_{idx}"
                )
            if st.button(("➕ Add another Detection Why" if lang=="en" else "➕ Agregar otro Porqué de Detección"), key=f"add_det_{sid}"):
                st.session_state.d5_det_whys.append("")

            # Combine Occurrence & Detection into the step answer (keeps same behavior as your original)
            st.session_state[sid]["answer"] = (
                ( "Occurrence Analysis:\n" if lang=="en" else "Análisis de Ocurrencia:\n" ) +
                "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\n" +
                ( "Detection Analysis:\n" if lang=="en" else "Análisis de Detección:\n" ) +
                "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )

            # Root Cause text area (keeps original field)
            st.session_state[sid]["extra"] = st.text_area(
                ("Root Cause (summary after 5-Whys)" if lang=="en" else "Causa Raíz (resumen después de 5-Whys)"),
                value=st.session_state[sid]["extra"],
                key=f"{sid}_root_cause"
            )
        else:
            # Other steps: plain answer text area like original
            st.session_state[sid]["answer"] = st.text_area(
                (f"Your Answer for {localized_title}" if lang=="en" else f"Su Respuesta para {localized_title}"),
                value=st.session_state[sid]["answer"],
                key=f"ans_{sid}"
            )

# ---------------------------
# Interactive 5-Why Tab (extra, isolated)
# ---------------------------
with interactive_tab_obj:
    st.header(t.get("interactive_5why_tab", "Interactive 5-Why"))

    # Simple AI suggestion logic (expandable later)
    def ai_suggestion_for_next_why(prev_answer):
        pa = prev_answer.lower()
        if "solder" in pa or "soldadur" in pa:  # covers 'soldadura'
            return ("Was the soldering process performed correctly?" if lang=="en" else "¿Se realizó correctamente el proceso de soldadura?")
        # default
        return ("Why did this happen?" if lang=="en" else "¿Por qué ocurrió esto?")

    def ai_root_cause_summary(whys_list):
        combined = " | ".join([w for w in whys_list if w.strip()])
        if not combined:
            return ""
        return (f"AI analysis of root cause based on: {combined}" if lang=="en" else f"Análisis AI de causa raíz basado en: {combined}")

    # Render dynamic whys
    for idx, val in enumerate(st.session_state.interactive_whys):
        st.session_state.interactive_whys[idx] = st.text_input(
            (f"Why {idx+1}" if lang=="en" else f"Porqué {idx+1}"),
            value=val,
            key=f"interactive_why_{idx}"
        )
        if st.session_state.interactive_whys[idx].strip():
            suggestion = ai_suggestion_for_next_why(st.session_state.interactive_whys[idx])
            st.markdown(f"*{('Suggested next question:' if lang=='en' else 'Siguiente sugerencia:')}* {suggestion}")

    if st.button(("➕ Add another Why" if lang=="en" else "➕ Agregar otro Porqué"), key="add_dynamic_why"):
        st.session_state.interactive_whys.append("")

    # Root cause suggestion area (AI)
    rc = ai_root_cause_summary(st.session_state.interactive_whys)
    st.session_state.interactive_root_cause = st.text_area(
        ( "AI Suggested Root Cause" if lang=="en" else "Causa Raíz sugerida por AI" ),
        value=rc,
        height=150,
        key="interactive_root_ca_area"
    )

# ---------------------------
# AI Helper Tab (extra)
# ---------------------------
with ai_tab_obj:
    st.header(t["ai_helper"])
    st.info(( "This tab can provide additional AI guidance or corrective action suggestions." if lang=="en" else "Esta pestaña puede proporcionar sugerencias adicionales de AI o acciones correctivas."))
    if st.button(( "Generate AI Suggestions" if lang=="en" else "Generar Sugerencias AI" )):
        whys_text = "\n".join([w for w in st.session_state.interactive_whys if w.strip()])
        st.text_area(( "AI Suggestions" if lang=="en" else "Sugerencias AI" ),
                     ( f"AI would analyze the following 5-Whys:\n{whys_text}\n\nand provide recommendations here." if lang=="en"
                       else f"AI analizaría los siguientes 5-Whys:\n{whys_text}\n\ny proporcionaría recomendaciones aquí." ),
                     height=200)

# ---------------------------
# Save / Export to Excel (preserve NPQP formatting + append interactive 5-Why)
# ---------------------------
if st.button(t["save_report"]):
    # Build rows using canonical IDs but write localized titles to Excel
    data_rows = []
    for sid, localized_title, guidance, example in npqp_steps:
        ans = st.session_state[sid]["answer"]
        extra = st.session_state[sid]["extra"]
        data_rows.append((sid, localized_title, ans, extra))

    # Append interactive 5-Why as final row (localized label)
    interactive_label = ("Interactive 5-Why" if lang=="en" else "5-Why Interactivo")
    interactive_answer = "\n".join([w for w in st.session_state.interactive_whys if w.strip()])
    interactive_extra = st.session_state.interactive_root_cause
    data_rows.append(("INTERACTIVE", interactive_label, interactive_answer, interactive_extra))

    # Check if there's at least one answer filled
    if not any(row[2].strip() or row[3].strip() for row in data_rows):
        st.error(t["no_answers"])
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
        ws["A3"] = "Report Date"
        ws["B3"] = st.session_state.report_date
        ws["A4"] = "Prepared By"
        ws["B4"] = st.session_state.prepared_by

        # Headers
        headers = ["Step", "Your Answer", "Root Cause / Extra"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        row = 6
        for col, header in enumerate(headers, start=1):
            cell = ws.cell(row=row, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        # Content rows
        row = 7
        for sid, localized_title, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=localized_title)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)

            # Choose color by sid if exists, else default
            fill_color = step_colors_by_id.get(sid, step_colors_by_id.get("Interactive 5-Why", "FFFFFF"))
            for col in range(1, 4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
            row += 1

        # Adjust column widths to match original look
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)

        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button(t["download"], f, file_name=xlsx_file)
