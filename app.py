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

# Hide default Streamlit menu, header, footer
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Language toggle
# ---------------------------
lang_choice = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
lang = "en" if lang_choice == "English" else "es"

# ---------------------------
# Language dictionaries
# ---------------------------
texts = {
    "en": {
        "header": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_report": "💾 Save 8D Report",
        "download": "📥 Download XLSX",
        "no_answers": "⚠️ No answers filled in yet. Please complete some fields before saving.",
        "ai_helper": "💡 AI Helper Suggestions"
    },
    "es": {
        "header": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado Por",
        "save_report": "💾 Guardar Reporte 8D",
        "download": "📥 Descargar XLSX",
        "no_answers": "⚠️ No se han completado respuestas. Por favor complete algunos campos antes de guardar.",
        "ai_helper": "💡 Sugerencias del Asistente AI"
    }
}
t = texts[lang]

# ---------------------------
# D1–D8 NPQP Steps per language
# ---------------------------
npqp_steps_texts = {
    "en": [
        ("D1: Concern Details", "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.", "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
        ("D2: Similar Part Considerations", "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.", "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
        ("D3: Initial Analysis", "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.", "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
        ("D4: Implement Containment", "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.", "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
        ("D5: Final Analysis", "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected). Add more Whys if needed.", ""),
        ("D6: Permanent Corrective Actions", "Define corrective actions that eliminate the root cause permanently and prevent recurrence.", "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
        ("D7: Countermeasure Confirmation", "Verify that corrective actions effectively resolve the issue long-term.", "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
        ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)", "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.", "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future.")
    ],
    "es": [
        ("D1: Detalles de la Preocupación", "Describa claramente las preocupaciones del cliente. Incluya cuál es el problema, dónde ocurrió, cuándo y cualquier dato de soporte.", "Ejemplo: El cliente reportó ruido estático en el amplificador durante la prueba final en Planta A."),
        ("D2: Consideraciones de Piezas Similares", "Revise piezas, modelos, colores, manos opuestas, frontales/traseras, etc. para ver si el problema es recurrente o aislado.", "Ejemplo: Mismo tipo de altavoz usado en otro modelo de radio; diferentes colores de amplificador; unidades de audio frontales vs traseras."),
        ("D3: Análisis Inicial", "Realice una investigación inicial para identificar problemas evidentes, recopilar datos y documentar hallazgos.", "Ejemplo: Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores."),
        ("D4: Implementar Contención", "Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes.", "Ejemplo: Inspección 100% de amplificadores antes del envío; uso de blindaje temporal; cuarentena de lotes afectados."),
        ("D5: Análisis Final", "Use análisis de 5 porqués para determinar la causa raíz. Separe por Ocurrencia (por qué sucedió) y Detección (por qué no se detectó). Agregue más porqués si es necesario.", ""),
        ("D6: Acciones Correctivas Permanentes", "Defina acciones correctivas que eliminen la causa raíz permanentemente y prevengan recurrencias.", "Ejemplo: Actualizar proceso de soldadura, reentrenar operadores, actualizar instrucciones de trabajo y agregar inspección automatizada."),
        ("D7: Confirmación de Contramedidas", "Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo.", "Ejemplo: Pruebas funcionales en amplificadores corregidos, pruebas aceleradas de vida útil, y monitoreo de primeras unidades de producción."),
        ("D8: Actividades de Seguimiento (Lecciones Aprendidas / Prevención de Recurrencias)", "Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y entrenamiento para prevenir recurrencia.", "Ejemplo: Actualizar SOPs, PFMEA, instrucciones de trabajo y capacitación de empleados para prevenir el mismo problema en el futuro.")
    ]
}

npqp_steps = npqp_steps_texts[lang]

# ---------------------------
# Initialize session state
# ---------------------------
for step, _, _ in npqp_steps:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)
st.session_state.setdefault("interactive_whys", [""])
st.session_state.setdefault("interactive_root_cause", "")

# Color dictionary for Excel
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)": "D3D3D3",
    # Spanish mapping
    "D1: Detalles de la Preocupación": "ADD8E6",
    "D2: Consideraciones de Piezas Similares": "90EE90",
    "D3: Análisis Inicial": "FFFF99",
    "D4: Implementar Contención": "FFD580",
    "D5: Análisis Final": "FF9999",
    "D6: Acciones Correctivas Permanentes": "D8BFD8",
    "D7: Confirmación de Contramedidas": "E0FFFF",
    "D8: Actividades de Seguimiento (Lecciones Aprendidas / Prevención de Recurrencias)": "D3D3D3",
    "Interactive 5-Why": "FFE4B5"
}

# ---------------------------
# Report Information
# ---------------------------
st.subheader("Report Information")
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input(t["report_date"], value=today_str)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# AI Helper functions
# ---------------------------
def ai_suggestion_for_next_why(prev_answer):
    if "solder" in prev_answer.lower() or "soldadura" in prev_answer.lower():
        return "Was the soldering process performed correctly?" if lang=="en" else "¿Se realizó correctamente el proceso de soldadura?"
    return "Why did this happen?" if lang=="en" else "¿Por qué ocurrió esto?"

def ai_root_cause_summary(whys_list):
    combined = " | ".join([w for w in whys_list if w.strip()])
    if combined.strip() == "":
        return ""
    return f"AI analysis of root cause based on: {combined}" if lang=="en" else f"Análisis AI de causa raíz basado en: {combined}"

# ---------------------------
# Tabs: Current Features, Interactive 5-Why, AI Helper
# ---------------------------
tab_current, tab_5why, tab_ai = st.tabs([
    "Current Features",
    "Interactive 5-Why",
    t["ai_helper"]
])

# ---------------------------
# Current Features Tab: D1–D8 inputs
# ---------------------------
with tab_current:
    for i, (step, note, example) in enumerate(npqp_steps):
        st.markdown(f"### {step}")
        if step.startswith("D5"):
            full_training_note = (
                "**Training Guidance / Guía de Entrenamiento:** Use 5-Why analysis to determine the root cause.\n\n"
                "**Occurrence Example / Ejemplo de Ocurrencia (5-Whys):**\n"
                "1. Cold solder joint on DSP chip\n"
                "2. Soldering temperature too low\n"
                "3. Operator didn’t follow profile\n"
                "4. Work instructions were unclear\n"
                "5. No visual confirmation step\n\n"
                "**Detection Example / Ejemplo de Detección (5-Whys):**\n"
                "1. QA inspection missed cold joint\n"
                "2. Inspection checklist incomplete\n"
                "3. No automated test step\n"
                "4. Batch testing not performed\n"
                "5. Early warning signal not tracked\n\n"
                "**Root Cause Example / Ejemplo de Causa Raíz:**\n"
                "Insufficient process control on soldering operation, combined with inadequate QA checklist, "
                "allowed defective DSP soldering to pass undetected."
            )
            st.info(full_training_note)
        else:
            st.info(f"**Training Guidance / Guía de Entrenamiento:** {note}\n\n💡 **Example / Ejemplo:** {example}")

        # Input fields
        if step.startswith("D5"):
            st.markdown("#### Occurrence Analysis / Análisis de Ocurrencia")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=val, key=f"{step}_occ_{idx}")
            if st.button("➕ Add another Occurrence Why / Agregar otro Porqué de Ocurrencia", key=f"add_occ_{step}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### Detection Analysis / Análisis de Detección")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", value=val, key=f"{step}_det_{idx}")
            if st.button("➕ Add another Detection Why / Agregar otro Porqué de Detección", key=f"add_det_{step}"):
                st.session_state.d5_det_whys.append("")

            # Combine Occurrence & Detection into answer
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )

            # Root Cause field (fixed syntax)
            st.session_state[step]["extra"] = st.text_area(
                "Root Cause (summary after 5-Whys) / Causa Raíz",
                value=st.session_state[step]["extra"],
                key=f"{step}_root_cause"
            )
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step} / Su Respuesta para {step}",
                                                             value=st.session_state[step]["answer"],
                                                             key=f"ans_{step}")

# ---------------------------
# Interactive 5-Why Tab
# ---------------------------
with tab_5why:
    st.header("Interactive 5-Why Analysis (AI-Powered)")

    for idx, val in enumerate(st.session_state.interactive_whys):
        placeholder = st.empty()
        st.session_state.interactive_whys[idx] = placeholder.text_input(
            f"Why {idx+1}", value=val, key=f"interactive_why_{idx}"
        )
        if val.strip() != "":
            suggestion = ai_suggestion_for_next_why(val)
            st.markdown(f"*Suggested next question / Siguiente sugerencia: {suggestion}*")

    if st.button("➕ Add another Why / Agregar otro Porqué", key="add_dynamic_why"):
        st.session_state.interactive_whys.append("")

    root_cause = ai_root_cause_summary(st.session_state.interactive_whys)
    st.session_state.interactive_root_cause = root_cause
    st.text_area("AI Suggested Root Cause / Causa Raíz sugerida por AI", value=root_cause, height=150)

# ---------------------------
# AI Helper Tab
# ---------------------------
with tab_ai:
    st.header(t["ai_helper"])
    st.info("This tab can provide additional AI guidance or corrective action suggestions / Este tab puede proporcionar sugerencias adicionales de AI.")
    if st.button("Generate AI Suggestions / Generar Sugerencias AI"):
        whys_text = "\n".join([w for w in st.session_state.interactive_whys if w.strip()])
        st.text_area("AI Suggestions / Sugerencias AI", f"AI would analyze the following 5-Whys:\n{whys_text}\n\nand provide recommendations here." if lang=="en" else f"AI analizaría los siguientes 5-Whys:\n{whys_text}\n\ny proporcionaría recomendaciones aquí.")

# ---------------------------
# Save Button (Excel Export)
# ---------------------------
if st.button(t["save_report"]):
    data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]
    data_rows.append(("Interactive 5-Why", "\n".join([w for w in st.session_state.interactive_whys if w.strip()]),
