import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import openai

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Training App",
    page_icon="https://raw.githubusercontent.com/marlon7820-lab/8d-streamlit-app/refs/heads/main/IMG_7771%20Small.png",
    layout="wide"
)

# Hide default menu/footer
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
lang_choice = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
lang = "en" if lang_choice == "English" else "es"
prev_lang = st.session_state.get("prev_lang", lang)
st.session_state["prev_lang"] = lang

# ---------------------------
# UI Texts
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
        "add_occ": "➕ Add another Occurrence Why",
        "add_det": "➕ Add another Detection Why"
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
        "add_occ": "➕ Agregar otro Porqué de Ocurrencia",
        "add_det": "➕ Agregar otro Porqué de Detección"
    }
}
t = ui_texts[lang]

st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['header']}</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP Steps
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
               ""),
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

npqp_steps = [(sid,) + npqp_steps_texts[lang][sid] for sid in step_ids]

# ---------------------------
# Session state init
# ---------------------------
for sid, _, _, _ in npqp_steps:
    if sid not in st.session_state:
        st.session_state[sid] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("interactive_whys", [""])
st.session_state.setdefault("interactive_root_cause", "")
st.session_state.setdefault("translations", {})

# ---------------------------
# Colors for Excel
# ---------------------------
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
# Translation helper with caching
# ---------------------------
def translate_text_cached(text, target_lang="es", field_id=None):
    if not text.strip():
        return text
    cache_key = f"{field_id}_{target_lang}" if field_id else f"default_{target_lang}"
    if cache_key in st.session_state["translations"]:
        return st.session_state["translations"][cache_key]
    key = st.secrets.get("OPENAI_API_KEY", "")
    if not key:
        return text
    openai.api_key = key
    prompt = f"Translate the following text to {target_lang}, keeping technical terms intact:\n{text}"
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role":"user","content":prompt}]
        )
        translation = response.choices[0].message.content
        st.session_state["translations"][cache_key] = translation
        return translation
    except:
        return text

# ---------------------------
# Auto-translate fields if language changed
# ---------------------------
def auto_translate_all(prev_lang, current_lang):
    if prev_lang == current_lang:
        return
    target_lang = "es" if current_lang == "es" else "en"

    # Translate D1–D8 answers and extra fields
    for sid, _, _, _ in npqp_steps:
        st.session_state[sid]["answer"] = translate_text_cached(st.session_state[sid]["answer"], target_lang, field_id=f"{sid}_answer")
        st.session_state[sid]["extra"] = translate_text_cached(st.session_state[sid]["extra"], target_lang, field_id=f"{sid}_extra")

    # Translate D5 Occurrence & Detection Whys
    st.session_state.d5_occ_whys = [
        translate_text_cached(w, target_lang, field_id=f"d5_occ_{i}") for i, w in enumerate(st.session_state.d5_occ_whys)
    ]
    st.session_state.d5_det_whys = [
        translate_text_cached(w, target_lang, field_id=f"d5_det_{i}") for i, w in enumerate(st.session_state.d5_det_whys)
    ]

    # Translate Interactive 5-Why
    st.session_state.interactive_whys = [
        translate_text_cached(w, target_lang, field_id=f"interactive_{i}") for i, w in enumerate(st.session_state.interactive_whys)
    ]
    st.session_state.interactive_root_cause = translate_text_cached(
        st.session_state.interactive_root_cause, target_lang, field_id="interactive_root_cause"
    )

    # Refresh Streamlit input fields
    for sid, _, _, _ in npqp_steps:
        st.session_state[f"ans_{sid}"] = st.session_state[sid]["answer"]
        st.session_state[f"{sid}_root_text"] = st.session_state[sid]["extra"]
    for i in range(len(st.session_state.d5_occ_whys)):
        st.session_state[f"D5_occ_{i}"] = st.session_state.d5_occ_whys[i]
    for i in range(len(st.session_state.d5_det_whys)):
        st.session_state[f"D5_det_{i}"] = st.session_state.d5_det_whys[i]
    for i in range(len(st.session_state.interactive_whys)):
        st.session_state[f"interactive_{i}"] = st.session_state.interactive_whys[i]
    st.session_state["interactive_root_text"] = st.session_state.interactive_root_cause

auto_translate_all(prev_lang, lang)
st.session_state["prev_lang"] = lang

# ---------------------------
# Report info
# ---------------------------
st.subheader(t["report_info"])
st.session_state.report_date = st.text_input(t["report_date"], st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs D1-D8
# ---------------------------
tabs = st.tabs([step for step,_,_,_ in npqp_steps])
for i, (sid, step_title, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step_title}")
        if sid == "D5":
            st.info(
                "**Training Guidance:** Use 5-Why analysis to determine the root cause.\n\n"
                "**Occurrence Example (5-Whys):**\n"
                "1. Cold solder joint on DSP chip\n2. Soldering temperature too low\n3. Operator didn’t follow profile\n4. Work instructions were unclear\n5. No visual confirmation step\n\n"
                "**Detection Example (5-Whys):**\n"
                "1. QA inspection missed cold joint\n2. Inspection checklist incomplete\n3. No automated test step\n4. Batch testing not performed\n5. Early warning signal not tracked\n\n"
                "**Root Cause Example:**\n"
                "Insufficient process control on soldering operation, combined with inadequate QA checklist, allowed defective DSP soldering to pass undetected."
            )
        else:
            st.info(f"**Training Guidance:** {note}\n\n💡 **Example:** {example}")

        # Inputs
        if sid == "D5":
            st.markdown("#### Occurrence Analysis")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", val, key=f"{sid}_occ_{idx}")
            if st.button(t["add_occ"], key=f"add_occ_{sid}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### Detection Analysis")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", val, key=f"{sid}_det_{idx}")
            if st.button(t["add_det"], key=f"add_det_{sid}"):
                st.session_state.d5_det_whys.append("")

            st.session_state[sid]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
            st.session_state
