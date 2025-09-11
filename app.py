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
st.session_state["prev_lang"] = lang  # track previous language

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
        "interactive_5why_tab": "5-Why Interactivo",
        "ai_tab": "Asistente AI"
    }
}
t = ui_texts[lang]

st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['header']}</h1>", unsafe_allow_html=True)

# ---------------------------
# Step definitions
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
# Initialize session_state
# ---------------------------
for sid, _, _, _ in npqp_steps:
    if sid not in st.session_state:
        st.session_state[sid] = {"answer": "", "extra": ""}

st.session_state.setdefault("report_date", "")
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""]*5)
st.session_state.setdefault("d5_det_whys", [""]*5)
st.session_state.setdefault("interactive_whys", [""])
st.session_state.setdefault("interactive_root_cause", "")

# ---------------------------
# Translation function
# ---------------------------
def translate_text(text, target_lang="es"):
    if not text.strip():
        return text
    key = st.secrets.get("OPENAI_API_KEY", "")
    if not key:
        st.warning("OpenAI API key missing. Cannot translate.")
        return text
    openai.api_key = key
    prompt = f"Translate the following text to {target_lang} keeping technical terms intact:\n{text}"
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role":"user","content":prompt}]
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Translation failed: {e}")
        return text

# ---------------------------
# Detect language change and translate all stored answers
# ---------------------------
