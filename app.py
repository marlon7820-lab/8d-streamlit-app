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

# Language selection
if "lang" not in st.session_state:
    st.session_state.lang = "en"
lang = st.radio("🌐 Language / Idioma", ["English", "Español"], index=0 if st.session_state.lang=="en" else 1)
st.session_state.lang = "en" if lang=="English" else "es"

# Bilingual labels and guidance
texts = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "d1": ("D1: Concern Details",
               "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
               "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
        "d2": ("D2: Similar Part Considerations",
               "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
               "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
        "d3": ("D3: Initial Analysis",
               "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
               "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
        "d4": ("D4: Implement Containment",
               "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
               "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
        "d5": ("D5: Root Cause Analysis (5-Why)",
               "Use 5-Why analysis to determine the root cause. Occurrence and Detection separate.",
               ""),
        "d6": ("D6: Permanent Corrective Actions",
               "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
               "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
        "d7": ("D7: Countermeasure Confirmation",
               "Verify that corrective actions effectively resolve the issue long-term.",
               "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
        "d8": ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
               "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
               "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future."),
        "occurrence": "Occurrence Analysis",
        "detection": "Detection Analysis",
        "why": "Why",
        "suggestions": "Suggestions based on previous answer",
    },
    "es": {
        "app_title": "📑 Aplicación de Entrenamiento 8D",
        "report_date": "📅 Fecha del Informe",
        "prepared_by": "✍️ Preparado Por",
        "d1": ("D1: Detalles del Problema",
               "Describa claramente las preocupaciones del cliente. Incluya qué es el problema, dónde ocurrió, cuándo y cualquier dato de apoyo.",
               "Ejemplo: El cliente reportó ruido estático en el amplificador durante la prueba final en la Planta A."),
        "d2": ("D2: Consideraciones de Partes Similares",
               "Verifique partes similares, modelos, partes genéricas, otros colores, mano opuesta, frente/atrás, etc., para ver si el problema es recurrente o aislado.",
               "Ejemplo: Mismo tipo de altavoz usado en otro modelo de radio; diferentes colores de amplificador; unidades de audio delanteras vs traseras."),
        "d3": ("D3: Análisis Inicial",
               "Realice una investigación inicial para identificar problemas obvios, recopilar datos y documentar hallazgos iniciales.",
               "Ejemplo: Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores."),
        "d4": ("D4: Implementar Contención",
               "Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes.",
               "Ejemplo: Inspección 100% de amplificadores antes del envío; uso de protección temporal; cuarentena de lotes afectados."),
        "d5": ("D5: Análisis de Causa Raíz (5-Why)",
               "Use análisis 5-Why para determinar la causa raíz. Ocurrencia y Detección separadas.",
               ""),
        "d6": ("D6: Acciones Correctivas Permanentes",
               "Defina acciones correctivas que eliminen permanentemente la causa raíz y eviten la recurrencia.",
               "Ejemplo: Actualizar proceso de soldadura, reentrenar operadores, actualizar instrucciones de trabajo y agregar inspección automatizada."),
        "d7": ("D7: Confirmación de Contramedidas",
               "Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo.",
               "Ejemplo: Pruebas funcionales en amplificadores corregidos, pruebas aceleradas de vida y monitoreo de las primeras unidades de producción."),
        "d8": ("D8: Seguimiento (Lecciones Aprendidas / Prevención de Recurrencia)",
               "Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y entrenamientos para prevenir recurrencias.",
               "Ejemplo: Actualizar SOPs, PFMEA, instrucciones de trabajo y capacitación de empleados para prevenir el mismo problema en el futuro."),
        "occurrence": "Análisis de Ocurrencia",
        "detection": "Análisis de Detección",
        "why": "Por qué",
        "suggestions": "Sugerencias basadas en la respuesta anterior",
    }
}

t = texts[st.session_state.lang]

# App title
st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['app_title']}</h1>", unsafe_allow_html=True)

# ---------------------------
# Report info
# ---------------------------
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.setdefault("report_date", today_str)
st.session_state.setdefault("prepared_by", "")

st.session_state.report_date = st.text_input(t["report_date"], value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], value=st.session_state.prepared_by)

# ---------------------------
# 8D Tabs setup
# ---------------------------
npqp_steps = ["d1","d2","d3","d4","d5","d6","d7","d8"]
step_colors = {
    "d1":"ADD8E6","d2":"90EE90","d3":"FFFF99","d4":"FFD580",
    "d5":"FF9999","d6":"D8BFD8","d7":"E0FFFF","d8":"D3D3D3"
}

for step in npqp_steps:
    st.session_state.setdefault(step, {"answer": "", "extra": ""})

tabs = st.tabs([texts[st.session_state.lang][s][0] for s in npqp_steps])

for i, step in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {texts[st.session_state.lang][step][0]}")
        st.info(f"**Guidance:** {texts[st.session_state.lang][step][1]}\n\n💡 **Example:** {texts[st.session_state.lang][step][2]}")

        # D5: Interactive 5-Why
        if step=="d5":
            st.session_state.setdefault("d5_occ", [""]*5)
            st.session_state.setdefault("d5_det", [""]*5)

            def get_suggestions(prev_answer):
                if not prev_answer:
                    return []
                keywords = prev_answer.lower().split()
                suggestions=[]
                if "operator" in keywords or "operador" in keywords:
                    suggestions = ["Operator skipped step","Operator misread instructions","Operator not trained properly"]
                elif "process" in keywords or "proceso" in keywords:
                    suggestions = ["Process not standardized","Process not monitored","Equipment settings incorrect"]
                elif "inspection" in keywords or "inspección" in keywords:
                    suggestions = ["Inspection step missing","Checklist incomplete","Test not performed"]
                return suggestions[:3]

            st.markdown(f"#### {t['occurrence']}")
            for idx in range(5):
                prev = st.session_state.d5_occ[idx-1] if idx>0 else ""
                suggestions = get_suggestions(prev)
                if suggestions:
                    st.markdown(f"💡 {t['suggestions']}: {', '.join(suggestions)}")
                st.session_state.d5_occ[idx] = st.text_input(f"{t['why']} {idx+1}", value=st.session_state.d5_occ[idx], key=f"occ_{idx}")

            st.markdown(f"#### {t['
