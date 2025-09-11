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
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ---------------------------
# Language selection
# ---------------------------
lang = st.selectbox("🌐 Language / Idioma", ["English", "Español"])

# Safe language dictionary
lang_map = {"English": "en", "Español": "es"}
selected_lang = lang_map.get(lang, "en")

# ---------------------------
# Text dictionary
# ---------------------------
t = {
    "en": {
        "app_title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_button": "💾 Save 8D Report",
        "download": "📥 Download XLSX",
        "tabs": ["D1: Concern Details", "D2: Similar Part Considerations", "D3: Initial Analysis",
                 "D4: Implement Containment", "D5: Final Analysis", "D6: Permanent Corrective Actions",
                 "D7: Countermeasure Confirmation", "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)"],
        "training_notes": [
            "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
            "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
            "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
            "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
            "Use 5-Why analysis to determine the root cause. Separate by Occurrence and Detection.",
            "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
            "Verify that corrective actions effectively resolve the issue long-term.",
            "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence."
        ],
        "examples": [
            "Example: Customer reported static noise in amplifier during end-of-line test at Plant A.",
            "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units.",
            "Example: Visual inspection of solder joints, initial functional tests, checking connectors.",
            "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches.",
            "Occurrence Example (5-Whys): 1. Cold solder joint on DSP chip\n2. Soldering temperature too low\n3. Operator didn’t follow profile\n4. Work instructions were unclear\n5. No visual confirmation step\n\nDetection Example (5-Whys): 1. QA inspection missed cold joint\n2. Inspection checklist incomplete\n3. No automated test step\n4. Batch testing not performed\n5. Early warning signal not tracked",
            "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection.",
            "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs.",
            "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future."
        ]
    },
    "es": {
        "app_title": "📑 App de Entrenamiento 8D",
        "report_date": "📅 Fecha",
        "prepared_by": "✍️ Preparado Por",
        "save_button": "💾 Guardar Reporte 8D",
        "download": "📥 Descargar XLSX",
        "tabs": ["D1: Detalles de la Preocupación", "D2: Consideraciones de Piezas Similares",
                 "D3: Análisis Inicial", "D4: Implementar Contención", "D5: Análisis Final",
                 "D6: Acciones Correctivas Permanentes", "D7: Confirmación de Contramedidas",
                 "D8: Seguimiento / Lecciones Aprendidas"],
        "training_notes": [
            "Describe claramente las preocupaciones del cliente. Incluye qué problema es, dónde ocurrió, cuándo y cualquier dato de apoyo.",
            "Verifica piezas similares, modelos, colores diferentes, mano opuesta, frente/trasero, etc. para ver si el problema se repite o es aislado.",
            "Realiza una investigación inicial para identificar problemas obvios, recopilar datos y documentar hallazgos iniciales.",
            "Define acciones de contención temporales para prevenir que el cliente vea el problema mientras se desarrollan acciones permanentes.",
            "Usa el análisis de 5-Why para determinar la causa raíz. Sepáralo en Ocurrencia y Detección.",
            "Define acciones correctivas que eliminen la causa raíz permanentemente y prevengan la recurrencia.",
            "Verifica que las acciones correctivas resuelvan efectivamente el problema a largo plazo.",
            "Documenta lecciones aprendidas, actualiza estándares, procedimientos, FMEAs y capacitación para prevenir recurrencia."
        ],
        "examples": [
            "Ejemplo: Cliente reportó ruido estático en el amplificador durante prueba final en Planta A.",
            "Ejemplo: Mismo tipo de altavoz usado en otro modelo de radio; colores de amplificador diferentes; unidades delanteras vs traseras.",
            "Ejemplo: Inspección visual de soldaduras, pruebas funcionales iniciales, revisión de conectores.",
            "Ejemplo: Inspección 100% de amplificadores antes del envío; uso de protección temporal; cuarentena de lotes afectados.",
            "Ejemplo Ocurrencia (5-Whys): 1. Soldadura fría en chip DSP\n2. Temperatura de soldadura demasiado baja\n3. Operador no siguió perfil\n4. Instrucciones poco claras\n5. Sin paso de verificación visual\n\nEjemplo Detección (5-Whys): 1. QA no detectó soldadura fría\n2. Checklist incompleto\n3. Sin prueba automatizada\n4. No se realizó prueba de lote\n5. Señal de alerta temprana no registrada",
            "Ejemplo: Actualizar proceso de soldadura, reentrenar operadores, actualizar instrucciones de trabajo y agregar inspección automatizada.",
            "Ejemplo: Pruebas funcionales en amplificadores corregidos, pruebas aceleradas de vida útil, monitoreo de primeras producciones.",
            "Ejemplo: Actualizar SOPs, PFMEA, instrucciones de trabajo y capacitación de empleados para prevenir el mismo problema."
        ]
    }
}[selected_lang]

# ---------------------------
# Add custom app header
# ---------------------------
st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['app_title']}</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = t["tabs"]

# ---------------------------
# Initialize session state
# ---------------------------
for i, step in enumerate(npqp_steps):
    if step not in st.session_state:
        st.session_state[step] = {"answer": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", [""] * 5)
st.session_state.setdefault("d5_det_whys", [""] * 5)

# ---------------------------
# Color dictionary for Excel
# ---------------------------
step_colors = {
    npqp_steps[0]: "ADD8E6",
    npqp_steps[1]: "90EE90",
    npqp_steps[2]: "FFFF99",
    npqp_steps[3]: "FFD580",
    npqp_steps[4]: "FF9999",
    npqp_steps[5]: "D8BFD8",
    npqp_steps[6]: "E0FFFF",
    npqp_steps[7]: "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.subheader(t["report_date"])
st.session_state.report_date = st.text_input(t["report_date"], st.session_state.report_date)
st.subheader(t["prepared_by"])
st.session_state.prepared_by = st.text_input(t["prepared_by"], st.session_state.prepared_by)

# ---------------------------
# Tabs
# ---------------------------
tabs = st.tabs(npqp_steps)
for i, step in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")
        st.info(f"**Training Guidance:** {t['training_notes'][i]}\n\n💡 **Example:** {t['examples'][i]}")

        # D5 interactive 5-Why
        if i == 4:
            st.markdown("#### Occurrence Analysis")
            for idx in range(5):
                if idx == 0:
                    st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=st.session_state.d5_occ_whys[idx], key=f"{step}_occ_{idx}")
                else:
                    prev = st.session_state.d5_occ_whys[idx - 1]
                    suggestions = [f"Follow-up on '{prev}' #{n}" for n in range(1,4)]
                    st.session_state.d5_occ_whys[idx] = st.selectbox(f"Occurrence
