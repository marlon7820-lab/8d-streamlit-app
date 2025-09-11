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
# Language selection
# ---------------------------
lang = st.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])

lang_map = {"English": "en", "Español": "es"}

t = {
    "en": {
        "title": "📑 8D Training App",
        "report_date": "📅 Report Date",
        "prepared_by": "✍️ Prepared By",
        "save_button": "💾 Save 8D Report",
        "download_button": "📥 Download XLSX",
        "steps": [
            "D1: Concern Details",
            "D2: Similar Part Considerations",
            "D3: Initial Analysis",
            "D4: Implement Containment",
            "D5: Final Analysis",
            "D6: Permanent Corrective Actions",
            "D7: Countermeasure Confirmation",
            "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)"
        ],
        "guidance": [
            "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
            "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
            "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
            "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
            "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected). Add more Whys if needed.",
            "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
            "Verify that corrective actions effectively resolve the issue long-term.",
            "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence."
        ],
        "examples": [
            "Example: Customer reported static noise in amplifier during end-of-line test at Plant A.",
            "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units.",
            "Example: Visual inspection of solder joints, initial functional tests, checking connectors.",
            "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches.",
            "Example: Occurrence: Cold solder joint → Solder temp too low → Operator didn’t follow profile → Unclear instructions → No visual check. Detection: QA missed joint → Checklist incomplete → No automated test → Batch not tested → Early warning not tracked.",
            "Example: Update soldering process, retrain operators, update work instructions, add automated inspection.",
            "Example: Functional tests on corrected amplifiers, accelerated life testing, monitoring first production runs.",
            "Example: Update SOPs, PFMEA, work instructions, employee training to prevent recurrence."
        ]
    },
    "es": {
        "title": "📑 App de Entrenamiento 8D",
        "report_date": "📅 Fecha del Reporte",
        "prepared_by": "✍️ Preparado Por",
        "save_button": "💾 Guardar Reporte 8D",
        "download_button": "📥 Descargar XLSX",
        "steps": [
            "D1: Detalles de la Preocupación",
            "D2: Consideraciones de Piezas Similares",
            "D3: Análisis Inicial",
            "D4: Implementar Contención",
            "D5: Análisis Final",
            "D6: Acciones Correctivas Permanentes",
            "D7: Confirmación de Contramedidas",
            "D8: Actividades de Seguimiento (Lecciones Aprendidas / Prevención de Recurrencia)"
        ],
        "guidance": [
            "Describa claramente las preocupaciones del cliente. Incluya qué es el problema, dónde ocurrió, cuándo y cualquier dato de soporte.",
            "Verifique piezas similares, modelos, piezas genéricas, otros colores, mano opuesta, frente/atrás, etc. para ver si el problema es recurrente o aislado.",
            "Realice una investigación inicial para identificar problemas obvios, recopile datos y documente hallazgos iniciales.",
            "Defina acciones de contención temporales para evitar que el cliente vea el problema mientras se desarrollan acciones permanentes.",
            "Use el análisis de 5 porqués para determinar la causa raíz. Separe por Ocurrencia (por qué ocurrió) y Detección (por qué no se detectó). Agregue más Porqués si es necesario.",
            "Defina acciones correctivas que eliminen permanentemente la causa raíz y prevengan recurrencia.",
            "Verifique que las acciones correctivas resuelvan efectivamente el problema a largo plazo.",
            "Documente lecciones aprendidas, actualice estándares, procedimientos, FMEAs y capacitación para prevenir recurrencia."
        ],
        "examples": [
            "Ejemplo: El cliente reportó ruido estático en el amplificador durante la prueba final en la Planta A.",
            "Ejemplo: Mismo tipo de altavoz usado en otro modelo de radio; diferentes colores de amplificador; unidades de audio delanteras vs traseras.",
            "Ejemplo: Inspección visual de soldaduras, pruebas funcionales iniciales, verificación de conectores.",
            "Ejemplo: Inspección 100% de amplificadores antes del envío; uso de blindaje temporal; cuarentena de lotes afectados.",
            "Ejemplo: Ocurrencia: Soldadura fría → Temperatura baja → Operador no siguió perfil → Instrucciones poco claras → Sin verificación visual. Detección: QA pasó por alto la unión → Lista de verificación incompleta → Sin prueba automatizada → Lote no probado → Señal temprana no rastreada.",
            "Ejemplo: Actualizar proceso de soldadura, capacitar operadores, actualizar instrucciones de trabajo, agregar inspección automatizada.",
            "Ejemplo: Pruebas funcionales en amplificadores corregidos, pruebas de vida aceleradas, monitoreo de primeras corridas de producción.",
            "Ejemplo: Actualizar SOPs, PFMEA, instrucciones de trabajo, capacitación de empleados para prevenir recurrencia."
        ]
    }
}[lang_map[lang]]

# ---------------------------
# Session state
# ---------------------------
if "report_date" not in st.session_state:
    st.session_state.report_date = datetime.datetime.today().strftime("%B %d, %Y")
if "prepared_by" not in st.session_state:
    st.session_state.prepared_by = ""
if "answers" not in st.session_state:
    st.session_state.answers = {step: "" for step in t["steps"]}
if "d5_occ" not in st.session_state:
    st.session_state.d5_occ = [""] * 5
if "d5_det" not in st.session_state:
    st.session_state.d5_det = [""] * 5

# ---------------------------
# Header
# ---------------------------
st.markdown(f"<h1 style='text-align: center; color: #1E90FF;'>{t['title']}</h1>", unsafe_allow_html=True)

# ---------------------------
# Report info
# ---------------------------
st.subheader(t["report_date"])
st.session_state.report_date = st.text_input("", value=st.session_state.report_date)
st.subheader(t["prepared_by"])
st.session_state.prepared_by = st.text_input("", value=st.session_state.prepared_by)

# ---------------------------
# Tabs
# ---------------------------
tabs = st.tabs(t["steps"])
for i, step in enumerate(t["steps"]):
    with tabs[i]:
        st.markdown(f"### {step}")
        st.info(f"**Guidance:** {t['guidance'][i]}\n\n💡 **Example:** {t['examples'][i]}")

        # D5 interactive 5-Why
        if "D5" in step:
            st.markdown("#### Occurrence Analysis")
            for idx in range(len(st.session_state.d5_occ)):
                st.session_state.d5_occ[idx] = st.text_input(f"Occurrence Why {idx+1}", value=st.session_state.d5_occ[idx], key=f"d5_occ_{idx}")

            st.markdown("#### Detection Analysis")
            for idx in range(len(st.session_state.d5_det)):
                st.session_state.d5_det[idx] = st.text_input(f"Detection Why {idx+1}", value=st.session_state.d5_det[idx], key=f"d5_det_{idx}")

            # Combine for storage
            st.session_state.answers[step] = (
                "Occurrence:\n" + "\n".join([w for w in st.session_state.d5_occ if w.strip()]) +
                "\n\nDetection:\n" + "\n".join([w for w in st.session_state.d5_det if w.strip()])
            )
        else:
            st.session_state.answers[step] = st.text_area(f"Your Answer for {step}", value=st.session_state.answers[step], key=f"ans_{step}")

# ---------------------------
# Save to Excel
# ---------------------------
if st.button(t["save_button"]):
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    # Title
    ws.merge_cells("A1:C1")
    ws["A1"] = t["title"]
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
    for step in t["steps"]:
        ans = st.session_state.answers.get(step, "")
        ws.cell(row=row, column=1, value=step)
        ws.cell(row=row, column=2, value=ans)
        ws.cell(row=row, column=3, value="" if step != t["steps"][0] else ans)  # Only D1 extra

        for col in range(1, 4):
            ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")
        row += 1

    # Column widths
    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    xlsx_file = "NPQP_8D_Report.xlsx"
    wb.save(xlsx_file)
    st.success("✅ NPQP 8D Report saved successfully.")
    with open(xlsx_file, "rb") as f:
        st.download_button(t["download_button"], f, file_name=xlsx_file)
