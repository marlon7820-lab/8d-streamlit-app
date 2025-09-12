import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from datetime import datetime

# ---------------------------
# Initialize session state
# ---------------------------
if "D1" not in st.session_state:
    for i in range(1, 9):
        st.session_state[f"D{i}"] = {"answer": "", "extra": ""}

# D5 interactive Whys storage
if "d5_occ_whys" not in st.session_state:
    st.session_state.d5_occ_whys = [""] * 5
if "d5_det_whys" not in st.session_state:
    st.session_state.d5_det_whys = [""] * 5

# Language selector
lang = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])
lang_key = "en" if lang.startswith("E") else "es"

# Titles, guidance, and labels
t = {
    "en": {
        "D1": "D1: Team",
        "D2": "D2: Problem Description",
        "D3": "D3: Containment Actions",
        "D4": "D4: Root Cause Analysis",
        "D5": "D5: 5-Why Analysis",
        "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Preventive Actions",
        "D8": "D8: Closure",
        "Root_Cause": "Root Cause (summary after 5-Whys)",
        "Occurrence_Why": "Why (Occurrence)",
        "Detection_Why": "Why (Detection)",
        "Training_Guidance": "Training Guidance"
    },
    "es": {
        "D1": "D1: Equipo",
        "D2": "D2: Descripción del Problema",
        "D3": "D3: Acciones de Contención",
        "D4": "D4: Análisis de Causa Raíz",
        "D5": "D5: Análisis 5-Why",
        "D6": "D6: Acciones Correctivas Permanentes",
        "D7": "D7: Acciones Preventivas",
        "D8": "D8: Cierre",
        "Root_Cause": "Causa Raíz (resumen después del 5-Whys)",
        "Occurrence_Why": "Por qué (Ocurrencia)",
        "Detection_Why": "Por qué (Detección)",
        "Training_Guidance": "Guía de Entrenamiento"
    }
}

# Training guidance samples for each D (simplified)
npqp_steps = {
    4: ["Use 5-Why methodology to identify root causes for occurrence and detection."]
}

# ---------------------------
# Tabs for D1–D8
# ---------------------------
tabs = st.tabs([t[lang_key][f"D{i}"] for i in range(1, 9)])

# Free-text input for all Ds except D5 handled separately
for i in range(8):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][f'D{i+1}']}")
        st.text_area(f"Enter your response / Ingrese su respuesta", 
                     value=st.session_state[f"D{i+1}"]['answer'], 
                     key=f"D{i+1}_answer")
        # Extra field only for D1
        if i == 0:
            st.session_state["D1"]['extra'] = st.text_area("Additional notes / Notas adicionales", 
                                                            value=st.session_state["D1"]['extra'], key="D1_extra")

# ---------------------------
# D5: Enhanced Interactive 5-Why
# ---------------------------
with tabs[4]:
    st.markdown(f"### {t[lang_key]['D5']}")
    st.info(f"**{t[lang_key]['Training_Guidance']}:** {npqp_steps[4][0]}")

    # Base hints for suggestions
    occurrence_hints = [
        "Equipment issue", "Process variation", "Material defect",
        "Method not followed", "Environment factor", "Human error", 
        "Operator training", "Setup mistake"
    ]
    detection_hints = [
        "Inspection gap", "Testing method inadequate", "Detection process missing",
        "Tool calibration", "Measurement error"
    ]

    def generate_suggestions(previous_answers, base_hints):
        suggestions = []
        for ans in previous_answers:
            if ans.strip():
                for part in ans.split(","):
                    part_clean = part.strip()
                    if part_clean and part_clean not in suggestions:
                        suggestions.append(part_clean)
        for hint in base_hints:
            if hint not in suggestions:
                suggestions.append(hint)
        return suggestions

    st.markdown("#### Occurrence Analysis")
    for idx in range(5):
        prompt = f"{t[lang_key]['Occurrence_Why']} {idx+1}"
        if idx == 0:
            st.session_state.d5_occ_whys[idx] = st.text_input(prompt, value=st.session_state.d5_occ_whys[idx], key=f"occ_{idx}")
        else:
            options = generate_suggestions(st.session_state.d5_occ_whys[:idx], occurrence_hints)
            choice = st.selectbox(prompt, options=options, index=0, key=f"occ_{idx}_select")
            free_text = st.text_input(f"{prompt} (free text)", value="", key=f"occ_{idx}_free")
            st.session_state.d5_occ_whys[idx] = free_text if free_text.strip() else choice

    st.markdown("#### Detection Analysis")
    for idx in range(5):
        prompt = f"{t[lang_key]['Detection_Why']} {idx+1}"
        if idx == 0:
            st.session_state.d5_det_whys[idx] = st.text_input(prompt, value=st.session_state.d5_det_whys[idx], key=f"det_{idx}")
        else:
            options = generate_suggestions(st.session_state.d5_det_whys[:idx], detection_hints)
            choice = st.selectbox(prompt, options=options, index=0, key=f"det_{idx}_select")
            free_text = st.text_input(f"{prompt} (free text)", value="", key=f"det_{idx}_free")
            st.session_state.d5_det_whys[idx] = free_text if free_text.strip() else choice

    # Aggregate D5 answers
    st.session_state.D5["answer"] = (
        "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
        "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
    )
    st.session_state.D5["extra"] = st.text_area(f"{t[lang_key]['Root_Cause']}", value=st.session_state.D5["extra"], key="root_cause")

# ---------------------------
# Add date and prepared by at the bottom
# ---------------------------
report_date = datetime.now().strftime("%Y-%m-%d")
prepared_by = st.text_input("Prepared by / Preparado por", value="Your Name", key="prepared_by")
st.markdown(f"**Date / Fecha:** {report_date}")

# ---------------------------
# Download as XLS
# ---------------------------
def save_xls():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"
    headers = [t[lang_key][f"D{i}"] for i in range(1, 9)]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h).fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
    # Add answers
    for col, i in enumerate(range(1, 9), start=1):
        answer = st.session_state[f"D{i}"]["answer"]
        extra = st.session_state[f"D{i}"]["extra"]
        ws.cell(row=2, column=col, value=answer + ("\n" + extra if extra else ""))
    # Date & prepared by
    ws.cell(row=3, column=1, value=f"Date: {report_date}")
    ws.cell(row=4, column=1, value=f"Prepared by: {prepared_by}")
    return wb

wb = save_xls()
from io import BytesIO
output = BytesIO()
wb.save(output)
output.seek(0)
st.download_button("📥 Download XLS", data=output, file_name="8D_Report.xlsx")
