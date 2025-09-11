import streamlit as st
import datetime
from openpyxl import Workbook
from deep_translator import GoogleTranslator

# ------------------------------
# Translation Helper
# ------------------------------
def translate_text(text, direction="en_to_es"):
    if not text.strip():
        return text
    try:
        if direction == "en_to_es":
            return GoogleTranslator(source="en", target="es").translate(text)
        else:
            return GoogleTranslator(source="es", target="en").translate(text)
    except Exception:
        return text

# ------------------------------
# Streamlit UI
# ------------------------------
st.set_page_config(page_title="8D Report", layout="wide")
st.title("8D Report Form")

# Language selector
lang = st.radio("Language / Idioma", ["English", "Español"], horizontal=True)
direction = "en_to_es" if lang == "Español" else "es_to_en"

def T(text):
    """Translate text if needed"""
    if lang == "English":
        return text
    return translate_text(text, direction)

# Store answers in session_state
if "answers" not in st.session_state:
    st.session_state.answers = {}

# ------------------------------
# Global Header
# ------------------------------
st.session_state.answers["reported_date"] = st.date_input(
    T("Reported Date"), value=datetime.date.today()
)
st.session_state.answers["prepared_by"] = st.text_input(T("Prepared By"))

# ------------------------------
# D1 - Team
# ------------------------------
st.header(T("D1: Establish the Team"))
st.session_state.answers["team_description"] = st.text_area(
    T("Describe the Team Members and Roles")
)

# ------------------------------
# D2 - Problem Description
# ------------------------------
st.header(T("D2: Describe the Problem"))
st.session_state.answers["problem_description"] = st.text_area(
    T("Problem Statement")
)

# ------------------------------
# D3 - Containment Actions
# ------------------------------
st.header(T("D3: Implement and Verify Containment Actions"))
st.session_state.answers["containment_actions"] = st.text_area(
    T("Containment Actions")
)

# ------------------------------
# D4 - Root Cause Analysis
# ------------------------------
st.header(T("D4: Root Cause Analysis"))

st.subheader(T("5-Why for Occurrence"))
if "occurrence_whys" not in st.session_state:
    st.session_state.occurrence_whys = [""]

for i in range(len(st.session_state.occurrence_whys)):
    st.session_state.occurrence_whys[i] = st.text_input(
        T(f"Why {i+1} (Occurrence)"), value=st.session_state.occurrence_whys[i], key=f"occ_{i}"
    )

if st.button(T("Add Why for Occurrence")):
    st.session_state.occurrence_whys.append("")

st.subheader(T("5-Why for Detection"))
if "detection_whys" not in st.session_state:
    st.session_state.detection_whys = [""]

for i in range(len(st.session_state.detection_whys)):
    st.session_state.detection_whys[i] = st.text_input(
        T(f"Why {i+1} (Detection)"), value=st.session_state.detection_whys[i], key=f"det_{i}"
    )

if st.button(T("Add Why for Detection")):
    st.session_state.detection_whys.append("")

# ------------------------------
# D5 - Permanent Corrective Actions
# ------------------------------
st.header(T("D5: Choose and Verify Permanent Corrective Actions"))
st.session_state.answers["corrective_actions"] = st.text_area(
    T("List Permanent Corrective Actions")
)

# ------------------------------
# Save Button
# ------------------------------
if st.button(T("Save 8D Report")):
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"
    row = 1
    for key, value in st.session_state.answers.items():
        ws.cell(row=row, column=1, value=key)
        ws.cell(row=row, column=2, value=str(value))
        row += 1
    # Save Why Analysis
    ws.cell(row=row, column=1, value="5-Why Occurrence")
    ws.cell(row=row, column=2, value="\n".join(st.session_state.occurrence_whys))
    row += 1
    ws.cell(row=row, column=1, value="5-Why Detection")
    ws.cell(row=row, column=2, value="\n".join(st.session_state.detection_whys))

    wb.save("8D_Report.xlsx")
    st.success(T("✅ 8D Report saved successfully!"))
