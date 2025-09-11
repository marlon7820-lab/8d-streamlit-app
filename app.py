import streamlit as st
import datetime
from transformers import pipeline

# ---------------------------
# 1️⃣ CONFIG & SETUP
# ---------------------------
st.set_page_config(page_title="8D Report Builder", layout="wide")

# Load translation pipelines (cached so they don't reload every time)
@st.cache_resource
def load_translators():
    return {
        "en_to_es": pipeline("translation", model="Helsinki-NLP/opus-mt-en-es"),
        "es_to_en": pipeline("translation", model="Helsinki-NLP/opus-mt-es-en")
    }

translators = load_translators()

# Free lightweight model for interactive suggestions
@st.cache_resource
def load_suggester():
    return pipeline("text2text-generation", model="google/flan-t5-small")

suggester = load_suggester()

# ---------------------------
# 2️⃣ SESSION STATE INIT
# ---------------------------
if "language" not in st.session_state:
    st.session_state.language = "en"

if "answers" not in st.session_state:
    st.session_state.answers = {f"D{i}": "" for i in range(1, 9)}
    st.session_state.whys_occ = ["" for _ in range(5)]
    st.session_state.whys_det = ["" for _ in range(5)]
    st.session_state.report_date = datetime.date.today().strftime("%B %d, %Y")
    st.session_state.prepared_by = ""

# ---------------------------
# 3️⃣ TRANSLATION UTILS
# ---------------------------
def translate_text(text, direction="en_to_es"):
    try:
        if text.strip():
            result = translators[direction](text)[0]['translation_text']
            return result
    except Exception:
        pass
    return text

def t(text):
    if st.session_state.language == "es":
        return translate_text(text, "en_to_es")
    return text

# ---------------------------
# 4️⃣ LANGUAGE SWITCHER
# ---------------------------
col1, col2 = st.columns([4,1])
with col2:
    lang = st.radio("🌐", ["English", "Español"], index=0 if st.session_state.language=="en" else 1)
    st.session_state.language = "en" if lang == "English" else "es"

# ---------------------------
# 5️⃣ HEADER
# ---------------------------
st.title(t("8D Report Builder"))
with st.container():
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.report_date = st.date_input(
            t("Report Date"), 
            datetime.datetime.strptime(st.session_state.report_date, "%B %d, %Y")
        ).strftime("%B %d, %Y")
    with c2:
        st.session_state.prepared_by = st.text_input(
            t("Prepared By"), 
            st.session_state.prepared_by
        )

# ---------------------------
# 6️⃣ D-TABS
# ---------------------------
tabs = st.tabs([f"D{i}" for i in range(1,9)])

# D1 — Team Definition
with tabs[0]:
    st.header(t("D1 – Define the Team"))
    st.session_state.answers["D1"] = st.text_area(
        t("Team Members & Roles"), 
        st.session_state.answers["D1"]
    )
    st.text_area(t("Additional Notes"), key="D1_extra")

# D2 – Problem Description
with tabs[1]:
    st.header(t("D2 – Describe the Problem"))
    st.session_state.answers["D2"] = st.text_area(
        t("Problem Description"), 
        st.session_state.answers["D2"]
    )

# D3 – Interim Containment
with tabs[2]:
    st.header(t("D3 – Interim Containment"))
    st.session_state.answers["D3"] = st.text_area(
        t("Containment Actions"), 
        st.session_state.answers["D3"]
    )

# D4 – Root Cause Analysis
with tabs[3]:
    st.header(t("D4 – Root Cause"))
    st.session_state.answers["D4"] = st.text_area(
        t("Root Cause"), 
        st.session_state.answers["D4"]
    )

# D5 – Interactive 5-Why
with tabs[4]:
    st.header(t("D5 – 5-Why Analysis"))

    st.subheader(t("Occurrence Analysis"))
    for i in range(5):
        st.session_state.whys_occ[i] = st.text_input(
            t(f"Why {i+1} (Occurrence)"),
            value=st.session_state.whys_occ[i],
            key=f"occ_{i}"
        )
        if st.session_state.whys_occ[i]:
            suggestions = suggester(f"Suggest possible next root causes based on: {st.session_state.whys_occ[i]}")
            st.caption("💡 " + suggestions[0]['generated_text'])

    st.subheader(t("Detection Analysis"))
    for i in range(5):
        st.session_state.whys_det[i] = st.text_input(
            t(f"Why {i+1} (Detection)"),
            value=st.session_state.whys_det[i],
            key=f"det_{i}"
        )
        if st.session_state.whys_det[i]:
            suggestions = suggester(f"Suggest possible next detection failure reasons based on: {st.session_state.whys_det[i]}")
            st.caption("💡 " + suggestions[0]['generated_text'])

# D6 – Permanent Actions
with tabs[5]:
    st.header(t("D6 – Permanent Actions"))
    st.session_state.answers["D6"] = st.text_area(
        t("Permanent Corrective Actions"), 
        st.session_state.answers["D6"]
    )

# D7 – Prevent Recurrence
with tabs[6]:
    st.header(t("D7 – Prevent Recurrence"))
    st.session_state.answers["D7"] = st.text_area(
        t("Systemic Fixes to Prevent Recurrence"), 
        st.session_state.answers["D7"]
    )

# D8 – Team Recognition
with tabs[7]:
    st.header(t("D8 – Team Recognition"))
    st.session_state.answers["D8"] = st.text_area(
        t("Team Recognition"), 
        st.session_state.answers["D8"]
    )

st.success(t("✅ All inputs are auto-saved in memory."))
