import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
from transformers import pipeline

# ---------------------------
# Load AI model (Hugging Face GPT-2)
# ---------------------------
@st.cache_resource
def load_ai_model():
    return pipeline('text-generation', model='gpt2')

generator = load_ai_model()

def suggest_root_cause_hf(whys_list, lang="en"):
    chain_text = "\n".join([f"{i+1}. {w}" for i, w in enumerate(whys_list) if w.strip()])
    if not chain_text:
        return ""
    prompt = f"Based on the following 5-Why analysis, suggest a concise root cause in {'English' if lang=='en' else 'Spanish'}:\n{chain_text}\nRoot Cause:"
    try:
        response = generator(prompt, max_length=150, do_sample=True, temperature=0.7)
        return response[0]['generated_text'].split("Root Cause:")[-1].strip()
    except Exception as e:
        st.warning(f"Root cause suggestion failed: {e}")
        return ""

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

st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📑 8D Training App</h1>", unsafe_allow_html=True)

# ---------------------------
# NPQP 8D Steps with guidance/examples
# ---------------------------
npqp_steps = [
    ("D1: Concern Details",
     "Describe the customer concerns clearly. Include what the issue is, where it occurred, when, and any supporting data.",
     "Example: Customer reported static noise in amplifier during end-of-line test at Plant A."),
    ("D2: Similar Part Considerations",
     "Check for similar parts, models, generic parts, other colors, opposite hand, front/rear, etc. to see if issue is recurring or isolated.",
     "Example: Same speaker type used in another radio model; different amplifier colors; front vs. rear audio units."),
    ("D3: Initial Analysis",
     "Perform an initial investigation to identify obvious issues, collect data, and document initial findings.",
     "Example: Visual inspection of solder joints, initial functional tests, checking connectors."),
    ("D4: Implement Containment",
     "Define temporary containment actions to prevent the customer from seeing the problem while permanent actions are developed.",
     "Example: 100% inspection of amplifiers before shipment; use of temporary shielding; quarantine of affected batches."),
    ("D5: Final Analysis",
     "Use 5-Why analysis to determine the root cause. Separate by Occurrence (why it happened) and Detection (why it wasn’t detected). Add more Whys if needed.",
     ""),  # D5 training guidance will be added dynamically
    ("D6: Permanent Corrective Actions",
     "Define corrective actions that eliminate the root cause permanently and prevent recurrence.",
     "Example: Update soldering process, retrain operators, update work instructions, and add automated inspection."),
    ("D7: Countermeasure Confirmation",
     "Verify that corrective actions effectively resolve the issue long-term.",
     "Example: Functional tests on corrected amplifiers, accelerated life testing, and monitoring of first production runs."),
    ("D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
     "Document lessons learned, update standards, procedures, FMEAs, and training to prevent recurrence.",
     "Example: Update SOPs, PFMEA, work instructions, and employee training to prevent the same issue in future.")
]

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
st.session_state.setdefault("prev_lang", "en")

# Color dictionary for Excel
step_colors = {
    "D1: Concern Details": "ADD8E6",
    "D2: Similar Part Considerations": "90EE90",
    "D3: Initial Analysis": "FFFF99",
    "D4: Implement Containment": "FFD580",
    "D5: Final Analysis": "FF9999",
    "D6: Permanent Corrective Actions": "D8BFD8",
    "D7: Countermeasure Confirmation": "E0FFFF",
    "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)": "D3D3D3"
}

# ---------------------------
# Report info
# ---------------------------
st.subheader("Report Information")
today_str = datetime.datetime.today().strftime("%B %d, %Y")
st.session_state.report_date = st.text_input("📅 Report Date", value=today_str)
st.session_state.prepared_by = st.text_input("✍️ Prepared By", st.session_state.prepared_by)

# ---------------------------
# Language selection
# ---------------------------
lang = st.radio("Language / Idioma", ["en", "es"], horizontal=True)
if lang != st.session_state.prev_lang:
    st.session_state.prev_lang = lang
    # Placeholder: could later implement translation of existing answers

# ---------------------------
# Tabs for each D-step
# ---------------------------
tabs = st.tabs([step for step, _, _ in npqp_steps])
for i, (step, note, example) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {step}")

        # Guidance text
        if step.startswith("D5"):
            full_training_note = (
                "**Training Guidance:** Use 5-Why analysis to determine the root cause.\n\n"
                "**Occurrence Example (5-Whys):**\n"
                "1. Cold solder joint on DSP chip\n2. Soldering temperature too low\n3. Operator didn’t follow profile\n4. Work instructions were unclear\n5. No visual confirmation step\n\n"
                "**Detection Example (5-Whys):**\n"
                "1. QA inspection missed cold joint\n2. Inspection checklist incomplete\n3. No automated test step\n4. Batch testing not performed\n5. Early warning signal not tracked\n\n"
                "**Root Cause Example:**\n"
                "Insufficient process control on soldering operation, combined with inadequate QA checklist, "
                "allowed defective DSP soldering to pass undetected."
            )
            st.info(full_training_note)
        else:
            st.info(f"**Training Guidance:** {note}\n\n💡 **Example:** {example}")

        # ---------------------------
        # Input fields
        # ---------------------------
        if step.startswith("D5"):
            st.markdown("#### Occurrence Analysis")
            for idx, val in enumerate(st.session_state.d5_occ_whys):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"Occurrence Why {idx+1}", value=val, key=f"{step}_occ_{idx}")
            if st.button("➕ Add another Occurrence Why", key=f"add_occ_{step}"):
                st.session_state.d5_occ_whys.append("")

            st.markdown("#### Detection Analysis")
            for idx, val in enumerate(st.session_state.d5_det_whys):
                st.session_state.d5_det_whys[idx] = st.text_input(f"Detection Why {idx+1}", value=val, key=f"{step}_det_{idx}")
            if st.button("➕ Add another Detection Why", key=f"add_det_{step}"):
                st.session_state.d5_det_whys.append("")

            st.markdown("#### Interactive 5-Why / AI Root Cause")
            for idx, val in enumerate(st.session_state.interactive_whys):
                st.session_state.interactive_whys[idx] = st.text_input(f"Why {idx+1}", value=val, key=f"interactive_{idx}")
                if idx == len(st.session_state.interactive_whys)-1 and val.strip():
                    st.session_state.interactive_whys.append("")

            if st.button("Generate Root Cause Suggestion / Sugerencia AI"):
                whys_input = [w for w in st.session_state.interactive_whys if w.strip()]
                suggested_cause = suggest_root_cause_hf(whys_input, lang)
                if suggested_cause:
                    st.success(f"Suggested Root Cause / Causa Raíz Sugerida: {suggested_cause}")
                    st.session_state[step]["extra"] = suggested_cause
                else:
                    st.warning("No suggestion generated. Fill in the Why fields first.")

            # Save answers in session state
            st.session_state[step]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state.d5_occ_whys if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state.d5_det_whys if w.strip()])
            )
        else:
            st.session_state[step]["answer"] = st.text_area(f"Your Answer for {step}", value=st.session_state[step]["answer"], key=f"ans_{step}")
            st.session_state[step]["extra"] = st.text_area(f"Root Cause / Extra (if applicable)", value=st.session_state[step]["extra"], key=f"extra_{step}")

# ---------------------------
# Collect answers for Excel
# ---------------------------
data_rows = [(step, st.session_state[step]["answer"], st.session_state[step]["extra"]) for step, _, _ in npqp_steps]

# ---------------------------
# Save button with styled Excel
# ---------------------------
if st.button("💾 Save 8D Report"):
    if not any(ans for _, ans, _ in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
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
        for step, ans, extra in data_rows:
            ws.cell(row=row, column=1, value=step)
            ws.cell(row=row, column=2, value=ans)
            ws.cell(row=row, column=3, value=extra)

            fill_color = step_colors.get(step, "FFFFFF")
            for col in range(1, 4):
                ws.cell(row=row, column=col).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=row, column=col).alignment = Alignment(wrap_text=True, vertical="top")

            row += 1

        # Adjust column widths
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 40

        wb.save(xlsx_file)

        st.success("✅ NPQP 8D Report saved successfully.")
        with open(xlsx_file, "rb") as f:
            st.download_button("📥 Download XLSX", f, file_name=xlsx_file)
