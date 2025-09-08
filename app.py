import streamlit as st
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.title("📑 Nissan NPQP 8D Report Trainer")

# -------------------------------------------------------------------
# Define NPQP 8D steps with training notes and examples
# -------------------------------------------------------------------
npqp_steps = [
    ("D0: Prepare and Plan",
     "Define the problem clearly and plan resources. This is the stage where you set the scope and urgency.",
     "Example: Customer complaint from Nissan for static noise in amplifier during end-of-line testing."),

    ("D1: Establish Team",
     "Form a cross-functional team with the knowledge, time, and authority to solve the problem.",
     "Example: SQE, Design Engineer, Manufacturing Engineer, Supplier Representative."),

    ("D2: Describe the Problem",
     "Use 5W2H (What, Where, When, Why, How, How many). Be specific and measurable.",
     "Example: 200 radios failed in Plant A during functional test due to distorted audio."),

    ("D3: Implement Containment",
     "Protect the customer immediately while you investigate. Containment is temporary, not the final fix.",
     "Example: Implement 100% inspection of amplifier boards before shipment."),

    ("D4: Identify Root Cause",
     "Use the 5-Why method to determine the root cause, separated into Occurrence (why the problem happened) and Detection (why it wasn’t detected). Start with 5 Whys but add more if needed.",
     "Training Example (Electronics):\nOccurrence:\nProblem: 100 radios fail functional test due to distorted audio.\n"
     "Why 1: Cold solder joint on DSP chip.\nWhy 2: Soldering process temperature too low.\nWhy 3: Operator did not follow soldering profile.\nWhy 4: Work instructions were unclear.\nWhy 5: SOP not updated after process change.\n"
     "Detection:\nWhy 1: Visual inspection not detailed enough.\nWhy 2: No automated solder check.\nWhy 3: QA checklist incomplete.\nRoot Cause: SOP not updated + inadequate inspection process"),

    ("D5: Choose Permanent Actions",
     "Define corrective actions that eliminate the root cause permanently.",
     "Example: Update soldering process parameters, retrain operators, and improve solder paste specification."),

    ("D6: Implement and Validate",
     "Put corrective actions in place and verify they solve the problem long-term.",
     "Example: Run accelerated life tests on corrected amplifiers to confirm no solder failures."),

    ("D7: Prevent Recurrence",
     "Update standards, procedures, training, and FMEAs to prevent the same issue in future.",
     "Example: Add automated solder inspection camera, update work instructions and PFMEA."),

    ("D8: Recognize the Team",
     "Celebrate success and acknowledge the team’s contribution.",
     "Example: Share results with management and recognize all engineers and operators involved.")
]

# -------------------------------------------------------------------
# Initialize session state
# -------------------------------------------------------------------
if "answers" not in st.session_state:
    st.session_state.answers = {}
if "extra_info" not in st.session_state:
    st.session_state.extra_info = {}
if "report_date" not in st.session_state:
    st.session_state.report_date = ""
if "prepared_by" not in st.session_state:
    st.session_state.prepared_by = ""
if "d4_occ_whys" not in st.session_state:
    st.session_state.d4_occ_whys = [""] * 5
if "d4_det_whys" not in st.session_state:
    st.session_state.d4_det_whys = [""] * 5

# Color dictionary for Excel rows
step_colors = {
    "D0: Prepare and Plan": "ADD8E6",      # Light Blue
    "D1: Establish Team": "90EE90",        # Light Green
    "D2: Describe the Problem": "FFFF99",  # Light Yellow
    "D3: Implement Containment": "FFD580", # Light Orange
    "D4: Identify Root Cause": "FF9999",   # Light Red
    "D5: Choose Permanent Actions": "D8BFD8", # Light Purple
    "D6: Implement and Validate": "E0FFFF",   # Light Cyan
    "D7: Prevent Recurrence": "D3D3D3",       # Light Gray
    "D8: Recognize the Team": "FFB6C1"        # Light Pink
}

# -------------------------------------------------------------------
# Report header
# -------------------------------------------------------------------
st.subheader("Report Information")
st.session_state.report_date = st.text_input("📅 Report Date (YYYY-MM-DD)", value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input("✍️ Prepared By", value=st.session_state.prepared_by)

# -------------------------------------------------------------------
# Form sections
# -------------------------------------------------------------------
st.subheader("NPQP 8D Steps")
for step, desc, example in npqp_steps:
    st.markdown(f"### {step}")
    st.info(f"**Training Note:** {desc}")
    st.write(f"💡 **Example:** {example}")

    # Special handling for D4 5-Why step
    if step.startswith("D4"):
        st.markdown("#### Occurrence Analysis")
        occ_whys = st.session_state.get("d4_occ_whys", [""] * 5)
        for i in range(len(occ_whys)):
            occ_whys[i] = st.text_input(f"Occurrence Why {i+1}", value=occ_whys[i], key=f"{step}_occ_{i}")
        if st.button("➕ Add another Occurrence Why", key=f"add_occ_{step}"):
            occ_whys.append("")
        st.session_state.d4_occ_whys = occ_whys

        st.markdown("#### Detection Analysis")
        det_whys = st.session_state.get("d4_det_whys", [""] * 5)
        for i in range(len(det_whys)):
            det_whys[i] = st.text_input(f"Detection Why {i+1}", value=det_whys[i], key=f"{step}_det_{i}")
        if st.button("➕ Add another Detection Why", key=f"add_det_{step}"):
            det_whys.append("")
        st.session_state.d4_det_whys = det_whys

        # Combine Occurrence and Detection into main answer
        combined_ans = "Occurrence Analysis:\n" + "\n".join([w for w in occ_whys if w.strip()]) + \
                       "\n\nDetection Analysis:\n" + "\n".join([w for w in det_whys if w.strip()])
        st.session_state.answers[step]
