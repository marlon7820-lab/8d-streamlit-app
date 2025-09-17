# ------------------- PART 1 -------------------

import streamlit as st
import datetime
import json
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

# ------------------- INITIAL SETUP -------------------

st.set_page_config(page_title="8D Report Assistant", layout="wide")

# Initialize session state safely
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.report_date = datetime.datetime.today().strftime("%B %d, %Y")
    st.session_state.prepared_by = ""
    st.session_state.d5_occ_whys = [""] * 5
    st.session_state.d5_det_whys = [""] * 5
    st.session_state.d5_sys_whys = [""] * 5
    st.session_state.d5_occ_selected = []
    st.session_state.d5_det_selected = []
    st.session_state.d5_sys_selected = []
    st.session_state.answers = {}

# Translation dictionary (example keys, expand as needed)
t = {
    "en": {
        "Report_Date": "Report Date",
        "Prepared_By": "Prepared By",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Systemic_Why": "Systemic Why",
    }
}
lang_key = "en"

# Define steps (step_id, step_name, description)
npqp_steps = [
    ("D1", "D1 – Team", "Form the team"),
    ("D2", "D2 – Problem", "Describe the problem"),
    ("D3", "D3 – Containment", "Containment actions"),
    ("D4", "D4 – Root Cause", "Root cause analysis"),
    ("D5", "D5 – Corrective Action", "Identify and verify corrective actions"),
    ("D6", "D6 – Implement", "Implement corrective actions"),
    ("D7", "D7 – Prevent Recurrence", "Prevent recurrence"),
    ("D8", "D8 – Congratulate", "Congratulate the team"),
]

st.title("8D Report Assistant")

# Prepare tab labels
tab_labels = [step_name for _, step_name, _ in npqp_steps]
tabs = st.tabs(tab_labels)

# ------------------- STEP LOOP (D1 - D4) -------------------

for i, (step, step_name, description) in enumerate(npqp_steps):
    with tabs[i]:
        st.subheader(step_name)
        st.write(description)

        # Basic text area for answers (except D5, handled separately)
        if step != "D5":
            if step not in st.session_state.answers:
                st.session_state.answers[step] = ""
            st.session_state.answers[step] = st.text_area(
                "Your Answer",
                value=st.session_state.answers[step],
                key=f"ans_{step}"
            )

# ------------------- START OF D5 SECTION -------------------

with tabs[4]:
    st.subheader("D5 – Corrective Action")
    st.write("Perform root cause analysis using 5-Why for Occurrence, Detection, and Systemic issues.")
    # ------------------- PART 2 -------------------
# Continuation inside D5 tab

    st.markdown("### 5-Why Analysis")

    # ---------- OCCURRENCE SECTION ----------
    st.markdown("#### Occurrence Analysis")
    occurrence_options = [
        "Incorrect process parameter",
        "Design issue",
        "Supplier defect",
        "Machine failure",
        "Training gap"
    ]
    for idx in range(5):
        col1, col2 = st.columns([3, 2])
        with col1:
            st.session_state.d5_occ_whys[idx] = st.text_input(
                f"{t[lang_key]['Occurrence_Why']} {idx+1}",
                value=st.session_state.d5_occ_whys[idx],
                key=f"occ_why_{idx}"
            )
        with col2:
            selected = st.selectbox(
                f"Select Occurrence Cause {idx+1}",
                [""] + occurrence_options,
                index=(occurrence_options.index(st.session_state.d5_occ_whys[idx]) + 1)
                if st.session_state.d5_occ_whys[idx] in occurrence_options else 0,
                key=f"occ_sel_{idx}"
            )
            if selected and selected != st.session_state.d5_occ_whys[idx]:
                st.session_state.d5_occ_whys[idx] = selected

    if st.button("Clear Occurrence Whys"):
        for idx in range(5):
            st.session_state.d5_occ_whys[idx] = ""
            st.session_state[f"occ_why_{idx}"] = ""
            st.session_state[f"occ_sel_{idx}"] = ""

    st.markdown("---")

    # ---------- DETECTION SECTION ----------
    st.markdown("#### Detection Analysis")
    detection_options = [
        "Inspection missed defect",
        "Test equipment out of calibration",
        "Lack of error proofing",
        "Sampling plan inadequate",
        "Operator did not detect"
    ]
    for idx in range(5):
        col1, col2 = st.columns([3, 2])
        with col1:
            st.session_state.d5_det_whys[idx] = st.text_input(
                f"{t[lang_key]['Detection_Why']} {idx+1}",
                value=st.session_state.d5_det_whys[idx],
                key=f"det_why_{idx}"
            )
        with col2:
            selected = st.selectbox(
                f"Select Detection Cause {idx+1}",
                [""] + detection_options,
                index=(detection_options.index(st.session_state.d5_det_whys[idx]) + 1)
                if st.session_state.d5_det_whys[idx] in detection_options else 0,
                key=f"det_sel_{idx}"
            )
            if selected and selected != st.session_state.d5_det_whys[idx]:
                st.session_state.d5_det_whys[idx] = selected

    if st.button("Clear Detection Whys"):
        for idx in range(5):
            st.session_state.d5_det_whys[idx] = ""
            st.session_state[f"det_why_{idx}"] = ""
            st.session_state[f"det_sel_{idx}"] = ""

    st.markdown("---")

    # ---------- SYSTEMIC SECTION ----------
    st.markdown("#### Systemic Analysis")
    systemic_options = [
        "Procedure not updated",
        "FMEA not reviewed",
        "Lessons learned not captured",
        "Lack of periodic audits",
        "Poor change management"
    ]
    for idx in range(5):
        col1, col2 = st.columns([3, 2])
        with col1:
            st.session_state.d5_sys_whys[idx] = st.text_input(
                f"{t[lang_key]['Systemic_Why']} {idx+1}",
                value=st.session_state.d5_sys_whys[idx],
                key=f"sys_why_{idx}"
            )
        with col2:
            selected = st.selectbox(
                f"Select Systemic Cause {idx+1}",
                [""] + systemic_options,
                index=(systemic_options.index(st.session_state.d5_sys_whys[idx]) + 1)
                if st.session_state.d5_sys_whys[idx] in systemic_options else 0,
                key=f"sys_sel_{idx}"
            )
            if selected and selected != st.session_state.d5_sys_whys[idx]:
                st.session_state.d5_sys_whys[idx] = selected

    if st.button("Clear Systemic Whys"):
        for idx in range(5):
            st.session_state.d5_sys_whys[idx] = ""
            st.session_state[f"sys_why_{idx}"] = ""
            st.session_state[f"sys_sel_{idx}"] = ""
            # ------------------- PART 3 -------------------
# Render D6–D8 tabs

for step in ["D6", "D7", "D8"]:
    st.markdown(f"### {t[lang_key][step]}")

    note_text = note_dict[lang_key] if step != "D8" else npqp_steps[7][1][lang_key]
    example_text = example_dict[lang_key] if step != "D8" else npqp_steps[7][2][lang_key]

    st.markdown(f"""
    <div style="
        background-color:#b3e0ff; 
        color:black; 
        padding:12px; 
        border-left:5px solid #1E90FF; 
        border-radius:6px;
        width:100%;
        font-size:14px;
        line-height:1.5;
    ">
    <b>{t[lang_key]['Training_Guidance']}:</b> {note_text}<br><br>
    💡 <b>{t[lang_key]['Example']}:</b> {example_text}
    </div>
    """, unsafe_allow_html=True)

    st.session_state[step]["answer"] = st.text_area(
        "Your Answer",
        value=st.session_state[step]["answer"],
        key=f"ans_{step}"
    )

# ------------------- COLLECT ALL 8D ANSWERS -------------------
all_data_rows = []

# Include D1–D8 answers
for step, _, _ in npqp_steps:
    answer = st.session_state[step]["answer"] if step in st.session_state else ""
    extra = st.session_state[step].get("extra", "") if step in st.session_state else ""
    all_data_rows.append((step, answer, extra))

# Include D5 5-Why details
d5_occ_str = "\n".join([w for w in st.session_state.d5_occ_whys if w])
d5_det_str = "\n".join([w for w in st.session_state.d5_det_whys if w])
d5_sys_str = "\n".join([w for w in st.session_state.d5_sys_whys if w])

st.session_state.D5["answer"] = (
    f"Occurrence:\n{d5_occ_str}\n\n"
    f"Detection:\n{d5_det_str}\n\n"
    f"Systemic:\n{d5_sys_str}"
)
# ------------------- PART 4 -------------------
# ------------------- EXCEL EXPORT -------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Logo
    if os.path.exists("logo.png"):
        try:
            img = XLImage("logo.png")
            img.width = 140
            img.height = 40
            ws.add_image(img, "A1")
        except:
            pass

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=3)
    ws.cell(row=3, column=1, value="📋 8D Report Assistant").font = Font(bold=True, size=14)

    # Report info
    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    # Headers
    header_row = ws.max_row + 1
    headers = ["Step", "Answer", "Extra / Notes"]
    fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    # Data
    for step, answer, extra in all_data_rows:
        ws.append([t[lang_key][step], answer, extra])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font = Font(bold=True if c == 2 else False)
            cell.border = border

    for col in range(1, 4):
        ws.column_dimensions[get_column_letter(col)].width = 40

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ', '_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ------------------- SIDEBAR: JSON BACKUP / RESTORE + RESET -------------------
with st.sidebar:
    st.markdown("## Backup / Restore")

    # Save JSON
    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
        mime="application/json"
    )

    # Restore JSON
    st.markdown("---")
    st.markdown("### Restore from JSON")
    uploaded_file = st.file_uploader("Upload JSON file to restore", type="json")
    if uploaded_file:
        try:
            restore_data = json.load(uploaded_file)
            for k, v in restore_data.items():
                st.session_state[k] = v
            st.success("✅ Session restored from JSON!")
        except Exception as e:
            st.error(f"Error restoring JSON: {e}")

    # Clear / Reset All
    st.markdown("---")
    st.markdown("### Reset All Data")
    if st.button("🗑️ Clear All"):
        for step, _, _ in npqp_steps:
            st.session_state[step] = {"answer": "", "extra": ""}
        # Reset D5 5-Why details
        st.session_state["d5_occ_whys"] = [""] * 5
        st.session_state["d5_det_whys"] = [""] * 5
        st.session_state["d5_sys_whys"] = [""] * 5
        st.session_state["d5_occ_selected"] = []
        st.session_state["d5_det_selected"] = []
        st.session_state["d5_sys_selected"] = []
        st.session_state["report_date"] = datetime.datetime.today().strftime("%B %d, %Y")
        st.session_state["prepared_by"] = ""
        st.success("✅ All data has been reset!")
