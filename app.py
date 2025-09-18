import streamlit as st
import pandas as pd
import datetime
import io
from openpyxl import Workbook

# ---------------------------
# LANGUAGE TRANSLATIONS
# ---------------------------
t = {
    "en": {
        "title": "8D Report Assistant",
        "select_language": "Select Language",
        "D1": "D1: Establish the Team",
        "D2": "D2: Describe the Problem",
        "D3": "D3: Implement Containment Actions",
        "D4": "D4: Identify Root Cause",
        "D5": "D5: Final Analysis",
        "Occurrence": "Occurrence Analysis",
        "Detection": "Detection Analysis",
        "Systemic": "Systemic Analysis",
        "D6": "D6: Implement Corrective Actions",
        "D7": "D7: Prevent Recurrence",
        "D8": "D8: Congratulate the Team",
        "Team_Members": "Team Members",
        "Problem_Description": "Problem Description",
        "Containment_Actions": "Containment Actions",
        "Root_Cause": "Root Cause",
        "Corrective_Actions": "Corrective Actions",
        "Preventive_Actions": "Preventive Actions",
        "Team_Congratulations": "Team Congratulations",
        "Report_Date": "Report Date",
        "Download_Report": "Download 8D Report",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Systemic_Why": "Systemic Why",
        "Add_Why": "Add Why",
        "Remove_Why": "Remove Why",
    }
}

# ---------------------------
# SESSION STATE INITIALIZATION
# ---------------------------
if "lang" not in st.session_state:
    st.session_state.lang = "en"

if "d1_team" not in st.session_state:
    st.session_state.d1_team = ""

if "d2_problem" not in st.session_state:
    st.session_state.d2_problem = ""

if "d3_containment" not in st.session_state:
    st.session_state.d3_containment = ""

if "d4_root_cause" not in st.session_state:
    st.session_state.d4_root_cause = ""

if "d5_occurrence" not in st.session_state:
    st.session_state.d5_occurrence = [""]

if "d5_detection" not in st.session_state:
    st.session_state.d5_detection = [""]

if "d5_systemic" not in st.session_state:
    st.session_state.d5_systemic = [""]

if "d6_corrective" not in st.session_state:
    st.session_state.d6_corrective = ""

if "d7_preventive" not in st.session_state:
    st.session_state.d7_preventive = ""

if "d8_congratulations" not in st.session_state:
    st.session_state.d8_congratulations = ""

if "report_date" not in st.session_state:
    st.session_state.report_date = datetime.datetime.today().strftime("%B %d, %Y")

# ---------------------------
# APP HEADER
# ---------------------------
lang_key = st.session_state.lang
st.title(t[lang_key]["title"])
st.session_state.lang = st.selectbox(t[lang_key]["select_language"], ["en"], index=0)
lang_key = st.session_state.lang

tab_d1, tab_d2, tab_d3, tab_d4, tab_d5, tab_d6, tab_d7, tab_d8 = st.tabs(
    [t[lang_key]["D1"], t[lang_key]["D2"], t[lang_key]["D3"], t[lang_key]["D4"],
     t[lang_key]["D5"], t[lang_key]["D6"], t[lang_key]["D7"], t[lang_key]["D8"]]
)
# ---------------------------
# D1: Establish the Team
# ---------------------------
with tab_d1:
    st.subheader(t[lang_key]["D1"])
    st.session_state.d1_team = st.text_area(
        t[lang_key]["Team_Members"], value=st.session_state.d1_team, key="d1_team"
    )

# ---------------------------
# D2: Describe the Problem
# ---------------------------
with tab_d2:
    st.subheader(t[lang_key]["D2"])
    st.session_state.d2_problem = st.text_area(
        t[lang_key]["Problem_Description"], value=st.session_state.d2_problem, key="d2_problem"
    )

# ---------------------------
# D3: Implement Containment Actions
# ---------------------------
with tab_d3:
    st.subheader(t[lang_key]["D3"])
    st.session_state.d3_containment = st.text_area(
        t[lang_key]["Containment_Actions"], value=st.session_state.d3_containment, key="d3_containment"
    )

# ---------------------------
# D4: Identify Root Cause
# ---------------------------
with tab_d4:
    st.subheader(t[lang_key]["D4"])
    st.session_state.d4_root_cause = st.text_area(
        t[lang_key]["Root_Cause"], value=st.session_state.d4_root_cause, key="d4_root_cause"
    )
    # ---------------------------
# D5: Final Analysis
# ---------------------------
with tab_d5:
    st.subheader(t[lang_key]["D5"])  # Only once, no duplicate

    # Training guidance
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
        <b>{t[lang_key]['Training_Guidance']}:</b> {d5_note[lang_key]}
        </div>
    """, unsafe_allow_html=True)

    # ---------------------------
    # Occurrence Analysis
    # ---------------------------
    st.markdown("#### Occurrence Analysis")
    for idx, val in enumerate(st.session_state.d5_occ_whys):
        st.session_state.d5_occ_whys[idx] = st.text_input(
            f"{t[lang_key]['Occurrence_Why']} {idx+1}",
            value=val,
            key=f"occ_{idx}"
        )

    # ---------------------------
    # Detection Analysis
    # ---------------------------
    st.markdown("#### Detection Analysis")
    for idx, val in enumerate(st.session_state.d5_det_whys):
        st.session_state.d5_det_whys[idx] = st.text_input(
            f"{t[lang_key]['Detection_Why']} {idx+1}",
            value=val,
            key=f"det_{idx}"
        )

    # ---------------------------
    # Systemic Analysis (New Section)
    # ---------------------------
    st.markdown("#### Systemic Analysis")
    st.session_state.d5_systemic = st.text_area(
        "Describe systemic factors contributing to the issue",
        value=st.session_state.get("d5_systemic", ""),
        key="d5_systemic"
    )

    # Suggested Root Causes
    suggested_occ_rc = "Occurrence-related root cause: " + ", ".join([w for w in st.session_state.d5_occ_whys if w])
    suggested_det_rc = "Detection-related root cause: " + ", ".join([w for w in st.session_state.d5_det_whys if w])

    st.session_state.D5["answer"] = st.text_area(
        f"{t[lang_key]['Root_Cause_Occ']}",
        value=suggested_occ_rc,
        key="root_cause_occ"
    )
    st.text_area(
        f"{t[lang_key]['Root_Cause_Det']}",
        value=suggested_det_rc,
        key="root_cause_det"
    )
    # ---------------------------
# D6–D8 Tabs
# ---------------------------
for step in ["D6", "D7", "D8"]:
    tab = {"D6": tab_d6, "D7": tab_d7, "D8": tab_d8}[step]
    with tab:
        st.subheader(t[lang_key][step])
        note_text = npqp_steps[[s[0] for s in npqp_steps].index(step)][1][lang_key]
        example_text = npqp_steps[[s[0] for s in npqp_steps].index(step)][2][lang_key]
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
            "Your Answer", value=st.session_state[step]["answer"], key=f"ans_{step}"
        )

# ---------------------------
# Excel Export
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

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

    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])

    header_row = ws.max_row + 1
    headers = ["Step", "Answer", "Extra / Notes"]
    fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=h)
        cell.fill = fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for step, ans in [(s[0], st.session_state[s[0]]["answer"]) for s in npqp_steps]:
        extra = st.session_state[step].get("extra", "")
        ws.append([t[lang_key][step], ans, extra])
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

# ---------------------------
# Sidebar: JSON Backup / Restore
# ---------------------------
with st.sidebar:
    st.markdown("## Backup / Restore")

    def generate_json():
        save_data = {k: v for k, v in st.session_state.items() if not k.startswith("_")}
        return json.dumps(save_data, indent=4)

    st.download_button(
        label="💾 Save Progress (JSON)",
        data=generate_json(),
        file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ', '_')}.json",
        mime="application/json"
    )

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

# ---------------------------
# Clear button removed
# ---------------------------
# Users can refresh the app to reset the form (no clear button needed)
