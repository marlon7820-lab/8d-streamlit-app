import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from openpyxl import Workbook

# ---------------------------
# Session State Initialization
# ---------------------------
if "initialized" not in st.session_state:
    st.session_state.initialized = True
    st.session_state.report_date = datetime.datetime.today().strftime("%B %d, %Y")
    st.session_state.team_members = ""
    st.session_state.problem_description = ""
    st.session_state.containment_actions = ""
    st.session_state.root_causes_occurrence = ["" for _ in range(5)]
    st.session_state.root_causes_detection = ["" for _ in range(5)]
    st.session_state.root_causes_systemic = ["" for _ in range(5)]
    st.session_state.corrective_actions = ""
    st.session_state.verification = ""
    st.session_state.preventive_actions = ""
    st.session_state.congratulations = ""

# ---------------------------
# Translation Dictionary
# ---------------------------
t = {
    "en": {
        "Report_Date": "Report Date",
        "Team_Members": "Team Members",
        "Problem_Description": "Problem Description",
        "Containment_Actions": "Containment Actions",
        "Occurrence_Why": "Occurrence Why",
        "Detection_Why": "Detection Why",
        "Systemic_Why": "Systemic Why",
        "Corrective_Actions": "Corrective Actions",
        "Verification": "Verification of Actions",
        "Preventive_Actions": "Preventive Actions",
        "Congratulations": "Team Recognition",
        "D1": "D1: Establish the Team",
        "D2": "D2: Describe the Problem",
        "D3": "D3: Implement and Verify Interim Containment Actions",
        "D4": "D4: Identify Root Causes",
        "D5": "D5: Final Analysis",
        "D6": "D6: Implement and Validate Corrective Actions",
        "D7": "D7: Prevent Recurrence",
        "D8": "D8: Congratulate the Team"
    }
}

lang_key = "en"

# ---------------------------
# Page Layout
# ---------------------------
st.set_page_config(page_title="8D Report Assistant", layout="wide")
st.title("8D Report Assistant")
# ---------------------------
# Tabs for 8D Steps
# ---------------------------
npqp_steps = ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8"]
tabs = st.tabs([t[lang_key][step] for step in npqp_steps])

for i, step in enumerate(npqp_steps):
    with tabs[i]:
        # ✅ FIX: Only skip duplicate header for D5
        if step != "D5":
            st.markdown(f"### {t[lang_key][step]}")

        if step == "D1":
            st.text_area(t[lang_key]["Team_Members"], key="team_members")
        elif step == "D2":
            st.text_area(t[lang_key]["Problem_Description"], key="problem_description")
        elif step == "D3":
            st.text_area(t[lang_key]["Containment_Actions"], key="containment_actions")
        elif step == "D4":
            st.subheader("Root Cause Analysis")
            st.markdown("#### Occurrence")
            for idx in range(5):
                st.session_state.root_causes_occurrence[idx] = st.text_input(
                    f"{t[lang_key]['Occurrence_Why']} {idx+1}",
                    value=st.session_state.root_causes_occurrence[idx],
                    key=f"occurrence_{idx}"
                )
            st.markdown("#### Detection")
            for idx in range(5):
                st.session_state.root_causes_detection[idx] = st.text_input(
                    f"{t[lang_key]['Detection_Why']} {idx+1}",
                    value=st.session_state.root_causes_detection[idx],
                    key=f"detection_{idx}"
                )
            st.markdown("#### Systemic")
            for idx in range(5):
                st.session_state.root_causes_systemic[idx] = st.text_input(
                    f"{t[lang_key]['Systemic_Why']} {idx+1}",
                    value=st.session_state.root_causes_systemic[idx],
                    key=f"systemic_{idx}"
                )
        elif step == "D5":
            st.text_area("Enter Final Analysis Details", key="final_analysis")
        elif step == "D6":
            st.text_area(t[lang_key]["Corrective_Actions"], key="corrective_actions")
        elif step == "D7":
            st.text_area(t[lang_key]["Preventive_Actions"], key="preventive_actions")
        elif step == "D8":
            st.text_area(t[lang_key]["Congratulations"], key="congratulations")
            # ---------------------------
# Excel Export Function
# ---------------------------
def create_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "8D Report"

    ws.append(["Report Date", st.session_state.report_date])
    ws.append(["Team Members", st.session_state.team_members])
    ws.append(["Problem Description", st.session_state.problem_description])
    ws.append(["Containment Actions", st.session_state.containment_actions])

    ws.append([])
    ws.append(["Root Cause - Occurrence"])
    for idx, cause in enumerate(st.session_state.root_causes_occurrence):
        ws.append([f"Why {idx+1}", cause])

    ws.append([])
    ws.append(["Root Cause - Detection"])
    for idx, cause in enumerate(st.session_state.root_causes_detection):
        ws.append([f"Why {idx+1}", cause])

    ws.append([])
    ws.append(["Root Cause - Systemic"])
    for idx, cause in enumerate(st.session_state.root_causes_systemic):
        ws.append([f"Why {idx+1}", cause])

    ws.append([])
    ws.append(["Final Analysis", st.session_state.get("final_analysis", "")])
    ws.append(["Corrective Actions", st.session_state.corrective_actions])
    ws.append(["Preventive Actions", st.session_state.preventive_actions])
    ws.append(["Congratulations", st.session_state.congratulations])

    output = BytesIO()
    wb.save(output)
    return output.getvalue()
    # ---------------------------
# Download Button
# ---------------------------
st.markdown("---")
excel_data = create_excel()
st.download_button(
    label="📥 Download 8D Report as Excel",
    data=excel_data,
    file_name="8D_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
