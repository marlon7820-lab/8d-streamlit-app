import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import datetime
import io
import json
import os

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="8D Report Assistant",
    page_icon="logo.png",
    layout="wide"
)

# ---------------------------
# App styles
# ---------------------------
st.markdown("""
<style>
.stApp {background: linear-gradient(to right, #f0f8ff, #e6f2ff); color: #000000 !important;}
.stTabs [data-baseweb="tab"] {font-weight: bold; color: #000000 !important;}
textarea {background-color: #ffffff !important; border: 1px solid #1E90FF !important; border-radius: 5px; color: #000000 !important;}
.stInfo {background-color: #e6f7ff !important; border-left: 5px solid #1E90FF !important; color: #000000 !important;}
button[kind="primary"] {background-color: #87AFC7 !important; color: white !important; font-weight: bold;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Main title
# ---------------------------
st.markdown("<h1 style='text-align: center; color: #1E90FF;'>📋 8D Report Assistant</h1>", unsafe_allow_html=True)

# ---------------------------
# Version info
# ---------------------------
version_number = "v1.0.9"
last_updated = "October 10, 2025"
st.markdown(f"""
<hr style='border:1px solid #1E90FF; margin-top:10px; margin-bottom:5px;'> 
<p style='font-size:12px; font-style:italic; text-align:center; color:#555555;'>
Version {version_number} | Last updated: {last_updated}
</p>
""", unsafe_allow_html=True)

# ---------------------------
# Sidebar: Language selection
# ---------------------------
st.sidebar.title("8D Report Assistant")
st.sidebar.markdown("---")
st.sidebar.header("Settings")

lang = st.sidebar.selectbox("Select Language / Seleccionar Idioma", ["English", "Español"])
lang_key = "en" if lang == "English" else "es"

# ---------------------------
# Language dictionary
# ---------------------------
t = {
    "en": {
        "D1": "D1: Concern Details", "D2": "D2: Similar Part Considerations",
        "D3": "D3: Initial Analysis", "D4": "D4: Implement Containment",
        "D5": "D5: Final Analysis", "D6": "D6: Permanent Corrective Actions",
        "D7": "D7: Countermeasure Confirmation", "D8": "D8: Follow-up Activities (Lessons Learned / Recurrence Prevention)",
        "Report_Date": "Report Date", "Prepared_By": "Prepared By",
        "Root_Cause_Occ": "Root Cause (Occurrence)", "Root_Cause_Det": "Root Cause (Detection)", "Root_Cause_Sys": "Root Cause (Systemic)",
        "Occurrence_Why": "Occurrence Why", "Detection_Why": "Detection Why", "Systemic_Why": "Systemic Why",
        "Save": "💾 Save 8D Report", "Download": "📥 Download XLSX",
        "Training_Guidance": "Training Guidance", "Example": "Example",
        "FMEA_Failure": "FMEA Failure Occurrence",
        "Location": "Material Location", "Status": "Activity Status", "Containment_Actions": "Containment Actions"
    },
    "es": {
        "D1": "D1: Detalles de la preocupación", "D2": "D2: Consideraciones de partes similares",
        "D3": "D3: Análisis inicial", "D4": "D4: Implementar contención",
        "D5": "D5: Análisis final", "D6": "D6: Acciones correctivas permanentes",
        "D7": "D7: Confirmación de contramedidas", "D8": "D8: Actividades de seguimiento (Lecciones aprendidas / Prevención de recurrencia)",
        "Report_Date": "Fecha del informe", "Prepared_By": "Preparado por",
        "Root_Cause_Occ": "Causa raíz (Ocurrencia)", "Root_Cause_Det": "Causa raíz (Detección)", "Root_Cause_Sys": "Causa raíz (Sistémica)",
        "Occurrence_Why": "Por qué Ocurrencia", "Detection_Why": "Por qué Detección", "Systemic_Why": "Por qué Sistémico",
        "Save": "💾 Guardar Informe 8D", "Download": "📥 Descargar XLSX",
        "Training_Guidance": "Guía de Entrenamiento", "Example": "Ejemplo",
        "FMEA_Failure": "Ocurrencia de falla FMEA",
        "Location": "Ubicación del material", "Status": "Estado de la actividad", "Containment_Actions": "Acciones de contención"
    }
}

# ---------------------------
# Initialize session state
# ---------------------------
default_whys = [""]*5
for step in ["D1","D2","D3","D4","D5","D6","D7","D8"]:
    if step not in st.session_state:
        st.session_state[step] = {"answer": "", "extra": ""}
st.session_state.setdefault("report_date", datetime.datetime.today().strftime("%B %d, %Y"))
st.session_state.setdefault("prepared_by", "")
st.session_state.setdefault("d5_occ_whys", default_whys.copy())
st.session_state.setdefault("d5_det_whys", default_whys.copy())
st.session_state.setdefault("d5_sys_whys", default_whys.copy())
st.session_state.setdefault("d4_location", "")
st.session_state.setdefault("d4_status", "")
st.session_state.setdefault("d4_containment", "")

# ---------------------------
# NPQP 8D Steps
# ---------------------------
npqp_steps = [
    ("D1", {"en":"Describe the customer concerns clearly.","es":"Describa claramente las preocupaciones del cliente."},
           {"en":"Customer reported static noise in amplifier during end-of-line test.","es":"El cliente reportó ruido estático en el amplificador durante la prueba final."}),
    ("D2", {"en":"Check for similar parts, models, generic parts, other colors, etc.","es":"Verifique partes similares, modelos, partes genéricas, otros colores, etc."},
           {"en":"Similar model radio, Front vs. rear speaker.","es":"Radio de modelo similar, altavoz delantero vs trasero."}),
    ("D3", {"en":"Perform an initial investigation to identify obvious issues.","es":"Realice una investigación inicial para identificar problemas evidentes."},
           {"en":"Visual inspection of solder joints, initial functional tests.","es":"Inspección visual de soldaduras, pruebas funcionales iniciales."}),
    ("D4", {"en":"Define temporary containment actions and material location.","es":"Defina acciones de contención temporales y ubicación del material."}, {"en":"","es":""}),
    ("D5", {"en":"Use 5-Why analysis to determine the root cause.","es":"Use el análisis de 5 Porqués para determinar la causa raíz."}, {"en":"","es":""}),
    ("D6", {"en":"Define corrective actions that eliminate the root cause permanently.","es":"Defina acciones correctivas que eliminen la causa raíz permanentemente."},
           {"en":"Update soldering process, redesign fixture.","es":"Actualizar proceso de soldadura, rediseñar herramienta."}),
    ("D7", {"en":"Verify that corrective actions effectively resolve the issue.","es":"Verifique que las acciones correctivas resuelvan efectivamente el problema."},
           {"en":"Functional tests on corrected amplifiers.","es":"Pruebas funcionales en amplificadores corregidos."}),
    ("D8", {"en":"Document lessons learned, update standards, FMEAs.","es":"Documente lecciones aprendidas, actualice estándares, FMEAs."},
           {"en":"Update SOPs, PFMEA, work instructions.","es":"Actualizar SOPs, PFMEA, instrucciones de trabajo."})
]

# ---------------------------
# Root cause suggestion helper
# ---------------------------
def suggest_root_cause(whys):
    text = " ".join(whys).lower()
    if any(word in text for word in ["training","knowledge","human error"]):
        return "Lack of proper training / knowledge gap"
    if any(word in text for word in ["equipment","tool","machine","fixture"]):
        return "Equipment, tooling, or maintenance issue"
    if any(word in text for word in ["procedure","process","standard"]):
        return "Procedure or process not followed or inadequate"
    if any(word in text for word in ["communication","information","handover"]):
        return "Poor communication or unclear information flow"
    if any(word in text for word in ["material","supplier","component","part"]):
        return "Material, supplier, or logistics-related issue"
    if any(word in text for word in ["design","specification","drawing"]):
        return "Design or engineering issue"
    if any(word in text for word in ["management","supervision","resource"]):
        return "Management or resource-related issue"
    if any(word in text for word in ["temperature","humidity","contamination","environment"]):
        return "Environmental or external factor"
    return "Systemic issue identified from analysis"

# ---------------------------
# Sidebar: JSON Save / Restore / Reset
# ---------------------------
st.sidebar.markdown("---")
st.sidebar.header("Backup / Restore / Reset")

def generate_json():
    save_data = {k:v for k,v in st.session_state.items() if not k.startswith("_")}
    return json.dumps(save_data, indent=4)

st.sidebar.download_button(
    label="💾 Save Progress (JSON)",
    data=generate_json(),
    file_name=f"8D_Report_Backup_{st.session_state.report_date.replace(' ','_')}.json",
    mime="application/json"
)

uploaded_file = st.sidebar.file_uploader("Upload JSON file to restore", type="json")
if uploaded_file:
    try:
        restore_data = json.load(uploaded_file)
        for k,v in restore_data.items():
            st.session_state[k] = v
        st.success("✅ Session restored from JSON!")
    except Exception as e:
        st.error(f"Error restoring JSON: {e}")

if st.sidebar.button("🧹 Reset Session"):
    preserve_keys = ["lang","lang_key"]
    preserved = {k: st.session_state[k] for k in preserve_keys if k in st.session_state}
    for key in list(st.session_state.keys()):
        if key not in preserve_keys and not key.startswith("_"):
            try:
                del st.session_state[key]
            except KeyError:
                pass
    for k,v in preserved.items():
        st.session_state[k] = v
    st.experimental_rerun()

# ---------------------------
# Render D1-D8 tabs
# ---------------------------
tab_labels = []
for step, _, _ in npqp_steps:
    tab_labels.append(f"{t[lang_key][step]}")

tabs = st.tabs(tab_labels)

for i, (step, note_dict, example_dict) in enumerate(npqp_steps):
    with tabs[i]:
        st.markdown(f"### {t[lang_key][step]}")
        st.markdown(f"<div style='background-color:#b3e0ff;color:black;padding:12px;border-left:5px solid #1E90FF;border-radius:6px;'>{note_dict[lang_key]}<br><br>💡 {example_dict[lang_key]}</div>", unsafe_allow_html=True)
        if step == "D4":
            st.session_state[step]["location"] = st.selectbox("Location of Material", ["","Work in Progress","Stores Stock","Warehouse Stock","Service Parts","Other"], index=0)
            st.session_state[step]["status"] = st.selectbox("Status of Activities", ["","Pending","In Progress","Completed","Other"], index=0)
            st.session_state[step]["answer"] = st.text_area("Containment Actions / Notes", value=st.session_state[step]["answer"])
        elif step == "D5":
            st.markdown("#### Occurrence Analysis")
            for idx in range(len(st.session_state.d5_occ_whys)):
                st.session_state.d5_occ_whys[idx] = st.text_input(f"{t[lang_key]['Occurrence_Why']} {idx+1}", value=st.session_state.d5_occ_whys[idx])
            st.markdown("#### Detection Analysis")
            for idx in range(len(st.session_state.d5_det_whys)):
                st.session_state.d5_det_whys[idx] = st.text_input(f"{t[lang_key]['Detection_Why']} {idx+1}", value=st.session_state.d5_det_whys[idx])
            st.markdown("#### Systemic Analysis")
            for idx in range(len(st.session_state.d5_sys_whys)):
                st.session_state.d5_sys_whys[idx] = st.text_input(f"{t[lang_key]['Systemic_Why']} {idx+1}", value=st.session_state.d5_sys_whys[idx])
            occ_rc_text = suggest_root_cause([w for w in st.session_state.d5_occ_whys if w.strip()])
            det_rc_text = suggest_root_cause([w for w in st.session_state.d5_det_whys if w.strip()])
            sys_rc_text = suggest_root_cause([w for w in st.session_state.d5_sys_whys if w.strip()])
            st.text_area(f"{t[lang_key]['Root_Cause_Occ']}", value=occ_rc_text, height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Det']}", value=det_rc_text, height=80, disabled=True)
            st.text_area(f"{t[lang_key]['Root_Cause_Sys']}", value=sys_rc_text, height=80, disabled=True)
        else:
            st.session_state[step]["answer"] = st.text_area("Your Answer", value=st.session_state[step]["answer"])

# ---------------------------
# Excel generation
# ---------------------------
def generate_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "NPQP 8D Report"
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin,right=thin,top=thin,bottom=thin)
    if os.path.exists("logo.png"):
        try:
            img = XLImage("logo.png")
            img.width = 140
            img.height = 40
            ws.add_image(img, "A1")
        except:
            pass
    ws.merge_cells(start_row=3,start_column=1,end_row=3,end_column=3)
    ws.cell(row=3,column=1,value="📋 8D Report Assistant").font = Font(bold=True,size=14)
    ws.append([t[lang_key]['Report_Date'], st.session_state.report_date])
    ws.append([t[lang_key]['Prepared_By'], st.session_state.prepared_by])
    ws.append([])
    # Header
    headers = ["Step","Answer","Extra / Notes"]
    fill = PatternFill(start_color="1E90FF", end_color="1E90FF", fill_type="solid")
    for c_idx,h in enumerate(headers,start=1):
        cell = ws.cell(row=ws.max_row+1,column=c_idx,value=h)
        cell.fill = fill
        cell.font = Font(bold=True,color="FFFFFF")
        cell.alignment = Alignment(horizontal="center",vertical="center")
        cell.border = border
    # Step answers
    for step, _, _ in npqp_steps:
        answer = st.session_state[step]["answer"]
        extra = ""
        if step=="D4":
            extra = f"Location: {st.session_state[step]['location']} | Status: {st.session_state[step]['status']}"
        elif step=="D5":
            extra = f"Occurrence: {' | '.join(st.session_state.d5_occ_whys)}; Detection: {' | '.join(st.session_state.d5_det_whys)}; Systemic: {' | '.join(st.session_state.d5_sys_whys)}"
        ws.append([t[lang_key][step], answer, extra])
        for col in range(1,4):
            cell = ws.cell(row=ws.max_row,column=col)
            cell.alignment = Alignment(wrap_text=True,vertical="top")
            cell.font = Font(bold=True if col==2 else False)
            cell.border = border
    for col in range(1,4):
        ws.column_dimensions[get_column_letter(col)].width = 40
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

st.download_button(
    label=f"{t[lang_key]['Download']}",
    data=generate_excel(),
    file_name=f"8D_Report_{st.session_state.report_date.replace(' ','_')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
