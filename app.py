# app.py - Full 8D NPQP app with bilingual guidance, interactive 5-Why, cached translation, AI helper, and Excel export.
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import datetime
import openai

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(page_title="8D Training App", page_icon="📑", layout="wide")

# Hide Streamlit chrome
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Language selector
# ---------------------------
lang_choice = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"], index=0)
lang = "en" if lang_choice == "English" else "es"
st.session_state.setdefault("prev_lang", lang)

# ---------------------------
# Translations caching & helper (uses OpenAI if key present)
# ---------------------------
st.session_state.setdefault("translations", {})  # cache: key -> translated_text

def translate_cached(text, src_target, field_key=None):
    """
    src_target: tuple like ("en","es") meaning translate from en->es
    field_key: unique identifier for caching
    """
    if not text or not text.strip():
        return text
    cache_key = f"{field_key}_{src_target[0]}_{src_target[1]}" if field_key else f"default_{src_target[0]}_{src_target[1]}_{hash(text)}"
    if cache_key in st.session_state["translations"]:
        return st.session_state["translations"][cache_key]

    api_key = st.secrets.get("OPENAI_API_KEY", "")
    if not api_key:
        # no API key: don't translate, return original
        return text

    openai.api_key = api_key
    # keep prompt simple and safe for translation
    prompt = f"Translate the following text from {'English' if src_target[0]=='en' else 'Spanish'} to {'English' if src_target[1]=='en' else 'Spanish'}. Keep technical terms and lists intact.\n\n{text}"
    try:
        resp = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[{"role":"user","content":prompt}],
            temperature=0.0,
            max_tokens=800
        )
        translated = resp.choices[0].message.content.strip()
        st.session_state["translations"][cache_key] = translated
        return translated
    except Exception as e:
        # translation failed: return original
        st.warning(f"Translation error (field {field_key}): {e}")
        return text

def maybe_translate_all(prev_lang, new_lang):
    """Translate all stored fields once on language switch and update widget keys so inputs refresh."""
    if prev_lang == new_lang:
        return
    src_target = (prev_lang, new_lang)
    api_key = st.secrets.get("OPENAI_API_KEY", "")
    # If no API key, skip translation but still update prev_lang so we don't loop
    if not api_key:
        st.session_state["prev_lang"] = new_lang
        return

    # Steps canonical IDs
    step_ids = ["D1","D2","D3","D4","D5","D6","D7","D8"]

    # Translate D1-D8 answer and extra fields
    for sid in step_ids:
        for field in ("answer","extra"):
            text = st.session_state[sid][field]
            if text and text.strip():
                translated = translate_cached(text, src_target, field_key=f"{sid}_{field}")
                # update both canonical storage and widget default key
                st.session_state[sid][field] = translated
                # corresponding widget keys:
                if field == "answer":
                    st.session_state[f"ans_{sid}"] = translated
                else:
                    st.session_state[f"{sid}_extra_widget"] = translated

    # Translate D5 occurrence/detection whys
    occ = st.session_state.get("d5_occ_whys", [])
    det = st.session_state.get("d5_det_whys", [])
    for i, t in enumerate(occ):
        if t and t.strip():
            tr = translate_cached(t, src_target, field_key=f"d5_occ_{i}")
            st.session_state["d5_occ_whys"][i] = tr
            st.session_state[f"d5_occ_{i}"] = tr
    for i, t in enumerate(det):
        if t and t.strip():
            tr = translate_cached(t, src_target, field_key=f"d5_det_{i}")
            st.session_state["d5_det_whys"][i] = tr
            st.session_state[f"d5_det_{i}"] = tr

    # Translate interactive whys (if any)
    for i, t in enumerate(st.session_state.get("interactive_whys", [])):
        if t and t.strip():
            tr = translate_cached(t, src_target, field_key=f"interactive_{i}")
            st.session_state["interactive_whys"][i] = tr
            st.session_state[f"interactive_{i}"] = tr

    # Translate interactive root cause
    ir = st.session_state.get("interactive_root_cause", "")
    if ir and ir.strip():
        tr = translate_cached(ir, src_target, field_key="interactive_root_cause")
        st.session_state["interactive_root_cause"] = tr
        st.session_state["interactive_root_widget"] = tr

    # update prev_lang
    st.session_state["prev_lang"] = new_lang

# run translation on language switch
maybe_translate_all(st.session_state.get("prev_lang", lang), lang)

# ---------------------------
# UI text/guidelines (bilingual) for each step
# ---------------------------
guidelines = {
    "D1": {
        "en": """**D1: Concern Details**
- State the customer concern clearly.
- Include when/where/how it was observed and any evidence (photos, test logs).
- Include customer part number, serial, and lot if available.""",
        "es": """**D1: Detalles de la Preocupación**
- Indique claramente la preocupación del cliente.
- Incluya cuándo/dónde/cómo se observó y evidencia (fotos, registros de prueba).
- Incluya número de pieza del cliente, serie y lote si está disponible."""
    },
    "D2": {
        "en": """**D2: Similar Part Considerations**
- Check similar parts, alternate colors, mirror-hand units, and other models.
- Determine if issue is isolated to a lot or recurring across variants.""",
        "es": """**D2: Consideraciones de Piezas Similares**
- Revise piezas similares, colores alternos, unidades espejo y otros modelos.
- Determine si el problema es aislado a un lote o recurrente en variantes."""
    },
    "D3": {
        "en": """**D3: Initial Analysis**
- Capture initial inspection findings (visual, functional).
- Collect process data, test results, operator, machine, lot numbers.""",
        "es": """**D3: Análisis Inicial**
- Registre hallazgos de inspección inicial (visual, funcional).
- Recoja datos de proceso, resultados de prueba, operador, máquina, números de lote."""
    },
    "D4": {
        "en": """**D4: Implement Containment**
- Define temporary actions to prevent customer exposure.
- Include labeling, quarantine, 100% inspection, or special testing.""",
        "es": """**D4: Implementar Contención**
- Defina acciones temporales para evitar la exposición del cliente.
- Incluya etiquetado, cuarentena, inspección 100% o pruebas especiales."""
    },
    "D5": {
        "en": """**D5: Final Analysis (5-Whys)**
- Use separate Occurrence and Detection tracks for 5-Why analysis.
- Keep each Why focused and causal (avoid solutions as whys).
- Document root cause summary after completing whys.""",
        "es": """**D5: Análisis Final (5-Whys)**
- Use pistas separadas de Ocurrencia y Detección para el análisis 5-Why.
- Mantenga cada porqué enfocado y causal (evite soluciones como porqués).
- Documente un resumen de la causa raíz después de completar los porqués."""
    },
    "D6": {
        "en": """**D6: Permanent Corrective Actions**
- Define actions that eliminate the root cause.
- Include owner, target completion date, and verification steps.""",
        "es": """**D6: Acciones Correctivas Permanentes**
- Defina acciones que eliminen la causa raíz.
- Incluya responsable, fecha objetivo y pasos de verificación."""
    },
    "D7": {
        "en": """**D7: Countermeasure Confirmation**
- Verify the permanent actions are effective over time.
- Use test plans, sampling, or metrics monitoring.""",
        "es": """**D7: Confirmación de Contramedidas**
- Verifique que las acciones permanentes sean efectivas en el tiempo.
- Use planes de prueba, muestreo o monitoreo de métricas."""
    },
    "D8": {
        "en": """**D8: Follow-up / Lessons Learned**
- Capture lessons, update PFMEA/SOPs, and train impacted teams.
- Close the loop with documentation and preventive controls.""",
        "es": """**D8: Seguimiento / Lecciones Aprendidas**
- Capture lecciones, actualice PFMEA/SOPs y capacite equipos.
- Cierre el ciclo con documentación y controles preventivos."""
    }
}

# ---------------------------
# NPQP steps canonical
# ---------------------------
steps_order = ["D1","D2","D3","D4","D5","D6","D7","D8"]

# Initialize any missing st.session_state keys for widget-binding
for sid in steps_order:
    st.session_state.setdefault(f"ans_{sid}", st.session_state[sid]["answer"])
    st.session_state.setdefault(f"{sid}_extra_widget", st.session_state[sid]["extra"])

# Also widget keys for d5 whys and interactive
for i in range(8):  # keep 8 possible interactive whys
    st.session_state.setdefault(f"interactive_{i}", st.session_state.get("interactive_whys", [""]*8)[i] if i < len(st.session_state.get("interactive_whys",[])) else "")
for i in range(8):
    st.session_state.setdefault(f"d5_occ_{i}", st.session_state.get("d5_occ_whys", [""]*8)[i] if i < len(st.session_state.get("d5_occ_whys",[])) else "")
    st.session_state.setdefault(f"d5_det_{i}", st.session_state.get("d5_det_whys", [""]*8)[i] if i < len(st.session_state.get("d5_det_whys",[])) else "")

st.session_state.setdefault("interactive_whys", st.session_state.get("interactive_whys", [""]))
st.session_state.setdefault("interactive_root_cause", st.session_state.get("interactive_root_cause", ""))
st.session_state.setdefault("interactive_root_widget", st.session_state.get("interactive_root_cause",""))

# ---------------------------
# Header / Report info
# ---------------------------
st.subheader(t["app_title"])
st.markdown("")  # small gap
st.subheader(t["report_info"])
st.session_state.report_date = st.text_input(t["report_date"], value=st.session_state.report_date)
st.session_state.prepared_by = st.text_input(t["prepared_by"], value=st.session_state.prepared_by)

# ---------------------------
# Tabs for each step
# ---------------------------
tabs = st.tabs([f"{sid}" for sid in steps_order])
for i, sid in enumerate(steps_order):
    title = sid
    with tabs[i]:
        # show bilingual guideline
        st.markdown(guidelines[sid][lang])
        st.markdown("---")

        if sid != "D5":
            # Main answer (text_area). Use widget key 'ans_{sid}' to reflect translations.
            st.session_state[sid]["answer"] = st.text_area(
                f"Your Answer for {sid}",
                value=st.session_state.get(f"ans_{sid}", st.session_state[sid]["answer"]),
                key=f"ans_{sid}",
                height=180
            )
            # Extra / Root cause notes for each non-D5 step (keeps parity)
            st.session_state[sid]["extra"] = st.text_area(
                f"Root Cause / Extra Notes for {sid}",
                value=st.session_state.get(f"{sid}_extra_widget", st.session_state[sid]["extra"]),
                key=f"{sid}_extra_widget",
                height=120
            )
        else:
            # D5: Occurrence + Detection interactive whys + interactive chain + AI helper
            st.markdown("#### Occurrence Analysis (5-Why - Occurrence)")
            # show 5 occurrence whys (but interactive - show up to 5 or allow add)
            occ_list = st.session_state.get("d5_occ_whys", [""] * 5)
            # ensure length
            if len(occ_list) < 5:
                occ_list += [""] * (5 - len(occ_list))
            for idx in range(5):
                # show field with value coming from widget key to allow refresh after translation
                val = st.text_input(f"Occurrence Why {idx+1}", value=st.session_state.get(f"d5_occ_{idx}", occ_list[idx]), key=f"d5_occ_{idx}")
                # save back to canonical list
                occ_list[idx] = val
            st.session_state["d5_occ_whys"] = occ_list
            if st.button(t["add_occ"], key="add_occ_button"):
                st.session_state.d5_occ_whys.append("")
                st.experimental_rerun()

            st.markdown("#### Detection Analysis (5-Why - Detection)")
            det_list = st.session_state.get("d5_det_whys", [""] * 5)
            if len(det_list) < 5:
                det_list += [""] * (5 - len(det_list))
            for idx in range(5):
                val = st.text_input(f"Detection Why {idx+1}", value=st.session_state.get(f"d5_det_{idx}", det_list[idx]), key=f"d5_det_{idx}")
                det_list[idx] = val
            st.session_state["d5_det_whys"] = det_list
            if st.button(t["add_det"], key="add_det_button"):
                st.session_state.d5_det_whys.append("")
                st.experimental_rerun()

            # Combine into D5 answer like your original app
            st.session_state["D5"]["answer"] = (
                "Occurrence Analysis:\n" + "\n".join([w for w in st.session_state["d5_occ_whys"] if w.strip()]) +
                "\n\nDetection Analysis:\n" + "\n".join([w for w in st.session_state["d5_det_whys"] if w.strip()])
            )

            # Root cause summary field (D5 extra)
            st.session_state["D5"]["extra"] = st.text_area(
                "Root Cause (summary after 5-Whys)",
                value=st.session_state["D5"].get("extra", ""),
                key="d5_root_text",
                height=120
            )

            # --- Interactive 5-Why chain (separate from occurrence/detection) ---
            st.markdown("### Interactive 5-Why (guided chain)")
            # keep list length
            inter = st.session_state.get("interactive_whys", [""])
            # show interactive chain: show Why 1 always, Why n appears only if Why n-1 filled
            max_whys = 8
            for idx in range(max_whys):
                show_field = False
                if idx == 0:
                    show_field = True
                else:
                    prev = inter[idx-1] if idx-1 < len(inter) else ""
                    show_field = bool(prev and prev.strip())
                current_val = inter[idx] if idx < len(inter) else ""
                if show_field:
                    val = st.text_input(f"Why {idx+1}", value=st.session_state.get(f"interactive_{idx}", current_val), key=f"interactive_{idx}")
                    # ensure list is long enough
                    if idx >= len(inter):
                        inter.extend([""]*(idx - len(inter) + 1))
                    inter[idx] = val
            # trim trailing empties
            while len(inter) > 1 and not inter[-1].strip():
                inter.pop()
            st.session_state["interactive_whys"] = inter

            # interactive root cause summary (editable)
            st.session_state["interactive_root_cause"] = st.text_area(
                "Interactive Root Cause Summary (editable)",
                value=st.session_state.get("interactive_root_cause", ""),
                key="interactive_root_widget",
                height=120
            )

            # AI Helper section (optional, does not overwrite fields)
            st.markdown(f"### {t['ai_helper']}")
            if st.button(t["ai_btn"], key="ai_generate"):
                api_key = st.secrets.get("OPENAI_API_KEY", "")
                if not api_key:
                    st.warning("OpenAI API key missing in Streamlit secrets — AI suggestions disabled.")
                else:
                    openai.api_key = api_key
                    # Compose prompt using current interactive whys + occurrence/detection
                    prompt_lang = "English" if lang == "English" else "Spanish"
                    occ_text = "\n".join([w for w in st.session_state["d5_occ_whys"] if w.strip()])
                    det_text = "\n".join([w for w in st.session_state["d5_det_whys"] if w.strip()])
                    inter_text = "\n".join([w for w in st.session_state["interactive_whys"] if w.strip()])
                    prompt = f"""You are an expert NPQP 8D analyst. Language: {prompt_lang}.
Given the following information, suggest up to 3 additional 'Why' questions for Occurrence and Detection (if helpful),
and provide a concise Root Cause summary and 2 suggested permanent corrective actions.

Occurrence Whys:
{occ_text}

Detection Whys:
{det_text}

Interactive Why chain:
{inter_text}

Respond in {prompt_lang}."""
                    try:
                        response = openai.ChatCompletion.create(
                            model="gpt-4",
                            messages=[{"role":"user","content":prompt}],
                            temperature=0.3,
                            max_tokens=500
                        )
                        ai_out = response.choices[0].message.content
                        st.text_area(t["ai_output"], value=ai_out, height=300, key="ai_output_widget")
                    except Exception as e:
                        st.error(f"AI call failed: {e}")

# ---------------------------
# Excel Export (Save & Download)
# ---------------------------
if st.button("💾 Save 8D Report"):
    # collect rows using canonical step titles and current answers
    data_rows = []
    for sid in steps_order:
        ans = st.session_state[sid]["answer"]
        extra = st.session_state[sid]["extra"]
        data_rows.append((sid, ans, extra))
    # append Interactive 5-Why as extra row
    inter_text = "\n".join([w for w in st.session_state.get("interactive_whys", []) if w.strip()])
    inter_extra = st.session_state.get("interactive_root_cause", "")
    data_rows.append(("Interactive 5-Why", inter_text, inter_extra))

    if not any(row[1].strip() or row[2].strip() for row in data_rows):
        st.error("⚠️ No answers filled in yet. Please complete some fields before saving.")
    else:
        fname = "NPQP_8D_Report.xlsx"
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
        headers = ["Step", "Your Answer", "Root Cause / Extra"]
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        r = 6
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=c, value=h)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = header_fill

        # Content rows
        r = 7
        for step, ans, extra in data_rows:
            ws.cell(row=r, column=1, value=step)
            ws.cell(row=r, column=2, value=ans)
            ws.cell(row=r, column=3, value=extra)
            # color by step prefix if available
            prefix = step if step.startswith("D") else step[:2]
            fill_color = {
                "D1":"ADD8E6","D2":"90EE90","D3":"FFFF99","D4":"FFD580",
                "D5":"FF9999","D6":"D8BFD8","D7":"E0FFFF","D8":"D3D3D3"
            }.get(prefix, "FFFFFF")
            for c in range(1,4):
                ws.cell(row=r, column=c).fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
            r += 1

        for c in range(1,4):
            ws.column_dimensions[get_column_letter(c)].width = 40

        wb.save(fname)
        st.success("✅ NPQP 8D Report saved successfully.")
        with open(fname, "rb") as f:
            st.download_button("📥 Download XLSX", f, file_name=fname)

# ---------------------------
# End of app
# ---------------------------
