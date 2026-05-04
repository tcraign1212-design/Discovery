import streamlit as st
import io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
from openai import OpenAI
import anthropic
import os

# ──────────────────────────────────────────────
# 1. INITIALIZE SESSION STATE (The "Memory" Sieve)
# ──────────────────────────────────────────────
if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""
if "comm_status" not in st.session_state:
    st.session_state["comm_status"] = "No"

# ──────────────────────────────────────────────
# 2. PAGE CONFIGURATION
# ──────────────────────────────────────────────
st.set_page_config(page_title="Case Intelligence Brief Generator", layout="wide")

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("*Strategic case framing for pre-litigation intake and active litigation workup*")
st.markdown("---")

# API Keys from Secrets/Env
env_gemini_key    = st.secrets.get("GEMINI_API_KEY")    or os.environ.get("GEMINI_API_KEY")
env_openai_key    = st.secrets.get("OPENAI_API_KEY")    or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# ──────────────────────────────────────────────
# 3. DOCUMENT GENERATION (Texas Format)
# ──────────────────────────────────────────────
def generate_brief_docx(brief_text: str, doc_title: str) -> io.BytesIO:
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Inches(1), Inches(1)
        section.left_margin, section.right_margin = Inches(1.25), Inches(1.25)

    style = doc.styles["Normal"]
    style.font.name, style.font.size = "Times New Roman", Pt(12)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(doc_title)
    run_title.font.bold, run_title.font.size = True, Pt(13)

    doc.add_paragraph("_" * 70)
    doc.add_paragraph()

    for line in brief_text.split("\n"):
        cleaned = line.strip()
        if not cleaned:
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        if cleaned.startswith("### "):
            run = p.add_run(cleaned.replace("### ", "").strip())
            run.font.bold, run.font.underline = True, True
        elif cleaned.startswith("## "):
            run = p.add_run(cleaned.replace("## ", "").strip())
            run.font.bold = True
            p.paragraph_format.space_before = Pt(10)
        elif cleaned.startswith(("- ", "* ")):
            p.paragraph_format.left_indent = Inches(0.25)
            p.add_run(cleaned[2:].strip())
        else:
            p.add_run(cleaned)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ──────────────────────────────────────────────
# 4. PROMPT BUILDER (Operational Strategy)
# ──────────────────────────────────────────────
def build_prompt(stage, ctype, dlevel, doi, sol, summary, gov, comm_analysis):
    gov_txt = "\n- GOVERNMENT ENTITY: Apply TTCA analysis and notice deadlines." if gov else ""
    cv_txt = "\n- COMMERCIAL VEHICLE: Identify carrier safety protocols, driver files, and preservation needs." if comm_analysis else ""
    
    params = f"Framework: {ctype}\nDOI: {doi}\nSOL: {sol}\nGov Entity: {gov}\nComm Analysis: {comm_analysis}\n\nSummary:\n{summary}"
    
    if stage == "Pre-Litigation":
        return f"Draft a Texas Pre-Suit Brief. {params} {gov_txt}{cv_txt} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Risk Flags, ## 4. Proof Gaps, ## 5. Defense Anticipation, ## 6. Action Items."
    return f"Draft a Texas Litigation Blueprint. {params} Discovery Level: {dlevel} {gov_txt}{cv_txt} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Proof Gaps, ## 4. Defense Anticipation, ## 5. Discovery Blueprint, ## 6. Strategic Flags."

# ──────────────────────────────────────────────
# 5. MAIN LAYOUT (Revised for Persistence)
# ──────────────────────────────────────────────
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)

    # Use a unique key for the API input to force persistence
    st.text_input(
        f"{ai_engine} API Key", 
        type="password", 
        key="active_key_input",
        value=env_gemini_key if ai_engine == "Gemini (Google)" else (env_openai_key if ai_engine == "ChatGPT (OpenAI)" else env_anthropic_key)
    )

    st.markdown("---")
    case_type = st.selectbox("Framework:", ["Standard MVA", "Trucking", "Premises", "Workplace", "UM/UIM", "TTCA"])
    
    discovery_level = "N/A"
    if case_stage == "Active Litigation":
        discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True)

    st.markdown("**Risk Flags**")
    government_entity = st.checkbox("Government Entity Involved")
    
    # Binding the commercial status to a key prevents it from disappearing
    st.radio(
        "Commercial Vehicle Involved?", 
        ["No", "Yes", "Unsure"], 
        key="comm_status_radio",
        horizontal=True
    )
    
    case_summary = st.text_area("Case Summary", height=160)
    
    with st.expander("Dates"):
        doi = st.text_input("Incident Date", placeholder="YYYY-MM-DD")
        sol = st.text_input("SOL Date", placeholder="YYYY-MM-DD")

    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)

# ──────────────────────────────────────────────
# 6. OUTPUT PANEL (Execution)
# ──────────────────────────────────────────────
with col2:
    st.subheader("2. Case Intelligence Brief")

    # Pulling values from session state to ensure they exist during rerun
    current_key = st.session_state.get("active_key_input", "")
    comm_flag = st.session_state.get("comm_status_radio", "No") in ["Yes", "Unsure"]

    if run_brief:
        if not case_summary.strip():
            st.warning("Veto: Case summary required.")
        elif not current_key:
            st.error("Engine Error: Missing API Key.")
        else:
            prompt = build_prompt(case_stage, case_type, discovery_level, doi, sol, case_summary, government_entity, comm_flag)
            
            try:
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Gemini thinking..."):
                        genai.configure(api_key=current_key)
                        model = genai.GenerativeModel("gemini-1.5-flash")
                        st.session_state["brief_content"] = model.generate_content(prompt).text
                # ... [Keep OpenAI/Claude blocks similar, using current_key] ...
            except Exception as e:
                st.error(f"Integrity Error: {e}")

    # Display & Export
    if st.session_state.get("brief_content"):
        st.text_area("Edit Brief:", value=st.session_state["brief_content"], height=500, key="brief_editor")
        # Update content from editor key
        st.session_state["brief_content"] = st.session_state["brief_editor"]
        
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        st.download_button("📥 Download (.docx)", data=docx_data, file_name="Case_Brief.docx", use_container_width=True)
