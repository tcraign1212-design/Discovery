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
# 5. MAIN LAYOUT
# ──────────────────────────────────────────────
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)

    # API Key Sieve
    if ai_engine == "Gemini (Google)":
        active_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password", key="gem_key")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password", key="gpt_key")
    else:
        active_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password", key="claude_key")

    st.markdown("---")
    case_type = st.selectbox("Framework:", ["Standard MVA", "Trucking", "Premises", "Workplace", "UM/UIM", "TTCA"])
    
    discovery_level = "N/A"
    if case_stage == "Active Litigation":
        discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True)

    st.markdown("**Risk Flags**")
    # 1. Government Entity Check
    government_entity = st.checkbox("Government Entity Involved")
    
    # 2. Universal Commercial Vehicle Check (Moved outside all IF blocks)
    comm_status = st.radio(
        "Commercial Vehicle Involved?", 
        ["No", "Yes", "Unsure"], 
        index=0, 
        horizontal=True,
        key="commercial_flag_permanent"  # Added key for session persistence
    )
    include_comm = (comm_status in ["Yes", "Unsure"])

    case_summary = st.text_area("Case Summary", height=160, placeholder="Detail the incident and liability theory...")
    
    with st.expander("Dates"):
        date_of_incident = st.text_input("Incident Date", placeholder="YYYY-MM-DD")
        sol_date = st.text_input("SOL Date", placeholder="YYYY-MM-DD")

    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)
