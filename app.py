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
# 1. INITIALIZE SESSION STATE (Prevent Rebooting)
# ──────────────────────────────────────────────
if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""
if "active_api_key" not in st.session_state:
    st.session_state["active_api_key"] = ""

# ──────────────────────────────────────────────
# 2. PAGE CONFIGURATION
# ──────────────────────────────────────────────
st.set_page_config(page_title="Case Intelligence Brief Generator", layout="wide")

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("*Strategic case framing for pre-litigation intake and active litigation workup*")
st.markdown("---")

# API Keys from Secrets
env_gemini_key    = st.secrets.get("GEMINI_API_KEY")    or os.environ.get("GEMINI_API_KEY")
env_openai_key    = st.secrets.get("OPENAI_API_KEY")    or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# ──────────────────────────────────────────────
# 3. DOCUMENT GENERATION
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
# 4. PROMPT BUILDER
# ──────────────────────────────────────────────
def build_prompt(case_stage, case_type, discovery_level, doi, sol, summary, gov, comm):
    gov_txt = "\n- GOVERNMENT ENTITY: Apply TTCA analysis & notice deadlines." if gov else ""
    cv_txt = "\n- COMMERCIAL/FMCSR: Identify carrier safety protocols & preservation." if comm else ""
    
    params = f"Framework: {case_type}\nDOI: {doi}\nSOL: {sol}\nGov: {gov}\nComm: {comm}\n\nSummary:\n{summary}"
    
    if case_stage == "Pre-Litigation":
        return f"Draft a Texas Pre-Suit Brief. {params} {gov_txt}{cv_txt} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Risk Flags, ## 4. Proof Gaps, ## 5. Defense Anticipation, ## 6. Action Items."
    return f"Draft a Texas Litigation Blueprint. {params} Discovery Level: {discovery_level} {gov_txt}{cv_txt} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Proof Gaps, ## 4. Defense Anticipation, ## 5. Discovery Blueprint, ## 6. Strategic Flags."

# ──────────────────────────────────────────────
# 5. MAIN LAYOUT
# ──────────────────────────────────────────────
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)

    # API Key Input
    if ai_engine == "Gemini (Google)":
        st.session_state["active_api_key"] = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password")
    elif ai_engine == "ChatGPT (OpenAI)":
        st.session_state["active_api_key"] = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password")
    else:
        st.session_state["active_api_key"] = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password")

    st.markdown("---")
    case_type = st.selectbox("Framework:", ["Standard MVA", "Trucking", "Premises", "Workplace", "UM/UIM", "TTCA"])
    
    discovery_level = "N/A"
    if case_stage == "Active Litigation":
        discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True)

    st.markdown("**Risk Flags**")
    government_entity = st.checkbox("Government Entity Involved")
    commercial_status = st.radio("Commercial Vehicle?", ["No", "Yes", "Unsure"], index=0, horizontal=True)
    include_commercial_analysis = (commercial_status in ["Yes", "Unsure"])

    case_summary = st.text_area("Case Summary", height=160)
    with st.expander("Dates"):
        date_of_incident = st.text_input("Incident Date", placeholder="YYYY-MM-DD")
        sol_date = st.text_input("SOL Date", placeholder="YYYY-MM-DD")

    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)

# ──────────────────────────────────────────────
# 6. OUTPUT PANEL
# ──────────────────────────────────────────────
with col2:
    st.subheader("2. Case Intelligence Brief")

    if run_brief:
        if not case_summary.strip():
            st.warning("Please provide a case summary.")
        elif not st.session_state["active_api_key"]:
            st.error("Missing API Key.")
        else:
            prompt = build_prompt(case_stage, case_type, discovery_level, date_of_incident, sol_date, case_summary, government_entity, include_commercial_analysis)
            
            try:
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Gemini thinking..."):
                        genai.configure(api_key=st.session_state["active_api_key"])
                        model = genai.GenerativeModel("gemini-1.5-flash")
                        st.session_state["brief_content"] = model.generate_content(prompt).text
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("OpenAI thinking..."):
                        client = OpenAI(api_key=st.session_state["active_api_key"])
                        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": prompt}])
                        st.session_state["brief_content"] = response.choices[0].message.content
                elif ai_engine == "Claude (Anthropic)":
                    with st.spinner("Claude thinking..."):
                        client = anthropic.Anthropic(api_key=st.session_state["active_api_key"])
                        response = client.messages.create(model="claude-3-5-sonnet-20240620", max_tokens=4000, messages=[{"role": "user", "content": prompt}])
                        st.session_state["brief_content"] = response.content[0].text
                
                st.session_state["brief_stage"] = case_stage
            except Exception as e:
                st.error(f"Engine Error: {e}")

    if st.session_state["brief_content"]:
        edited_text = st.text_area("Edit Brief:", value=st.session_state["brief_content"], height=500)
        st.session_state["brief_content"] = edited_text
        
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        st.download_button("📥 Download (.docx)", data=docx_data, file_name="Case_Brief.docx", use_container_width=True)
