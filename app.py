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
# 1. PAGE CONFIG & SESSION STATE
# ──────────────────────────────────────────────
st.set_page_config(page_title="Case Intelligence Brief", layout="wide")

if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("---")

# API Keys
env_gemini_key    = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
env_openai_key    = st.secrets.get("OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# ──────────────────────────────────────────────
# 2. DOCUMENT GENERATION (Texas Standards)
# ──────────────────────────────────────────────
def generate_brief_docx(text, title):
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = Inches(1)
        s.left_margin = s.right_margin = Inches(1.25)
    
    style = doc.styles["Normal"]
    style.font.name, style.font.size = "Times New Roman", Pt(12)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.font.bold, run.font.size = True, Pt(13)
    
    doc.add_paragraph("_" * 70)
    
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        if line.startswith("### "):
            r = p.add_run(line.replace("### ", ""))
            r.font.bold, r.font.underline = True, True
        elif line.startswith("## "):
            r = p.add_run(line.replace("## ", ""))
            r.font.bold = True
        else:
            p.add_run(line)
            
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ──────────────────────────────────────────────
# 3. PROMPT BUILDER
# ──────────────────────────────────────────────
def build_prompt(stage, ctype, dlevel, doi, sol, summary, gov, cv_jurisdiction):
    gov_txt = "\n- GOV FLAG: Apply TTCA analysis/notice deadlines." if gov else ""
    
    # Restored logic for Commercial Vehicle / Trucking flags
    cv_analysis = ""
    if ctype in ["Trucking", "Commercial Vehicle"]:
        jurisdiction = f"Jurisdiction: {cv_jurisdiction}"
        cv_analysis = f"\n- {ctype.upper()} FLAG: Apply {jurisdiction} analysis. Focus on driver logs, qualification files, and vehicle maintenance under relevant FMCSR/TX-DOT standards."
    
    params = f"Framework: {ctype}\nDOI: {doi}\nSOL: {sol}\nGov: {gov}\n\nSummary:\n{summary}"
    
    if stage == "Pre-Litigation":
        return f"Draft Texas Pre-Suit Brief. {params} {gov_txt}{cv_analysis} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Risk Flags, ## 4. Proof Gaps, ## 5. Defense Anticipation, ## 6. Action Items."
    return f"Draft Texas Litigation Blueprint. {params} Discovery: {dlevel} {gov_txt}{cv_analysis} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Proof Gaps, ## 4. Defense Anticipation, ## 5. Discovery Blueprint, ## 6. Strategic Flags."

# ──────────────────────────────────────────────
# 4. MAIN LAYOUT
# ──────────────────────────────────────────────
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)

    # Key Persistence
    if ai_engine == "Gemini (Google)":
        active_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password")
    else:
        active_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password")

    st.markdown("---")
    
    # ADDED "Commercial Vehicle" TO THE DROPDOWN LIST HERE
    case_type = st.selectbox(
        "Framework:", 
        ["Standard MVA", "Commercial Vehicle", "Trucking", "Premises", "Workplace", "UM/UIM", "TTCA"]
    )
    
    discovery_level = "Level 2"
    if case_stage == "Active Litigation":
        discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True)

    st.markdown("**Risk Flags**")
    government_entity = st.checkbox("Government Entity Involved")
    
    # REMOVED the separate commercial status radio button to keep the UI clean

    case_summary = st.text_area("Case Summary", height=160)
    
    with st.expander("Dates"):
        date_of_incident = st.text_input("Incident Date", placeholder="YYYY-MM-DD")
        sol_date = st.text_input("SOL Date", placeholder="YYYY-MM-DD")

    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)

# ──────────────────────────────────────────────
# 5. OUTPUT PANEL
# ──────────────────────────────────────────────
with col2:
    st.subheader("2. Case Intelligence Brief")

    if run_brief:
        if not case_summary.strip():
            st.warning("Veto: Provide a case summary.")
        elif not active_key:
            st.error("Missing API Key.")
        else:
            prompt = build_prompt(case_stage, case_type, discovery_level, date_of_incident, sol_date, case_summary, government_entity, include_comm)
            
            try:
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Gemini thinking..."):
                        genai.configure(api_key=active_key)
                        model = genai.GenerativeModel("gemini-1.5-flash")
                        st.session_state["brief_content"] = model.generate_content(prompt).text
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("OpenAI thinking..."):
                        client = OpenAI(api_key=active_key)
                        resp = client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": prompt}])
                        st.session_state["brief_content"] = resp.choices[0].message.content
                elif ai_engine == "Claude (Anthropic)":
                    with st.spinner("Claude thinking..."):
                        client = anthropic.Anthropic(api_key=active_key)
                        resp = client.messages.create(model="claude-3-5-sonnet-20240620", max_tokens=4000, messages=[{"role": "user", "content": prompt}])
                        st.session_state["brief_content"] = resp.content[0].text
            except Exception as e:
                st.error(f"Error: {e}")

    if st.session_state["brief_content"]:
        edited = st.text_area("Edit Brief:", value=st.session_state["brief_content"], height=500)
        st.session_state["brief_content"] = edited
        
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        st.download_button("📥 Download (.docx)", data=docx_data, file_name="Case_Brief.docx", use_container_width=True)
