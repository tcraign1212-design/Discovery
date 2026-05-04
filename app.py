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

# Ensure brief content persists through reruns
if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("*Strategic case framing for pre-litigation intake and active litigation workup*")
st.markdown("---")

# API Keys from Secrets/Env
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
    gov_txt = "\n- GOV FLAG: Apply TTCA analysis and notice deadlines." if gov else ""
    
    cv_analysis = ""
    if ctype in ["Trucking", "Commercial Vehicle"]:
        jurisdiction = f"Jurisdiction: {cv_jurisdiction}"
        cv_analysis = f"\n- {ctype.upper()} FLAG: Apply {jurisdiction} analysis. Focus on driver qualification, maintenance logs, and preservation under relevant FMCSR/TX-DOT standards."
    
    params = f"Framework: {ctype}\nDOI: {doi}\nSOL: {sol}\nGov Entity: {gov}\n\nSummary:\n{summary}"
    
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

    # API Key Selection
    if ai_engine == "Gemini (Google)":
        active_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password", key="gem_p")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password", key="gpt_p")
    else:
        active_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password", key="claude_p")

    st.markdown("---")
    
    case_type = st.selectbox(
        "Framework:", 
        ["Standard MVA", "Commercial Vehicle", "Trucking", "Premises", "Workplace", "UM/UIM", "TTCA"]
    )
    
    # Restored Jurisdiction Toggle
    cv_jurisdiction = "N/A"
    if case_type in ["Trucking", "Commercial Vehicle"]:
        cv_jurisdiction = st.radio(
            "FMCSR Scope:", 
            ["Interstate (Federal)", "Intrastate (Texas)", "Unsure"], 
            index=2, 
            horizontal=True,
            key="cv_scope"
        )

    discovery_level = "Level 2"
    if case_stage == "Active Litigation":
        discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True)

    st.markdown("**Risk Flags**")
    government_entity = st.checkbox("Government Entity Involved")

    case_summary = st.text_area("Case Summary", height=160, placeholder="Identify parties and incident mechanics...")
    
    with st.expander("Dates"):
        doi_input = st.text_input("Incident Date", placeholder="YYYY-MM-DD")
        sol_input = st.text_input("SOL Date", placeholder="YYYY-MM-DD")

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
            st.error("Integrity Error: Missing API Key.")
        else:
            prompt = build_prompt(
                case_stage, 
                case_type, 
                discovery_level, 
                doi_input, 
                sol_input, 
                case_summary, 
                government_entity,
                cv_jurisdiction
            )
            
            try:
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Gemini analyzing..."):
                        genai.configure(api_key=active_key)
                        # Changed from 'gemini-1.5-flash' to 'gemini-1.5-flash-latest'
                        model = genai.GenerativeModel("gemini-1.5-flash-latest")
                        st.session_state["brief_content"] = model.generate_content(prompt).text
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("OpenAI analyzing..."):
                        client = OpenAI(api_key=active_key)
                        resp = client.chat.completions.create(
                            model="gpt-4o", 
                            messages=[{"role": "user", "content": prompt}]
                        )
                        st.session_state["brief_content"] = resp.choices[0].message.content
                elif ai_engine == "Claude (Anthropic)":
                    with st.spinner("Claude analyzing..."):
                        client = anthropic.Anthropic(api_key=active_key)
                        resp = client.messages.create(
                            model="claude-3-5-sonnet-20240620", 
                            max_tokens=4000, 
                            messages=[{"role": "user", "content": prompt}]
                        )
                        st.session_state["brief_content"] = resp.content[0].text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    # Render brief and download button if content exists
    if st.session_state["brief_content"]:
        edited_text = st.text_area(
            "Strategic Brief Editor:", 
            value=st.session_state["brief_content"], 
            height=600
        )
        st.session_state["brief_content"] = edited_text
        
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        st.download_button(
            label="📥 Download Strategic Case Brief (.docx)",
            data=docx_data,
            file_name="Case_Intelligence_Brief.docx",
            use_container_width=True
        )
