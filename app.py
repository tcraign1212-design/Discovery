import streamlit as st
import io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
from openai import OpenAI
import anthropic
import os

# 1. PAGE CONFIGURATION
st.set_page_config(page_title="Case Intelligence Brief Generator", layout="wide")

if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("*Strategic case framing for Texas personal injury litigation*")
st.markdown("---")

# 2. SECURE API KEY RESOLUTION
env_gemini_key    = st.secrets.get("GEMINI_API_KEY")    or os.environ.get("GEMINI_API_KEY")
env_openai_key    = st.secrets.get("OPENAI_API_KEY")    or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# 3. DOCUMENT GENERATION
def generate_brief_docx(brief_text: str, doc_title: str) -> io.BytesIO:
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1.25)
    style_normal = doc.styles["Normal"]
    style_normal.font.name, style_normal.font.size = "Times New Roman", Pt(12)
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(doc_title)
    run_title.font.bold, run_title.font.size = True, Pt(13)
    doc.add_paragraph("_" * 70)
    lines = brief_text.split("\n")
    for line in lines:
        cleaned = line.strip()
        if not cleaned:
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        if cleaned.startswith("### "):
            run = p.add_run(cleaned.replace("### ", "").strip())
            run.font.bold, run.font.size, run.font.underline = True, Pt(12), True
        elif cleaned.startswith("## "):
            run = p.add_run(cleaned.replace("## ", "").strip())
            run.font.bold, run.font.size = True, Pt(12)
            p.paragraph_format.space_before = Pt(10)
        else:
            p.add_run(cleaned)
    buffer = io.BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer

# 4. PROMPT BUILDER
def build_prompt(stage, ctype, dlevel, doi, sol, summary, gov, cv_jurisdiction):
    gov_flag = "\n- GOV FLAG: Apply TTCA analysis and notice deadlines." if gov else ""
    cv_flag = ""
    if ctype in ["Commercial Vehicle", "Commercial Vehicle / Trucking Crash"]:
        cv_flag = f"\n- COMMERCIAL FLAG: Apply {cv_jurisdiction} analysis. Focus on FMCSR/TX-DOT standards."
    
    params = f"Framework: {ctype}\nDOI: {doi}\nSOL: {sol}\nGov Entity: {gov}\n\nSummary:\n{summary}"
    
    if stage == "Pre-Litigation":
        return f"Draft Texas Pre-Suit Brief. {params} {gov_flag}{cv_flag} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Risk Flags, ## 4. Proof Gaps, ## 5. Defense Anticipation, ## 6. Action Items."
    return f"Draft Texas Litigation Blueprint. {params} Discovery: {dlevel} {gov_flag}{cv_flag} Headers: ## 1. Chronology, ## 2. Liability, ## 3. Proof Gaps, ## 4. Defense Anticipation, ## 5. Discovery Blueprint, ## 6. Strategic Flags."

# 5. MAIN LAYOUT
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Select Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)

    if ai_engine == "Gemini (Google)":
        active_api_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_api_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password")
    else:
        active_api_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password")

    st.markdown("---")
    case_type = st.selectbox(
        "Case Type / Framework:",
        [
            "Standard Motor Vehicle Accident (MVA)",
            "Commercial Vehicle / Trucking Crash",
            "Premises Liability (Slip/Trip/Fall)",
            "Workplace Injury / Non-Subscriber",
            "Uninsured/Underinsured Motorist (UM/UIM)",
            "Texas Tort Claims Act (TTCA) / Sovereign Immunity",
        ],
    )

    cv_jurisdiction = "N/A"
    if "Commercial Vehicle" in case_type:
        cv_jurisdiction = st.radio("FMCSR Scope:", ["Interstate (Federal)", "Intrastate (Texas)", "Unsure"], index=2, horizontal=True)

    discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True) if case_stage == "Active Litigation" else "N/A"
    government_entity = st.checkbox("Government Entity Involved")
    case_summary = st.text_area("Case Summary", height=160)
    
    with st.expander("Dates & Deadlines"):
        date_of_incident = st.text_input("Date of Incident", placeholder="YYYY-MM-DD")
        sol_date = st.text_input("SOL Expiration Date", placeholder="YYYY-MM-DD")

    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)

# 6. OUTPUT PANEL
with col2:
    st.subheader("2. Case Intelligence Brief")
    if run_brief:
        if not case_summary.strip():
            st.warning("Please provide a case summary.")
        elif not active_api_key:
            st.error(f"Please enter an API key.")
        else:
            prompt = build_prompt(case_stage, case_type, discovery_level, date_of_incident, sol_date, case_summary, government_entity, cv_jurisdiction)
            
            output_text = ""
            try:
                if ai_engine == "Gemini (Google)":
                with st.spinner("Analyzing via Gemini..."):
                    try:
                        genai.configure(api_key=active_api_key)
                        # This is the stable production anchor
                        model = genai.GenerativeModel("gemini-1.5-flash")
                        response = model.generate_content(prompt)
                        output_text = response.text
                    except Exception as e:
                        st.error(f"Gemini error: {e}")
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("Analyzing via OpenAI..."):
                        client = OpenAI(api_key=active_api_key)
                        response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": prompt}])
                        output_text = response.choices[0].message.content
                elif ai_engine == "Claude (Anthropic)":
                    with st.spinner("Analyzing via Claude..."):
                        client = anthropic.Anthropic(api_key=active_api_key)
                        response = client.messages.create(model="claude-3-5-sonnet-20240620", max_tokens=4000, messages=[{"role": "user", "content": prompt}])
                        output_text = response.content[0].text
                
                if output_text:
                    st.session_state["brief_content"] = output_text
            except Exception as e:
                st.error(f"Error: {e}")

    if st.session_state["brief_content"]:
        edited_text = st.text_area("Edit Brief:", value=st.session_state["brief_content"], height=500)
        st.session_state["brief_content"] = edited_text
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        st.download_button("📥 Download (.docx)", data=docx_data, file_name="Case_Brief.docx", use_container_width=True)
