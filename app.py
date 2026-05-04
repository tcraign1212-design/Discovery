import streamlit as st
import io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai  # The modern SDK
from openai import OpenAI
import anthropic
import os

# 1. PAGE CONFIGURATION
st.set_page_config(page_title="Case Intelligence Brief Generator", layout="wide")
if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("---")

# 2. API KEY RESOLUTION
env_gemini_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
env_openai_key = st.secrets.get("OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# 3. DOCUMENT GENERATION
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
            r = p.add_run(line.replace("### ", "")); r.font.bold, r.font.underline = True, True
        elif line.startswith("## "):
            r = p.add_run(line.replace("## ", "")); r.font.bold = True
        else:
            p.add_run(line)
    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# 4. PROMPT BUILDER
def build_prompt(stage, ctype, dlevel, doi, sol, summary, gov, cv_jurisdiction):
    gov_txt = "\n- GOV FLAG: Apply TTCA analysis." if gov else ""
    cv_txt = f"\n- COMMERCIAL FLAG: Apply {cv_jurisdiction} analysis." if "Commercial" in ctype else ""
    params = f"Framework: {ctype}\nDOI: {doi}\nSOL: {sol}\n\nSummary:\n{summary}"
    return f"Draft Texas {'Pre-Suit Brief' if stage == 'Pre-Litigation' else 'Litigation Blueprint'}. {params} {gov_txt}{cv_txt}"

# 5. MAIN LAYOUT
col1, col2 = st.columns([2, 3])
with col1:
    st.subheader("1. Case Details")
    case_stage = st.radio("Case Stage:", ["Pre-Litigation", "Active Litigation"], horizontal=True)
    ai_engine = st.radio("Select Model Engine:", ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"], horizontal=True)
    
    if ai_engine == "Gemini (Google)":
        active_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password", key="gem_k")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password", key="gpt_k")
    else:
        active_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password", key="ant_k")

    case_type = st.selectbox("Framework:", ["Standard MVA", "Commercial Vehicle / Trucking Crash", "Premises", "Workplace", "UM/UIM", "TTCA"])
    cv_jurisdiction = st.radio("FMCSR Scope:", ["Interstate (Federal)", "Intrastate (Texas)", "Unsure"], index=2, horizontal=True) if "Commercial" in case_type else "N/A"
    discovery_level = st.radio("Discovery Level:", ["Level 1", "Level 2", "Level 3"], index=1, horizontal=True) if case_stage == "Active Litigation" else "N/A"
    government_entity = st.checkbox("Government Entity Involved")
    case_summary = st.text_area("Case Summary", height=160)
    doi_in = st.text_input("Incident Date")
    sol_in = st.text_input("SOL Date")
    run_brief = st.button("Generate Case Intelligence Brief", type="primary", use_container_width=True)

# 6. OUTPUT PANEL
with col2:
    st.subheader("2. Case Intelligence Brief")
    if run_brief:
        if not case_summary.strip() or not active_key:
            st.warning("Veto: Provide a case summary and API key.")
        else:
            prompt = build_prompt(case_stage, case_type, discovery_level, doi_in, sol_in, case_summary, government_entity, cv_jurisdiction)
            try:
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Gemini thinking..."):
                        # Fix: Initialize modern client
                        client = genai.Client(api_key=active_key)
                        # Fix: Generate content using modern SDK path
                        response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
                        st.session_state["brief_content"] = response.text
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
                st.error(f"Execution Error: {e}")

    if st.session_state["brief_content"]:
        st.session_state["brief_content"] = st.text_area("Edit Brief:", value=st.session_state["brief_content"], height=500)
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        st.download_button("📥 Download (.docx)", data=docx_data, file_name="Case_Brief.docx", use_container_width=True)
