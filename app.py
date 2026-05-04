import streamlit as st
import io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from openai import OpenAI
import anthropic
import os

# 1. PAGE CONFIG & SESSION STATE
st.set_page_config(page_title="Case Intelligence Brief", layout="wide")
if "brief_content" not in st.session_state:
    st.session_state["brief_content"] = ""

st.title("Legal Utility: Case Intelligence Brief Generator")
st.markdown("---")

# API Keys
env_gemini_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
env_openai_key = st.secrets.get("OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# 2. DOCUMENT GENERATION
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

# ──────────────────────────────────────────────
# 3. STRATEGIC PROMPT BUILDER
# ──────────────────────────────────────────────
def build_prompt(stage, ctype, dlevel, doi, sol, summary, gov, cv_jurisdiction):
    gov_flag = (
        "\n   - GOVERNMENT ENTITY FLAG: Apply TTCA / sovereign immunity analysis. "
        "Identify specific agency rules and notice deadlines (e.g., 6-month notice)."
        if gov else ""
    )

    cv_flag = (
        f"\n   - COMMERCIAL VEHICLE FLAG: Apply FMCSR / {cv_jurisdiction} analysis. "
        "Flag DQF, Hours of Service, and black box preservation obligations."
        if "Commercial" in ctype or "Trucking" in ctype else ""
    )

    shared_params = f"""
CASE PARAMETERS:
- Framework: {ctype}
- DOI: {doi if doi else "Not provided"}
- SOL: {sol if sol else "Not provided"}
- Gov Involvement: {"YES" if gov else "No"}
- Commercial Involvement: {"YES" if cv_jurisdiction != "N/A" else "No"}

CASE SUMMARY:
\"\"\"{summary}\"\"\"
"""

    if stage == "Pre-Litigation":
        return f"""
You are a Texas personal injury litigation strategist conducting a pre-suit intake review. 
Your output is a Case Intelligence Brief—a structured strategic framework for early case workup.
Do NOT draft discovery requests. This matter has not been filed.

{shared_params}
{gov_flag}{cv_flag}

DIRECTIVES:
1. Identify specific agency notice deadlines (TTCA vs. Home-Rule).
2. List FOIA / Texas Public Information Act targets and record categories.
3. Identify spoliation risks and what preservation demand letters should issue immediately.
4. Be precise on dates. If the SOL date is provided, flag approaching windows.
5. Use these exact headers:
## 1. Chronology & Case Metrics
## 2. Liability Theory & Exposure Analysis
## 3. Intake Risk Flags
## 4. Proof Gap Analysis
## 5. Defense Anticipation
## 6. Pre-Suit Action Items
"""

    else:  # Active Litigation
        return f"""
You are a Texas personal injury litigation strategist conducting a case workup review for
an active suit. Your output is a Case Intelligence Brief — a Discovery Blueprint and
strategic framework. Do NOT draft actual discovery requests. 

{shared_params}
- Discovery Level: {dlevel}
{gov_flag}{cv_flag}

DIRECTIVES:
1. Do not include items covered by TRCP Rule 194.2 Required Disclosures.
2. Organize the Discovery Blueprint by target and strategic purpose.
3. Identify expert witness needs, spoliation risks, and insurance/coverage issues.
4. Use these exact headers:
## 1. Chronology & Case Metrics
## 2. Liability Theory & Exposure Analysis
## 3. Proof Gap Analysis
## 4. Defense Anticipation & Contention Points
## 5. Discovery Blueprint
  ### Requests for Admission
  ### Requests for Production
  ### Interrogatories
## 6. Strategic Flags
"""

# 4. MAIN LAYOUT
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

# 5. OUTPUT PANEL
with col2:
    st.subheader("2. Case Intelligence Brief")
    if run_brief:
        if not case_summary.strip() or not active_key:
            st.warning("Veto: Provide summary and key.")
        else:
            prompt = build_prompt(case_stage, case_type, discovery_level, doi_in, sol_in, case_summary, government_entity, cv_jurisdiction)
            try:
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Gemini thinking..."):
                        # Client uses the modern google-genai library
                        client = genai.Client(api_key=active_key)
                        response = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
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
