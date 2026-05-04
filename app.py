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
# 5. OUTPUT PANEL
# ──────────────────────────────────────────────
with col2:
    st.subheader("2. Case Intelligence Brief")

    if run_brief:
        # Data Integrity Check: The Sieve must have content to work
        if not case_summary.strip():
            st.warning("Veto: Provide a case summary to filter for causation gaps.")
        elif not active_key:
            st.error("Integrity Error: Missing API Key for selected engine.")
        else:
            # Passing all parameters including the restored FMCSR Jurisdiction
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
                    with st.spinner("Gemini analyzing case logic..."):
                        genai.configure(api_key=active_key)
                        model = genai.GenerativeModel("gemini-1.5-flash")
                        st.session_state["brief_content"] = model.generate_content(prompt).text
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("OpenAI analyzing case logic..."):
                        client = OpenAI(api_key=active_key)
                        resp = client.chat.completions.create(
                            model="gpt-4o", 
                            messages=[{"role": "user", "content": prompt}]
                        )
                        st.session_state["brief_content"] = resp.choices[0].message.content
                elif ai_engine == "Claude (Anthropic)":
                    with st.spinner("Claude analyzing case logic..."):
                        client = anthropic.Anthropic(api_key=active_key)
                        resp = client.messages.create(
                            model="claude-3-5-sonnet-20240620", 
                            max_tokens=4000, 
                            messages=[{"role": "user", "content": prompt}]
                        )
                        st.session_state["brief_content"] = resp.content[0].text
            except Exception as e:
                st.error(f"Execution Error: {e}")

    # Display the result in a persistent editor
    if st.session_state["brief_content"]:
        # Standardize formatting to Times New Roman style in the UI
        edited_text = st.text_area(
            "Strategic Brief Editor:", 
            value=st.session_state["brief_content"], 
            height=600
        )
        st.session_state["brief_content"] = edited_text
        
        # Mechanical Necessity: Generate the Word Doc for final output
        docx_data = generate_brief_docx(st.session_state["brief_content"], "CASE INTELLIGENCE BRIEF")
        
        st.download_button(
            label="📥 Download Strategic Case Brief (.docx)",
            data=docx_data,
            file_name="Case_Intelligence_Brief.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
