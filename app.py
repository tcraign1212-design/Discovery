import streamlit as st
import io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
from openai import OpenAI
import anthropic
import os

# 1. Page Configuration
st.set_page_config(page_title="Case Review & Discovery Generator", layout="wide")

st.title("Legal Utility: Case Review & Advanced Discovery Planner")
st.markdown("---")

# --- SECURE API KEY RESOLUTION ---
env_gemini_key = st.secrets.get("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY")
env_openai_key = st.secrets.get("OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY")
env_anthropic_key = st.secrets.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_API_KEY")

# --- DOCUMENT GENERATION FUNCTIONS ---

def generate_audit_docx(audit_text):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("CASE AUDIT & TARGETED DISCOVERY PLAN")
    run_title.font.bold = True
    run_title.font.size = Pt(13)
    
    doc.add_paragraph("______________________________________________________________________")
    doc.add_paragraph()

    lines = audit_text.split('\n')
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
        p = doc.add_paragraph()
        if cleaned_line.startswith('###') or cleaned_line.startswith('##'):
            run = p.add_run(cleaned_line.replace('#', '').strip())
            run.font.bold = True
        elif cleaned_line.startswith('-') or cleaned_line.startswith('*'):
            run = p.add_run(cleaned_line)
        else:
            run = p.add_run(cleaned_line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# 2. Main Input & Review Interface
col1, col2 = st.columns([2, 3])

with col1:
    st.subheader("1. Case Details (Minimal Inputs)")
    
    ai_engine = st.radio(
        "Select Model Engine:",
        ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"],
        horizontal=True
    )
    
    active_api_key = ""
    if ai_engine == "Gemini (Google)":
        active_api_key = st.text_input("Gemini API Key", value=env_gemini_key or "", type="password")
    elif ai_engine == "ChatGPT (OpenAI)":
        active_api_key = st.text_input("OpenAI API Key", value=env_openai_key or "", type="password")
    elif ai_engine == "Claude (Anthropic)":
        active_api_key = st.text_input("Anthropic API Key", value=env_anthropic_key or "", type="password")

    st.markdown("---")
    
    case_type = st.selectbox(
        "Select Case Type & Framework:",
        [
            "Standard Motor Vehicle Accident (MVA)",
            "Commercial Vehicle / Trucking Crash",
            "Premises Liability (Slip/Trip/Fall)",
            "Workplace Injury / Non-Subscriber",
            "Uninsured/Underinsured Motorist (UM/UIM)",
            "Texas Tort Claims Act (TTCA) / Sovereign Immunity"
        ]
    )
    
    discovery_level = st.radio(
        "Discovery Level (Texas Rules of Civil Procedure):",
        ["Level 1 (Expedited, up to $250k)", "Level 2 (Standard)", "Level 3 (Custom/Complex)"],
        index=1,
        horizontal=True
    )
    
    case_summary = st.text_area(
        "Brief Case Summary / Fact Pattern",
        height=140,
        placeholder="Enter key details here (e.g., rear-end collision, commercial truck lane change, slip on clear liquid...)"
    )
    
    with st.expander("Optional Dates & Details"):
        c1, c2 = st.columns(2)
        with c1:
            date_of_incident = st.text_input("Date of Incident", placeholder="YYYY-MM-DD")
        with c2:
            sol_date = st.text_input("Statute of Limitations", placeholder="YYYY-MM-DD")

    run_audit = st.button("Generate Plan & Discovery", type="primary", use_container_width=True)

with col2:
    st.subheader("2. Audit Preview & Edit")
    
    if run_audit:
        if not case_summary:
            st.warning("Please provide a case summary to proceed.")
        elif not active_api_key:
            st.error(f"Please enter an API key for {ai_engine} above.")
        else:
            prompt = f"""
            You are an expert personal injury litigation analyst and master discovery drafter. 
            Your task is to perform an objective case review and draft highly detailed, case-specific discovery requests categorized exactly like formal Texas litigation requests.

            CASE PARAMETERS:
            - Framework Type: {case_type}
            - Discovery Level: {discovery_level}
            - Date of Incident: {date_of_incident if date_of_incident else "Not provided"}
            - Statute of Limitations: {sol_date if sol_date else "Not provided"}
            
            CASE SUMMARY:
            "{case_summary}"

            RULES & DIRECTIVES:
            1. Focus strictly on critical chronology, liability exposure, and strategic proof. Do not over-index on damages unless directly relevant to liability proof.
            2. For discovery, DO NOT include items covered under TRCP Rule 194.2 Required Disclosures. 
            3. Limit requests to exactly 3 case-specific Interrogatories and 3 targeted Requests for Production that directly address the gaps in the fact pattern.
            4. Format discovery with clear, topical category headers just like a formal request (e.g., 'Defendant's Conduct at the Time of the Collision', 'Causation & Background').
            5. Ensure questions are highly descriptive, formal, and thorough—not generic.

            Structure the output clearly using these exact headers:
            - ## 1. Chronology & Case Metrics
            - ## 2. Merits & Liability Exposure Analysis
            - ## 3. Targeted Pre-Litigation Steps
            - ## 4. Custom Discovery Requests (Excludes Initial Disclosures)
            """
            
            output_text = ""
            
            if ai_engine == "Gemini (Google)":
                with st.spinner("Analyzing case via Gemini..."):
                    try:
                        genai.configure(api_key=active_api_key)
                        model = genai.GenerativeModel("gemini-2.5-flash")
                        response = model.generate_content(prompt)
                        output_text = response.text
                    except Exception as e:
                        st.error(f"Inference error: {e}")
                            
            elif ai_engine == "ChatGPT (OpenAI)":
                with st.spinner("Analyzing case via OpenAI..."):
                    try:
                        client = OpenAI(api_key=active_api_key)
                        response = client.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": prompt}]
                        )
                        output_text = response.choices[0].message.content
                    except Exception as e:
                        st.error(f"Inference error: {e}")

            elif ai_engine == "Claude (Anthropic)":
                with st.spinner("Analyzing case via Anthropic..."):
                    try:
                        client = anthropic.Anthropic(api_key=active_api_key)
                        response = client.messages.create(
                            model="claude-3-5-sonnet-20241022",
                            max_tokens=2500,
                            messages=[{"role": "user", "content": prompt}]
                        )
                        output_text = response.content[0].text
                    except Exception as e:
                        st.error(f"Inference error: {e}")
            
            if output_text:
                st.session_state["review_content"] = output_text

    # Let the user view, edit, and download the report directly on screen
    if "review_content" in st.session_state:
        edited_text = st.text_area(
            "Review and edit your report below before exporting:",
            value=st.session_state["review_content"],
            height=400
        )
        st.session_state["review_content"] = edited_text
        
        st.markdown("### Export Report to Word")
        docx_data = generate_audit_docx(st.session_state["review_content"])
        st.download_button(
            label="📥 Download Audit & Discovery Plan (.docx)",
            data=docx_data,
            file_name="Case_Audit_And_Discovery_Plan.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
