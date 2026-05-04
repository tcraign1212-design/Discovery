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
st.set_page_config(page_title="Case Review & Veto Auditor", layout="wide")

st.title("Legal Utility: Case Review & Workflow Auditor")
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
    run_title = p_title.add_run("CASE AUDIT & WORKFLOW VETO REPORT")
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


# 2. Tabs and Layout
tab1, tab2 = st.tabs(["🔍 Case Review & Audit", "ℹ️ About the Veto Philosophy"])

with tab1:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Case & Workflow Inputs")
        
        # User AI Engine Selection
        ai_engine = st.radio(
            "Select Inference Model Engine:",
            ["Gemini (Google)", "ChatGPT (OpenAI)", "Claude (Anthropic)"],
            horizontal=True
        )
        
        active_api_key = ""
        
        if ai_engine == "Gemini (Google)":
            if env_gemini_key:
                active_api_key = st.text_input(
                    "Optional: Enter your own Gemini API Key to override default settings",
                    value=env_gemini_key,
                    type="password"
                )
            else:
                active_api_key = st.text_input(
                    "Required: Enter your personal Gemini API Key",
                    type="password",
                    placeholder="AI key is not logged or stored on the server"
                )
        elif ai_engine == "ChatGPT (OpenAI)":
            if env_openai_key:
                active_api_key = st.text_input(
                    "Optional: Enter your own OpenAI API Key to override default settings",
                    value=env_openai_key,
                    type="password"
                )
            else:
                active_api_key = st.text_input(
                    "Required: Enter your personal OpenAI (ChatGPT) API Key",
                    type="password",
                    placeholder="sk-... (AI key is not logged or stored on the server)"
                )
        elif ai_engine == "Claude (Anthropic)":
            if env_anthropic_key:
                active_api_key = st.text_input(
                    "Optional: Enter your own Anthropic API Key to override default settings",
                    value=env_anthropic_key,
                    type="password"
                )
            else:
                active_api_key = st.text_input(
                    "Required: Enter your personal Anthropic (Claude) API Key",
                    type="password",
                    placeholder="sk-ant-... (AI key is not logged or stored on the server)"
                )
        
        st.markdown("---")
        
        # Restored Case Type Selection
        case_type = st.selectbox(
            "Select Case Type & Framework:",
            [
                "Standard Motor Vehicle Accident (MVA)",
                "Commercial Vehicle / Trucking Crash",
                "Premises Liability (Slip/Trip and Fall)",
                "Workplace Injury / Non-Subscriber Claim",
                "Uninsured/Underinsured Motorist (UM/UIM)",
                "Texas Tort Claims Act (TTCA) / Sovereign Immunity"
            ]
        )
        
        # Case Details Input fields
        case_summary = st.text_area(
            "Case Summary / Fact Pattern",
            height=120,
            placeholder="e.g., Rear-end collision on 4/15/2024. Defendant driver claims brakes failed..."
        )
        
        # Advanced Data Fields for Legal Integrity Screening
        c1, c2 = st.columns(2)
        with c1:
            date_of_incident = st.text_input("Date of Incident (DOI)", placeholder="YYYY-MM-DD")
            policy_limits = st.text_input("Insurance Limits ($)", placeholder="e.g., 30k/60k")
        with c2:
            sol_date = st.text_input("Statute of Limitations (SOL)", placeholder="YYYY-MM-DD")
            health_lien = st.text_input("Known Liens ($)", placeholder="e.g., ER Lien 12k")
            
        screening_modules = st.multiselect(
            "Select Specific Veto Sieve Checks to Execute:",
            [
                "Statute of Limitations (SOL) Calculation & Threat Analysis",
                "Causation Gaps (Pre-
