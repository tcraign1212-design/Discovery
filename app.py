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
                "Causation Gaps (Pre-existing injuries / Gaps in treatment)",
                "Coverage Squeeze (Policy Limits vs. Potential Case Valuation)",
                "Lien Squeeze (Health Insurance & Medical Liens vs. Gross Settlement)"
            ],
            default=["Statute of Limitations (SOL) Calculation & Threat Analysis", "Causation Gaps (Pre-existing injuries / Gaps in treatment)"]
        )
        
        run_audit = st.button("Generate Case Audit Report", type="primary")

    with col2:
        st.subheader("2. Audit & Risk Evaluation")
        
        if run_audit:
            if not case_summary:
                st.warning("Please provide case details or a summary to analyze.")
            elif not active_api_key:
                st.error(f"Please provide a valid {ai_engine} API key to continue.")
            else:
                modules_to_run = "\n".join([f"- {m}" for m in screening_modules])
                
                # Sieve Audit Prompt
                prompt = f"""
                You are a defense-minded legal auditor and personal injury workflow analyst. 
                Your core operational strategy is the 'Veto Philosophy'—treating every review as a sieve to catch fatal case errors before proceeding to the next litigation phase.

                CASE SPECIFICS PROVIDED:
                - Date of Incident: {date_of_incident if date_of_incident else "Not provided"}
                - Statute of Limitations Date: {sol_date if sol_date else "Not provided"}
                - Insurance Limits: {policy_limits if policy_limits else "Not provided"}
                - Known Liens/ER Bills: {health_lien if health_lien else "Not provided"}
                
                CASE SUMMARY:
                "{case_summary}"

                SPECIFIC RISK MODULES TO EVALUATE:
                {modules_to_run if modules_to_run else "Standard comprehensive case review checks."}

                TASK RULES:
                1. Critically analyze the details above. Treat the case with a skeptical eye, seeking where it might fail down the line.
                2. Explicitly review the active Modules requested. Call out missing data, insurance or lien discrepancies, and temporal risks.
                3. Structure output strictly with clear headers:
                   - ## 1. Summary of Exposure & Phase-Gate Readiness
                   - ## 2. Core Veto Sieve Violations (Red Flags)
                   - ## 3. Actionable Mitigation Tasks (Pre-Litigation or Pre-Phase Move)
                4. Maintain a strategic, sharp, and data-integrity focused tone.
                """
                
                output_text = ""
                
                # Inference Routing
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Processing through Gemini Veto Sieve..."):
                        try:
                            genai.configure(api_key=active_api_key)
                            model = genai.GenerativeModel("gemini-2.5-flash")
                            response = model.generate_content(prompt)
                            output_text = response.text
                        except Exception as e:
                            st.error(f"Error calling Gemini AI: {e}")
                                
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("Processing through ChatGPT Veto Sieve..."):
                        try:
                            client = OpenAI(api_key=active_api_key)
                            response = client.chat.completions.create(
                                model="gpt-4o",
                                messages=[{"role": "user", "content": prompt}]
                            )
                            output_text = response.choices[0].message.content
                        except Exception as e:
                            st.error(f"Error calling ChatGPT AI: {e}")

                elif ai_engine == "Claude (Anthropic)":
                    with st.spinner("Processing through Claude Veto Sieve..."):
                        try:
                            client = anthropic.Anthropic(api_key=active_api_key)
                            response = client.messages.create(
                                model="claude-3-5-sonnet-20241022",
                                max_tokens=2500,
                                messages=[{"role": "user", "content": prompt}]
                            )
                            output_text = response.content[0].text
                        except Exception as e:
                            st.error(f"Error calling Claude AI: {e}")
                
                if output_text:
                    st.success("Case Audit Generated Successfully")
                    st.session_state["last_audit_output"] = output_text

        # 4. Display Persistent Downloads
        if "last_audit_output" in st.session_state:
            output_text = st.session_state["last_audit_output"]
            
            st.markdown("### Export Report to Word")
            docx_data = generate_audit_docx(output_text)
            st.download_button(
                label="📥 Download Audit Report (.docx)",
                data=docx_data,
                file_name="Case_Workflow_Veto_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            st.markdown("---")
            st.markdown(output_text)

with tab2:
    st.markdown("""
    ### The Veto Philosophy
    Building elite case-management processes means setting up mechanical safeguards that catch issues before they turn into major liabilities.
    
    * **Statute of Limitations:** Tracks critical countdown metrics to stop procedural defaults.
    * **Causation Integrity:** Catches treatment gaps and pre-existing injury issues before defense discovery can exploit them.
    * **Lien/Coverage Squeezes:** Compares total medical costs against available policy coverage early, avoiding unprofitable settlements down the line.
    """)
