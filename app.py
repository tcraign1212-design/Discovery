import streamlit as st
import os
import google.generativeai as genai
from pypdf import PdfReader
import docx2txt

# 1. Page Configuration
st.set_page_config(page_title="Discovery Auditor & Drafter", layout="wide")

# Pull master key if available in Streamlit Secrets
master_api_key = os.environ.get("GEMINI_API_KEY")

st.title("Case Discovery Auditor & Drafter")
st.markdown("---")

# 2. Hardcoded Checklist Database
CHECKLISTS = {
    "Standard MVA": [
        "Crash report (CR-3) with unredacted officer narrative",
        "TRCP 194 mandatory disclosures (Insurance policies/limits)",
        "Cell phone records / Call Detail Records (CDR) of defendant",
        "18.001 billing affidavits for all medical providers",
        "Property damage photos and repair estimates"
    ],
    "Premises Liability": [
        "Surveillance video covering the incident window",
        "Sweep logs, checklists, and maintenance logs for day of/prior",
        "Prior incident reports or 911 call logs for the location",
        "18.001 billing affidavits for all treating providers",
        "30(b)(6) Corporate Rep Topics on safety/inspection policies"
    ],
    "Commercial Vehicle": [
        "FMCSR Driver Qualification (DQ) File",
        "ECM / Black Box / Telematics data",
        "Hours of Service (HOS) / ELD logs for 7 days prior",
        "Post-accident toxicology screening",
        "30(b)(6) Corporate Rep Topics on hiring and maintenance"
    ]
}

# Helper function to extract text from files
def extract_text_from_file(uploaded_file):
    try:
        if uploaded_file.name.lower().endswith('.pdf'):
            reader = PdfReader(uploaded_file)
            extracted = []
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    extracted.append(text)
            return "\n".join(extracted)
        elif uploaded_file.name.lower().endswith('.docx'):
            return docx2txt.process(uploaded_file)
    except Exception as e:
        return f"Error reading file: {str(e)}"
    return ""

# 3. Two-Column User Interface
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Case Details & Evidence Upload")
    case_type = st.selectbox("Select Case Type", list(CHECKLISTS.keys()))
    
    # TRCP Discovery Level Selection
    trcp_level = st.radio(
        "TRCP Discovery Control Plan",
        ["Level 1 (TRCP 190.2)", "Level 2 (TRCP 190.3)"],
        help="Level 1: Max 15 RFPs, 15 ROGs, 15 RFAs. Level 2: Max 25 ROGs."
    )
    
    defense_theory = st.text_input(
        "What is the defense's theory?",
        placeholder="e.g., Open and obvious defect; Sudden emergency"
    )
    
    # Text area for additional notes
    case_notes = st.text_area(
        "Case Notes / File Status Details",
        height=120,
        placeholder="Paste any additional notes or file updates here..."
    )

    # Document Uploader Widget
    uploaded_files = st.file_uploader(
        "Upload Case Documents (CR-3, Defense Answers, Pleadings, etc.)", 
        type=["pdf", "docx"], 
        accept_multiple_files=True
    )

    uploaded_text_summary = ""
    if uploaded_files:
        st.info(f"Loaded {len(uploaded_files)} file(s).")
        for f in uploaded_files:
            file_text = extract_text_from_file(f)
            uploaded_text_summary += f"\n--- EXTRACTED FROM FILE: {f.name} ---\n{file_text}\n"
    
    user_api_key = st.text_input(
        "Individual Gemini API Key (Optional)",
        type="password",
        placeholder="Paste your individual key here to override the team's master key"
    )
    
    engine_choice = st.radio(
        "Choose Your Output Format",
        ["Gemini (Audit + Draft Discovery)", "Midpage (Export Prompt to Claude)"]
    )
    
    run_button = st.button("Generate Audit & Discovery", type="primary")

with col2:
    st.subheader("2. Strategic Output & Drafts")
    
    if run_button:
        # We need either manual notes OR uploaded documents to process.
        if not case_notes and not uploaded_text_summary:
            st.warning("Please paste some case notes or upload a file first.")
        else:
            selected_checklist = "\n".join([f"- {item}" for item in CHECKLISTS[case_type]])
            
            # Formulate the explicit rule boundaries based on Level
            if trcp_level == "Level 1 (TRCP 190.2)":
                limit_text = "Exactly 15 Requests for Production (RFPs), exactly 15 Interrogatories (ROGs), and exactly 15 Requests for Admissions (RFAs)."
            else:
                limit_text = "Exactly 25 Interrogatories (ROGs), and up to 25 each of Requests for Production (RFPs) and Requests for Admissions (RFAs)."

            # Merge manual case notes and document contents
            full_case_data = f"{case_notes}\n\n{uploaded_text_summary}"

            # OPTION 1: Gemini In-App Execution
            if engine_choice == "Gemini (Audit + Draft Discovery)":
                active_key = user_api_key if user_api_key else master_api_key
                
                if not active_key:
                    st.error("No API key detected. Please enter your individual key or add the master key in the app settings.")
                else:
                    with st.spinner("Analyzing case data and drafting targeted discovery package..."):
                        try:
                            genai.configure(api_key=active_key)
                            
                            prompt = f"""
                            You are an expert defense-minded legal auditor and discovery drafter checking a plaintiff's personal injury case file.
                            
                            CASE TYPE: {case_type}
                            DEFENSE THEORY: {defense_theory}
                            DISCOVERY CONTROL PLAN: {trcp_level}
                            MANDATORY CHECKLIST FOR THIS CASE TYPE:
                            {selected_checklist}
                            
                            CURRENT FILE STATUS, NOTES, AND EXTRACTED FILE TEXT:
                            {full_case_data}
                            
                            TASK:
                            Part 1: Case Audit
                            - Compare current file status against the mandatory checklist. List each item as [COMPLETED] or [MISSING].
                            - Note any statute of limitations (SOL) threats or structural causation gaps.
                            - Identify key vulnerabilities where the Defense Theory threatens liability.
                            
                            Part 2: Draft Targeted Discovery (Strictly bound by TRCP limits)
                            Draft exactly the following number of items based on {trcp_level}:
                            {limit_text}
                            
                            - Your output MUST contain:
                              1. Requests for Production (RFPs)
                              2. Interrogatories (ROGs)
                              3. Requests for Admissions (RFAs)
                              4. Deposition Topics
                            
                            - Ensure every question is highly specific, completely surgical, and directly targets the elements required to dismantle the defense theory. Avoid generic boilerplate.
                            """
                            
                            model = genai.GenerativeModel("gemini-2.5-flash")
                            response = model.generate_content(prompt)
                            
                            st.success("Analysis and Drafting Complete")
                            st.markdown(response.text)
                            
                        except Exception as e:
                            st.error(f"Error calling AI: {e}")
                            
            # OPTION 2: Midpage Integration
            elif engine_choice == "Midpage (Export Prompt to Claude)":
                st.info("Direct Midpage integration requires Claude Desktop or Web with the Midpage plugin.")
                
                midpage_prompt = f"""Use the Midpage tool to pull binding Texas case law and analyze this case:

CASE TYPE: {case_type}
DEFENSE THEORY: {defense_theory}
DISCOVERY LEVEL: {trcp_level}
MANDATORY DISCOVERY CHECKLIST:
{selected_checklist}
EVIDENCE CONTENT GATHERED:
{full_case_data}

INSTRUCTIONS:
1. Conduct a gap analysis and pinpoint vulnerabilities related to the defense theory.
2. Under the strict limits of {trcp_level}, draft exactly {limit_text} to secure the missing evidence.
3. Search Midpage case law for the best binding Texas precedent that defeats the defense theory, and cite it within the discovery strategy."""
                
                st.markdown("### Copy the prompt below and paste it directly into Claude:")
                st.code(midpage_prompt, language="text")
                st.success("Prompt generated above. Use the copy icon in the top right of the gray box.")
