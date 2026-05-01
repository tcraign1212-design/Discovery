import streamlit as st
import os
import google.generativeai as genai

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

# 3. Two-Column User Interface
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("1. Case Details & Settings")
    case_type = st.selectbox("Select Case Type", list(CHECKLISTS.keys()))
    
    defense_theory = st.text_input(
        "What is the defense's theory?",
        placeholder="e.g., Open and obvious defect; Sudden emergency"
    )
    
    case_notes = st.text_area(
        "Paste Case Notes / Evidence Gathered",
        height=200,
        placeholder="Paste current file status, what discovery has been served, etc."
    )
    
    # User Key Override Field
    user_api_key = st.text_input(
        "Individual Gemini API Key (Optional)",
        type="password",
        placeholder="Paste your individual key here to override the team's master key"
    )
    
    # Engine Picker
    engine_choice = st.radio(
        "Choose Your Output Format",
        ["Gemini (Audit + Draft Discovery)", "Midpage (Export Prompt to Claude)"]
    )
    
    run_button = st.button("Generate Audit & Discovery", type="primary")

with col2:
    st.subheader("2. Strategic Output & Drafts")
    
    if run_button:
        if not case_notes:
            st.warning("Please paste some case notes first.")
        else:
            selected_checklist = "\n".join([f"- {item}" for item in CHECKLISTS[case_type]])
            
            # OPTION 1: Gemini In-App Execution
            if engine_choice == "Gemini (Audit + Draft Discovery)":
                active_key = user_api_key if user_api_key else master_api_key
                
                if not active_key:
                    st.error("No API key detected. Please enter your individual key or have an admin add the master key in the app settings.")
                else:
                    with st.spinner("Analyzing and drafting targeted discovery..."):
                        try:
                            genai.configure(api_key=active_key)
                            
                            prompt = f"""
                            You are an expert defense-minded legal auditor and discovery drafter checking a plaintiff's personal injury case file.
                            
                            CASE TYPE: {case_type}
                            DEFENSE THEORY: {defense_theory}
                            MANDATORY CHECKLIST FOR THIS CASE TYPE:
                            {selected_checklist}
                            
                            CURRENT FILE STATUS / EVIDENCE IN HAND:
                            {case_notes}
                            
                            TASK:
                            Part 1: Case Audit
                            - Compare file status against the mandatory checklist. List each item as [COMPLETED] or [MISSING].
                            - Identify vulnerabilities where the Defense Theory threatens liability.
                            
                            Part 2: Draft Targeted Discovery (Texas Rules of Civil Procedure)
                            - For the [MISSING] items and defense vulnerabilities, draft the exact text for:
                              1. Requests for Production (RFPs)
                              2. Interrogatories (ROGs)
                              3. Deposition Topics
                            - Ensure the drafted discovery is surgical and specifically designed to dismantle the defense theory.
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
MANDATORY DISCOVERY CHECKLIST:
{selected_checklist}
EVIDENCE GATHERED:
{case_notes}

INSTRUCTIONS:
1. Conduct a gap analysis and pinpoint vulnerabilities related to the defense theory.
2. Draft targeted Texas-compliant Requests for Production, Interrogatories, and Deposition Topics to secure the missing evidence.
3. Search Midpage case law for the best binding Texas precedent that defeats the defense theory, and cite it within the discovery strategy."""
                
                st.markdown("### Copy the prompt below and paste it directly into Claude:")
                st.code(midpage_prompt, language="text")
                st.success("Prompt generated above. Use the copy icon in the top right of the gray box.")
