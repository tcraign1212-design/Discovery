import streamlit as st
import os
import google.generativeai as genai

# 1. Page Configuration
st.set_page_config(page_title="Discovery Auditor", layout="wide")

# Pull free Google API key from Streamlit secrets
api_key = os.environ.get("GEMINI_API_KEY")

st.title("Case Discovery Auditor & Gap Analyst")
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
    st.subheader("1. Case Details")
    case_type = st.selectbox("Select Case Type", list(CHECKLISTS.keys()))
    
    defense_theory = st.text_input(
        "What is the defense's theory?",
        placeholder="e.g., Open and obvious defect; Sudden emergency"
    )
    
    case_notes = st.text_area(
        "Paste Case Notes / Evidence Gathered",
        height=250,
        placeholder="Paste current file status, what discovery has been served, etc."
    )
    
    # NEW ENGINE PICKER
    engine_choice = st.radio(
        "Choose Your Analysis Engine",
        ["Gemini (In-App Output)", "Midpage (Export Prompt to Claude + Midpage)"]
    )
    
    run_button = st.button("Run Audit", type="primary")

with col2:
    st.subheader("2. Audit Output")
    
    if run_button:
        if not case_notes:
            st.warning("Please paste some case notes first.")
        else:
            selected_checklist = "\n".join([f"- {item}" for item in CHECKLISTS[case_type]])
            
            # OPTION 1: Gemini In-App Execution
            if engine_choice == "Gemini (In-App Output)":
                if not api_key:
                    st.error("Missing API Key. Add your GEMINI_API_KEY to the Streamlit secrets.")
                else:
                    with st.spinner("Analyzing case against the mandatory ruleset via Gemini..."):
                        try:
                            genai.configure(api_key=api_key)
                            
                            prompt = f"""
                            You are a defense-minded legal auditor checking a plaintiff's case file for discovery gaps.
                            CASE TYPE: {case_type}
                            DEFENSE THEORY: {defense_theory}
                            MANDATORY CHECKLIST FOR THIS CASE TYPE:
                            {selected_checklist}
                            
                            CURRENT FILE STATUS / EVIDENCE IN HAND:
                            {case_notes}
                            
                            TASK:
                            1. Cross-reference the File Status against the Mandatory Checklist. List each item as [COMPLETED] or [MISSING].
                            2. Identify any strategic vulnerabilities or blind spots. Focus on what is needed to overcome the Defense Theory under Texas law.
                            3. Provide immediate corrective action items.
                            """
                            
                            model = genai.GenerativeModel("gemini-2.5-flash")
                            response = model.generate_content(prompt)
                            
                            st.success("Gemini Analysis Complete")
                            st.markdown(response.text)
                            
                        except Exception as e:
                            st.error(f"Error calling AI: {e}")
                            
            # OPTION 2: Midpage + Claude Integration
            elif engine_choice == "Midpage (Export Prompt to Claude + Midpage)":
                st.info("Direct Midpage integration requires Claude Desktop or Web with the Midpage MCP plugin.")
                
                # Format a hyper-specific Midpage/Claude prompt
                midpage_prompt = f"""Use the Midpage tool to pull binding Texas case law and analyze the following case file:

CASE TYPE: {case_type}
DEFENSE THEORY: {defense_theory}

MANDATORY DISCOVERY CHECKLIST TO ENFORCE:
{selected_checklist}

EVIDENCE GATHERED IN THIS FILE:
{case_notes}

INSTRUCTIONS:
1. Identify which mandatory checklist items are missing.
2. Search Midpage case law for the best binding Texas precedent that neutralizes the defense's specific theory ({defense_theory}).
3. Draft a strategic gap analysis memo, hyperlinking the cases found via Midpage."""
                
                st.markdown("### Copy the prompt below and paste it directly into Claude:")
                st.code(midpage_prompt, language="text")
                st.success("Prompt generated above. Use the copy icon in the top right of the gray box.")
