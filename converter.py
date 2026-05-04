import streamlit as st
import os
import io
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import google.generativeai as genai
from openai import OpenAI

# 1. Page Configuration
st.set_page_config(page_title="Discovery Drafter & Utility", layout="wide")

st.title("Legal Utility: PDF Converter & Discovery Response Drafter")
st.markdown("---")

# 2. Canonical Texas 42-Point Objection Taxonomy
TAXONOMY_OBJECTIONS = {
    # 1. Core Scope and Form Objections
    "OB-TX-001: Relevance / Outside Scope": (
        "Plaintiff objects to this request under TRCP 193.2 to the extent it seeks information "
        "not relevant to any party's claim or defense and is not within the permissible scope of discovery."
    ),
    "OB-TX-002: Overbroad as to Time": (
        "Plaintiff objects that this request is overly broad because it is not reasonably limited in time "
        "to the matters at issue in this litigation."
    ),
    "OB-TX-003: Overbroad as to Subject Matter": (
        "Plaintiff objects that this request is overly broad because it is not reasonably tailored to include "
        "only matters relevant to the issues in dispute."
    ),
    "OB-TX-004: Vague / Ambiguous / Undefined Terms": (
        "Plaintiff objects that this request is vague and ambiguous because it fails to define the key terms "
        "with reasonable certainty, such that Plaintiff cannot determine the exact information sought."
    ),
    "OB-TX-005: Lack of Reasonable Particularity": (
        "Plaintiff objects that this request fails to describe the items or documents sought with reasonable "
        "particularity and therefore imposes an improper burden on the responding party."
    ),
    "OB-TX-006: Improper Fishing Expedition": (
        "Plaintiff objects that this request constitutes an improper fishing expedition and is not reasonably "
        "tailored to obtain information directly relevant to the claims or defenses at issue."
    ),
    
    # 2. Burden, Proportionality, and Procedural Objections
    "OB-TX-007: Undue Burden / Expense": (
        "Plaintiff objects that complying with this request would impose an undue burden and expense "
        "that is completely disproportionate to the likely benefit of the discovery."
    ),
    "OB-TX-008: More Convenient / Less Burdensome Source": (
        "Plaintiff objects to this request to the extent the information or responsive documents sought are "
        "obtainable from a source that is more convenient, less burdensome, or less expensive."
    ),
    "OB-TX-009: Duplicative / Previously Produced": (
        "Plaintiff objects that this request is unreasonably cumulative or duplicative and, to the extent responsive "
        "material exists, it has already been produced or identified in prior productions."
    ),
    "OB-TX-010: Premature / Discovery Ongoing": (
        "Plaintiff objects that this request is premature to the extent it seeks a complete factual or evidentiary "
        "statement before discovery is sufficiently developed."
    ),
    "OB-TX-011: Mandatory Initial Disclosures (TRCP 194)": (
        "Plaintiff objects that this request seeks information through an improper discovery vehicle as the "
        "subject matter is directly governed by required initial disclosures under TRCP 194.1."
    ),
    
    # 3. Privilege and Protected-Information Objections
    "OB-TX-012: TRCP 193.3 Withholding Statement": (
        "Responsive material has been withheld pursuant to Tex. R. Civ. P. 193.3. The withheld material is responsive "
        "to this request and is withheld on the basis of applicable privileges."
    ),
    "OB-TX-013: Attorney-Client Privilege": (
        "Plaintiff withholds responsive material protected by the attorney-client privilege and provides this "
        "withholding statement pursuant to Rule 193.3."
    ),
    "OB-TX-014: Work Product Privilege": (
        "Plaintiff objects and withholds responsive material to the extent it constitutes work product or material "
        "prepared in anticipation of litigation under TRCP 192.5."
    ),
    "OB-TX-015: Consulting Expert Protection": (
        "Plaintiff objects to this request to the extent it seeks the identity, mental impressions, or opinions of "
        "consulting experts not expected to testify."
    ),
    "OB-TX-016: Spousal Privilege": (
        "Plaintiff objects to the extent this request seeks confidential communications between spouses protected "
        "by the spousal communications privilege."
    ),
    "OB-TX-017: Mental Health Records Not at Issue": (
        "Plaintiff objects to this request to the extent it seeks highly sensitive mental-health information where "
        "Plaintiff has not affirmatively placed their mental condition at issue in this litigation."
    ),
    "OB-TX-018: Privacy / Sensitive Personal Info": (
        "Plaintiff objects to this request to the extent it seeks highly sensitive personal identifiers or private "
        "information without a demonstrated need that is proportional to the case."
    ),
    
    # 4. Interrogatory-Specific Objections
    "OB-TX-019: Narrative / Marshaling Proof": (
        "Plaintiff objects to this interrogatory to the extent it requires a narrative response or detailed marshaling "
        "of proof more appropriately developed through deposition."
    ),
    "OB-TX-020: Marshal All Evidence": (
        "Plaintiff objects that this interrogatory improperly seeks to force Plaintiff to marshal all evidence "
        "supporting its claims or defenses."
    ),
    "OB-TX-021: Medical Opinion from Lay Party": (
        "Plaintiff objects to the extent this interrogatory requires a lay party Plaintiff to provide medical opinions "
        "beyond Plaintiff's personal knowledge or expert qualifications."
    ),
    "OB-TX-022: Improper Expert Discovery by Interrogatory": (
        "Plaintiff objects to this interrogatory to the extent it seeks expert information outside the scope or "
        "manner authorized by the expert discovery rules."
    ),
    "OB-TX-023: Beyond Current Knowledge": (
        "Plaintiff objects to the extent this interrogatory seeks information beyond Plaintiff's present knowledge "
        "and attempts to bind Plaintiff to a complete evidentiary statement before discovery is complete."
    ),
    "OB-TX-024: Exceeds Numerical Discovery Plan Limits": (
        "Plaintiff objects because this set exceeds the maximum number of permissible requests or answers under the "
        "governing TRCP discovery control plan."
    ),
    
    # 5. RFP / Authorization / Records Objections
    "OB-TX-025: Not in Possession, Custody, or Control": (
        "Plaintiff objects to the extent this request seeks materials not within Plaintiff's possession, custody, "
        "or control."
    ),
    "OB-TX-026: Request Requires Creation of a Document": (
        "Plaintiff objects because this request improperly requires Plaintiff to create a document that does "
        "not presently exist."
    ),
    "OB-TX-027: Blank Authorization / Lack of Specificity": (
        "Plaintiff objects to signing the requested authorization in blank because it fails to specify the records "
        "sought or specific providers, depriving Plaintiff of a meaningful opportunity to evaluate relevance."
    ),
    "OB-TX-028: Medical Authorization Improper / Records Forthcoming": (
        "Plaintiff objects to the extent this request seeks a blanket medical authorization rather than relevant "
        "medical records properly subject to production or disclosure."
    ),
    "OB-TX-029: Tax Returns / Heightened Financial Privacy": (
        "Plaintiff objects that this request seeks highly confidential tax information and is overbroad, intrusive, "
        "and not shown to be necessary in the form requested."
    ),
    "OB-TX-030: Social Security / Identifier Forms": (
        "Plaintiff objects to this request to the extent it seeks Social Security records or identifiers that are "
        "irrelevant, overbroad, and unduly intrusive."
    ),
    "OB-TX-031: Employment Records Limited to Wage Period": (
        "Plaintiff objects to this request to the extent it seeks personnel and employment records beyond those "
        "reasonably related to any claimed wage loss."
    ),
    "OB-TX-032: Phone Records Lack Nexus": (
        "Plaintiff objects because this request seeks telephone records without a pleaded or otherwise demonstrated "
        "nexus to any claim or defense in the case."
    ),
    
    # 6. Damages, Medical-Billing, and Collateral-Source Modules
    "OB-TX-033: Collateral Source Rule": (
        "Plaintiff objects to this request to the extent it seeks information regarding non-utilized health insurance, "
        "private health plans, or collateral benefits barred by the Texas Collateral Source Rule."
    ),
    "OB-TX-034: Medical Billing Creation Expansion": (
        "Plaintiff objects to the extent this request improperly expands the burden of medical billing disclosure "
        "by requiring the creation of an itemized analysis beyond the records themselves."
    ),
    "OB-TX-035: Premature Damage Computation": (
        "Plaintiff objects to the extent this request seeks a premature, exhaustive, or artificially fixed statement "
        "of damages before discovery is fully developed."
    ),
    
    # 7. Requests for Admission (RFAs)
    "OB-TX-036: Core Issue / Disputed Merits RFA": (
        "Plaintiff objects because this request for admission improperly seeks to establish a disputed merits issue "
        "rather than narrow an uncontroverted fact."
    ),
    "OB-TX-037: Compound RFA": (
        "Plaintiff objects because this request for admission is compound and does not permit a fair admission "
        "or denial of a single proposition."
    ),
    "OB-TX-038: Calls for Legal Conclusion": (
        "Plaintiff objects to this request for admission to the extent it seeks a pure legal conclusion."
    ),
    "OB-TX-039: Cannot Admit/Deny After Reasonable Inquiry": (
        "After reasonable inquiry, Plaintiff lacks sufficient information to admit or deny this request and "
        "therefore denies it."
    ),
    
    # 8. Practice-Area Modules
    "OB-TX-040: Prejudicial Terms / Controverted Text": (
        "Plaintiff objects to the use of highly subjective or prejudicial terms within the request to the extent "
        "they assume a disputed characterization of the incident."
    ),
    "OB-TX-041: All Written Complaints (Overbroad)": (
        "Plaintiff objects because the request for all written complaints across unrelated matters is overly broad, "
        "vague, and constitutes a prohibited fishing expedition."
    ),
    "OB-TX-042: Ex Parte Provider Communications": (
        "Plaintiff objects to the extent the requested authorization would permit ex parte communications with "
        "healthcare providers rather than the production of defined records."
    )
}

# --- IMPROVED DOCUMENT GENERATION ENGINE ---

def convert_pdf_text_to_docx(pdf_file):
    reader = PdfReader(pdf_file)
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

    for i, page in enumerate(reader.pages):
        page_text = page.extract_text()
        if page_text:
            lines = page_text.split('\n')
            for line in lines:
                cleaned_line = line.strip()
                if cleaned_line:
                    doc.add_paragraph(cleaned_line)
        else:
            doc.add_paragraph(f"[Page {i+1} contained no directly extractable text]")
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generate_response_docx(content_text):
    """
    Advanced discovery drafting engine. Automatically detects legal patterns 
    and applies court-standard bolding, spacing, and clean styling.
    """
    doc = docx.Document()
    
    # Set 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Base font style: Times New Roman 12pt
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # 1. Pleading Title Block
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(6)
    run_title = p_title.add_run("PLAINTIFF’S RESPONSES AND OBJECTIONS TO DEFENDANT’S DISCOVERY")
    run_title.font.bold = True
    run_title.font.size = Pt(12)
    
    # 2. Add structural separator line
    p_sep = doc.add_paragraph()
    p_sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sep.paragraph_format.space_after = Pt(18)
    p_sep.add_run("______________________________________________________________________")

    # 3. Text parsing and pattern detection
    lines = content_text.split('\n')
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        p = doc.add_paragraph()
        # Apply standard legal 1.15 line spacing and space after paragraph
        p.paragraph_format.space_after = Pt(12)
        p.paragraph_format.line_spacing = 1.15
        
        # Pattern A: Detect discovery headers (e.g., INTERROGATORY NO. 1:, REQUEST NO. 2:)
        if (cleaned_line.upper().startswith("INTERROGATORY") or 
            cleaned_line.upper().startswith("REQUEST FOR PRODUCTION") or 
            cleaned_line.upper().startswith("REQUEST NO.")):
            
            # Bold the header entirely for scannability
            run = p.add_run(cleaned_line)
            run.font.bold = True
            
        # Pattern B: Detect answer or response flags (e.g., ANSWER:, RESPONSE:)
        elif cleaned_line.upper() in ["ANSWER:", "RESPONSE:"]:
            run = p.add_run(cleaned_line)
            run.font.bold = True
            # Add a slight indent to simulate standard pleading alignment
            p.paragraph_format.left_indent = Inches(0.5)
            
        # Pattern C: Remove markdown heading characters but keep text bolded
        elif cleaned_line.startswith('###') or cleaned_line.startswith('##'):
            text_run = cleaned_line.replace('#', '').strip()
            run = p.add_run(text_run)
            run.font.bold = True
            
        # Pattern D: Parse any inline markdown bolding markers (e.g., **Subject to...**)
        elif '**' in cleaned_line:
            parts = cleaned_line.split('**')
            for index, part in enumerate(parts):
                if index % 2 == 1:  # Inside double asterisks -> bold
                    run = p.add_run(part)
                    run.font.bold = True
                else:  # Outside double asterisks -> normal text
                    run = p.add_run(part)
        else:
            # Standard, unformatted text line
            p.add_run(cleaned_line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# 3. Streamlit Tab UI
tab1, tab2 = st.tabs(["📄 Convert PDF to Word", "⚖️ Discovery Response Drafter"])

# --- TAB 1: PDF TO WORD ---
with tab1:
    st.subheader("Direct PDF to Word Converter")
    st.info("Upload any plain text PDF to convert it directly into an editable Word (.docx) file.")
    
    uploaded_pdf = st.file_uploader("Upload PDF File", type=["pdf"])
    
    if uploaded_pdf:
        with st.spinner("Parsing PDF text and building Word document..."):
            docx_buffer = convert_pdf_text_to_docx(uploaded_pdf)
            st.success("Conversion successful!")
            
            base_name = os.path.splitext(uploaded_pdf.name)[0]
            st.download_button(
                label=f"📥 Download '{base_name}.docx'",
                data=docx_buffer,
                file_name=f"{base_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

# --- TAB 2: DISCOVERY RESPONSE DRAFTER ---
with tab2:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1. Setup Objections & Workflow")
        
        ai_engine = st.radio(
            "Select Inference Model Engine:",
            ["Gemini (Google)", "ChatGPT (OpenAI)"],
            horizontal=True
        )
        
        user_api_key = st.text_input(
            f"Required: Enter your personal {ai_engine} API Key",
            type="password",
            placeholder="Key is not logged or stored"
        )
        
        st.markdown("---")
        
        request_type = st.selectbox(
            "Discovery Type", 
            ["Requests for Production (RFPs)", "Interrogatories (ROGs)", "Requests for Admissions (RFAs)"]
        )
        
        selected_objections = st.multiselect(
            "Select Candidate Objections to Screen & Apply:",
            list(TAXONOMY_OBJECTIONS.keys())
        )
        
        incoming_request = st.text_area(
            "Paste Defendant's Exact Question",
            height=130,
            placeholder="e.g., Interrogatory No. 1: Please state your full name..."
        )
        
        factual_basis = st.text_area(
            "Enter Factual Answer/Response details",
            height=100,
            placeholder="What actually occurred or what do we have?"
        )
        
        run_button = st.button("Generate Final Discovery Response", type="primary")

    with col2:
        st.subheader("2. Live Response Preview")
        
        if run_button:
            if not incoming_request:
                st.warning("Please paste the defendant's request first.")
            elif not user_api_key:
                st.error(f"Please provide your personal {ai_engine} API Key.")
            else:
                objection_text = "\n".join([f"- {TAXONOMY_OBJECTIONS[obj]}" for obj in selected_objections])
                
                # Strict drafting layout via engineering prompt
                prompt = f"""
                You are a meticulous legal discovery drafting engine. Your job is to output a direct legal response.

                INCOMING REQUEST TYPE: {request_type}
                DEFENDANT'S REQUEST:
                "{incoming_request}"

                PLAINTIFF'S FACTUAL RESPONSE:
                "{factual_basis if factual_basis else "Plaintiff has provided no additional information at this time."}"

                THE SELECTED CANDIDATE OBJECTIONS TO APPLY:
                {objection_text if objection_text else "None."}

                STRICT RESPONSE RULES:
                1. Avoid general opening context statements or conversation. Do not say "Here is your response".
                2. Do not use Markdown formatting characters like '#' or '*' in your output unless wrapping a heading.
                3. First, type out the exact text of the relevant objections applied.
                4. Directly after the objections, output: "**Subject to and without waiving the foregoing, Plaintiff responds as follows:**"
                5. Add the factual response text.
                """
                
                output_text = ""
                
                # Execute inference using the explicit User Key
                if ai_engine == "Gemini (Google)":
                    with st.spinner("Processing objections via Gemini..."):
                        try:
                            genai.configure(api_key=user_api_key)
                            model = genai.GenerativeModel("gemini-2.5-flash")
                            response = model.generate_content(prompt)
                            output_text = response.text
                        except Exception as e:
                            st.error(f"Error calling Gemini AI: {e}")
                                
                elif ai_engine == "ChatGPT (OpenAI)":
                    with st.spinner("Processing objections via ChatGPT..."):
                        try:
                            client = OpenAI(api_key=user_api_key)
                            response = client.chat.completions.create(
                                model="gpt-4o",
                                messages=[{"role": "user", "content": prompt}]
                            )
                            output_text = response.choices[0].message.content
                        except Exception as e:
                            st.error(f"Error calling ChatGPT AI: {e}")
                
                if output_text:
                    st.success("Draft Generated Successfully")
                    st.session_state["last_responder_output"] = output_text

        # 4. Display Persistent Downloads
        if "last_responder_output" in st.session_state:
            output_text = st.session_state["last_responder_output"]
            
            st.markdown("### Export Word Document")
            docx_data = generate_response_docx(output_text)
            st.download_button(
                label="📥 Download Word File (.docx)",
                data=docx_data,
                file_name="Compliant_Discovery_Response.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
            st.markdown("---")
            st.markdown(output_text)
