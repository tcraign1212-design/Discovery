import streamlit as st
import io
import re
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import google.generativeai as genai
from openai import OpenAI

# 1. Page Configuration
st.set_page_config(page_title="Discovery Drafter & Utility", layout="wide")

st.title("Legal Utility: Bulk Discovery Response Drafter")
st.markdown("---")

# 2. Complete Taxonomy of Objections
TAXONOMY_OBJECTIONS = {
    "OB-TX-001: Relevance / Outside Scope": "Plaintiff objects to this request under TRCP 193.2 to the extent it seeks information not relevant to any party's claim or defense and is not within the permissible scope of discovery.",
    "OB-TX-002: Overbroad as to Time": "Plaintiff objects that this request is overly broad because it is not reasonably limited in time to the matters at issue in this litigation.",
    "OB-TX-003: Overbroad as to Subject Matter": "Plaintiff objects that this request is overly broad because it is not reasonably tailored to include only matters relevant to the issues in dispute.",
    "OB-TX-004: Vague / Ambiguous / Undefined Terms": "Plaintiff objects that this request is vague and ambiguous because it fails to define the key terms with reasonable certainty, such that Plaintiff cannot determine the exact information sought.",
    "OB-TX-005: Lack of Reasonable Particularity": "Plaintiff objects that this request fails to describe the items or documents sought with reasonable particularity and therefore imposes an improper burden on the responding party.",
    "OB-TX-006: Improper Fishing Expedition": "Plaintiff objects that this request constitutes an improper fishing expedition and is not reasonably tailored to obtain information directly relevant to the claims or defenses at issue.",
    "OB-TX-007: Undue Burden / Expense": "Plaintiff objects that complying with this request would impose an undue burden and expense that is completely disproportionate to the likely benefit of the discovery.",
    "OB-TX-013: Attorney-Client Privilege": "Plaintiff withholds responsive material protected by the attorney-client privilege and provides this withholding statement pursuant to Rule 193.3.",
    "OB-TX-014: Work Product Privilege": "Plaintiff objects and withholds responsive material to the extent it constitutes work product or material prepared in anticipation of litigation under TRCP 192.5.",
    "OB-TX-018: Privacy / Sensitive Personal Info": "Plaintiff objects to this request to the extent it seeks highly sensitive personal identifiers or private information without a demonstrated need that is proportional to the case.",
    "OB-TX-019: Narrative / Marshaling Proof": "Plaintiff objects to this interrogatory to the extent it requires a narrative response or detailed marshaling of proof more appropriately developed through deposition.",
    "OB-TX-025: Not in Possession, Custody, or Control": "Plaintiff objects to the extent this request seeks materials not within Plaintiff's possession, custody, or control.",
    "OB-TX-026: Request Requires Creation of a Document": "Plaintiff objects because this request improperly requires Plaintiff to create a document that does not presently exist."
}

# 3. Secure Document Generation Logic
def generate_precise_docx(full_content, use_styling=False, case_info=None):
    doc = docx.Document()
    
    # 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Base Font Setup (Times New Roman 12pt)
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Add Case Style Box if requested
    if use_styling and case_info:
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        col_widths = [Inches(3.25), Inches(3.25)]
        for i, col in enumerate(table.columns):
            col.width = col_widths[i]

        p_left = table.cell(0, 0).paragraphs[0]
        p_left.paragraph_format.space_after = Pt(0)
        p_left.paragraph_format.line_spacing = 1.15
        p_left.add_run(f"CAUSE NO. {case_info.get('cause_no', '')}\n").bold = True
        p_left.add_run(f"{case_info.get('plaintiff', '')}\nPlaintiff,\nvs.\n{case_info.get('defendant', '')}\nDefendant.")

        p_right = table.cell(0, 1).paragraphs[0]
        p_right.paragraph_format.space_after = Pt(0)
        p_right.paragraph_format.line_spacing = 1.15
        p_right.add_run(f"§\tIN THE DISTRICT COURT\n§\n§\n§\t{case_info.get('court', '')}\n§\n§\t{case_info.get('county', '')} COUNTY, TEXAS")
        
        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(12)

    # Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(18)
    p_title.paragraph_format.space_before = Pt(6)
    run_title = p_title.add_run("PLAINTIFF’S RESPONSES AND OBJECTIONS TO DEFENDANT’S DISCOVERY")
    run_title.font.bold = True
    run_title.font.size = Pt(12)

    # Split output by exact discovery item markers
    item_regex = re.compile(r'(?=(?:INTERROGATORY|REQUEST FOR PRODUCTION|REQUEST FOR ADMISSION)\s+NO\.\s*\d+[:\s])', re.IGNORECASE)
    items = item_regex.split(full_content)
    
    subpart_regex = re.compile(r'^(?:[a-g1-9]\.|\([a-g1-9]\))\s', re.IGNORECASE)

    for item in items:
        cleaned_item = item.strip()
        if not cleaned_item:
            continue

        lines = cleaned_item.split('\n')
        for line in lines:
            cleaned_line = line.strip()
            if not cleaned_line:
                continue

            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(6)

            upper_line = cleaned_line.upper()

            # Format 1: Discovery Item Headers
            if any(x in upper_line for x in ["INTERROGATORY NO.", "REQUEST FOR PRODUCTION NO.", "REQUEST FOR ADMISSION NO."]):
                # Bold only the item prefix up to the colon
                match = re.match(r'^((?:INTERROGATORY|REQUEST FOR PRODUCTION|REQUEST FOR ADMISSION)\s+NO\.\s*\d+[:\s]*)\s*(.*)$', cleaned_line, re.IGNORECASE)
                if match:
                    run_bold = p.add_run(match.group(1) + " ")
                    run_bold.font.bold = True
                    p.add_run(match.group(2))
                else:
                    p.add_run(cleaned_line)
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(6)

            # Format 2: Direct Answers or Response Strings
            elif upper_line.startswith("ANSWER:") or upper_line.startswith("RESPONSE:"):
                run = p.add_run(cleaned_line)
                run.font.bold = True
                p.paragraph_format.left_indent = Inches(0)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(6)

            # Format 3: Subparts
            elif subpart_regex.match(cleaned_line):
                p.paragraph_format.left_indent = Inches(0.5)
                p.paragraph_format.space_after = Pt(4)
                p.add_run(cleaned_line)

            # Format 4: Normal Content
            else:
                p.paragraph_format.left_indent = Inches(0)
                # Parse markdown boldings if present
                if "**" in cleaned_line:
                    parts = cleaned_line.split('**')
                    for index, part in enumerate(parts):
                        if index % 2 == 1:
                            p.add_run(part).font.bold = True
                        else:
                            p.add_run(part)
                else:
                    p.add_run(cleaned_line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- TAB INTERFACE ---
tab1, tab2 = st.tabs(["⚖️ Bulk Discovery Drafter", "🛠 Setup Settings"])

with tab2:
    st.subheader("Model and Case Information Setup")
    
    ai_engine = st.radio("Select Inference Model Engine:", ["Gemini", "ChatGPT"], horizontal=True)
    api_key = st.text_input(f"Enter your {ai_engine} API Key", type="password")
    
    st.markdown("---")
    use_case = st.checkbox("Include Style Box at Top?", value=True)
    
    c1, c2 = st.columns(2)
    with c1:
        cause = st.text_input("Cause No.", value="DC-25-23025")
        plaintiff = st.text_input("Plaintiff", value="MIGUEL ORTIZ")
        court = st.text_input("Court", value="160th Judicial District")
    with c2:
        defendant = st.text_input("Defendant", value="GAMALIEL MORALES BAUTISTA")
        county = st.text_input("County", value="Dallas")


with tab1:
    st.subheader("1. Input Discovery Request Block")
    
    selected_objections = st.multiselect(
        "Apply Candidate Objections:", 
        list(TAXONOMY_OBJECTIONS.keys())
    )
    
    raw_discovery_input = st.text_area(
        "Paste Raw Incoming Discovery Block", 
        height=250, 
        placeholder="Paste multiple Interrogatories, RFPs, or RFAs here exactly as received."
    )
    
    factual_basis_input = st.text_area(
        "Enter Factual Basis Details", 
        height=100, 
        placeholder="What actually occurred or what do we have? (Applies to all questions in this batch)"
    )
    
    if st.button("Generate Final Discovery Document", type="primary"):
        if not raw_discovery_input:
            st.warning("Please paste raw discovery text first.")
        elif not api_key:
            st.error("Please configure your API Key in the Setup Settings tab.")
        else:
            objection_text = "\n".join([f"- {TAXONOMY_OBJECTIONS[obj]}" for obj in selected_objections])
            
            prompt = f"""You are a legal discovery drafting engine. Output the completely formatted discovery text.

RAW INCOMING DISCOVERY INPUT:
\"\"\"{raw_discovery_input}\"\"\"

PLAINTIFF'S FACTUAL BASIS OR DIRECT RESPONSE:
"{factual_basis_input if factual_basis_input else "None."}"

CANDIDATE OBJECTIONS TO BE INCLUDED BEFORE THE FACTUAL RESPONSE:
{objection_text if objection_text else "None."}

STRICT OUTPUT STRUCTURING RULES:
1. Output ALL discovery requests exactly as received.
2. Under each distinct discovery request, format the response exactly in this order:
   A. Write "ANSWER:" for Interrogatories/Admissions or "RESPONSE:" for Requests for Production.
   B. Output the applied objections text.
   C. Write the phrase on a single line exactly: "**Subject to and without waiving the foregoing, Plaintiff responds as follows:**"
   D. Output the Factual Basis details cleanly.
3. No opening conversational phrases or conversational text. Start directly with the first discovery header.
"""
            
            output_text = ""
            
            if ai_engine == "Gemini":
                with st.spinner("Generating output via Gemini AI..."):
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel("gemini-2.5-flash")
                        response = model.generate_content(prompt)
                        output_text = response.text
                    except Exception as e:
                        st.error(f"Error calling Gemini AI: {e}")
                            
            elif ai_engine == "ChatGPT":
                with st.spinner("Generating output via ChatGPT AI..."):
                    try:
                        client = OpenAI(api_key=api_key)
                        response = client.chat.completions.create(
                            model="gpt-4o",
                            messages=[{"role": "user", "content": prompt}]
                        )
                        output_text = response.choices[0].message.content
                    except Exception as e:
                        st.error(f"Error calling ChatGPT AI: {e}")
            
            if output_text:
                st.session_state["final_doc_output"] = output_text
                st.success("Responses processed and built.")

    if "final_doc_output" in st.session_state:
        st.markdown("---")
        st.subheader("2. Export Document Layout")
        
        case_dict = {"cause_no": cause, "plaintiff": plaintiff, "defendant": defendant, "court": court, "county": county} if use_case else None
        
        doc_buffer = generate_precise_docx(
            full_content=st.session_state["final_doc_output"],
            use_styling=use_case,
            case_info=case_dict
        )
        
        st.download_button(
            label="📥 Download Correctly Formatted Word File",
            data=doc_buffer,
            file_name="Compliant_Discovery_Document.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
        st.markdown("#### Preview Output")
        st.text(st.session_state["final_doc_output"])
