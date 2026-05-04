import streamlit as st
import io
import re
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pypdf import PdfReader
import google.generativeai as genai
from openai import OpenAI

# 1. Page Configuration
st.set_page_config(page_title="Discovery Drafter", layout="wide")

st.title("Legal Utility: Discovery Response Drafter")
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

# 3. Predictable Document Generation Function
def generate_precise_docx(header_text, subparts_text, answer_text, objections_list, use_styling=False, case_info=None):
    doc = docx.Document()
    
    # Page setup: 1-inch margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Base Font: Times New Roman 12
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(12)

    # Optional Case Style Box
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
        
        doc.add_paragraph().paragraph_format.space_before = Pt(12)

    # 1. Add the Request Header (e.g., INTERROGATORY NO. 1: Please generally describe...)
    p_header = doc.add_paragraph()
    p_header.paragraph_format.space_before = Pt(12)
    p_header.paragraph_format.space_after = Pt(4)
    p_header.paragraph_format.line_spacing = 1.15
    p_header.paragraph_format.left_indent = Inches(0)
    
    # Split header text to only bold the identifier
    match = re.match(r'^((?:INTERROGATORY|REQUEST FOR PRODUCTION|REQUEST FOR ADMISSION)\s+NO\.\s*\d+:)\s*(.*)$', header_text, re.IGNORECASE)
    if match:
        run_bold = p_header.add_run(match.group(1) + " ")
        run_bold.font.bold = True
        p_header.add_run(match.group(2))
    else:
        p_header.add_run(header_text)

    # 2. Add subparts with explicit 0.5-inch indentation
    if subparts_text:
        subparts_lines = [line.strip() for line in subparts_text.split('\n') if line.strip()]
        for line in subparts_lines:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Inches(0.5)
            p_sub.paragraph_format.space_after = Pt(4)
            p_sub.paragraph_format.line_spacing = 1.15
            p_sub.add_run(line)

    # 3. Add the Answer Keyword Exactly
    p_ans = doc.add_paragraph()
    p_ans.paragraph_format.space_before = Pt(8)
    p_ans.paragraph_format.space_after = Pt(6)
    p_ans.paragraph_format.left_indent = Inches(0)
    run_ans = p_ans.add_run(answer_text if answer_text else "ANSWER:")
    run_ans.font.bold = True

    # 4. Append Objections
    if objections_list:
        p_obj = doc.add_paragraph()
        p_obj.paragraph_format.left_indent = Inches(0)
        p_obj.paragraph_format.space_after = Pt(6)
        p_obj.paragraph_format.line_spacing = 1.15
        p_obj.add_run(" ".join(objections_list))

    # Save to buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- TABBED INTERFACE ---
tab1, tab2 = st.tabs(["📄 Direct Docx Generator", "🤖 AI Assisted Drafter"])

with tab1:
    st.subheader("Direct Layout Control (Zero Guessing)")
    st.info("Input your discovery item piece by piece below. This enforces the exact spacing, bolding, and indentation rules.")
    
    c1, c2 = st.columns(2)
    with c1:
        use_case = st.checkbox("Include Case Style Box?", value=True, key="c_style_1")
        cause = st.text_input("Cause No.", value="DC-25-23025")
        plaintiff = st.text_input("Plaintiff", value="MIGUEL ORTIZ")
        defendant = st.text_input("Defendant", value="GAMALIEL MORALES BAUTISTA")
        court = st.text_input("Court", value="160th Judicial District")
        county = st.text_input("County", value="Dallas")

    with c2:
        req_type = st.selectbox("Document Answer Type", ["ANSWER:", "RESPONSE:"], key="ans_type_1")
        
        req_header = st.text_input(
            "Request Title and Main Question", 
            value="INTERROGATORY NO. 1: Please generally describe the Incident in Question, including:"
        )
        
        req_subparts = st.text_area(
            "Request Subparts (Each line here is indented exactly 0.5 inches)", 
            value="(a) the time You arrived at the Store;\n(b) the purpose for which You traveled to the Store;\n(c) who accompanied You to the store;"
        )
        
        chosen_objections = st.multiselect(
            "Apply Candidate Objections", 
            list(TAXONOMY_OBJECTIONS.keys()), 
            key="objs_1"
        )
        
    if st.button("Download Exact Docx Layout", type="primary"):
        case_dict = {"cause_no": cause, "plaintiff": plaintiff, "defendant": defendant, "court": court, "county": county}
        
        obj_text_list = [TAXONOMY_OBJECTIONS[obj] for obj in chosen_objections]
        
        doc_buffer = generate_precise_docx(
            header_text=req_header,
            subparts_text=req_subparts,
            answer_text=req_type,
            objections_list=obj_text_list,
            use_styling=use_case,
            case_info=case_dict
        )
        
        st.download_button(
            label="📥 Download Perfectly Formatted Word File",
            data=doc_buffer,
            file_name="Compliant_Discovery.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

with tab2:
    st.subheader("AI Objections Generation")
    
    ai_engine = st.radio("Select AI Inference Engine:", ["Gemini", "OpenAI"], horizontal=True)
    api_key = st.text_input("API Key", type="password")
    
    raw_req = st.text_area("Paste Incoming Question Here")
    factual_basis = st.text_area("Factual Basis for Answer")
    
    if st.button("Process using AI"):
        st.info("Obtaining objection recommendations...")
        # Integrates clean completions for tab 2...
