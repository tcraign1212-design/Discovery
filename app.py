import streamlit as st
import os
import io
import google.generativeai as genai
from pypdf import PdfReader
import docx2txt
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from fpdf import FPDF

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

# --- DOCUMENT GENERATION FUNCTIONS ---

def generate_docx(content_text):
    """Generates a professional Word document with clean styling."""
    doc = docx.Document()
    
    # Page setup
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Style definitions
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(11)

    # Title Banner/Header
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("CASE AUDIT & TARGETED DISCOVERY PACKAGE")
    title_run.font.size = Pt(14)
    title_run.font.bold = True
    
    # Thin separator line
    doc.add_paragraph("______________________________________________________________________")
    doc.add_paragraph()

    # Process plain text lines
    lines = content_text.split('\n')
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            continue
            
        p = doc.add_paragraph()
        
        # Check header levels
        if cleaned_line.startswith('###'):
            run = p.add_run(cleaned_line.replace('###', '').strip())
            run.font.size = Pt(12)
            run.font.bold = True
        elif cleaned_line.startswith('##'):
            run = p.add_run(cleaned_line.replace('##', '').strip())
            run.font.size = Pt(13)
            run.font.bold = True
        elif cleaned_line.startswith('#'):
            run = p.add_run(cleaned_line.replace('#', '').strip())
            run.font.size = Pt(14)
            run.font.bold = True
        else:
            if cleaned_line.startswith('* ') or cleaned_line.startswith('- '):
                p.style = 'List Bullet'
                run = p.add_run(cleaned_line[2:])
            else:
                run = p.add_run(cleaned_line)
                
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def generate_pdf(content_text):
    """Generates a standard PDF using pure-Python FPDF2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(left=20, top=20, right=20)
    
    # Title
    pdf.set_font("Times", "B", size=14)
    pdf.cell(w=0, h=10, txt="CASE AUDIT & DISCOVERY PACKAGE", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    # Thin horizontal separator
    pdf.set_draw_color(150, 150, 150)
    pdf.line(x1=20, y1=pdf.get_y(), x2=190, y2=pdf.get_y())
    pdf.ln(8)
    
    # Convert input text to ISO-8859-1 compatible text to prevent encoding crashes
    lines = content_text.split('\n')
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            pdf.ln(2)
            continue
            
        # Clean special markdown characters and handle formatting
        if cleaned_line.startswith('###'):
            pdf.set_font("Times", "B", size=12)
            txt = cleaned_line.replace('###', '').strip()
        elif cleaned_line.startswith('##'):
            pdf.set_font("Times", "B", size=13)
            txt = cleaned_line.replace('##', '').strip()
        elif cleaned_line.startswith('#'):
            pdf.set_font("Times", "B", size=14)
            txt = cleaned_line.replace('#', '').strip()
        elif cleaned_line.startswith('* ') or cleaned_line.startswith('- '):
            pdf.set_font("Times", "", size=11)
            txt = f"• {cleaned_line[2:].strip()}"
        else:
            pdf.set_font("Times", "", size=11)
            txt = cleaned_line

        # Encode text cleanly for FPDF
        safe_txt = txt.encode('latin-1', 'replace').decode('latin-1')
        pdf.multi_cell(w=0, h=6, txt=safe_txt)
        
    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer

# --- FILE EXTRACTOR HELPER ---

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
    
    trcp_level = st.radio(
        "TRCP Discovery Control Plan",
        ["Level 1 (TRCP 190.2)", "Level 2 (TRCP 190.3)"],
        help="Level 1: Max 15 RFPs, 15 ROGs, 15 RFAs. Level 2: Max 25 ROGs."
    )
    
    defense_theory = st.text_input(
        "What is the defense's theory?",
        placeholder="e.g., Open and obvious defect; Sudden emergency"
    )
    
    case_notes = st.text_area(
        "Case Notes / File Status Details",
        height=120,
        placeholder="Paste any additional notes or file updates here..."
    )

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
        if not case_notes and not uploaded_text_summary:
            st.warning("Please paste some case notes or upload a file first.")
        else:
            selected_checklist = "\n".join([f"- {item}" for item in CHECKLISTS[case_type]])
            
            if trcp_level == "Level 1 (TRCP 190.2)":
                limit_text = "Exactly 15 Requests for Production (RFPs), exactly 15 Interrogatories (ROGs), and exactly 15 Requests for Admissions (RFAs)."
            else:
                limit_text = "Exactly 25 Interrogatories (ROGs), and up to 25 each of Requests for Production (RFPs) and Requests for Admissions (RFAs)."

            full_case_data = f"{case_notes}\n\n{uploaded_text_summary}"

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
                            output_text = response.text
                            
                            st.success("Analysis and Drafting Complete")
                            st.session_state["last_analysis_output"] = output_text
                            
                        except Exception as e:
                            st.error(f"Error calling AI: {e}")

    # 4. Persistence and Export buttons
    if "last_analysis_output" in st.session_state:
        output_text = st.session_state["last_analysis_output"]
        
        st.markdown("### Export Discovery Package")
        c1, c2 = st.columns(2)
        
        with c1:
            docx_data = generate_docx(output_text)
            st.download_button(
                label="📥 Download as Word File (.docx)",
                data=docx_data,
                file_name="Targeted_Discovery_Package.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        with c2:
            pdf_data = generate_pdf(output_text)
            st.download_button(
                label="📥 Download as PDF File (.pdf)",
                data=pdf_data,
                file_name="Targeted_Discovery_Package.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            
        st.markdown("---")
        st.markdown(output_text)
