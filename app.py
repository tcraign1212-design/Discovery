def generate_pdf(content_text):
    """Generates a standard PDF using pure-Python FPDF2 with safe fixed widths."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(left=20, top=20, right=20)
    
    # Calculate exact usable width: 210mm total - 40mm margins = 170mm
    usable_width = 170 
    
    # Title Banner
    pdf.set_font("Times", "B", size=14)
    pdf.cell(w=usable_width, h=10, txt="CASE AUDIT & DISCOVERY PACKAGE", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    # Thin horizontal separator
    pdf.set_draw_color(150, 150, 150)
    pdf.line(x1=20, y1=pdf.get_y(), x2=190, y2=pdf.get_y())
    pdf.ln(8)
    
    # Process the AI output line by line
    lines = content_text.split('\n')
    for line in lines:
        cleaned_line = line.strip()
        if not cleaned_line:
            pdf.ln(3)
            continue
            
        # Clear specific markdown formatting characters before rendering
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
            txt = f"- {cleaned_line[2:].strip()}"
        else:
            pdf.set_font("Times", "", size=11)
            txt = cleaned_line

        # Replace any residual unencodable characters to avoid crashes
        txt = txt.replace('**', '').replace('__', '')
        safe_txt = txt.encode('latin-1', 'replace').decode('latin-1')
        
        # Use fixed width (usable_width) instead of 0 to stop the truncation error
        pdf.multi_cell(w=usable_width, h=6, txt=safe_txt)
        
    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer
