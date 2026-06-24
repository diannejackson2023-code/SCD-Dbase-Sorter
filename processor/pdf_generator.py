from fpdf import FPDF
import os

class SCD_PDF(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 15)
        self.cell(0, 10, 'SCD Dbase Sorter - Project Documentation', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def convert_md_to_pdf(md_path, pdf_path):
    """
    Converts a Markdown file to a PDF.
    Very basic implementation for text and headers.
    """
    if not os.path.exists(md_path):
        print(f"Error: {md_path} not found.")
        return False
        
    pdf = SCD_PDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)
    
    filename = os.path.basename(md_path)
    pdf.set_font('Helvetica', 'B', 14)
    pdf.cell(0, 10, f"Document: {filename}", 0, 1, 'L')
    pdf.ln(5)
    
    effective_page_width = pdf.w - 2 * pdf.l_margin
    
    with open(md_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # Sanitize for latin-1
            line = line.encode('latin-1', 'replace').decode('latin-1')
            
            if not line:
                pdf.ln(5)
                continue
                
            # Basic Header handling
            if line.startswith('# '):
                pdf.set_font('Helvetica', 'B', 16)
                pdf.multi_cell(effective_page_width, 10, line[2:])
                pdf.set_font('Helvetica', size=10)
            elif line.startswith('## '):
                pdf.set_font('Helvetica', 'B', 14)
                pdf.multi_cell(effective_page_width, 10, line[3:])
                pdf.set_font('Helvetica', size=10)
            elif line.startswith('### '):
                pdf.set_font('Helvetica', 'B', 12)
                pdf.multi_cell(effective_page_width, 10, line[4:])
                pdf.set_font('Helvetica', size=10)
            elif line.startswith('- '):
                pdf.multi_cell(effective_page_width, 7, f"  * {line[2:]}")
            else:
                pdf.multi_cell(effective_page_width, 7, line)
                
    pdf.output(pdf_path)
    return True

if __name__ == "__main__":
    print("PDF Generator starting...")
    docs = [
        "/home/team/shared/SCD_Dbase_Sorter/TECHNICAL_MANUAL.md",
        "/home/team/shared/SCD_Dbase_Sorter/USER_MANUAL.md",
        "/home/team/shared/SCD_Dbase_Sorter/CHAT_HISTORY.md",
        "/home/team/shared/SCD_Dbase_Sorter/DEPLOYMENT_SECURITY.md"
    ]
    
    for doc in docs:
        out = doc.replace(".md", ".pdf")
        print(f"Converting {doc} to {out}...")
        success = convert_md_to_pdf(doc, out)
        if success:
            print(f"Successfully created {out}")
        else:
            print(f"Failed to create {out}")
