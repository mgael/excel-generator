import streamlit as st
import pandas as pd
from docx import Document
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import io

# --- INTEGRITY STYLING CONSTANTS ---
HEADER_FILL = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid") # Professional Blue
HEADER_FONT = Font(color="FFFFFF", bold=True, size=12)
BORDER_STYLE = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def process_document(word_file):
    doc = Document(word_file)
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # We start with a blank sheet
        df_empty = pd.DataFrame()
        df_empty.to_excel(writer, sheet_name="Full_Report", index=False)
        worksheet = writer.sheets["Full_Report"]
        
        current_row = 1

        # 1. EXTRACT HEADER INTEGRITY
        for section in doc.sections:
            for para in section.header.paragraphs:
                if para.text.strip():
                    cell = worksheet.cell(row=current_row, column=1, value=para.text.strip())
                    cell.font = Font(italic=True, color="808080")
                    current_row += 1

        # 2. EXTRACT BODY (Text and Tables)
        for block in doc.element.body.iterchildren():
            # Handle Paragraphs
            if block.tag.endswith('p'):
                from docx.text.paragraph import Paragraph
                p = Paragraph(block, doc)
                if p.text.strip():
                    cell = worksheet.cell(row=current_row, column=1, value=p.text.strip())
                    # Check for Bold/Heading style
                    if any(run.bold for run in p.runs) or "Heading" in p.style.name:
                        cell.font = Font(bold=True, size=14)
                    current_row += 1
            
            # Handle Tables (Color & Border Integrity)
            elif block.tag.endswith('tbl'):
                from docx.table import Table
                t = Table(block, doc)
                for r_idx, row in enumerate(t.rows):
                    for c_idx, cell_obj in enumerate(row.cells):
                        cell = worksheet.cell(row=current_row, column=c_idx + 1, value=cell_obj.text.strip())
                        cell.border = BORDER_STYLE
                        
                        # If it's the first row of a table, treat it as a Header
                        if r_idx == 0:
                            cell.fill = HEADER_FILL
                            cell.font = HEADER_FONT
                        else:
                            cell.alignment = Alignment(wrap_text=True, vertical='center')
                    current_row += 1
                current_row += 1 # Space after table

        # 3. EXTRACT FOOTER INTEGRITY
        for section in doc.sections:
            for para in section.footer.paragraphs:
                if para.text.strip():
                    worksheet.cell(row=current_row, column=1, value=f"FOOTER: {para.text.strip()}").font = Font(size=8)
                    current_row += 1

        # Column Formatting
        for col in worksheet.columns:
            worksheet.column_dimensions[col[0].column_letter].width = 30

    return output.getvalue()

# --- STREAMLIT UI ---
st.set_page_config(page_title="Visual Integrity Converter")
st.title("📄 Mirror-Image Converter")

uploaded_word = st.file_uploader("Upload Word Doc", type="docx")
if uploaded_word:
    result = process_document(uploaded_word)
    st.success("Visual mapping applied (Headers, Footers, and Table Colors).")
    st.download_button("Download High-Integrity Excel", result, "Final_Report.xlsx")

# (Task 2 Merge logic remains same as previous stable version)
