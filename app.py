import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import io

# --- 1. THE CORE ENGINE: EXTRACTING WITH STYLE ---
def get_styled_content(doc):
    content = []
    # We iterate through paragraphs and tables in order
    for block in doc.element.body.iterchildren():
        if block.tag.endswith('p'): # Paragraph
            from docx.text.paragraph import Paragraph
            p = Paragraph(block, doc)
            if p.text.strip():
                # Store text + basic bold/italic metadata
                is_bold = any(run.bold for run in p.runs)
                content.append({'data': [p.text.strip()], 'bold': is_bold, 'type': 'text'})
        
        elif block.tag.endswith('tbl'): # Table
            from docx.table import Table
            t = Table(block, doc)
            for row in t.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                content.append({'data': row_data, 'bold': False, 'type': 'table'})
            content.append({'data': [], 'bold': False, 'type': 'spacer'})
    return content

# --- 2. STREAMLIT INTERFACE ---
st.set_page_config(page_title="High-Integrity Converter", layout="wide")
st.title("📄 Professional Document Integrity System")

st.header("1. Word to Excel (Format & Style Preservation)")
word_file = st.file_uploader("Upload Word Document", type=["docx"])

if word_file:
    try:
        doc = Document(word_file)
        styled_data = get_styled_content(doc)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Create a dummy dataframe for the structure
            df = pd.DataFrame([item['data'] for item in styled_data])
            df.to_excel(writer, index=False, header=False, sheet_name="Integrity_Export")
            
            workbook = writer.book
            worksheet = writer.sheets["Integrity_Export"]

            # --- 3. APPLYING THE INTEGRITY STYLING ---
            for idx, item in enumerate(styled_data):
                row_num = idx + 1
                for col_num, value in enumerate(item['data']):
                    cell = worksheet.cell(row=row_num, column=col_num + 1)
                    
                    # Apply Bold if it was bold in Word
                    if item.get('bold'):
                        cell.font = Font(bold=True, size=12)
                    
                    # Layout Arrangement: Center text if it's in a table
                    if item['type'] == 'table':
                        cell.alignment = Alignment(vertical='center', wrap_text=True)
                        # Light shading for tables to keep them distinct
                        cell.fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")

            # Autofit columns for professional look
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                worksheet.column_dimensions[column].width = min(max_length + 5, 50)

        st.success("Structure and Style mapping complete.")
        st.download_button("Download Styled Excel", output.getvalue(), "formatted_report.xlsx")
        
    except Exception as e:
        st.error(f"Integrity Error: {e}")

st.divider()

# --- 4. TASK 2: MULTI-SHEET MERGE ---
