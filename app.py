import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Data Integrity Converter", layout="wide")

st.title("📄 High-Integrity Document System")
st.markdown("Preserving data structure and merging files securely.")

# --- TASK 1: WORD TO EXCEL ---
st.header("1. Convert Word to Excel")
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if word_file:
    try:
        # Load the document from the upload buffer
        doc = Document(word_file)
        data = []

        # We loop through every element (paragraph and table) 
        # to ensure the order (integrity) is maintained.
        for block in doc.iter_block_items():
            if hasattr(block, 'text'):  # Paragraph
                if block.text.strip():
                    data.append([block.text.strip()])
            elif hasattr(block, 'rows'):  # Table
                for row in block.rows:
                    row_data = [cell.text.strip() for cell in row.cells]
                    data.append(row_data)

        if data:
            df_word = pd.DataFrame(data)
            st.success("File processed! Content found.")
            st.dataframe(df_word.head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # We write to the Excel file without losing structure
                df_word.to_excel(writer, index=False, header=False)
                
                # Basic styling for integrity: Autofit columns
                worksheet = writer.sheets['Sheet1']
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    worksheet.column_dimensions[column].width = max_length + 2

            st.download_button(
                label="Download Excel File",
                data=output.getvalue(),
                file_name="converted_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"System Error: {e}. Please ensure the file is not corrupted.")

# --- Helper function for element ordering ---
def iter_block_items(parent):
    from docx.document import Document as _Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    parent_elm = parent.element.body if isinstance(parent, _Document) else parent._tc
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

Document.iter_block_items = iter_block_items

st.divider()

# --- TASK 2: MERGE EXCEL ---
st.header("2. Merge Excel Files")
excel_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if excel_files:
    merged_out = io.BytesIO()
    with pd.ExcelWriter(merged_out, engine='openpyxl') as writer:
        for f in excel_files:
            # Integrity: Load every single sheet in the workbook
            xls = pd.ExcelFile(f)
            for sheet in xls.sheet_names:
                df = pd.read_excel(f, sheet_name=sheet)
                name = f"{f.name[:10]}_{sheet}"[:31]
                df.to_excel(writer, sheet_name=name, index=False)
    
    st.success("All sheets merged successfully.")
    st.download_button(
        label="Download Merged File",
        data=merged_out.getvalue(),
        file_name="merged_data.xlsx"
    )
