import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Document Converter & Merger", layout="wide")

st.title("📄 Professional Document Converter")
st.markdown("Maintaining data integrity for Word-to-Excel conversion.")

# --- TASK 1: WORD TO EXCEL (INTEGRITY FOCUSED) ---
st.header("1. Convert Word to Excel")
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if word_file:
    doc = Document(word_file)
    all_data = []

    # Iterate through all elements in the document in order
    for block in doc.iter_block_items():
        # If it's a paragraph, add it to a new row
        if hasattr(block, 'text'):
            if block.text.strip():
                all_data.append([block.text])
        
        # If it's a table, extract all rows and cells
        elif hasattr(block, 'rows'):
            for row in block.rows:
                row_data = [cell.text.strip() for cell in row.cells]
                all_data.append(row_data)

    if all_data:
        # Create DataFrame - using max columns found to avoid alignment issues
        df_word = pd.DataFrame(all_data)
        
        st.success("Content extracted! Review the preview below:")
        st.dataframe(df_word.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_word.to_excel(writer, index=False, header=False)
        
        st.download_button(
            label="Download Converted Excel",
            data=output.getvalue(),
            file_name="converted_document.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("The document appears to be empty or unreadable.")

st.divider()

# --- TASK 2: MERGE EXCEL FILES ---
st.header("2. Merge Excel Files into Separate Sheets")
uploaded_files = st.file_uploader("Upload multiple Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    output_merged = io.BytesIO()
    
    with pd.ExcelWriter(output_merged, engine='openpyxl') as writer:
        for file in uploaded_files:
            # Integrity check: read all sheets from the uploaded file
            xls = pd.ExcelFile(file)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet_name)
                # Clean sheet name: combine filename and original sheet name
                final_sheet_name = f"{file.name[:15]}_{sheet_name}"[:31]
                df.to_excel(writer, sheet_name=final_sheet_name, index=False)
    
    st.success(f"Successfully merged {len(uploaded_files)} file(s)!")
    
    st.download_button(
        label="Download Merged Workbook",
        data=output_merged.getvalue(),
        file_name="merged_workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Logic to handle the sequence of document elements
def iter_block_items(parent):
    from docx.document import Document as _Document
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import _Cell, Table
    from docx.text.paragraph import Paragraph

    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise TypeError("Unknown parent type")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# Inject the helper function into Document
Document.iter_block_items = iter_block_items
