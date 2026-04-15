import streamlit as st
import pandas as pd
from docx import Document
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
import io

# 1. Define the Helper Function FIRST
def iter_block_items(parent):
    """
    Yields each paragraph and table child within a docx document, 
    preserving the original order.
    """
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    else:
        parent_elm = parent._tc

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# 2. Set up Page Config
st.set_page_config(page_title="High-Integrity Converter", layout="wide")

st.title("📄 Professional Document Converter")
st.info("Ensuring data integrity by preserving document sequence.")

# --- TASK 1: WORD TO EXCEL ---
st.header("1. Convert Word to Excel")
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if word_file:
    try:
        doc = Document(word_file)
        data = []

        # Now we use the function defined above
        for block in iter_block_items(doc):
            if isinstance(block, Paragraph):
                if block.text.strip():
                    data.append([block.text.strip()])
            elif isinstance(block, Table):
                for row in block.rows:
                    row_data = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                    data.append(row_data)
                data.append([]) # Add a spacer row after tables

        if data:
            # Finding the max columns to ensure the DataFrame is consistent
            max_cols = max(len(row) for row in data)
            df_word = pd.DataFrame(data, columns=[f"Col {i+1}" for i in range(max_cols)])
            
            st.success("Document analyzed successfully!")
            st.dataframe(df_word.head(10))

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_word.to_excel(writer, index=False, header=False)
            
            st.download_button(
                label="Download Converted Excel",
                data=output.getvalue(),
                file_name="converted_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No readable text or tables found in this document.")

    except Exception as e:
        st.error(f"System Error: {e}")

st.divider()

# --- TASK 2: MERGE EXCEL ---
st.header("2. Merge Excel Files into Separate Sheets")
excel_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if excel_files:
    try:
        merged_out = io.BytesIO()
        with pd.ExcelWriter(merged_out, engine='openpyxl') as writer:
            for f in excel_files:
                xls = pd.ExcelFile(f)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(f, sheet_name=sheet)
                    # Create unique sheet names: Filename_Sheetname
                    clean_name = f"{f.name[:15]}_{sheet}"[:31]
                    df.to_excel(writer, sheet_name=clean_name, index=False)
        
        st.success(f"Merged {len(excel_files)} file(s) into one workbook.")
        st.download_button(
            label="Download Merged Workbook",
            data=merged_out.getvalue(),
            file_name="merged_data.xlsx"
        )
    except Exception as e:
        st.error(f"Error during merge: {e}")
