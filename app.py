import streamlit as st
import pandas as pd
from docx import Document
import io

# Set page config
st.set_page_config(page_title="Doc-to-Excel Converter", layout="wide")

st.title("📄 Professional Document Converter")
st.info("Upload a .docx to convert it to Excel, or multiple .xlsx files to merge them.")

# --- TASK 1: WORD TO EXCEL ---
st.header("1. Convert Word to Excel")
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="word_upload")

if word_file is not None:
    try:
        doc = Document(word_file)
        data = []

        # Logic: Extract every table and paragraph to maintain integrity
        for table in doc.tables:
            for row in table.rows:
                # Extract text from each cell in the row
                row_data = [cell.text.strip() for cell in row.cells]
                data.append(row_data)
            # Add an empty row between tables for clarity
            data.append([])

        # Also grab any text not in tables
        if not data:
            for para in doc.paragraphs:
                if para.text.strip():
                    data.append([para.text.strip()])

        if data:
            df_word = pd.DataFrame(data)
            st.success("File processed successfully!")
            st.dataframe(df_word.head(10))

            # Conversion logic
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_word.to_excel(writer, index=False, header=False)
            
            st.download_button(
                label="Download Excel File",
                data=output.getvalue(),
                file_name="converted_word_doc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("The document appears to be empty.")
            
    except Exception as e:
        st.error(f"Error processing Word file: {e}")

st.divider()

# --- TASK 2: MERGE EXCEL FILES ---
st.header("2. Merge Multiple Excel Sheets")
excel_files = st.file_uploader("Upload Excel files to merge", type=["xlsx"], accept_multiple_files=True, key="excel_upload")

if excel_files:
    try:
        merged_output = io.BytesIO()
        with pd.ExcelWriter(merged_output, engine='openpyxl') as writer:
            for f in excel_files:
                # Load the workbook to see all sheets
                temp_df_dict = pd.read_excel(f, sheet_name=None)
                for sheet_name, df in temp_df_dict.items():
                    # Create a unique sheet name based on filename + sheetname
                    clean_name = f"{f.name[:10]}_{sheet_name}"[:31]
                    df.to_excel(writer, sheet_name=clean_name, index=False)
        
        st.success(f"Merged {len(excel_files)} files!")
        st.download_button(
            label="Download Merged Workbook",
            data=merged_output.getvalue(),
            file_name="merged_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error merging files: {e}")
