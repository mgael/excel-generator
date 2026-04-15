import streamlit as st
import pandas as pd
from spire.doc import Document, FileFormat
import io
import os

st.set_page_config(page_title="High-Fidelity Converter", layout="wide")

st.title("📄 High-Fidelity Document System")
st.markdown("Preserving colors, formatting, and structural integrity.")

# --- TASK 1: WORD TO EXCEL (FORMAT PRESERVATION) ---
st.header("1. Convert Word to Excel (Keep Formatting)")
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if word_file:
    # Save the uploaded file temporarily because Spire needs a path
    with open("temp_doc.docx", "wb") as f:
        f.write(word_file.getbuffer())
    
    try:
        # Initialize Spire.Doc
        doc = Document()
        doc.LoadFromFile("temp_doc.docx")
        
        # Convert Word to Excel format directly to preserve integrity
        output_path = "converted_result.xlsx"
        doc.SaveToFile(output_path, FileFormat.XLSX)
        
        with open(output_path, "rb") as f:
            excel_data = f.read()
            
        st.success("Conversion complete! Formatting and colors preserved.")
        
        st.download_button(
            label="Download Excel with Formatting",
            data=excel_data,
            file_name="formatted_document.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Cleanup
        os.remove("temp_doc.docx")
        os.remove(output_path)
        
    except Exception as e:
        st.error(f"Integrity Error: {e}")

st.divider()

# --- TASK 2: MERGE EXCEL (KEEPING SHEET INTEGRITY) ---
st.header("2. Merge Excel Files (Separate Sheets)")
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    output_merged = io.BytesIO()
    with pd.ExcelWriter(output_merged, engine='openpyxl') as writer:
        for file in uploaded_files:
            # Load all sheets from the file to ensure nothing is missed
            all_sheets = pd.read_excel(file, sheet_name=None)
            for sheet_name, df in all_sheets.items():
                # Name sheet by: FileName_OriginalSheetName
                final_name = f"{file.name[:10]}_{sheet_name}"[:31]
                df.to_excel(writer, sheet_name=final_name, index=False)
                
    st.success(f"Merged {len(uploaded_files)} workbooks successfully.")
    st.download_button(
        label="Download Merged Workbook",
        data=output_merged.getvalue(),
        file_name="merged_integrity_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
