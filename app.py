import streamlit as st
import pandas as pd
from docx2python import docx2python
import io

st.set_page_config(page_title="Document Converter & Merger", layout="wide")

st.title("📄 Document Processing System")
st.markdown("Convert Word to Excel and Merge Multiple Sheets")

# --- TASK 1: WORD TO EXCEL ---
st.header("1. Convert Word to Excel")
word_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"])

if word_file:
    # Extract data from Word
    with docx2python(word_file) as doc:
        # doc.body is a nested list: [sheet][table][row][cell]
        # We simplify it to extract the first table found or all text
        content = doc.body
        
    # Flattening logic: Taking the first table and making it a DataFrame
    if content:
        # Assuming table 1, row 1 is data
        try:
            data = content[0][0] # First table
            df_word = pd.DataFrame(data)
            
            st.success("Word Document parsed successfully!")
            st.dataframe(df_word.head())

            # Convert to Excel download link
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_word.to_excel(writer, index=False, header=False)
            
            st.download_button(
                label="Download as Excel",
                data=output.getvalue(),
                file_name="converted_word.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error("Could not find a clear table structure in this Word doc.")

st.divider()

# --- TASK 2: MERGE EXCEL FILES ---
st.header("2. Merge Excel Files into Sheets")
uploaded_files = st.file_uploader("Upload multiple Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    output_merged = io.BytesIO()
    
    with pd.ExcelWriter(output_merged, engine='openpyxl') as writer:
        for file in uploaded_files:
            # Read each uploaded file
            df = pd.read_excel(file)
            # Use the filename (minus .xlsx) as the sheet name
            sheet_name = file.name.split('.')[0][:30] 
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    st.success(f"Merged {len(uploaded_files)} files into one workbook!")
    
    st.download_button(
        label="Download Merged Workbook",
        data=output_merged.getvalue(),
        file_name="merged_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
