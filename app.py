import streamlit as st
import pandas as pd
import io
import os

def extract_name(expression, kind):
    if not isinstance(expression, str):
        return ""
    kind_str = f',Kind="{kind}"'
    try:
        pos_kind = expression.index(kind_str)
        before = expression[:pos_kind]
        search = "[Name="
        last_name_pos = before.rfind(search)
        if last_name_pos == -1:
            return ""
        value_start = last_name_pos + len(search)
        value_len = pos_kind - last_name_pos - len(search)
        value = expression[value_start:value_start + value_len]
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]
        return value
    except ValueError:
        return ""

st.title("Excel Processor for Databricks Expressions")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Expressions")
        
        # Filter rows where 'Expression' contains 'Databricks'
        df_filtered = df[df['Expression'].str.contains('Databricks', na=False)]
        
        # Extract schema and table name for each filtered row
        schemas = []
        table_names = []
        for _, row in df_filtered.iterrows():
            schema = extract_name(row['Expression'], "Schema")
            table_name = extract_name(row['Expression'], "Table")
            schemas.append(schema)
            table_names.append(table_name)
        
        # Create output DataFrame
        df_output = pd.DataFrame({
            'Schema': schemas,
            'Table Name': table_names
        })
        
        # Remove duplicates based on Schema and Table Name
        df_output = df_output.drop_duplicates(subset=['Schema', 'Table Name'])
        
        # Get filename without extension and trim to first 10 characters if longer than 31
        filename_without_ext = os.path.splitext(uploaded_file.name)[0]
        if len(filename_without_ext) > 31:
            filename_without_ext = filename_without_ext[:10]
        sheet_name = f"{filename_without_ext}_datasources"
        output_filename = f"{filename_without_ext}_datasources.xlsx"
        
        # Create in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_output.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        
        st.download_button(
            label="Download Output Excel",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("Processing complete. Click the button to download the output.")
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")