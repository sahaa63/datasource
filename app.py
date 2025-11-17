import streamlit as st
import pandas as pd
import io
import os

def extract_name(expression, kind):
    if not isinstance(expression, str):
        return ""
    kind_str = f',Kind="{kind}"'
    try:
        # 1. Find the position of the specific kind (e.g., Kind="View")
        pos_kind = expression.index(kind_str)
        before = expression[:pos_kind]
        search = "[Name="
        
        # 2. Search backward from the kind position for the preceding [Name=
        last_name_pos = before.rfind(search)
        if last_name_pos == -1:
            return ""
        
        # 3. Calculate the start and length of the value
        value_start = last_name_pos + len(search)
        value_len = pos_kind - last_name_pos - len(search)
        
        # 4. Extract the value
        value = expression[value_start:value_start + value_len]
        
        # 5. Clean up surrounding quotes if present (e.g., "name" -> name)
        if value.startswith('"') and value.endswith('"'):
            value = value[1:-1]
        
        return value
    except ValueError:
        # 'kind_str' was not found in the expression
        return ""

st.title("Excel Processor for Databricks Expressions")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="Expressions")
        
        # Filter rows where 'Expression' contains 'Databricks'
        df_filtered = df[df['Expression'].str.contains('Databricks', na=False)]
        
        # Extract schema, table name, and view name for each filtered row
        # These lists are temporary for processing before concatenation
        schemas = []
        final_names = []
        
        for _, row in df_filtered.iterrows():
            expression = row['Expression']
            
            # Extract Schema Name (Kind="Schema")
            schema = extract_name(expression, "Schema")
            
            # Extract View Name (Kind="View")
            view_name = extract_name(expression, "View")
            
            # Extract Table Name (Kind="Table")
            table_name = extract_name(expression, "Table")
            
            # Determine the final name: Prioritize View over Table
            final_name = view_name if view_name else table_name
            
            schemas.append(schema)
            final_names.append(final_name)
        
        # Create an interim DataFrame for concatenation and deduplication
        df_interim = pd.DataFrame({
            'Schema': schemas,
            'Name': final_names 
        })
        
        # --- CONCATENATION AND DEDUPLICATION CHANGE ---
        
        # 1. Combine Schema and Name into a single string column: 'Schema.Table/View Name'
        # This is where the old final output format is created
        df_interim['Schema.Table/View Name'] = df_interim['Schema'] + '.' + df_interim['Name']
        
        # 2. Create the final output DataFrame with only the required single column
        df_output = df_interim[['Schema.Table/View Name']]
        
        # 3. Remove duplicates based on the combined column
        df_output = df_output.drop_duplicates().reset_index(drop=True)

        # ---------------------------------------------
        
        # --- PREVIEW SECTION ---
        st.subheader("Extracted Data Sources Preview 🔎")
        
        # Only display the single column
        st.dataframe(df_output)
        
        # Get filename without extension and trim for sheet name
        filename_without_ext = os.path.splitext(uploaded_file.name)[0]
        if len(filename_without_ext) > 31:
            filename_without_ext = filename_without_ext[:10]
        sheet_name = f"{filename_without_ext}_datasources"
        output_filename = f"{filename_without_ext}_datasources.xlsx"
        
        # Create in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write the single-column DataFrame
            df_output.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        
        st.success("Processing complete. Click the button to download the full list.")
        
        st.download_button(
            label="Download Output Excel",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
