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
        # Assuming the sheet name is "Expressions" as in the original code
        df = pd.read_excel(uploaded_file, sheet_name="Expressions")
        
        # Filter rows where 'Expression' contains 'Databricks'
        df_filtered = df[df['Expression'].astype(str).str.contains('Databricks', na=False)]
        
        # Extract schema and table name for each filtered row
        schemas = []
        table_names = []
        for _, row in df_filtered.iterrows():
            schema = extract_name(row['Expression'], "Schema")
            table_name = extract_name(row['Expression'], "Table")
            schemas.append(schema)
            table_names.append(table_name)
        
        # Create output DataFrame with two columns
        df_temp = pd.DataFrame({
            'Schema': schemas,
            'Table Name': table_names
        })
        
        # Remove duplicates based on Schema and Table Name
        df_temp = df_temp.drop_duplicates(subset=['Schema', 'Table Name']).reset_index(drop=True)
        
        # Combine into the final desired single-column format
        df_output = pd.DataFrame({
            'Schema.Table Name': df_temp['Schema'] + '.' + df_temp['Table Name']
        })

        if df_output.empty:
            st.warning("No Databricks expressions found or no Schema/Table Name could be extracted.")
        else:
            # --- Online Preview ---
            st.subheader("Data Source Preview")
            st.dataframe(df_output)

            # --- Download Logic ---
            
            # Get filename without extension and trim to first 10 characters if longer than 31
            filename_without_ext = os.path.splitext(uploaded_file.name)[0]
            if len(filename_without_ext) > 31:
                filename_without_ext = filename_without_ext[:10]
            sheet_name = f"{filename_without_ext}_datasources"
            output_filename = f"{filename_without_ext}_datasources.xlsx"
            
            # Create in-memory Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Use df_output which has the single combined column
                df_output.to_excel(writer, sheet_name=sheet_name, index=False)
            output.seek(0)
            
            st.download_button(
                label="Download Output Excel",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.success("Processing complete. The output is displayed above. Click the button to download the Excel file.")

    except KeyError:
        st.error("Error: The uploaded Excel file must contain a sheet named 'Expressions' and a column named 'Expression'.")
    except ValueError as ve:
        st.error(f"Error reading Excel file: Check if the file is a valid Excel format. Details: {str(ve)}")
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
