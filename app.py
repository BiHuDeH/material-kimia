import streamlit as st
import pandas as pd
from io import BytesIO

# Load the file from an upload
uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Replace blank cells with zero for calculations
    df.fillna(0, inplace=True)

    # Ensure unique column names for "گرماژ" and "تعداد" columns
    columns = list(df.columns)
    new_columns = []
    gramaj_counter = 1  # Counter for unique naming of "گرماژ"
    for col in columns:
        if col == "تعداد":
            # Generate unique names for "گرماژ" columns
            new_columns.append(f"گرماژ_{gramaj_counter}")
            gramaj_counter += 1
        new_columns.append(col)
    df.columns = new_columns  # Update columns in DataFrame

    # Add calculated "گرماژ" columns before each "تعداد" column
    for idx, col in enumerate(df.columns):
        if "تعداد" in col:
            gramaj_col = df.columns[idx - 1]
            sale_col = df.columns[idx + 1] if (idx + 1 < len(df.columns) and "فروش" in df.columns[idx + 1]) else None
            if sale_col:
                df[gramaj_col] = df[sale_col] * df[col]

    # Display the updated DataFrame in Streamlit
    st.write("Updated DataFrame:")
    st.dataframe(df)

    # Define a function to convert DataFrame to Excel format for download
    @st.cache_data
    def convert_df_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        processed_data = output.getvalue()
        return processed_data

    # Provide a download button for the updated DataFrame
    st.download_button(
        label="Download updated Excel file",
        data=convert_df_to_excel(df),
        file_name="updated_material_kimia.xlsx",
        mime="application/vnd.ms-excel"
    )
