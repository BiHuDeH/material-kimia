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

    # Create a new list to hold the modified column names
    new_columns = []
    gramaj_counter = 1  # Counter for unique naming

    # Loop through the columns and insert "گرماژ" before each "تعداد" column
    for col in df.columns:
        if "تعداد" in col:
            # Generate a unique name for the new "گرماژ" column
            gramaj_col = f"گرماژ_{gramaj_counter}"
            new_columns.append(gramaj_col)
            gramaj_counter += 1
        new_columns.append(col)

    # Reindex DataFrame with new columns for "گرماژ" where needed
    df = df.reindex(columns=new_columns, fill_value=0)

    # Perform the calculations and populate the "گرماژ" columns
    for idx, col in enumerate(df.columns):
        if "تعداد" in col:
            # The "گرماژ" column should be to the left of the "تعداد" column
            gramaj_col = df.columns[idx - 1]

            # Locate the "فروش" column for the same section
            sale_col = df.columns[idx + 1] if (idx + 1 < len(df.columns) and "فروش" in df.columns[idx + 1]) else None
            
            # Calculate "گرماژ" values if "فروش" is found
            if sale_col:
                df[gramaj_col] = df[sale_col] * df[col]

    # Display the updated DataFrame in Streamlit
    st.write("Updated DataFrame:")
    st.dataframe(df)

    # Provide a download option for the updated DataFrame
    @st.cache_data
    def convert_df(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        processed_data = output.getvalue()
        return processed_data

    st.download_button(
        label="Download updated Excel file",
        data=convert_df(df),
        file_name="updated_material_kimia.xlsx",
        mime="application/vnd.ms-excel"
    )
