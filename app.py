import streamlit as st
import pandas as pd

# Load the file from an upload
uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Identify and insert "گرماژ" columns to the left of each "تعداد" column
    new_columns = []
    gramaj_count = 1  # Counter for numbering "گرماژ" columns if duplicates are found
    for col in df.columns:
        # Check if the column is a "تعداد" column
        if "تعداد" in col:
            # Create a unique name for the new "گرماژ" column
            gramaj_col = f"گرماژ_{gramaj_count}" if f"گرماژ" in df.columns else "گرماژ"
            new_columns.append(gramaj_col)
            gramaj_count += 1
        new_columns.append(col)

    # Reorder columns to insert "گرماژ" columns before each "تعداد"
    df = df.reindex(columns=new_columns, fill_value=0)

    # Perform the calculations and populate the "گرماژ" columns
    for i, col in enumerate(df.columns):
        if "تعداد" in col:
            # Get the corresponding "گرماژ" column on the left
            gramaj_col = new_columns[new_columns.index(col) - 1]

            # Get the "فروش" column for the section
            sale_col = df.columns[new_columns.index(col) + 1]

            # Calculate values by multiplying "فروش" by "تعداد" and store in "گرماژ"
            df[gramaj_col] = df[sale_col] * df[col]

    # Display the updated DataFrame in Streamlit
    st.write("Updated DataFrame:")
    st.dataframe(df)

    # Provide a download option for the updated DataFrame
    @st.cache_data
    def convert_df(df):
        return df.to_excel(index=False)

    st.download_button(
        label="Download updated Excel file",
        data=convert_df(df),
        file_name="updated_material_kimia.xlsx",
        mime="application/vnd.ms-excel"
    )
