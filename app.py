import streamlit as st
import pandas as pd

# Load the file from an upload
uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)

    # Initialize a counter for unique "گرماژ" columns
    gramaj_counter = 1

    # Create a new list of columns to handle insertions
    new_columns = []
    for col in df.columns:
        # Check if the column is "تعداد"
        if "تعداد" in col:
            # Create a unique name for the "گرماژ" column
            gramaj_col = f"گرماژ_{gramaj_counter}" if f"گرماژ" in df.columns else "گرماژ"
            new_columns.append(gramaj_col)
            gramaj_counter += 1  # Increment the counter for unique naming
        new_columns.append(col)

    # Reindex DataFrame with new columns to create empty "گرماژ" columns where needed
    df = df.reindex(columns=new_columns, fill_value=0)

    # Perform the calculations and populate the "گرماژ" columns
    for col in df.columns:
        if "تعداد" in col:
            # Get the index of the "تعداد" column
            idx = df.columns.get_loc(col)
            
            # Find the associated "گرماژ" column on the left
            gramaj_col = df.columns[idx - 1]
            
            # Get the associated "فروش" column, which should be to the right
            sale_col = df.columns[idx + 1] if idx + 1 < len(df.columns) else None
            
            # If a corresponding "فروش" column exists, calculate "گرماژ"
            if sale_col and "فروش" in sale_col:
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
