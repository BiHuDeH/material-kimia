import streamlit as st
import pandas as pd

# Load the file from an upload
uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file:
    # Read Excel file
    df = pd.read_excel(uploaded_file)
    
    # Determine the number of material sections (each starting at a new material name in Row 1)
    num_sections = len([col for col in df.columns if "تعداد" in col])

    # Process each material section
    for section in range(num_sections):
        # Get the current column positions for "تعداد" and "فروش"
        qty_col = f"تعداد_{section+1}"
        sale_col = f"فروش_{section+1}"

        # Insert "گرماژ" column after "تعداد"
        gram_col = f"گرماژ_{section+1}"
        df[gram_col] = df[sale_col] * df[qty_col]

    # Add a summary row for each "گرماژ" column
    summary_row = {}
    for col in df.columns:
        if "گرماژ" in col:
            summary_row[col] = df[col].sum()
        else:
            summary_row[col] = ""  # Keep other columns empty in the summary row
    
    # Append summary row to the DataFrame
    df = df.append(summary_row, ignore_index=True)

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
        file_name="updated_product_material_kimia.xlsx",
        mime="application/vnd.ms-excel"
    )
