# Save this as app.py
import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.title("Excel Sheet Combiner")

# Let users upload multiple files
uploaded_files = st.file_uploader("Upload Excel files", type="xlsx", accept_multiple_files=True)

# Let users enter the sheet names they want to combine
sheet_names_input = st.text_input(
    "Enter comma-separated sheet names to combine",
    "Tiktok,Facebook,News,YouTube,Linkedin,Twitter,Instagram"
)

if st.button("Combine Files"):
    if not uploaded_files:
        st.warning("Please upload at least one Excel file.")
    else:
        # Parse sheet names (convert to lowercase for comparison)
        sheet_names = [s.strip().lower() for s in sheet_names_input.split(",")]
        combined_data = {sheet: [] for sheet in sheet_names}

        for uploaded_file in uploaded_files:
            try:
                company_name = os.path.splitext(uploaded_file.name)[0]
                xls = pd.ExcelFile(uploaded_file)

                # Normalize sheet names from the Excel file
                excel_sheets = {s.lower(): s for s in xls.sheet_names}

                for sheet in sheet_names:
                    if sheet in excel_sheets:
                        df = pd.read_excel(xls, sheet_name=excel_sheets[sheet])
                        df['Company'] = company_name
                        combined_data[sheet].append(df)
                    else:
                        st.warning(f"Sheet '{sheet}' not found in {uploaded_file.name}. Skipping this sheet.")
            except Exception as e:
                st.error(f"Error processing file {uploaded_file.name}: {e}")

        # Save the combined file to a BytesIO buffer
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for sheet, data in combined_data.items():
                if data:
                    combined_df = pd.concat(data, ignore_index=True)
                    # Use original casing from user input when saving
                    combined_df.to_excel(writer, sheet_name=sheet.capitalize(), index=False)
                    st.success(f"Sheet '{sheet}' combined successfully.")
                else:
                    st.warning(f"No data for sheet '{sheet}'.")

        # Provide download link
        output.seek(0)
        st.download_button(
            label="Download Combined Excel File",
            data=output,
            file_name="combined_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
