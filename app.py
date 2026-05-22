import streamlit as st
import pandas as pd
from io import BytesIO

from process_welding_db import process_file


st.title("Welding DB Weight Processor")

# ---------------------------------
# Upload Welding File
# ---------------------------------

welding_file = st.file_uploader(
    "Upload Welding DB File",
    type=["xlsb", "xlsx"],
    key="welding"
)

# ---------------------------------
# Upload ID Base File
# ---------------------------------

id_base_file = st.file_uploader(
    "Upload ID BASE WEIGHT CALCULATION File",
    type=["xlsb", "xlsx"],
    key="idbase"
)

# ---------------------------------
# Process
# ---------------------------------

if welding_file is not None and id_base_file is not None:

    st.success("Both files uploaded successfully")

    if st.button("Process File"):

        processed_df = process_file(
            welding_file,
            id_base_file
        )

        # ---------------------------------
        # Data Preview
        # ---------------------------------

        st.subheader("Processed Data Preview")

        st.write(
            "Total Rows:",
            len(processed_df)
        )

        st.write(
            "Total Columns:",
            len(processed_df.columns)
        )

        st.dataframe(
            processed_df.head(10)
        )

        # ---------------------------------
        # Download Output
        # ---------------------------------

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            processed_df.to_excel(
                writer,
                index=False,
                sheet_name="Processed_Data"
            )

        output.seek(0)

        st.download_button(
            label="Download Processed File",
            data=output,
            file_name="Welding_DB_Processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
