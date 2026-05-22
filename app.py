import streamlit as st
import pandas as pd
from io import BytesIO

from process_welding_db import process_file


st.title("Welding DB Processor")

uploaded_file = st.file_uploader(
    "Upload Welding DB File",
    type=["xlsb"]
)

if uploaded_file is not None:

    st.success("File uploaded successfully")

    if st.button("Process File"):

        processed_df = process_file(uploaded_file)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
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
