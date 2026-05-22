import streamlit as st
import pandas as pd
from io import BytesIO

from process_welding_db import process_file


st.title("Welding DB Processor")

uploaded_file = st.file_uploader(
    "Upload Welding DB File",
    type=["xlsb", "xlsx"]
)

if uploaded_file is not None:

    st.success("File uploaded successfully")

    if st.button("Process File"):

        # Process File
        processed_df = process_file(uploaded_file)

        # -----------------------------
        # Data Preview
        # -----------------------------
        st.subheader("Data Preview")

        st.write("Total Rows:", len(processed_df))
        st.write("Total Columns:", len(processed_df.columns))

        st.dataframe(processed_df.head(10))

        # -----------------------------
        # Column Names
        # -----------------------------
        st.subheader("Column Names")

        for col in processed_df.columns:
            st.write(col)

        # -----------------------------
        # Validation Summary
        # -----------------------------
        st.subheader("Validation Summary")

        total_rows = len(processed_df)

        pass_count = (
            processed_df["Validation_Status"]
            .eq("PASS")
            .sum()
        )

        fail_count = (
            processed_df["Validation_Status"]
            .eq("FAIL")
            .sum()
        )

        duplicate_count = (
            processed_df["Duplicate_Flag"]
            .eq("YES")
            .sum()
        )

        st.write("Total Rows:", total_rows)
        st.write("PASS Records:", pass_count)
        st.write("FAIL Records:", fail_count)
        st.write(
            "Duplicate Barcode Records:",
            duplicate_count
        )

        # -----------------------------
        # Download Output File
        # -----------------------------
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
