import streamlit as st

st.title("Welding DB Processor")

uploaded_file = st.file_uploader(
    "Upload Welding DB File",
    type=["xlsb"]
)

if uploaded_file is not None:
    st.success("File uploaded successfully")
    st.write("Filename:", uploaded_file.name)
