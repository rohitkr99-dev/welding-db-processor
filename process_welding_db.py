import pandas as pd


def process_file(uploaded_file):

    file_name = uploaded_file.name.lower()

    # Read XLSB
    if file_name.endswith(".xlsb"):

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            engine="pyxlsb"
        )

    # Read XLSX
    elif file_name.endswith(".xlsx"):

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            engine="openpyxl"
        )

    else:
        raise ValueError("Unsupported file format")

    # Example processing columns
    df["Match_Status"] = ""
    df["Remarks"] = ""

    return df
