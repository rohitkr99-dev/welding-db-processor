import pandas as pd


def process_file(uploaded_file):

    # Read XLSB file
    df = pd.read_excel(
        uploaded_file,
        sheet_name=0,
        engine="pyxlsb"
    )

    # Example processing columns
    df["Match_Status"] = ""
    df["Remarks"] = ""

    return df
