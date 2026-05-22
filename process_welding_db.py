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

        # Validation Columns
    df["Validation_Status"] = "PASS"
    df["Validation_Remarks"] = ""

    # Missing Welder Check
    welder_missing = (
        df["Welder1"].isna()
    )

    df.loc[
        welder_missing,
        "Validation_Status"
    ] = "FAIL"

    df.loc[
        welder_missing,
        "Validation_Remarks"
    ] += "Missing Welder1; "

    # Missing WPS Check
    wps_missing = (
        df["WPS No."].isna()
    )

    df.loc[
        wps_missing,
        "Validation_Status"
    ] = "FAIL"

    df.loc[
        wps_missing,
        "Validation_Remarks"
    ] += "Missing WPS No.; "

    # Duplicate Barcode Check
    duplicate_barcode = (
        df["Barcode"].duplicated(keep=False)
    )

    df["Duplicate_Flag"] = duplicate_barcode.map(
        {
            True: "YES",
            False: "NO"
        }
    )

    return df
