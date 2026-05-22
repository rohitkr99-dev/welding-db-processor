import pandas as pd


def process_file(uploaded_file):

    file_name = uploaded_file.name.lower()

    # Read XLSB File
    if file_name.endswith(".xlsb"):

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            engine="pyxlsb"
        )

    # Read XLSX File
    elif file_name.endswith(".xlsx"):

        df = pd.read_excel(
            uploaded_file,
            sheet_name=0,
            engine="openpyxl"
        )

    else:
        raise ValueError("Unsupported file format")

    # Clean Column Names
    df.columns = df.columns.str.strip()

    # Create Validation Columns
    df["Validation_Status"] = "PASS"
    df["Validation_Remarks"] = ""

    # -----------------------------
    # Missing Welder1 Check
    # -----------------------------
    if "Welder1" in df.columns:

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

    # -----------------------------
    # Missing WPS No. Check
    # -----------------------------
    if "WPS No." in df.columns:

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

    # -----------------------------
    # Duplicate Barcode Check
    # -----------------------------
    if "Barcode" in df.columns:

        duplicate_barcode = (
            df["Barcode"].duplicated(keep=False)
        )

        df["Duplicate_Flag"] = duplicate_barcode.map(
            {
                True: "YES",
                False: "NO"
            }
        )

    else:
        df["Duplicate_Flag"] = "COLUMN NOT FOUND"

    return df
