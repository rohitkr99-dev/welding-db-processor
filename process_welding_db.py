import pandas as pd
import numpy as np


MASTER_FILE = "master/ID_BASE_WEIGHT_CALCULATION.xlsx"
MASTER_SHEET = "WEIGHT CALCULATION (2)"


# ---------------------------------------------------
# SAFE FLOAT CONVERSION
# ---------------------------------------------------

def safe_float(value):

    try:

        if pd.isna(value):
            return None

        value = str(value).strip()

        if value == "":
            return None

        value = value.replace('"', '')

        return float(value)

    except:
        return None


# ---------------------------------------------------
# FIND NEAREST HIGHER INCH DIA
# ---------------------------------------------------

def get_nearest_higher_dia(
    available_inches,
    target
):

    for dia in available_inches:

        if dia >= target:
            return dia

    return None


# ---------------------------------------------------
# PROCESS FILE
# ---------------------------------------------------

def process_file(welding_file):

    # ---------------------------------------------------
    # READ INPUT FILE
    # ---------------------------------------------------

    welding_name = welding_file.name.lower()

    if welding_name.endswith(".xlsb"):

        welding_df = pd.read_excel(
            welding_file,
            sheet_name=0,
            engine="pyxlsb"
        )

    else:

        welding_df = pd.read_excel(
            welding_file,
            sheet_name=0,
            engine="openpyxl"
        )

    welding_df.columns = welding_df.columns.str.strip()

    # ---------------------------------------------------
    # DYNAMIC COLUMN DETECTION
    # ---------------------------------------------------

    inch_col = None
    thickness_col = None

    for col in welding_df.columns:

        col_name = str(col).strip().lower()

        if col_name == "inch dia":
            inch_col = col

        if col_name in [
            "thickness value",
            "thickness"
        ]:
            thickness_col = col

    if inch_col is None:
        raise Exception("Inch Dia column not found")

    if thickness_col is None:
        raise Exception("Thickness column not found")

    # ---------------------------------------------------
    # READ MASTER FILE
    # ---------------------------------------------------

    id_df = pd.read_excel(
        MASTER_FILE,
        sheet_name=MASTER_SHEET,
        header=None,
        engine="openpyxl"
    )

    # ---------------------------------------------------
    # BUILD MASTER LOOKUP
    # ONLY ROWS WHERE COLUMN C = Thick.
    # ---------------------------------------------------

    inch_map = {}

    for idx in range(len(id_df)):

        marker = str(
            id_df.iloc[idx, 2]
        ).strip()

        if marker == "Thick.":

            inch_val = safe_float(
                id_df.iloc[idx, 0]
            )

            if inch_val is not None:

                inch_map[inch_val] = idx

    available_inches = sorted(
        inch_map.keys()
    )

    # ---------------------------------------------------
    # GLOBAL LOWEST THICKNESS
    # FOR FALLBACK FORMULA
    # ---------------------------------------------------

    all_pairs = []

    for dia in available_inches:

        row_idx = inch_map[dia]

        thickness_row = id_df.iloc[row_idx]
        weight_row = id_df.iloc[row_idx + 1]

        for col_idx in range(3, 16):

            t = safe_float(
                thickness_row.iloc[col_idx]
            )

            w = safe_float(
                weight_row.iloc[col_idx]
            )

            if (
                t is not None
                and
                w is not None
            ):

                all_pairs.append(
                    (t, w)
                )

    lowest_pair = min(
        all_pairs,
        key=lambda x: x[0]
    )

    global_low_thickness = lowest_pair[0]
    global_low_weight = lowest_pair[1]

    fallback_weight = (
        global_low_weight
        /
        global_low_thickness
    ) * 6

    # ---------------------------------------------------
    # OUTPUT COLUMNS
    # ---------------------------------------------------

    welding_df["Consumable Weight"] = np.nan
    welding_df["Data Selection Logic"] = ""
    welding_df["Column D Thickness"] = np.nan
    welding_df["Column D Weight"] = np.nan

    # ---------------------------------------------------
    # PROCESS ROWS
    # ---------------------------------------------------

    for row_idx in welding_df.index:

        try:

            inch_dia = safe_float(
                welding_df.at[
                    row_idx,
                    inch_col
                ]
            )

            thickness = safe_float(
                welding_df.at[
                    row_idx,
                    thickness_col
                ]
            )

            if inch_dia is None:

                welding_df.at[
                    row_idx,
                    "Data Selection Logic"
                ] = "No data available"

                continue

            if thickness is None:

                welding_df.at[
                    row_idx,
                    "Data Selection Logic"
                ] = "No data available"

                continue

            # ---------------------------------------------------
            # FIND MATCHING INCH DIA
            # ---------------------------------------------------

            status_parts = []

            if inch_dia in inch_map:

                matched_inch = inch_dia

            else:

                matched_inch = get_nearest_higher_dia(
                    available_inches,
                    inch_dia
                )

                if matched_inch is not None:

                    status_parts.append(
                        "Higher Inch Dia"
                    )

            if matched_inch is None:

                welding_df.at[
                    row_idx,
                    "Data Selection Logic"
                ] = "No data available"

                welding_df.at[
                    row_idx,
                    "Column D Thickness"
                ] = 6

                welding_df.at[
                    row_idx,
                    "Column D Weight"
                ] = round(
                    fallback_weight,
                    4
                )

                continue

            # ---------------------------------------------------
            # GET THICKNESS ROWS
            # ---------------------------------------------------

            base_row = inch_map[
                matched_inch
            ]

            thickness_row = id_df.iloc[
                base_row
            ]

            weight_row = id_df.iloc[
                base_row + 1
            ]

            # ---------------------------------------------------
            # BUILD THICKNESS-WEIGHT PAIRS
            # ONLY D:P
            # ---------------------------------------------------

            pairs = []

            for col_idx in range(3, 16):

                t = safe_float(
                    thickness_row.iloc[col_idx]
                )

                w = safe_float(
                    weight_row.iloc[col_idx]
                )

                if (
                    t is not None
                    and
                    w is not None
                ):

                    pairs.append(
                        (
                            t,
                            w
                        )
                    )

            pairs.sort(
                key=lambda x: x[0]
            )

            final_weight = None

            # ---------------------------------------------------
            # EXACT MATCH
            # ---------------------------------------------------

            exact_match = next(
                (
                    p for p in pairs
                    if abs(p[0] - thickness) < 0.0001
                ),
                None
            )

            if exact_match:

                final_weight = exact_match[1]

                status_parts.append(
                    "Direct Match"
                )

            else:

                # ---------------------------------------------------
                # HIGHER THICKNESS
                # ---------------------------------------------------

                higher_match = next(
                    (
                        p for p in pairs
                        if p[0] > thickness
                    ),
                    None
                )

                if higher_match:

                    final_weight = higher_match[1]

                    status_parts.append(
                        "Higher Thickness"
                    )

                else:

                    # ---------------------------------------------------
                    # CALCULATED
                    # ---------------------------------------------------

                    highest = max(
                        pairs,
                        key=lambda x: x[0]
                    )

                    final_weight = (
                        highest[1]
                        /
                        highest[0]
                    ) * thickness

                    status_parts.append(
                        "Calculated Data"
                    )

            # ---------------------------------------------------
            # COLUMN D LOGIC
            # ---------------------------------------------------

            d_thickness = safe_float(
                thickness_row.iloc[3]
            )

            d_weight = safe_float(
                weight_row.iloc[3]
            )

            if d_thickness is None:

                d_thickness = 6

            if d_weight is None:

                d_weight = fallback_weight

            # ---------------------------------------------------
            # WRITE OUTPUT
            # ---------------------------------------------------

            welding_df.at[
                row_idx,
                "Consumable Weight"
            ] = round(
                final_weight,
                4
            )

            welding_df.at[
                row_idx,
                "Data Selection Logic"
            ] = " + ".join(
                status_parts
            )

            welding_df.at[
                row_idx,
                "Column D Thickness"
            ] = d_thickness

            welding_df.at[
                row_idx,
                "Column D Weight"
            ] = round(
                d_weight,
                4
            )

        except Exception as e:

            welding_df.at[
                row_idx,
                "Data Selection Logic"
            ] = f"Error: {str(e)}"

    # ---------------------------------------------------
    # FINAL VALIDATION
    # NO BLANKS ALLOWED
    # ---------------------------------------------------

    for row_idx in welding_df.index:

        d_t = welding_df.at[
            row_idx,
            "Column D Thickness"
        ]

        d_w = welding_df.at[
            row_idx,
            "Column D Weight"
        ]

        if pd.isna(d_t):

            welding_df.at[
                row_idx,
                "Column D Thickness"
            ] = 6

        if pd.isna(d_w):

            welding_df.at[
                row_idx,
                "Column D Weight"
            ] = round(
                fallback_weight,
                4
            )

    return welding_df
