import pandas as pd
import numpy as np


MASTER_FILE = "master/ID_BASE_WEIGHT_CALCULATION.xlsx"
MASTER_SHEET = "WEIGHT CALCULATION (2)"


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


def process_file(welding_file):

    # ---------------------------------
    # READ WELDING FILE
    # ---------------------------------

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

    # ---------------------------------
    # READ MASTER FILE
    # ---------------------------------

    id_df = pd.read_excel(
        MASTER_FILE,
        sheet_name=MASTER_SHEET,
        header=None,
        engine="openpyxl"
    )

    # ---------------------------------
    # OUTPUT COLUMNS
    # ---------------------------------

    welding_df["Final Weight"] = np.nan
    welding_df["Match Type"] = ""
    welding_df["Schedule 10 Thickness"] = np.nan
    welding_df["Schedule 10 Weight"] = np.nan
    welding_df["Data Status / Error Reason"] = ""

    # ---------------------------------
    # BUILD INCH DIA MAP
    # ---------------------------------

    inch_map = {}

    for idx in range(len(id_df)):

        value = safe_float(
            id_df.iloc[idx, 0]
        )

        if value is not None:

            inch_map[value] = idx

    available_inches = sorted(
        inch_map.keys()
    )

    # ---------------------------------
    # PROCESS EACH ROW
    # ---------------------------------

    for row_idx in welding_df.index:

        error_reason = ""

        try:

            # AI = Inch Dia
            inch_dia = safe_float(
                welding_df.iloc[row_idx, 34]
            )

            # AL = Thickness
            thickness = safe_float(
                welding_df.iloc[row_idx, 37]
            )

            # ---------------------------------
            # VALIDATION
            # ---------------------------------

            if inch_dia is None:

                error_reason = (
                    "Inch Dia blank in input file"
                )

                welding_df.at[
                    row_idx,
                    "Data Status / Error Reason"
                ] = error_reason

                continue

            if thickness is None:

                error_reason = (
                    "Original thickness blank in input file"
                )

                welding_df.at[
                    row_idx,
                    "Data Status / Error Reason"
                ] = error_reason

                continue

            # ---------------------------------
            # FIND MATCHED INCH DIA
            # ---------------------------------

            matched_inch = None
            nearest_higher_used = False

            for inch in available_inches:

                if inch >= inch_dia:

                    matched_inch = inch

                    if inch != inch_dia:
                        nearest_higher_used = True

                    break

            if matched_inch is None:

                error_reason = (
                    "No higher Inch Dia available"
                )

                welding_df.at[
                    row_idx,
                    "Data Status / Error Reason"
                ] = error_reason

                continue

            base_row = inch_map[
                matched_inch
            ]

            thickness_row = id_df.iloc[
                base_row
            ]

            weight_row = id_df.iloc[
                base_row + 1
            ]

            # ---------------------------------
            # BUILD THICKNESS-WEIGHT PAIRS
            # ---------------------------------

            thickness_candidates = []

            for col_idx in range(
                1,
                len(thickness_row)
            ):

                t_value = safe_float(
                    thickness_row.iloc[col_idx]
                )

                w_value = safe_float(
                    weight_row.iloc[col_idx]
                )

                if (
                    t_value is not None
                    and
                    w_value is not None
                ):

                    thickness_candidates.append(
                        (
                            t_value,
                            w_value,
                            col_idx
                        )
                    )

            if len(thickness_candidates) == 0:

                error_reason = (
                    "No valid thickness-weight pair found"
                )

                welding_df.at[
                    row_idx,
                    "Data Status / Error Reason"
                ] = error_reason

                continue

            thickness_candidates.sort(
                key=lambda x: x[0]
            )

            matched_weight = None
            match_type = ""

            # ---------------------------------
            # EXACT MATCH
            # ---------------------------------

            exact_found = False

            for (
                t_value,
                w_value,
                col_idx
            ) in thickness_candidates:

                if abs(t_value - thickness) < 0.0001:

                    matched_weight = w_value

                    match_type = "exact match"

                    exact_found = True

                    break

            # ---------------------------------
            # NEAREST HIGHER THICKNESS
            # ---------------------------------

            if not exact_found:

                for (
                    t_value,
                    w_value,
                    col_idx
                ) in thickness_candidates:

                    if t_value > thickness:

                        matched_weight = w_value

                        match_type = (
                            "nearest higher thickness"
                        )

                        exact_found = True

                        break

            # ---------------------------------
            # CALCULATED FIELD
            # ---------------------------------

            if not exact_found:

                max_thickness = (
                    thickness_candidates[-1][0]
                )

                max_weight = (
                    thickness_candidates[-1][1]
                )

                if max_thickness == 0:

                    error_reason = (
                        "Calculation failed"
                    )

                    welding_df.at[
                        row_idx,
                        "Data Status / Error Reason"
                    ] = error_reason

                    continue

                matched_weight = (
                    max_weight / max_thickness
                ) * thickness

                matched_weight = round(
                    matched_weight,
                    2
                )

                match_type = (
                    "calculated field"
                )

            # ---------------------------------
            # OVERRIDE MATCH TYPE
            # ---------------------------------

            if nearest_higher_used:

                match_type = (
                    "nearest higher Inch Dia"
                )

            # ---------------------------------
            # SCHEDULE 10 LOGIC
            # COLUMN D ONLY
            # ---------------------------------

            schedule10_thickness = safe_float(
                thickness_row.iloc[3]
            )

            schedule10_weight = safe_float(
                weight_row.iloc[3]
            )

            if (
                schedule10_thickness is None
                or
                schedule10_weight is None
            ):

                schedule10_thickness = 6

                schedule10_weight = round(
                    (
                        matched_weight / thickness
                    ) * 6,
                    2
                )

            # ---------------------------------
            # WRITE OUTPUT
            # ---------------------------------

            welding_df.at[
                row_idx,
                "Final Weight"
            ] = round(
                matched_weight,
                2
            )

            welding_df.at[
                row_idx,
                "Match Type"
            ] = match_type

            welding_df.at[
                row_idx,
                "Schedule 10 Thickness"
            ] = schedule10_thickness

            welding_df.at[
                row_idx,
                "Schedule 10 Weight"
            ] = schedule10_weight

            welding_df.at[
                row_idx,
                "Data Status / Error Reason"
            ] = "SUCCESS"

        except Exception as e:

            welding_df.at[
                row_idx,
                "Data Status / Error Reason"
            ] = f"Calculation failed: {str(e)}"

    return welding_df
