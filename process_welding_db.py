import pandas as pd

print("Welding DB Processor Started")

INPUT_FILE = "input/Welding DB.xlsb"

try:
    df = pd.read_excel(
        INPUT_FILE,
        sheet_name=0,
        engine="pyxlsb"
    )

    print("File Loaded Successfully")
    print("Total Rows:", len(df))
    print("Columns Found:")

    for col in df.columns:
        print("-", col)

except Exception as e:
    print("ERROR:")
    print(e)
