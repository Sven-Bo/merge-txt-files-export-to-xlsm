from pathlib import Path

import pandas as pd
import xlwings as xw

# Path settings
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "INPUT"
output_dir = current_dir / "OUTPUT"
output_dir.mkdir(exist_ok=True, parents=True)
files = list(input_dir.rglob("*.txt"))
excel_template = current_dir / "Test_template.xlsm"


# The 'key' of 'Student_Class1A_23224_Eng.txt' is 'Student_Class1A_23224'
keys = set("_".join(file.stem.split("_")[:3]) for file in files)

with xw.App(visible=False) as app:
    for key in keys:
        wb = app.books.open(excel_template)
        for file in files:
            if file.stem.startswith(key):
                if file.stem.endswith("_Eng"):
                    df_eng = pd.read_csv(file, sep="\t")
                    wb.sheets("Test_Eng").range("A8").options(index=False).value = df_eng
                if file.stem.endswith("_Math"):
                    df_math = pd.read_csv(file, sep="\t")
                    wb.sheets("Test_Math").range("A8").options(index=False).value = df_math
        wb.save(output_dir / f"{key}_Results.xlsm")
