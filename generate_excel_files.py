from pathlib import Path

import pandas as pd


current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "INPUT"
output_dir = current_dir / "OUTPUT"
output_dir.mkdir(exist_ok=True, parents=True)
files = list(input_dir.rglob("*.txt"))

# The 'key' of 'Student_Class1A_23224_Eng.txt' is 'Student_Class1A_23224'
keys = set("_".join(file.stem.split("_")[:3]) for file in files)

for key in keys:
    with pd.ExcelWriter(
        output_dir / f"{key}_Results.xlsx", engine="xlsxwriter"
    ) as writer:
        for file in files:
            if file.stem.startswith(key):
                if file.stem.endswith("_Eng"):
                    df_eng = pd.read_csv(file, sep="\t")
                    df_eng.to_excel(writer, sheet_name="Test_Eng", startrow=7, index=False)
                if file.stem.endswith("_Math"):
                    df_math = pd.read_csv(file, sep="\t")
                    df_math.to_excel(writer, sheet_name="Test_Math", startrow=7, index=False)
        # Pandas can only export to xls, xlsx
        # To export to xlsm, we need to inject a template vbaProject.bin using xlsxwriter
        workbook = writer.book
        workbook.filename = output_dir / f"{key}_Results.xlsm"
        workbook.add_vba_project(current_dir / "vbaProject.bin")

# Clean up the temp xlsx files in the output dir
tmp_files = list(output_dir.rglob("*.xlsx"))
for tmp_file in tmp_files:
    tmp_file.unlink()
