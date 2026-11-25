"""Cross-check the Excel workbook against the independent Python implementation.

Usage (from the project root):
    python src/verify_workbook.py [path/to/workbook.xlsx]

The workbook must contain calculated values. Files written by openpyxl have none until
they are opened and saved once in Excel or LibreOffice (the copy in this repo is already
calculated). Every value that exists in both calcs.py and the workbook is compared.
"""
import math
import sys
from pathlib import Path

from openpyxl import load_workbook

import build_workbook
import calcs


def main(path=None):
    root = Path(__file__).resolve().parent.parent
    path = Path(path) if path else root / "Substation_Design_Calculations.xlsx"
    _, reg = build_workbook.build(None)  # only needed for the key -> cell map
    wb = load_workbook(path, data_only=True)
    res = calcs.compute()

    checked, bad, empty = 0, [], []
    for key, ref in reg.items():
        if key not in res:
            continue
        sheet, cell = ref.split("!")
        got = wb[sheet][cell.replace("$", "")].value
        want = res[key]
        if got is None:
            empty.append(key)
            continue
        checked += 1
        if isinstance(want, str) or isinstance(got, str):
            ok = str(got) == str(want)
        else:
            ok = math.isclose(float(got), float(want), rel_tol=1e-6, abs_tol=1e-9)
        if not ok:
            bad.append((key, got, want))

    print(f"Compared {checked} values between {path.name} and calcs.py")
    if empty:
        print(f"No calculated value found for {len(empty)} cells (open and save the file in Excel/LibreOffice first): {empty[:5]}...")
    for key, got, want in bad:
        print(f"  MISMATCH {key}: workbook={got}  python={want}")
    if not bad and not empty:
        print("All values match.")
    return 1 if (bad or empty) else 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:2]))
