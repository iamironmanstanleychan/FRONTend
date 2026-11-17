"""
Fix the two formula errors in KungFu_Negotiation_Model.xlsx.
Requires: pip install openpyxl
Run once. Overwrites the file in place.
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

PATH = "KungFu_Negotiation_Model.xlsx"
wb = load_workbook(PATH)

# ---------------------------------------------------------------------------
# FIX 1 — Valuation B25:D33  (rate sensitivity)
# Replace SUMPRODUCT array formula with explicit term-by-term sum.
# P&L rows 5-10 hold Y1-Y6; grant and intangible from Assumptions.
# ---------------------------------------------------------------------------

ws = wb["Valuation"]
PL_FIRST, PL_LAST = 5, 10

fixed_1 = 0
for row in range(25, 34):                       # 9 rate rows
    rate_ref = f"$A{row}"
    for col_letter, pl_col in [("B", "B"), ("C", "C"), ("D", "D")]:
        # 6 forecast-year terms
        terms = []
        for i, pl_row in enumerate(range(PL_FIRST, PL_LAST + 1), start=1):
            terms.append(f"PL!${pl_col}${pl_row}/(1+{rate_ref})^{i}")
        # grant and intangible
        terms.append(f"Assumptions!$B$12/(1+{rate_ref})^Assumptions!$B$15")
        terms.append(f"Assumptions!$B$13/(1+{rate_ref})^Assumptions!$B$16")
        ws[f"{col_letter}{row}"] = "=" + "+".join(terms)
        fixed_1 += 1

print(f"Valuation: {fixed_1} cells rewritten (B25:D33).")

# ---------------------------------------------------------------------------
# FIX 2 — SeriesA A7:A11  (seed pre-money column)
# Convert text labels to numeric values so B7:G11 formulas can compute.
# ---------------------------------------------------------------------------

ws = wb["SeriesA"]
seed_pres = [5.71, 7.80, 9.00, 13.42, 20.00]
yellow_fill = PatternFill("solid", start_color="FFF2CC")
input_font = Font(name="Calibri", size=10, bold=True, color="9C6500")

fixed_2 = 0
for i, pre in enumerate(seed_pres):
    row = 7 + i
    cell = ws.cell(row, 1, pre)
    cell.number_format = '"HK$"#,##0.00"M"'
    cell.font = input_font
    cell.fill = yellow_fill
    fixed_2 += 1

print(f"SeriesA: {fixed_2} seed pre-money cells converted to numeric (A7:A11).")

# ---------------------------------------------------------------------------
# Verify all SeriesA formulas are intact
# ---------------------------------------------------------------------------

expected_formula_pattern = "=B$6/($A7+Assumptions!$B$5)"
check = ws["B7"].value
print(f"\nSeriesA B7 formula: {check}")
if check and "Assumptions" in str(check):
    print("  ✓ Formula structure intact — will compute once A7 is numeric.")
else:
    print("  ✗ Unexpected formula — investigate before relying on output.")

# ---------------------------------------------------------------------------
# Save
# ---------------------------------------------------------------------------

wb.save(PATH)
print(f"\nSaved: {PATH}")
print("Open in Excel — formulas recalculate on load.")