"""
Kung Fu AI Football — Negotiation Model (.xlsx)
Requires: pip install openpyxl
Standalone. Outputs KungFu_Negotiation_Model.xlsx

Eight working sheets:
  README        — how to use, colour legend
  Assumptions   — all inputs (yellow cells)
  PL            — Y1-Y6 net profit, three cases
  Valuation     — Y0, Y7Q2, midpoint, rate sensitivity
  SeedLadder    — 20x return ladder across pre-money prices
  Structures    — single, two-seed, two-tranche outcomes
  SeriesA       — return matrix (seed price x Series A price)
  Scenarios     — S0-S5 global stress test
  Dashboard     — key outputs, live-linked

Note: formulas require Excel or LibreOffice to calculate.
      The file is set to full-recalc on open.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# ===========================================================================
# STYLE CONSTANTS
# ===========================================================================

NAVY       = "1F3864"
LIGHT_BLUE = "D9E2F3"
YELLOW     = "FFF2CC"
YELLOW_B   = "BF9000"
GREEN      = "E2EFDA"
GREEN_D    = "0F6B3C"
RED        = "C00000"
RED_LIGHT  = "FCE4E4"
GREY       = "F2F2F2"
WHITE      = "FFFFFF"

FONT_TITLE  = Font(name="Calibri", size=16, bold=True, color=WHITE)
FONT_H1     = Font(name="Calibri", size=11, bold=True, color=WHITE)
FONT_H2     = Font(name="Calibri", size=10, bold=True, color=NAVY)
FONT_BODY   = Font(name="Calibri", size=10)
FONT_INPUT  = Font(name="Calibri", size=10, bold=True, color="9C6500")
FONT_SMALL  = Font(name="Calibri", size=9, italic=True, color="595959")
FONT_KEY    = Font(name="Calibri", size=11, bold=True, color=NAVY)
FONT_RED    = Font(name="Calibri", size=9, bold=True, color=RED)

def fill(color):
    return PatternFill("solid", start_color=color)

THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)

# ===========================================================================
# NUMBER FORMATS
# ===========================================================================

FMT_MONEY    = '"HK$"#,##0.00'
FMT_MONEY_0  = '"HK$"#,##0'
FMT_PCT      = '0.00%'
FMT_PCT_1    = '0.0%'
FMT_MULT     = '0.00"×"'
FMT_SHARES   = '#,##0.00'
FMT_PRICE    = '"HK$"#,##0'

# ===========================================================================
# HELPERS
# ===========================================================================

def title_band(ws, text, ncols=10, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    c.font = FONT_TITLE
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 28

def subtitle(ws, text, ncols=10, row=2):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[row].height = 28

def section(ws, row, text, ncols=10):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row, 1, text)
    c.font = FONT_H1
    c.fill = fill(NAVY)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20

def columns(ws, row, headers, fill_color=LIGHT_BLUE):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row, col, h)
        c.font = FONT_H2
        c.fill = fill(fill_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER
    ws.row_dimensions[row].height = 26

def inp(ws, row, col, value, fmt=None):
    c = ws.cell(row, col, value)
    c.font = FONT_INPUT
    c.fill = fill(YELLOW)
    c.border = Border(
        left=Side(style="thin", color=YELLOW_B),
        right=Side(style="thin", color=YELLOW_B),
        top=Side(style="thin", color=YELLOW_B),
        bottom=Side(style="thin", color=YELLOW_B),
    )
    if fmt: c.number_format = fmt
    c.alignment = Alignment(horizontal="center", vertical="center")
    return c

def frm(ws, row, col, formula, fmt=None, bold=False, fill_color=None):
    c = ws.cell(row, col, formula)
    c.font = Font(name="Calibri", size=10, bold=bold) if not bold else FONT_H2
    if fmt: c.number_format = fmt
    if fill_color: c.fill = fill(fill_color)
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal="center", vertical="center")
    return c

def lbl(ws, row, col, text, bold=False):
    c = ws.cell(row, col, text)
    c.font = FONT_H2 if bold else FONT_BODY
    return c

def note(ws, row, col, text, ncols=10):
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col+ncols-1)
    c = ws.cell(row, col, text)
    c.font = FONT_SMALL
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[row].height = 18

def widths(ws, w):
    for i, width in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

# ===========================================================================
# WORKBOOK
# ===========================================================================

wb = Workbook()
wb.calculation.fullCalcOnLoad = True
wb.calculation.calcMode = "auto"

# ===========================================================================
# SHEET 1: README
# ===========================================================================

ws = wb.active
ws.title = "README"
title_band(ws, "KUNG FU AI FOOTBALL — NEGOTIATION MODEL", ncols=6)
subtitle(ws, "Formula-driven. Change yellow cells; everything else recalculates.", ncols=6)

r = 4
section(ws, r, "HOW TO USE", ncols=6); r += 1

instructions = [
    ("Yellow cells", "Inputs. Change these to test scenarios."),
    ("White cells", "Formulas. Do not edit; they recalculate from inputs."),
    ("Navy headers", "Section titles. Click any cell to see the formula in the formula bar."),
    ("", ""),
    ("To test a scenario:", "Change inputs on the Assumptions sheet first. Then check SeedLadder, Structures, SeriesA."),
    ("To find the implied rate:", "Valuation sheet has a rate-sensitivity table. Use Excel's Data → What-If → Goal Seek."),
    ("To model a new Series A price:", "Change the Series A row header on the SeriesA sheet (yellow)."),
    ("", ""),
    ("Workbook formula recalculation:", "Formulas require Excel or LibreOffice. If you open in a viewer, values may show 0 until the file is opened in Excel."),
    ("", ""),
    ("Six key numbers to remember:", "1. Post-money ceiling = Series A / required multiple"),
    ("", "2. Founder dilution = cheque × multiple / Series A price"),
    ("", "3. Two-tranche structure beats single cheque by ~3.75 points of dilution"),
    ("", "4. In downturn (S1/S3), do not raise a priced round"),
    ("", "5. Never accept committed dilution with optional capital"),
    ("", "6. Ask fund minimum position size before proposing tranches"),
]

for label, desc in instructions:
    if label:
        lbl(ws, r, 1, label, bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c = ws.cell(r, 2, desc)
    c.font = FONT_BODY
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    r += 1

r += 2
section(ws, r, "SHEET MAP", ncols=6); r += 1
sheet_map = [
    ("Assumptions", "All inputs. Everything references here."),
    ("PL", "Y1-Y6 net profit, three cases, plus Y7 Q2 grant."),
    ("Valuation", "Y0 adjusted, Y7 Q2, midpoint matrix, rate sensitivity."),
    ("SeedLadder", "20x return ladder across seed pre-money prices."),
    ("Structures", "Single cheque, two seeds, two tranches — side by side."),
    ("SeriesA", "Investor return matrix: seed price × Series A price."),
    ("Scenarios", "S0-S5 global stress test with required multiple."),
    ("Dashboard", "Live-linked summary of the key outputs."),
]
for name, desc in sheet_map:
    lbl(ws, r, 1, name, bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c = ws.cell(r, 2, desc)
    c.font = FONT_BODY
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    r += 1

r += 2
section(ws, r, "CLASSIFICATION", ncols=6); r += 1
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
c = ws.cell(r, 1, "INTERNAL — CONFIDENTIAL — NOT FOR DISTRIBUTION")
c.font = FONT_RED
c.fill = fill(RED_LIGHT)
c.alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[r].height = 22

widths(ws, [30, 22, 22, 22, 22, 22])

# ===========================================================================
# SHEET 2: ASSUMPTIONS
# ===========================================================================

ws = wb.create_sheet("Assumptions")
title_band(ws, "ASSUMPTIONS", ncols=6)
subtitle(ws, "Change yellow cells. Everything else references here.", ncols=6)

r = 4
section(ws, r, "SEED PARAMETERS", ncols=6); r += 1
seed_params = [
    ("Seed cheque (HK$M)", 1.0, FMT_MONEY),
    ("Required multiple (x)", 20, FMT_MULT),
    ("Share count (authorised)", 100, "0"),
    ("Tranche 2 pre-money (HK$M)", 12.0, FMT_MONEY),
    ("Series A reference (HK$M)", 134.23, FMT_MONEY),
]
seed_cells = {}
for label, val, fmt in seed_params:
    lbl(ws, r, 1, label)
    seed_cells[label] = (r, 2)
    inp(ws, r, 2, val, fmt)
    r += 1

r += 1
section(ws, r, "FINANCIAL LINE ITEMS", ncols=6); r += 1
fin_params = [
    ("Grant (Y7 Q2, HK$M)", 0.24, FMT_MONEY),
    ("Intangible Y6 (HK$M)", 0.58, FMT_MONEY),
    ("Intangible Y3 (HK$M)", 0.29, FMT_MONEY),
    ("Grant years from Y0", 6.25, "0.00"),
    ("Intangible years from Y0", 6, "0"),
]
for label, val, fmt in fin_params:
    lbl(ws, r, 1, label)
    inp(ws, r, 2, val, fmt)
    r += 1

r += 1
section(ws, r, "DISCOUNT RATES", ncols=6); r += 1
for label, val in [("Rate 1", 0.25), ("Rate 2", 0.50), ("Rate 3", 0.75)]:
    lbl(ws, r, 1, label)
    inp(ws, r, 2, val, FMT_PCT)
    r += 1

r += 1
section(ws, r, "FUND PARAMETERS", ncols=6); r += 1
fund_params = [
    ("Survival rate", 0.05, FMT_PCT),
    ("Fund size ($M)", 100, FMT_MONEY_0),
    ("Number of positions", 100, "0"),
]
fund_cells = {}
for label, val, fmt in fund_params:
    lbl(ws, r, 1, label)
    fund_cells[label] = (r, 2)
    inp(ws, r, 2, val, fmt)
    r += 1

r += 1
section(ws, r, "DERIVED VALUES (auto-calculated)", ncols=6); r += 1
derived_start = r
derived = [
    ("Post-money target", f"=B{seed_cells['Series A reference (HK$M)'][0]}/B{seed_cells['Required multiple (x)'][0]}", FMT_MONEY),
    ("Price per share at target", f"=B{r}*1000000/B{seed_cells['Share count (authorised)'][0]}", FMT_PRICE),
    ("Pre-money at target", f"=B{r}-B{seed_cells['Seed cheque (HK$M)'][0]}", FMT_MONEY),
    ("Stake at target", f"=B{seed_cells['Seed cheque (HK$M)'][0]}/B{r}", FMT_PCT),
    ("Fund absolute return per winner ($M)",
     f"=B{fund_cells['Fund size ($M)'][0]}/"
     f"(B{fund_cells['Number of positions'][0]}*B{fund_cells['Survival rate'][0]})",
     FMT_MONEY),
]
for label, formula, fmt in derived:
    lbl(ws, r, 1, label, bold=True)
    frm(ws, r, 2, formula, fmt, bold=True, fill_color=GREEN)
    r += 1

r += 1
note(ws, r, 1, "The derived values above update automatically when you change the seed parameters. "
               "They are the anchor for every other sheet.", ncols=6)

widths(ws, [38, 18, 18, 18, 18, 18])

# Add data validation on key inputs
dv_pct = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                         allow_blank=False, showErrorMessage=True,
                         errorTitle="Invalid", error="Enter a value between 0 and 1")
dv_pct.error = "Enter a percentage as a decimal between 0 and 1"
dv_pct.errorTitle = "Percentage out of range"
dv_pct.error = "Enter a value between 0 and 1 (i.e. 0.05 for 5%)"
dv_pct.promptTitle = "Percentage"
dv_pct.prompt = "Enter a decimal (e.g. 0.05 for 5%)"
ws.add_data_validation(dv_pct)
for r in [derived_start - 5, derived_start - 4, derived_start - 3]:  # rate cells
    dv_pct.add(ws.cell(r, 2))

# ===========================================================================
# SHEET 3: PL
# ===========================================================================

ws = wb.create_sheet("PL")
title_band(ws, "P&L PROJECTIONS", ncols=5)
subtitle(ws, "Net profit, HK$M, after all operating costs, HKAS 38 amortisation and HKSAR taxes. "
             "Change yellow cells to test alternative forecasts.", ncols=5)

r = 4
columns(ws, r, ["Year", "Lower", "Centre", "Ceiling"]); r += 1

pl_data = [
    ("Y1", 0.58, 2.08, 5.24),
    ("Y2", 2.01, 5.19, 12.61),
    ("Y3", 12.43, 31.04, 74.31),
    ("Y4", 17.25, 44.20, 107.63),
    ("Y5", 23.72, 61.05, 149.13),
    ("Y6", 33.74, 86.74, 211.79),
]
pl_first = r
for year, lo, ce, hi in pl_data:
    lbl(ws, r, 1, year, bold=True)
    inp(ws, r, 2, lo, FMT_MONEY)
    inp(ws, r, 3, ce, FMT_MONEY)
    inp(ws, r, 4, hi, FMT_MONEY)
    r += 1
pl_last = r - 1

# Grant row (references Assumptions)
lbl(ws, r, 1, "Y7 Q2 grant", bold=True)
frm(ws, r, 2, "=Assumptions!B12", FMT_MONEY)
frm(ws, r, 3, "=Assumptions!B12", FMT_MONEY)
frm(ws, r, 4, "=Assumptions!B12", FMT_MONEY)
grant_row = r
r += 1

# Cumulative Y1-Y6
lbl(ws, r, 1, "Cumulative Y1-Y6", bold=True)
frm(ws, r, 2, f"=SUM(B{pl_first}:B{pl_last})", FMT_MONEY, bold=True, fill_color=LIGHT_BLUE)
frm(ws, r, 3, f"=SUM(C{pl_first}:C{pl_last})", FMT_MONEY, bold=True, fill_color=LIGHT_BLUE)
frm(ws, r, 4, f"=SUM(D{pl_first}:D{pl_last})", FMT_MONEY, bold=True, fill_color=LIGHT_BLUE)
cumul_row = r
r += 1

# Cumulative + grant
lbl(ws, r, 1, "Cumulative + grant", bold=True)
frm(ws, r, 2, f"=B{cumul_row}+B{grant_row}", FMT_MONEY, bold=True, fill_color=GREEN)
frm(ws, r, 3, f"=C{cumul_row}+C{grant_row}", FMT_MONEY, bold=True, fill_color=GREEN)
frm(ws, r, 4, f"=D{cumul_row}+D{grant_row}", FMT_MONEY, bold=True, fill_color=GREEN)
total_row = r

r += 2
note(ws, r, 1, "Yellow cells are inputs. Blue and green cells are formulas. "
               "Cumulative + grant is the Y7 Q2 undiscounted value before adding the intangible.", ncols=5)

# Store row references for later sheets
PL_FIRST = pl_first
PL_LAST = pl_last
PL_CUMUL = cumul_row
PL_TOTAL = total_row

widths(ws, [22, 16, 16, 16, 16])

# ===========================================================================
# SHEET 4: VALUATION
# ===========================================================================

ws = wb.create_sheet("Valuation")
title_band(ws, "VALUATION — Y0, Y7 Q2, MIDPOINT", ncols=6)
subtitle(ws, "Y0 recomputed from P&L. Midpoint = average of Y0 and Y7 Q2. "
             "Rate sensitivity below for implied discount rate.", ncols=6)

r = 4
section(ws, r, "Y0 ADJUSTED — RECOMPUTED FROM P&L", ncols=6); r += 1
columns(ws, r, ["Case", "Rate 1 (25%)", "Rate 2 (50%)", "Rate 3 (75%)"]); r += 1

# Each Y0 cell = sum(P&L_t/(1+r)^t) + PV grant + PV intangible
# References: P&L columns B/C/D, Assumptions rates in B19/B20/B21
def y0_formula(pl_col, rate_cell):
    pl_range = f"PL!${pl_col}${PL_FIRST}:${pl_col}${PL_LAST}"
    row_start = f"PL!${pl_col}${PL_FIRST}"
    formula = (
        f"=SUMPRODUCT({pl_range}/(1+Assumptions!{rate_cell})^"
        f"(ROW({pl_range})-ROW({row_start})+1))"
        f"+Assumptions!$B$12/(1+Assumptions!{rate_cell})^Assumptions!$B$15"
        f"+Assumptions!$B$13/(1+Assumptions!{rate_cell})^Assumptions!$B$16"
    )
    return formula

y0_rows = {}
for case_label, pl_col in [("Lower", "B"), ("Centre", "C"), ("Ceiling", "D")]:
    lbl(ws, r, 1, case_label, bold=True)
    frm(ws, r, 2, y0_formula(pl_col, "$B$19"), FMT_MONEY)
    frm(ws, r, 3, y0_formula(pl_col, "$B$20"), FMT_MONEY, fill_color=LIGHT_BLUE)
    frm(ws, r, 4, y0_formula(pl_col, "$B$21"), FMT_MONEY)
    y0_rows[case_label] = r
    r += 1

r += 1
section(ws, r, "Y7 Q2 UNDISCOUNTED", ncols=6); r += 1
columns(ws, r, ["Case", "Cumulative + grant", "+ Intangible", "Y7 Q2 value", "Per share"]); r += 1

y7q2_rows = {}
for case_label, pl_col in [("Lower", "B"), ("Centre", "C"), ("Ceiling", "D")]:
    lbl(ws, r, 1, case_label, bold=True)
    frm(ws, r, 2, f"=PL!{pl_col}{PL_TOTAL}", FMT_MONEY)
    frm(ws, r, 3, "=Assumptions!B13", FMT_MONEY)
    frm(ws, r, 4, f"=B{r}+C{r}", FMT_MONEY, bold=True, fill_color=GREEN)
    frm(ws, r, 5, f"=D{r}*1000000/Assumptions!B7", FMT_PRICE)
    y7q2_rows[case_label] = r
    r += 1

r += 1
section(ws, r, "MIDPOINT MATRIX (D = average of C and B)", ncols=6); r += 1
columns(ws, r, ["Case", "Rate 1 (25%)", "Rate 2 (50%)", "Rate 3 (75%)", "Per share at 50%"]); r += 1

mid_rows = {}
for case_label in ["Lower", "Centre", "Ceiling"]:
    y0_row = y0_rows[case_label]
    y7_row = y7q2_rows[case_label]
    lbl(ws, r, 1, case_label, bold=True)
    frm(ws, r, 2, f"=(B{y0_row}+D{y7_row})/2", FMT_MONEY)
    frm(ws, r, 3, f"=(C{y0_row}+D{y7_row})/2", FMT_MONEY, bold=True, fill_color=GREEN)
    frm(ws, r, 4, f"=(D{y0_row}+D{y7_row})/2", FMT_MONEY)
    frm(ws, r, 5, f"=C{r}*1000000/Assumptions!B7", FMT_PRICE)
    mid_rows[case_label] = r
    r += 1

r += 1
section(ws, r, "RATE SENSITIVITY — for implied discount rate", ncols=6); r += 1
note(ws, r, 1, "Use Goal Seek: set target cell to midpoint, variable cell to Assumptions!B19. "
               "The implied rate appears below.", ncols=6); r += 1
columns(ws, r, ["Rate", "Lower Y0", "Centre Y0", "Ceiling Y0"]); r += 1

sens_start = r
for rate in [0.02, 0.05, 0.08, 0.10, 0.12, 0.15, 0.20, 0.25, 0.30]:
    inp(ws, r, 1, rate, FMT_PCT)
    for col_pl, col_out in [("B", 2), ("C", 3), ("D", 4)]:
        formula = (
            f"=SUMPRODUCT(PL!${col_pl}${PL_FIRST}:${col_pl}${PL_LAST}/"
            f"(1+$A{r})^(ROW(PL!${col_pl}${PL_FIRST}:${col_pl}${PL_LAST})"
            f"-ROW(PL!${col_pl}${PL_FIRST})+1))"
            f"+Assumptions!$B$12/(1+$A{r})^Assumptions!$B$15"
            f"+Assumptions!$B$13/(1+$A{r})^Assumptions!$B$16"
        )
        frm(ws, r, col_out, formula, FMT_MONEY)
    r += 1

widths(ws, [18, 18, 18, 18, 20, 12])

# ===========================================================================
# SHEET 5: SEED LADDER
# ===========================================================================

ws = wb.create_sheet("SeedLadder")
title_band(ws, "SEED PRICE LADDER", ncols=8)
subtitle(ws, "For each seed pre-money: what the fund gets for the cheque, what Series A "
             "price delivers 20×, and the founder's retained stake.", ncols=8)

r = 4
columns(ws, r, ["Label", "Seed pre-money", "Cheque", "Post-money",
                "HK$1M buys", "Multiple at SA ref", "Series A for 20×",
                "Shares (of 100)", "Founder retains"]); r += 1

ladder_start = r
ladder = [
    ("A — 20× at midpoint", 5.71),
    ("B — 20× at SA 176M", 7.80),
    ("C — 20× at SA 200M", 9.00),
    ("D — 10× at midpoint", 13.42),
]
for label, pre in ladder:
    lbl(ws, r, 1, label)
    inp(ws, r, 2, pre, FMT_MONEY)
    frm(ws, r, 3, "=Assumptions!B5", FMT_MONEY)
    frm(ws, r, 4, f"=B{r}+C{r}", FMT_MONEY)
    frm(ws, r, 5, f"=C{r}/D{r}", FMT_PCT, fill_color=GREEN)
    frm(ws, r, 6, f"=Assumptions!B9/D{r}", FMT_MULT, bold=True, fill_color=LIGHT_BLUE)
    frm(ws, r, 7, f"=Assumptions!B6*D{r}", FMT_MONEY)
    frm(ws, r, 8, f"=Assumptions!B7*E{r}", FMT_SHARES)
    frm(ws, r, 9, f"=Assumptions!B7-H{r}", FMT_SHARES)
    r += 1

# Add two blank custom rows
for i in range(2):
    lbl(ws, r, 1, f"Custom {i+1}")
    inp(ws, r, 2, None, FMT_MONEY)
    frm(ws, r, 3, "=Assumptions!B5", FMT_MONEY)
    frm(ws, r, 4, f"=IF(B{r}=\"\",\"\",B{r}+C{r})", FMT_MONEY)
    frm(ws, r, 5, f"=IF(B{r}=\"\",\"\",C{r}/D{r})", FMT_PCT)
    frm(ws, r, 6, f"=IF(B{r}=\"\",\"\",Assumptions!B9/D{r})", FMT_MULT)
    frm(ws, r, 7, f"=IF(B{r}=\"\",\"\",Assumptions!B6*D{r})", FMT_MONEY)
    frm(ws, r, 8, f"=IF(B{r}=\"\",\"\",Assumptions!B7*E{r})", FMT_SHARES)
    frm(ws, r, 9, f"=IF(B{r}=\"\",\"\",Assumptions!B7-H{r})", FMT_SHARES)
    r += 1

r += 1
note(ws, r, 1, "Column F (Multiple at SA ref) is the investor's return multiple if the Series A "
               "clears at the reference price on Assumptions. Column G shows the Series A "
               "price required for exactly 20×.", ncols=9)

# Conditional format on multiple column
ws.conditional_formatting.add(
    f"F{ladder_start}:F{ladder_start+len(ladder)-1}",
    CellIsRule(operator="greaterThanOrEqual", formula=["20"],
               fill=fill(GREEN))
)

widths(ws, [22, 16, 12, 14, 14, 16, 16, 16, 16])

# ===========================================================================
# SHEET 6: STRUCTURES
# ===========================================================================

ws = wb.create_sheet("Structures")
title_band(ws, "SEED STRUCTURE COMPARISON", ncols=11)
subtitle(ws, "Single cheque vs two seeds vs two tranches. Post-money differs only where "
             "prices differ between tranches.", ncols=11)

r = 4
columns(ws, r, ["Structure", "Cheque 1", "Pre 1", "Cheque 2", "Pre 2",
                "Total raised", "Stake 1", "Stake 2",
                "Combined dilution", "Founder retains", "Largest holder"]); r += 1

structures = [
    ("One cheque (HK$1.0M)",    1.0, 5.71,  0.0, 0,   0),
    ("One cheque (HK$0.5M)",    0.5, 6.21,  0.0, 0,   0),
    ("Two seeds (equal, simultaneous)", 0.5, 5.71, 0.5, 5.71, 0),
    ("Two tranches (staged)",   0.5, 6.21,  0.5, 12.0, 1),
    ("Lead + follower",         0.75, 5.71, 0.25, 5.71, 0),
]

struct_start = r
for label, ch1, pre1, ch2, pre2, tranche in structures:
    lbl(ws, r, 1, label, bold=True)
    inp(ws, r, 2, ch1, FMT_MONEY)
    inp(ws, r, 3, pre1, FMT_MONEY)
    inp(ws, r, 4, ch2 if ch2 else None, FMT_MONEY)
    inp(ws, r, 5, pre2 if pre2 else None, FMT_MONEY)

    # Total raised
    frm(ws, r, 6, f"=B{r}+IF(D{r}=\"\",0,D{r})", FMT_MONEY)

    # Stake 1 (against own post-money)
    frm(ws, r, 7, f"=B{r}/(B{r}+C{r})", FMT_PCT)

    # Stake 2 — depends on whether tranche (staged) or simultaneous
    if tranche:
        # On post-tranche-1 company
        frm(ws, r, 8, f"=IF(D{r}=\"\",\"\",D{r}/(D{r}+E{r}))", FMT_PCT)
        # Combined: 1 - (1-s1)*(1-s2)
        frm(ws, r, 9, f"=IF(D{r}=\"\",G{r},1-(1-G{r})*(1-H{r}))", FMT_PCT,
            bold=True, fill_color=GREEN)
        frm(ws, r, 10, f"=1-I{r}", FMT_PCT)
    else:
        # Simultaneous — both priced off the same post-money
        if ch2:
            frm(ws, r, 8, f"=D{r}/(D{r}+E{r})", FMT_PCT)
        else:
            frm(ws, r, 8, "—")
        frm(ws, r, 9, f"=IF(ISNUMBER(H{r}),G{r}+H{r},G{r})", FMT_PCT,
            bold=True, fill_color=GREEN)
        frm(ws, r, 10, f"=1-I{r}", FMT_PCT)

    # Largest holder
    if ch2:
        frm(ws, r, 11, f"=MAX(G{r},IF(ISNUMBER(H{r}),H{r},0))", FMT_PCT)
    else:
        frm(ws, r, 11, f"=G{r}", FMT_PCT)
    r += 1

r += 1
note(ws, r, 1, "For simultaneous rounds: both investors pay the same price per share, so "
               "post-money is shared. For staged tranches: tranche 2 prices against the "
               "post-tranche-1 company, so combined dilution < sum of individual stakes.", ncols=11)

r += 2
section(ws, r, "INTERPRETATION", ncols=11); r += 1
interpretations = [
    "Two simultaneous seeds give the founder 85.1% vs 85.1% for a single HK$1.0M — same dilution, but two 7.45% holders sit below the 10% board-seat threshold.",
    "Two tranches give the founder ~88.9% vs 85.1% for a single HK$1.0M — a saving of ~3.75 percentage points, because tranche 2 prices against post-pilot evidence.",
    "The tranche structure is only viable if the first cheque funds 9+ months of runway.",
    "Never accept a structure where dilution is committed but tranche 2 is optional — this is deferred funding with investor option, not a real commitment.",
]
for line in interpretations:
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
    c = ws.cell(r, 1, "•  " + line)
    c.font = FONT_BODY
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)
    ws.row_dimensions[r].height = 28
    r += 1

widths(ws, [28, 14, 12, 14, 12, 14, 12, 12, 16, 16, 14])

# ===========================================================================
# SHEET 7: SERIES A
# ===========================================================================

ws = wb.create_sheet("SeriesA")
title_band(ws, "SERIES A RETURN MATRIX", ncols=7)
subtitle(ws, "Investor's return multiple = Series A pre-money ÷ Seed post-money. "
             "Cells ≥ 20× highlighted green. Change the Series A row headers (yellow).", ncols=7)

r = 4
note(ws, r, 1, "Seed post-money = seed pre-money + cheque (from Assumptions). "
               "Green = fund reaches 20× breakeven at that Series A price.", ncols=7); r += 2

# Header row with SA prices as inputs
lbl(ws, r, 1, "Seed pre-money", bold=True)
sa_prices = [70, 100, 134.23, 200, 300, 500]
for i, sa in enumerate(sa_prices):
    inp(ws, r, i + 2, sa, FMT_MONEY)
r += 1
header_row = r - 1

# Rows: seed pre-money values
seed_pres = [5.71, 7.80, 9.00, 13.42, 20.00]
matrix_start = r
for pre in seed_pres:
    lbl(ws, r, 1, f"HK${pre:.2f}M", bold=True)
    for col, sa in enumerate(sa_prices, 2):
        col_l = get_column_letter(col)
        formula = f"={col_l}${header_row}/($A{r}+Assumptions!$B$5)"
        frm(ws, r, col, formula, FMT_MULT)
    r += 1
matrix_end = r - 1

r += 1
note(ws, r, 1, "Note: Series A pre-money in the header is the absolute pre-money the "
               "company clears at, not the post-money of the seed round.", ncols=7)

# Conditional formatting: 20x or higher = green
ws.conditional_formatting.add(
    f"B{matrix_start}:{get_column_letter(1+len(sa_prices))}{matrix_end}",
    CellIsRule(operator="greaterThanOrEqual", formula=["20"],
               fill=fill(GREEN))
)
ws.conditional_formatting.add(
    f"B{matrix_start}:{get_column_letter(1+len(sa_prices))}{matrix_end}",
    CellIsRule(operator="lessThan", formula=["10"],
               fill=fill(RED_LIGHT))
)

widths(ws, [18, 14, 14, 14, 14, 14, 14])

# ===========================================================================
# SHEET 8: SCENARIOS
# ===========================================================================

ws = wb.create_sheet("Scenarios")
title_band(ws, "GLOBAL SCENARIOS (S0–S5)", ncols=11)
subtitle(ws, "Six macro scenarios. Inputs (yellow) drive the dilution and pre-money "
             "columns. Survival and required multiple are portfolio-level parameters.", ncols=11)

r = 4
columns(ws, r, ["Scenario", "SA low", "SA base", "SA high", "Survival",
                "Required mult", "Pre at base", "Dilution at base",
                "Pre at low", "Dilution at low"]); r += 1

scenarios = [
    ("S0 Base",                134, 200, 400, 0.05, 20),
    ("S1 Extended winter",      70, 100, 150, 0.04, 25),
    ("S2 Tailwind",            200, 300, 450, 0.06, 15),
    ("S3 Contraction",          50,  70, 100, 0.03, 30),
    ("S4 Compute constraint",   90, 130, 180, 0.05, 22),
    ("S5 HK financial tailwind", 150, 250, 400, 0.06, 15),
]

for name, lo, base, hi, surv, mult in scenarios:
    lbl(ws, r, 1, name, bold=True)
    inp(ws, r, 2, lo, FMT_MONEY)
    inp(ws, r, 3, base, FMT_MONEY)
    inp(ws, r, 4, hi, FMT_MONEY)
    inp(ws, r, 5, surv, FMT_PCT)
    inp(ws, r, 6, mult, FMT_MULT)

    # Pre at base = SA_base / mult - cheque
    frm(ws, r, 7, f"=C{r}/F{r}-Assumptions!$B$5", FMT_MONEY, bold=True, fill_color=LIGHT_BLUE)
    # Dilution at base = cheque × mult / SA_base
    frm(ws, r, 8, f"=Assumptions!$B$5*F{r}/C{r}", FMT_PCT, bold=True, fill_color=LIGHT_BLUE)
    # Pre at low
    frm(ws, r, 9, f"=B{r}/F{r}-Assumptions!$B$5", FMT_MONEY)
    # Dilution at low
    frm(ws, r, 10, f"=Assumptions!$B$5*F{r}/B{r}", FMT_PCT)
    r += 1

r += 1
note(ws, r, 1, "The fund's required multiple rises when survival falls, because fewer winners "
               "must return the fund. In S1 (25×) and S3 (30×), dilution at the base Series A "
               "price rises to 25–43% — the seed price that the fund can pay collapses.", ncols=11)

widths(ws, [26, 12, 12, 12, 11, 13, 14, 16, 14, 16])

# ===========================================================================
# SHEET 9: DASHBOARD
# ===========================================================================

ws = wb.create_sheet("Dashboard")
title_band(ws, "DASHBOARD — KEY OUTPUTS", ncols=6)
subtitle(ws, "All cells reference other sheets. Change Assumptions and this updates.", ncols=6)

r = 4
section(ws, r, "CORE PARAMETERS", ncols=6); r += 1
params = [
    ("Seed cheque", "=Assumptions!B5", FMT_MONEY),
    ("Required multiple", "=Assumptions!B6", FMT_MULT),
    ("Series A reference", "=Assumptions!B9", FMT_MONEY),
    ("Post-money ceiling", "=Assumptions!B29", FMT_MONEY),
    ("Price per share at ceiling", "=Assumptions!B30", FMT_PRICE),
    ("Stake at midpoint",
     f"=Assumptions!B5/Valuation!C{mid_rows['Centre']}", FMT_PCT),
]
for label, formula, fmt in params:
    lbl(ws, r, 1, label, bold=True)
    frm(ws, r, 2, formula, fmt, bold=True, fill_color=GREEN)
    r += 1

r += 1
section(ws, r, "THE 20× LADDER", ncols=6); r += 1
columns(ws, r, ["Scenario", "Pre-money", "HK$1M buys", "Multiple at SA ref",
                "SA for 20×"]); r += 1
for i, (label, pre) in enumerate(ladder):
    src_row = ladder_start + i
    lbl(ws, r, 1, label)
    frm(ws, r, 2, f"=SeedLadder!B{src_row}", FMT_MONEY)
    frm(ws, r, 3, f"=SeedLadder!E{src_row}", FMT_PCT, bold=True, fill_color=GREEN)
    frm(ws, r, 4, f"=SeedLadder!F{src_row}", FMT_MULT, bold=True)
    frm(ws, r, 5, f"=SeedLadder!G{src_row}", FMT_MONEY)
    r += 1

r += 1
section(ws, r, "PROPOSED STRUCTURE — TWO TRANCHES", ncols=6); r += 1
tranche_start = r
tranche_data = [
    ("Tranche 1: cheque", "=Structures!B9", FMT_MONEY),
    ("Tranche 1: pre-money", "=Structures!C9", FMT_MONEY),
    ("Tranche 1: stake", "=Structures!G9", FMT_PCT),
    ("Tranche 2: cheque", "=Structures!D9", FMT_MONEY),
    ("Tranche 2: pre-money", "=Structures!E9", FMT_MONEY),
    ("Tranche 2: stake", "=Structures!H9", FMT_PCT),
    ("Combined dilution", "=Structures!I9", FMT_PCT),
    ("Founder retains", "=Structures!J9", FMT_PCT),
]
for label, formula, fmt in tranche_data:
    lbl(ws, r, 1, label, bold=True)
    frm(ws, r, 2, formula, fmt, bold=True,
        fill_color=GREEN if "Combined" in label or "Founder" in label else None)
    r += 1

r += 1
section(ws, r, "SCENARIO QUICK-REFERENCE", ncols=6); r += 1
columns(ws, r, ["Scenario", "SA base", "Required mult",
                "Pre at base", "Dilution at base"]); r += 1
scenario_start_row = r
for i, name in enumerate([s[0] for s in scenarios]):
    src_row = 5 + i   # first data row on Scenarios sheet
    lbl(ws, r, 1, name)
    frm(ws, r, 2, f"=Scenarios!C{src_row}", FMT_MONEY)
    frm(ws, r, 3, f"=Scenarios!F{src_row}", FMT_MULT)
    frm(ws, r, 4, f"=Scenarios!G{src_row}", FMT_MONEY)
    frm(ws, r, 5, f"=Scenarios!H{src_row}", FMT_PCT, bold=True, fill_color=LIGHT_BLUE)
    r += 1

r += 1
section(ws, r, "DECISION RULE", ncols=6); r += 1
decision_text = [
    ("S0 (base)", "Raise HK$1.0M; target scenario B or C from ladder."),
    ("S1 (winter)", "Raise HK$0.5M, or defer and use PolyU/BUD. Do not set a low anchor."),
    ("S2 (tailwind)", "Raise HK$1.0M at higher end; use tailwind to justify higher SA."),
    ("S3 (contraction)", "Do not raise a priced round. Retain earnings; use non-dilutive capital."),
    ("S4 (compute)", "Raise but ring-fence capex; model margin compression explicitly."),
    ("S5 (HK tailwind)", "Best environment — raise at higher end."),
]
for label, action in decision_text:
    lbl(ws, r, 1, label, bold=True)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c = ws.cell(r, 2, action)
    c.font = FONT_BODY
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
    r += 1

r += 1
note(ws, r, 1, "Governing equation: Founder dilution = (cheque × required multiple) ÷ Series A price. "
               "Dilution is fixed by the Series A price and fund multiple — not by the pre-money "
               "you negotiate.", ncols=6)

widths(ws, [30, 20, 20, 20, 20, 20])

# ===========================================================================
# SAVE
# ===========================================================================

wb.save("KungFu_Negotiation_Model.xlsx")
print("Saved: KungFu_Negotiation_Model.xlsx")
print(f"Sheets: {', '.join(s.title for s in wb)}")
print("Open in Excel — formulas will recalculate on load.")