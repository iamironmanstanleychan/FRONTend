"""
Fix v3.1 Matrix tab — replace SUMPRODUCT(range, scalar) with SUM(range)*scalar
Run: python fix_matrix_v3_1.py
Input:  KungFu_AI_Football_Financial_Model_v3.1.xlsx
Output: KungFu_AI_Football_Financial_Model_v3.1_fixed.xlsx
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
CALC_FILL   = PatternFill("solid", fgColor="E2EFDA")
SUB_FILL    = PatternFill("solid", fgColor="D9E1F2")
SUB_FONT    = Font(bold=True, size=10)
NUM_FONT    = Font(size=9)
BOLD_NUM    = Font(size=9, bold=True)
ITALIC      = Font(italic=True, size=8)
TITLE_FONT  = Font(bold=True, size=14)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def hdr(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER


def widths(ws, arr):
    for i, w in enumerate(arr, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ─── CORRECTED FORMULA BUILDERS: SUM(range) × scalar ───
def b2b_f(som_row, dsom_row, split_name):
    return (
        f"=SUM($B${som_row}:$G${som_row})*arpu_annual*{split_name}/1000000"
        f"+SUM($B${dsom_row}:$G${dsom_row})*ball_attach*ball_price*ball_share_b2b/1000000"
        f"+SUM($B$40:$G$40)+6*license_fee+SUM($B$41:$G$41)+grant_y7q2"
        f"-6*platform_base"
        f"-SUM($B${som_row}:$G${som_row})*platform_per_user/1000000"
        f"-SUM($B${som_row}:$G${som_row})*user_enhancement/1000000"
        f"-SUM($B${dsom_row}:$G${dsom_row})*ball_attach*(ball_enhancement+ball_cost)/1000000"
        f"-6*manpower_annual"
    )


def b2c_f(som_row, dsom_row, split_name, mkt_name):
    return (
        f"=SUM($B${som_row}:$G${som_row})*arpu_annual*{split_name}/1000000"
        f"+SUM($B${dsom_row}:$G${dsom_row})*ball_attach*ball_price*(1-ball_share_b2b)/1000000"
        f"-SUM($B${som_row}:$G${som_row})*arpu_annual*{mkt_name}/1000000"
        f"-6*license_fee"
        f"-SUM($B$40:$G$40)"
    )


# ─── LOAD WORKBOOK AND REBUILD MATRIX ───
INFILE = "KungFu_AI_Football_Financial_Model_v3.1.xlsx"
wb = openpyxl.load_workbook(INFILE)

if "Matrix" in wb.sheetnames:
    del wb["Matrix"]

ws = wb.create_sheet("Matrix")

# Title
ws['A1'] = "Matrix — 3 Marketing Levels × 3 Scenarios (all 9 computed in parallel)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "Formula pattern: SUM(range) × scalar — avoids SUMPRODUCT dimension errors."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

# ── B2B grid (rows 4–8) ──
ws.cell(row=4, column=1, value="B2B FULL-CYCLE CUMULATIVE (Y6 + Y7 Q2 grant), HK$m").font = SUB_FONT
ws.cell(row=4, column=1).fill = SUB_FILL
ws.merge_cells('A4:E4')
ws.cell(row=5, column=1, value="Marketing \\ Scenario")
ws.cell(row=5, column=2, value="P10")
ws.cell(row=5, column=3, value="P50")
ws.cell(row=5, column=4, value="P90")
hdr(ws, 5, 4)

for i, (label, split_name) in enumerate([
    ("30%", "calib_b2b_30"),
    ("35%", "calib_b2b_35"),
    ("40%", "calib_b2b_40"),
], start=6):
    ws.cell(row=i, column=1, value=label).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    ws.cell(row=i, column=2, value=b2b_f(34, 37, split_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=3, value=b2b_f(35, 38, split_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=4, value=b2b_f(36, 39, split_name)).number_format = '#,##0.00'
    for c in range(2, 5):
        ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=c).fill = CALC_FILL
        ws.cell(row=i, column=c).alignment = CENTER
        ws.cell(row=i, column=c).font = NUM_FONT

# ── B2C grid (rows 10–14) ──
ws.cell(row=10, column=1, value="B2C Y6 CUMULATIVE, HK$m").font = SUB_FONT
ws.cell(row=10, column=1).fill = SUB_FILL
ws.merge_cells('A10:E10')
ws.cell(row=11, column=1, value="Marketing \\ Scenario")
ws.cell(row=11, column=2, value="P10")
ws.cell(row=11, column=3, value="P50")
ws.cell(row=11, column=4, value="P90")
hdr(ws, 11, 4)

for i, (label, split_name, mkt_name) in enumerate([
    ("30%", "calib_b2c_30", "calib_mkt_30"),
    ("35%", "calib_b2c_35", "calib_mkt_35"),
    ("40%", "calib_b2c_40", "calib_mkt_40"),
], start=12):
    ws.cell(row=i, column=1, value=label).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    ws.cell(row=i, column=2, value=b2c_f(34, 37, split_name, mkt_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=3, value=b2c_f(35, 38, split_name, mkt_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=4, value=b2c_f(36, 39, split_name, mkt_name)).number_format = '#,##0.00'
    for c in range(2, 5):
        ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=c).fill = CALC_FILL
        ws.cell(row=i, column=c).alignment = CENTER
        ws.cell(row=i, column=c).font = NUM_FONT

# ── Fairness grid (rows 16–20) ──
ws.cell(row=16, column=1, value="FAIRNESS RATIO (B2C / B2B)").font = SUB_FONT
ws.cell(row=16, column=1).fill = SUB_FILL
ws.merge_cells('A16:E16')
ws.cell(row=17, column=1, value="Marketing \\ Scenario")
ws.cell(row=17, column=2, value="P10")
ws.cell(row=17, column=3, value="P50")
ws.cell(row=17, column=4, value="P90")
hdr(ws, 17, 4)

for i in range(18, 21):
    b2c_row = i - 6   # 12, 13, 14
    b2b_row = i - 12  # 6, 7, 8
    ws.cell(row=i, column=1, value=["30%", "35%", "40%"][i - 18]).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    for j, col in enumerate(['B', 'C', 'D'], 2):
        ws.cell(row=i, column=j, value=f"={col}{b2c_row}/{col}{b2b_row}").number_format = '0.0000'
        ws.cell(row=i, column=j).border = BORDER
        ws.cell(row=i, column=j).fill = CALC_FILL
        ws.cell(row=i, column=j).alignment = CENTER
        ws.cell(row=i, column=j).font = NUM_FONT

ws.cell(row=22, column=1,
        value='="ACTIVE on P&L sheets: marketing = "&TEXT(marketing_pct,"0%")&", scenario = "&CHOOSE(scenario_selector,"P10","P50","P90")'
        ).font = Font(bold=True, size=10, color="1F4E78")
ws.merge_cells('A22:E22')

# ── Helper arrays (rows 30–41) ──
ws.cell(row=30, column=1, value="HELPER ARRAYS — auto-computed, do not edit").font = SUB_FONT
ws.cell(row=30, column=1).fill = SUB_FILL
ws.merge_cells('A30:G30')
ws.cell(row=31, column=1, value="Array")
for j in range(1, 7):
    ws.cell(row=31, column=j+1, value=f"Y{j}")
hdr(ws, 31, 7)

# Row 32: sam
ws.cell(row=32, column=1, value="sam").font = ITALIC
for j in range(1, 7):
    ws.cell(row=32, column=j+1, value=f"='Market Sizing'!D{4+j}").number_format = '#,##0.00'

# Row 33: base_pen
ws.cell(row=33, column=1, value="base_pen").font = ITALIC
for j in range(1, 7):
    ws.cell(row=33, column=j+1, value=f"=base_pen_Y{j}").number_format = '0.00000%'

# Rows 34–36: som P10/P50/P90
for i, (label, mult) in enumerate([
    ("som_p10", "mult_p10"),
    ("som_p50", "mult_p50"),
    ("som_p90", "mult_p90"),
], start=34):
    ws.cell(row=i, column=1, value=label).font = ITALIC
    for j in range(1, 7):
        col = get_column_letter(j+1)
        ws.cell(row=i, column=j+1, value=f"={col}32*1000000*{col}33*{mult}").number_format = '#,##0'

# Rows 37–39: dsom (Y1 = same; Y2–Y6 = current − previous)
for i, (label, som_row) in enumerate([
    ("dsom_p10", 34),
    ("dsom_p50", 35),
    ("dsom_p90", 36),
], start=37):
    ws.cell(row=i, column=1, value=label).font = ITALIC
    ws.cell(row=i, column=2, value=f"=B{som_row}").number_format = '#,##0'
    for j in range(2, 7):
        col = get_column_letter(j+1)
        prev = get_column_letter(j)
        ws.cell(row=i, column=j+1, value=f"={col}{som_row}-{prev}{som_row}").number_format = '#,##0'

# Row 40: custom
ws.cell(row=40, column=1, value="custom").font = ITALIC
ws.cell(row=40, column=2, value="=customization_y1").number_format = '#,##0.000'
for j in range(2, 7):
    ws.cell(row=40, column=j+1, value="=customization_annual").number_format = '#,##0.000'

# Row 41: grant
ws.cell(row=41, column=1, value="grant").font = ITALIC
ws.cell(row=41, column=2, value="=manpower_annual*grant_pct*grant_lag_frac").number_format = '#,##0.00'
for j in range(2, 7):
    ws.cell(row=41, column=j+1, value="=manpower_annual*grant_pct").number_format = '#,##0.00'

# Style helper rows
for i in range(32, 42):
    for j in range(1, 8):
        ws.cell(row=i, column=j).border = BORDER
        ws.cell(row=i, column=j).font = NUM_FONT if j > 1 else ITALIC
        if j > 1:
            ws.cell(row=i, column=j).fill = CALC_FILL
            ws.cell(row=i, column=j).alignment = CENTER

ws.cell(row=43, column=1,
        value="Rows 34–39 feed the 3×3 grids above. Formula pattern: SUM(range) × scalar — no SUMPRODUCT dimension issues."
        ).font = ITALIC
ws.merge_cells('A43:G43')

widths(ws, [22, 14, 14, 14, 14, 14, 14])

# ── Save ──
OUT = "KungFu_AI_Football_Financial_Model_v3.1_fixed.xlsx"
wb.save(OUT)
print(f"✅ Fixed workbook saved: {OUT}")
print("   Matrix sheet rebuilt with SUM(range)*scalar formulas")
print("   All other sheets preserved unchanged")
print("   Open in Excel/LibreOffice to verify all 27 cells compute correctly")