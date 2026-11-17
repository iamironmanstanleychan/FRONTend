"""
Surgical Update for KungFu_AI_Football_Financial_Model_v3.6.1
1. Renames 'Matrix' to 'Matrix_old'
2. Inserts a new dynamic 'Matrix' tab with a safe Helper Table
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

INFILE = "KungFu_AI_Football_Financial_Model_v3.6.1.xlsx"
OUTFILE = "KungFu_AI_Football_Financial_Model_v3.6.3.xlsx"

# Load workbook
wb = openpyxl.load_workbook(INFILE)

# 1. Rename existing Matrix
if "Matrix" in wb.sheetnames:
    wb["Matrix"].title = "Matrix_old"

# 2. Create new Matrix at the position before Sources
sources_idx = wb.sheetnames.index("Sources") if "Sources" in wb.sheetnames else len(wb.sheetnames)
ws = wb.create_sheet("Matrix", index=sources_idx)

# ═══════════════════════════════════════════════════════════
# STYLES
# ═══════════════════════════════════════════════════════════
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
NUM_FONT = Font(size=9)
BOLD_NUM = Font(size=9, bold=True)
ITALIC = Font(italic=True, size=8)
TITLE_FONT = Font(bold=True, size=14)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

def apply_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER

def apply_style(ws, row, ncols, bold=False):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BOLD_NUM if bold else NUM_FONT
        cell.border = BORDER
        cell.alignment = CENTER if c > 1 else CENTER
        if c > 1:
            cell.fill = CALC_FILL

# ═══════════════════════════════════════════════════════════
# BUILD THE VISIBLE GRIDS (Rows 4-20)
# ═══════════════════════════════════════════════════════════
ws['A1'] = "Matrix — 3 Marketing Levels × 3 Scenarios (Net Profit)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "Dynamic formulas via Helper Table below. Matches P&L logic exactly."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

# B2B Grid Title
ws.cell(row=4, column=1, value="B2B FULL-CYCLE CUMULATIVE (Net Profit), HK$m").font = BOLD_NUM
ws.cell(row=4, column=1).fill = SUB_FILL
ws.merge_cells('A4:E4')

# B2B Headers
ws.cell(row=5, column=1, value="Marketing \\ Scenario")
ws.cell(row=5, column=2, value="P10")
ws.cell(row=5, column=3, value="P50")
ws.cell(row=5, column=4, value="P90")
apply_header(ws, 5, 4)

# B2C Grid Title
ws.cell(row=10, column=1, value="B2C Y6 CUMULATIVE (Net Profit), HK$m").font = BOLD_NUM
ws.cell(row=10, column=1).fill = SUB_FILL
ws.merge_cells('A10:E10')

# B2C Headers
ws.cell(row=11, column=1, value="Marketing \\ Scenario")
ws.cell(row=11, column=2, value="P10")
ws.cell(row=11, column=3, value="P50")
ws.cell(row=11, column=4, value="P90")
apply_header(ws, 11, 4)

# Fairness Grid Title
ws.cell(row=16, column=1, value="FAIRNESS RATIO (B2C / B2B)").font = BOLD_NUM
ws.cell(row=16, column=1).fill = SUB_FILL
ws.merge_cells('A16:E16')

# Fairness Headers
ws.cell(row=17, column=1, value="Marketing \\ Scenario")
ws.cell(row=17, column=2, value="P10")
ws.cell(row=17, column=3, value="P50")
ws.cell(row=17, column=4, value="P90")
apply_header(ws, 17, 4)

# ═══════════════════════════════════════════════════════════
# BASE INPUTS (Rows 100-127)
# ═══════════════════════════════════════════════════════════
ws.cell(row=99, column=1, value="BASE INPUTS (For Helper Table)").font = BOLD_NUM
ws.cell(row=99, column=1).fill = SUB_FILL
ws.merge_cells('A99:H99')

# Row 100: Year labels
ws.cell(row=100, column=1, value="Year")
for i in range(1, 7):
    ws.cell(row=100, column=i+1, value=f"Y{i}")

# Rows 101-127: Base data
base_data = [
    ("SOM P10", 101, "='Market Sizing'!{}5", None),
    ("SOM P50", 102, "='Market Sizing'!{}6", None),
    ("SOM P90", 103, "='Market Sizing'!{}7", None),
    ("dSOM P10", 104, "dSOM", 101),
    ("dSOM P50", 105, "dSOM", 102),
    ("dSOM P90", 106, "dSOM", 103),
    ("Base Sub Rev P10", 107, "sub", 101),
    ("Base Sub Rev P50", 108, "sub", 102),
    ("Base Sub Rev P90", 109, "sub", 103),
    ("Ball Rev B2B P10", 110, "ball_b2b", 104),
    ("Ball Rev B2B P50", 111, "ball_b2b", 105),
    ("Ball Rev B2B P90", 112, "ball_b2b", 106),
    ("Ball Rev B2C P10", 113, "ball_b2c", 104),
    ("Ball Rev B2C P50", 114, "ball_b2c", 105),
    ("Ball Rev B2C P90", 115, "ball_b2c", 106),
    ("Platform P10", 116, "plat", 101),
    ("Platform P50", 117, "plat", 102),
    ("Platform P90", 118, "plat", 103),
    ("User Enh P10", 119, "user", 101),
    ("User Enh P50", 120, "user", 102),
    ("User Enh P90", 121, "user", 103),
    ("Ball Enh P10", 122, "ball_enh", 104),
    ("Ball Enh P50", 123, "ball_enh", 105),
    ("Ball Enh P90", 124, "ball_enh", 106),
    ("Ball Cost P10", 125, "ball_cost", 104),
    ("Ball Cost P50", 126, "ball_cost", 105),
    ("Ball Cost P90", 127, "ball_cost", 106),
]

for label, row_idx, mode, src_row in base_data:
    ws.cell(row=row_idx, column=1, value=label).font = ITALIC
    
    for i in range(1, 7):
        col = get_column_letter(i+1)  # B, C, D, E, F, G
        prev_col = get_column_letter(i) if i > 1 else None
        
        # Source Market Sizing row: E5 for P10, F5 for P50, etc.
        # But we want 3 different columns for each scenario.
        # Simplification: we store SOM P10/P50/P90 as separate rows, each referencing Market Sizing col E/F/G
        if mode == "='Market Sizing'!{}5":
            # Not used here since we'll handle below
            pass
    
    # Let's rewrite more cleanly
    pass

# ─── REDO BASE INPUTS WITH DIRECT FORMULAS ───
# Clear previous attempts
for r in range(101, 128):
    for c in range(2, 8):
        ws.cell(row=r, column=c).value = None

# Row 101: SOM P10 (References 'Market Sizing' column E)
ws.cell(row=101, column=1, value="SOM P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=101, column=i+1, value=f"='Market Sizing'!E{4+i}").number_format = '#,##0'

# Row 102: SOM P50
ws.cell(row=102, column=1, value="SOM P50").font = ITALIC
for i in range(1, 7):
    ws.cell(row=102, column=i+1, value=f"='Market Sizing'!F{4+i}").number_format = '#,##0'

# Row 103: SOM P90
ws.cell(row=103, column=1, value="SOM P90").font = ITALIC
for i in range(1, 7):
    ws.cell(row=103, column=i+1, value=f"='Market Sizing'!G{4+i}").number_format = '#,##0'

# Row 104: dSOM P10
ws.cell(row=104, column=1, value="dSOM P10").font = ITALIC
ws.cell(row=104, column=2, value="=B101").number_format = '#,##0'
for i in range(2, 7):
    col = get_column_letter(i+1)
    prev_col = get_column_letter(i)
    ws.cell(row=104, column=i+1, value=f"={col}101-{prev_col}101").number_format = '#,##0'

# Row 105: dSOM P50
ws.cell(row=105, column=1, value="dSOM P50").font = ITALIC
ws.cell(row=105, column=2, value="=B102").number_format = '#,##0'
for i in range(2, 7):
    col = get_column_letter(i+1)
    prev_col = get_column_letter(i)
    ws.cell(row=105, column=i+1, value=f"={col}102-{prev_col}102").number_format = '#,##0'

# Row 106: dSOM P90
ws.cell(row=106, column=1, value="dSOM P90").font = ITALIC
ws.cell(row=106, column=2, value="=B103").number_format = '#,##0'
for i in range(2, 7):
    col = get_column_letter(i+1)
    prev_col = get_column_letter(i)
    ws.cell(row=106, column=i+1, value=f"={col}103-{prev_col}103").number_format = '#,##0'

# Row 107: Base Sub Rev P10 (SOM * ARPU / 1e6)
ws.cell(row=107, column=1, value="Base Sub Rev P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=107, column=i+1, value=f"={col}101*arpu_annual/1000000").number_format = '#,##0.00'

# Row 108: Base Sub Rev P50
ws.cell(row=108, column=1, value="Base Sub Rev P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=108, column=i+1, value=f"={col}102*arpu_annual/1000000").number_format = '#,##0.00'

# Row 109: Base Sub Rev P90
ws.cell(row=109, column=1, value="Base Sub Rev P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=109, column=i+1, value=f"={col}103*arpu_annual/1000000").number_format = '#,##0.00'

# Row 110: Ball Rev B2B P10
ws.cell(row=110, column=1, value="Ball Rev B2B P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=110, column=i+1, value=f"={col}104*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'

# Row 111: Ball Rev B2B P50
ws.cell(row=111, column=1, value="Ball Rev B2B P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=111, column=i+1, value=f"={col}105*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'

# Row 112: Ball Rev B2B P90
ws.cell(row=112, column=1, value="Ball Rev B2B P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=112, column=i+1, value=f"={col}106*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'

# Row 113: Ball Rev B2C P10
ws.cell(row=113, column=1, value="Ball Rev B2C P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=113, column=i+1, value=f"={col}104*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'

# Row 114: Ball Rev B2C P50
ws.cell(row=114, column=1, value="Ball Rev B2C P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=114, column=i+1, value=f"={col}105*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'

# Row 115: Ball Rev B2C P90
ws.cell(row=115, column=1, value="Ball Rev B2C P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=115, column=i+1, value=f"={col}106*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'

# Row 116: Platform P10
ws.cell(row=116, column=1, value="Platform P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=116, column=i+1, value=f"=platform_base+platform_per_user*{col}101/1000000").number_format = '#,##0.00'

# Row 117: Platform P50
ws.cell(row=117, column=1, value="Platform P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=117, column=i+1, value=f"=platform_base+platform_per_user*{col}102/1000000").number_format = '#,##0.00'

# Row 118: Platform P90
ws.cell(row=118, column=1, value="Platform P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=118, column=i+1, value=f"=platform_base+platform_per_user*{col}103/1000000").number_format = '#,##0.00'

# Row 119: User Enh P10
ws.cell(row=119, column=1, value="User Enh P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=119, column=i+1, value=f"=user_enhancement*{col}101/1000000").number_format = '#,##0.00'

# Row 120: User Enh P50
ws.cell(row=120, column=1, value="User Enh P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=120, column=i+1, value=f"=user_enhancement*{col}102/1000000").number_format = '#,##0.00'

# Row 121: User Enh P90
ws.cell(row=121, column=1, value="User Enh P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=121, column=i+1, value=f"=user_enhancement*{col}103/1000000").number_format = '#,##0.00'

# Row 122: Ball Enh P10
ws.cell(row=122, column=1, value="Ball Enh P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=122, column=i+1, value=f"={col}104*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'

# Row 123: Ball Enh P50
ws.cell(row=123, column=1, value="Ball Enh P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=123, column=i+1, value=f"={col}105*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'

# Row 124: Ball Enh P90
ws.cell(row=124, column=1, value="Ball Enh P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=124, column=i+1, value=f"={col}106*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'

# Row 125: Ball Cost P10
ws.cell(row=125, column=1, value="Ball Cost P10").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=125, column=i+1, value=f"={col}104*ball_attach*ball_cost/1000000").number_format = '#,##0.00'

# Row 126: Ball Cost P50
ws.cell(row=126, column=1, value="Ball Cost P50").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=126, column=i+1, value=f"={col}105*ball_attach*ball_cost/1000000").number_format = '#,##0.00'

# Row 127: Ball Cost P90
ws.cell(row=127, column=1, value="Ball Cost P90").font = ITALIC
for i in range(1, 7):
    col = get_column_letter(i+1)
    ws.cell(row=127, column=i+1, value=f"={col}106*ball_attach*ball_cost/1000000").number_format = '#,##0.00'

# Style the Base Input rows
for r in range(100, 128):
    apply_style(ws, r, 8, bold=(r in [100, 107, 108, 109]))

# ═══════════════════════════════════════════════════════════
# HELPER TABLE (Rows 30-80)
# ═══════════════════════════════════════════════════════════
ws.cell(row=29, column=1, value="HELPER TABLE (Year-by-Year Net Profit)").font = BOLD_NUM
ws.cell(row=29, column=1).fill = SUB_FILL
ws.merge_cells('A29:I29')

# Headers for Helper Table
headers = ["Scenario", "Mkt %", "Y1", "Y2", "Y3", "Y4", "Y5", "Y6", "Full-Cycle Total"]
for c, h in enumerate(headers, 1):
    cell = ws.cell(row=30, column=c, value=h)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = CENTER
    cell.border = BORDER

# Configuration for Helper Rows
# Format: (scenario_name, mult_col, mkt_pct_name, split_name, subrev_row, ballrev_b2b_row, ballrev_b2c_row, plat_row, user_row, ballenh_row, ballcost_row, is_b2b, helper_start_row)
scenarios = [
    # B2B scenarios: Rows 31-39
    ("P10", "30%", "calib_b2b_30", 107, 110, 116, 119, 122, 125, 31, True),
    ("P50", "30%", "calib_b2b_30", 108, 111, 117, 120, 123, 126, 32, True),
    ("P90", "30%", "calib_b2b_30", 109, 112, 118, 121, 124, 127, 33, True),
    ("P10", "35%", "calib_b2b_35", 107, 110, 116, 119, 122, 125, 34, True),
    ("P50", "35%", "calib_b2b_35", 108, 111, 117, 120, 123, 126, 35, True),
    ("P90", "35%", "calib_b2b_35", 109, 112, 118, 121, 124, 127, 36, True),
    ("P10", "40%", "calib_b2b_40", 107, 110, 116, 119, 122, 125, 37, True),
    ("P50", "40%", "calib_b2b_40", 108, 111, 117, 120, 123, 126, 38, True),
    ("P90", "40%", "calib_b2b_40", 109, 112, 118, 121, 124, 127, 39, True),
    
    # B2C scenarios: Rows 41-49
    ("P10", "30%", "calib_b2c_30", 107, 113, 0, 0, 0, 0, 41, False),
    ("P50", "30%", "calib_b2c_30", 108, 114, 0, 0, 0, 0, 42, False),
    ("P90", "30%", "calib_b2c_30", 109, 115, 0, 0, 0, 0, 43, False),
    ("P10", "35%", "calib_b2c_35", 107, 113, 0, 0, 0, 0, 44, False),
    ("P50", "35%", "calib_b2c_35", 108, 114, 0, 0, 0, 0, 45, False),
    ("P90", "35%", "calib_b2c_35", 109, 115, 0, 0, 0, 0, 46, False),
    ("P10", "40%", "calib_b2c_40", 107, 113, 0, 0, 0, 0, 47, False),
    ("P50", "40%", "calib_b2c_40", 108, 114, 0, 0, 0, 0, 48, False),
    ("P90", "40%", "calib_b2c_40", 109, 115, 0, 0, 0, 0, 49, False),
]

for scen, mkt, split, subrev_row, ballrev_b2b_row, plat_row, user_row, ballenh_row, ballcost_row, helper_row, is_b2b in scenarios:
    ws.cell(row=helper_row, column=1, value=scen).font = BOLD_NUM
    ws.cell(row=helper_row, column=2, value=mkt).font = BOLD_NUM
    
    for i in range(1, 7):
        col = get_column_letter(i+2)  # C, D, E, F, G, H
        custom = "customization_y1" if i == 1 else "customization_annual"
        grant = "grant_annual*grant_lag_frac" if i == 1 else "grant_annual"
        
        if is_b2b:
            rev = f"({col}{subrev_row}*{split} + {col}{ballrev_b2b_row} + {custom} + license_fee + {grant})"
            cost = f"({col}{plat_row} + {col}{user_row} + {col}{ballenh_row} + {col}{ballcost_row} + manpower_annual + uni_rd_cash)"
            profit_expr = f"({rev} - {cost})"
            taxable_expr = f"({profit_expr} - uni_rd_cash*(rd_deduct-1))"
            tax_expr = f"IF({taxable_expr}<=0, 0, IF({taxable_expr}<=tax_limit, {taxable_expr}*tax_r1, tax_limit*tax_r1 + ({taxable_expr}-tax_limit)*tax_r2))"
            net_expr = f"{profit_expr} - {tax_expr}"
            ws.cell(row=helper_row, column=i+2, value=f"={net_expr}").number_format = '#,##0.00'
        else:
            # B2C Formula
            # Marketing: subrev_row * calib_mkt (which is mkt_name)
            mkt_name = f"calib_mkt_{mkt.replace('%', '')}"
            rev = f"({col}{subrev_row}*{split} + {col}{ballrev_b2b_row})"  # ballrev_b2b_row here is Ball Rev B2C
            cost = f"({col}{subrev_row}*{mkt_name} + license_fee + {custom})"
            profit_expr = f"({rev} - {cost})"
            tax_expr = f"IF({profit_expr}<=0, 0, IF({profit_expr}<=tax_limit, {profit_expr}*tax_r1, tax_limit*tax_r1 + ({profit_expr}-tax_limit)*tax_r2))"
            net_expr = f"{profit_expr} - {tax_expr}"
            ws.cell(row=helper_row, column=i+2, value=f"={net_expr}").number_format = '#,##0.00'
    
    # Total column (I)
    if is_b2b:
        ws.cell(row=helper_row, column=9, value=f"=SUM(C{helper_row}:H{helper_row})+grant_y7q2").number_format = '#,##0.00'
    else:
        ws.cell(row=helper_row, column=9, value=f"=SUM(C{helper_row}:H{helper_row})").number_format = '#,##0.00'

# Style helper rows
for r in range(31, 50):
    apply_style(ws, r, 9)

# ═══════════════════════════════════════════════════════════
# MATRIX GRIDS (References to Helper Table)
# ═══════════════════════════════════════════════════════════
# B2B Matrix (rows 6-8)
for i, mkt_label in enumerate(["30%", "35%", "40%"]):
    row = 6 + i
    ws.cell(row=row, column=1, value=mkt_label).font = BOLD_NUM
    ws.cell(row=row, column=1).alignment = CENTER
    ws.cell(row=row, column=1).border = BORDER
    
    # P10, P50, P90 for this marketing level
    # Helper rows are 31-39 (B2B)
    # P10 30%: row 31, P50 30%: row 32, P90 30%: row 33
    # P10 35%: row 34, P50 35%: row 35, P90 35%: row 36
    # P10 40%: row 37, P50 40%: row 38, P90 40%: row 39
    
    base_helper_row = 31 + i * 3
    ws.cell(row=row, column=2, value=f"=I{base_helper_row}").number_format = '#,##0.00'
    ws.cell(row=row, column=3, value=f"=I{base_helper_row + 1}").number_format = '#,##0.00'
    ws.cell(row=row, column=4, value=f"=I{base_helper_row + 2}").number_format = '#,##0.00'
    
    for c in range(2, 5):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).fill = CALC_FILL
        ws.cell(row=row, column=c).alignment = CENTER
        ws.cell(row=row, column=c).font = NUM_FONT

# B2C Matrix (rows 12-14)
for i, mkt_label in enumerate(["30%", "35%", "40%"]):
    row = 12 + i
    ws.cell(row=row, column=1, value=mkt_label).font = BOLD_NUM
    ws.cell(row=row, column=1).alignment = CENTER
    ws.cell(row=row, column=1).border = BORDER
    
    # Helper rows 41-49 (B2C)
    base_helper_row = 41 + i * 3
    ws.cell(row=row, column=2, value=f"=I{base_helper_row}").number_format = '#,##0.00'
    ws.cell(row=row, column=3, value=f"=I{base_helper_row + 1}").number_format = '#,##0.00'
    ws.cell(row=row, column=4, value=f"=I{base_helper_row + 2}").number_format = '#,##0.00'
    
    for c in range(2, 5):
        ws.cell(row=row, column=c).border = BORDER
        ws.cell(row=row, column=c).fill = CALC_FILL
        ws.cell(row=row, column=c).alignment = CENTER
        ws.cell(row=row, column=c).font = NUM_FONT

# Fairness Matrix (rows 18-20)
for i, mkt_label in enumerate(["30%", "35%", "40%"]):
    row = 18 + i
    ws.cell(row=row, column=1, value=mkt_label).font = BOLD_NUM
    ws.cell(row=row, column=1).alignment = CENTER
    ws.cell(row=row, column=1).border = BORDER
    
    b2b_row = 6 + i
    b2c_row = 12 + i
    for j, col in enumerate(['B', 'C', 'D'], 2):
        ws.cell(row=row, column=j, value=f"={col}{b2c_row}/{col}{b2b_row}").number_format = '0.0000'
        ws.cell(row=row, column=j).border = BORDER
        ws.cell(row=row, column=j).fill = CALC_FILL
        ws.cell(row=row, column=j).alignment = CENTER
        ws.cell(row=row, column=j).font = NUM_FONT

# Column widths
ws.column_dimensions['A'].width = 22
for c in range(2, 10):
    ws.column_dimensions[get_column_letter(c)].width = 14

# Save
wb.save(OUTFILE)
print(f"✅ Saved: {OUTFILE}")
print(f"   Matrix tab rebuilt with dynamic Helper Table.")
print(f"   Old Matrix tab preserved as 'Matrix_old'.")
print(f"   All other tabs untouched.")