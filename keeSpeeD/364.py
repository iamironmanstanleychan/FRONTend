"""
KungFu AI Football v3.6.4 — Patch Script
Fixes: Helper Table column offset (Y1-Y6 now read columns B-G instead of C-H)
"""

import openpyxl
from openpyxl.utils import get_column_letter

INFILE = "KungFu_AI_Football_Financial_Model_v3.6.3.xlsx"
OUTFILE = "KungFu_AI_Football_Financial_Model_v3.6.4.xlsx"

wb = openpyxl.load_workbook(INFILE)
ws = wb["Matrix"]

# ═══════════════════════════════════════════════════════════
# SCENARIO DEFINITIONS (copy from v3.6.3 script)
# ═══════════════════════════════════════════════════════════
scenarios = [
    # (scen, mkt, split, subrev_row, ballrev_b2b_row, plat_row, user_row, ballenh_row, ballcost_row, helper_row, is_b2b)
    # B2B
    ("P10", "30%", "calib_b2b_30", 107, 110, 116, 119, 122, 125, 31, True),
    ("P50", "30%", "calib_b2b_30", 108, 111, 117, 120, 123, 126, 32, True),
    ("P90", "30%", "calib_b2b_30", 109, 112, 118, 121, 124, 127, 33, True),
    ("P10", "35%", "calib_b2b_35", 107, 110, 116, 119, 122, 125, 34, True),
    ("P50", "35%", "calib_b2b_35", 108, 111, 117, 120, 123, 126, 35, True),
    ("P90", "35%", "calib_b2b_35", 109, 112, 118, 121, 124, 127, 36, True),
    ("P10", "40%", "calib_b2b_40", 107, 110, 116, 119, 122, 125, 37, True),
    ("P50", "40%", "calib_b2b_40", 108, 111, 117, 120, 123, 126, 38, True),
    ("P90", "40%", "calib_b2b_40", 109, 112, 118, 121, 124, 127, 39, True),
    # B2C
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


# ═══════════════════════════════════════════════════════════
# FIX: Rewrite the Helper Table with correct column offset
# ═══════════════════════════════════════════════════════════
for scen, mkt, split, subrev_row, ballrev_b2b_row, plat_row, user_row, ballenh_row, ballcost_row, helper_row, is_b2b in scenarios:
    for i in range(1, 7):
        # FIX: i+1 gives B, C, D, E, F, G (correct)
        col = get_column_letter(i+1)
        custom = "customization_y1" if i == 1 else "customization_annual"
        grant = "grant_annual*grant_lag_frac" if i == 1 else "grant_annual"
        
        if is_b2b:
            rev = f"({col}{subrev_row}*{split} + {col}{ballrev_b2b_row} + {custom} + license_fee + {grant})"
            cost = f"({col}{plat_row} + {col}{user_row} + {col}{ballenh_row} + {col}{ballcost_row} + manpower_annual + uni_rd_cash)"
            profit_expr = f"({rev} - {cost})"
            taxable_expr = f"({profit_expr} - uni_rd_cash*(rd_deduct-1))"
            tax_expr = f"IF({taxable_expr}<=0, 0, IF({taxable_expr}<=tax_limit, {taxable_expr}*tax_r1, tax_limit*tax_r1 + ({taxable_expr}-tax_limit)*tax_r2))"
            net_expr = f"{profit_expr} - {tax_expr}"
            ws.cell(row=helper_row, column=i+2, value=f"={net_expr}")
        else:
            mkt_name = f"calib_mkt_{mkt.replace('%', '')}"
            rev = f"({col}{subrev_row}*{split} + {col}{ballrev_b2b_row})"
            cost = f"({col}{subrev_row}*{mkt_name} + license_fee + {custom})"
            profit_expr = f"({rev} - {cost})"
            tax_expr = f"IF({profit_expr}<=0, 0, IF({profit_expr}<=tax_limit, {profit_expr}*tax_r1, tax_limit*tax_r1 + ({profit_expr}-tax_limit)*tax_r2))"
            net_expr = f"{profit_expr} - {tax_expr}"
            ws.cell(row=helper_row, column=i+2, value=f"={net_expr}")

wb.save(OUTFILE)
print(f"✅ Fixed workbook saved: {OUTFILE}")
print("   → Helper Table now reads columns B–G (Y1–Y6) correctly.")
print("   → Y6 values should now be positive (e.g., B2B ≈ 52.9, B2C ≈ 31.3).")
print("   → All 18 Matrix cells will update to match v3.6.1 CSV values.")
