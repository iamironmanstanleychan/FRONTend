"""
KungFu AI Football — 6-Year Financial Model v3.1
Live workbook: VLOOKUP split + CHOOSE scenario + 3x3 matrix
Run: python generate_v3_1.py
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ═══════════════════════════════════════════════════════════
# STYLES
# ═══════════════════════════════════════════════════════════
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL  = PatternFill("solid", fgColor="FFF2CC")
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
LEFT   = Alignment(horizontal='left', vertical='center', wrap_text=True)


def set_name(wb, name, sheet, cell):
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!{cell}")


def hdr(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER


def row_style(ws, row, ncols, fill=None, bold=False):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = BOLD_NUM if bold else NUM_FONT
        cell.border = BORDER
        cell.alignment = LEFT if c == 1 else CENTER
        if fill: cell.fill = fill


def widths(ws, arr):
    for i, w in enumerate(arr, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ═══════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)

# =====================================================================
# 1. COVER
# =====================================================================
ws = wb.create_sheet("Cover")
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 75

ws['B2'] = "Kung Fu AI Football — 6-Year Financial Model v3.1"
ws['B2'].font = Font(bold=True, size=18)
ws.merge_cells('B2:C2')

info = [
    ("Type", "Live Excel model — formulas throughout, VLOOKUP + CHOOSE driven"),
    ("Base case", "P50 scenario at 30% B2C marketing"),
    ("Marketing control", "Assumptions!C17 — dropdown 30% / 35% / 40%"),
    ("Scenario control", "Assumptions!C28 — 1 = P10, 2 = P50, 3 = P90"),
    ("Split mechanism", "VLOOKUP on Assumptions rows 35–37 — auto-rebalances to preserve ~1.10 fairness"),
    ("Market sizing", "TAM 300m → SAM 13.89m → SOM ≈604,200 (P50, Y6)"),
    ("Matrix sheet", "All 9 combinations (3 marketing × 3 scenarios) computed via SUMPRODUCT"),
    ("Yellow cells", "INPUT — editable"),
    ("Green cells", "CALCULATED — do not overwrite"),
    ("Key checks", "Summary! row 14–22 — PASS/FAIL invariants"),
]
for i, (k, v) in enumerate(info, start=4):
    ws.cell(row=i, column=2, value=k).font = SUB_FONT
    ws.cell(row=i, column=2).fill = SUB_FILL
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=3, value=v).font = NUM_FONT
    ws.cell(row=i, column=3).border = BORDER
    ws.cell(row=i, column=3).alignment = LEFT

ws['B16'] = "HOW TO USE"
ws['B16'].font = SUB_FONT
howto = [
    "1. Open Assumptions — change marketing_pct (C17) to 30/35/40%",
    "2. Change scenario_selector (C28) to 1 (P10), 2 (P50), or 3 (P90)",
    "3. Every P&L line recalculates — B2B, B2C, Summary, Matrix",
    "4. Open Matrix to see all 9 combinations side by side",
    "5. Open Summary — PASS/FAIL checks confirm the model is internally consistent",
]
for i, t in enumerate(howto, start=17):
    ws.cell(row=i, column=2, value=t).font = NUM_FONT
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)


# =====================================================================
# 2. ASSUMPTIONS
# =====================================================================
ws = wb.create_sheet("Assumptions")
ws['A1'] = "Assumptions — INPUT SHEET"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:F1')
ws['A2'] = "Yellow = editable. Green = formula. Named ranges in column D."
ws['A2'].font = ITALIC
ws.merge_cells('A2:F2')

ws.append([])
ws.append(["Category", "Parameter", "Value", "Named range", "Unit", "Note"])
hdr(ws, 4, 6)

# Main inputs rows 5–28
inputs = [
    # (cat, param, value, name, unit, note, format, is_input)
    ("Subscription", "Coach/Player share",           0.50,  "split_coach",       "%",       "Contract",                        "0.00%",  True),
    ("Subscription", "B2B subscription share",       "=VLOOKUP(marketing_pct,$A$35:$C$37,2,FALSE)",
                                                            "split_b2b",         "%",       "AUTO from calibration",            "0.0000%",False),
    ("Subscription", "B2C subscription share",       "=VLOOKUP(marketing_pct,$A$35:$C$37,3,FALSE)",
                                                            "split_b2c",         "%",       "AUTO from calibration",            "0.0000%",False),
    ("Pricing",      "Weighted ARPU (annual)",       1592,  "arpu_annual",       "HK$/yr",  "Tier mix 60/25/12/3",              "#,##0",  True),
    ("Ball",         "Selling price",                980,   "ball_price",        "HK$",     "Price anchor",                     "#,##0",  True),
    ("Ball",         "Manufacturing cost",           250,   "ball_cost",         "HK$",     "B2B-borne",                        "#,##0",  True),
    ("Ball",         "B2B revenue share",            0.35,  "ball_share_b2b",    "%",       "Min for B2B ball-positive",        "0.00%",  True),
    ("Ball",         "Attach rate (new users)",      0.50,  "ball_attach",       "%",       "One ball per two new users",       "0.00%",  True),
    ("Platform",     "Base (fixed)",                 0.20,  "platform_base",     "HK$m/yr", "Model D Hybrid",                   "#,##0.00",True),
    ("Platform",     "Per-user compute",             2.00,  "platform_per_user", "HK$/user","Model D Hybrid",                   "#,##0.00",True),
    ("Platform",     "User enhancement",             5.00,  "user_enhancement",  "HK$/user","Proportional to SOM",              "#,##0.00",True),
    ("Platform",     "Ball enhancement",             10.00, "ball_enhancement",  "HK$/ball","Proportional to balls",            "#,##0.00",True),
    ("B2C",          "Marketing % of gross sub",     0.30,  "marketing_pct",     "%",       "DROPDOWN: 30/35/40%",              "0.00%",  True),
    ("B2C",          "Exclusive license fee",        0.20,  "license_fee",       "HK$m/yr", "Paid to B2B",                      "#,##0.00",True),
    ("Customization","Y1 fee",                       0.10,  "customization_y1",  "HK$m",    "Both B2B rev & B2C cost",          "#,##0.000",True),
    ("Customization","Y2-Y6 annual fee",             0.025, "customization_annual","HK$m",  "Hybrid structure",                 "#,##0.000",True),
    ("ITF",          "Manpower annual (gross)",      0.48,  "manpower_annual",   "HK$m/yr", "2 BSc × HK$20k × 12",              "#,##0.00",True),
    ("ITF",          "Grant coverage",               1.00,  "grant_pct",         "%",       "100% eligible",                    "0.00%",  True),
    ("ITF",          "Reimbursement lag",            0.5,   "grant_lag_frac",    "fraction","6 months",                         "0.00",   True),
    ("ITF",          "Y7 Q2 catch-up",               0.24,  "grant_y7q2",        "HK$m",    "Final reimbursement",              "#,##0.00",True),
    ("Scenario",     "P10 multiplier",               0.50,  "mult_p10",          "×",       "8.7 / 4",                          "0.00",   True),
    ("Scenario",     "P50 multiplier",               1.00,  "mult_p50",          "×",       "8.7 / 2",                          "0.00",   True),
    ("Scenario",     "P90 multiplier",               2.00,  "mult_p90",          "×",       "8.7 / 1",                          "0.00",   True),
    ("Scenario",     "Selector (1=P10, 2=P50, 3=P90)", 2,   "scenario_selector", "1-3",     "DROPDOWN",                         "#,##0",  True),
]

r = 5
for cat, param, val, name, unit, note, fmt, is_input in inputs:
    ws.cell(row=r, column=1, value=cat).font = NUM_FONT
    ws.cell(row=r, column=2, value=param).font = NUM_FONT
    c = ws.cell(row=r, column=3, value=val)
    c.number_format = fmt
    c.alignment = CENTER
    c.font = BOLD_NUM
    c.fill = INPUT_FILL if is_input else CALC_FILL
    ws.cell(row=r, column=4, value=name).font = ITALIC
    ws.cell(row=r, column=5, value=unit).font = ITALIC
    ws.cell(row=r, column=6, value=note).font = ITALIC
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
        if col != 3:
            ws.cell(row=r, column=col).alignment = LEFT if col in (2, 6) else CENTER
    set_name(wb, name, "Assumptions", f"$C${r}")
    r += 1

# Active multiplier — rows 30-31
ws.cell(row=30, column=1, value="ACTIVE SCENARIO").font = SUB_FONT
ws.cell(row=30, column=1).fill = SUB_FILL
ws.merge_cells('A30:F30')
ws.cell(row=31, column=2, value="Active scenario multiplier").font = NUM_FONT
c = ws.cell(row=31, column=3, value="=CHOOSE(scenario_selector,mult_p10,mult_p50,mult_p90)")
c.number_format = '0.00'; c.font = BOLD_NUM; c.fill = CALC_FILL; c.alignment = CENTER
ws.cell(row=31, column=4, value="scenario_multiplier").font = ITALIC
ws.cell(row=31, column=6, value="= 0.50 / 1.00 / 2.00").font = ITALIC
for col in range(1, 7): ws.cell(row=31, column=col).border = BORDER
set_name(wb, "scenario_multiplier", "Assumptions", "$C$31")

# Split calibration — rows 33-37
ws.cell(row=33, column=1, value="SPLIT CALIBRATION (drives VLOOKUP)").font = SUB_FONT
ws.cell(row=33, column=1).fill = SUB_FILL
ws.merge_cells('A33:F33')
ws.cell(row=34, column=1, value="Marketing")
ws.cell(row=34, column=2, value="B2B split")
ws.cell(row=34, column=3, value="B2C split")
ws.cell(row=34, column=4, value="Fairness (P50)")
ws.cell(row=34, column=5, value="Sum (must = 50%)")
ws.cell(row=34, column=6, value="Status")
hdr(ws, 34, 6)

calib = [
    (0.30, 0.1286, 0.3714, 1.0997),
    (0.35, 0.1048, 0.3952, 1.0990),
    (0.40, 0.0809, 0.4191, 1.0990),
]
for i, (mkt, b2b, b2c, fair) in enumerate(calib, start=35):
    c1 = ws.cell(row=i, column=1, value=mkt); c1.number_format = '0.00%'; c1.fill = INPUT_FILL
    c2 = ws.cell(row=i, column=2, value=b2b); c2.number_format = '0.0000%'; c2.fill = INPUT_FILL
    c3 = ws.cell(row=i, column=3, value=b2c); c3.number_format = '0.0000%'; c3.fill = INPUT_FILL
    c4 = ws.cell(row=i, column=4, value=fair); c4.number_format = '0.0000'; c4.fill = CALC_FILL
    c5 = ws.cell(row=i, column=5, value=f"=B{i}+C{i}"); c5.number_format = '0.0000%'; c5.fill = CALC_FILL
    c6 = ws.cell(row=i, column=6, value=f'=IF(ROUND(E{i},4)=0.5,"OK","ERR")')
    c6.font = BOLD_NUM
    for col in range(1, 7):
        ws.cell(row=i, column=col).border = BORDER
        ws.cell(row=i, column=col).alignment = CENTER

# Named ranges for calibration
set_name(wb, "calib_b2b_30", "Assumptions", "$B$35")
set_name(wb, "calib_b2b_35", "Assumptions", "$B$36")
set_name(wb, "calib_b2b_40", "Assumptions", "$B$37")
set_name(wb, "calib_b2c_30", "Assumptions", "$C$35")
set_name(wb, "calib_b2c_35", "Assumptions", "$C$36")
set_name(wb, "calib_b2c_40", "Assumptions", "$C$37")
set_name(wb, "calib_mkt_30", "Assumptions", "$A$35")
set_name(wb, "calib_mkt_35", "Assumptions", "$A$36")
set_name(wb, "calib_mkt_40", "Assumptions", "$A$37")

# Market sizing inputs — rows 39-46
ws.cell(row=39, column=1, value="MARKET SIZING INPUTS").font = SUB_FONT
ws.cell(row=39, column=1).fill = SUB_FILL
ws.merge_cells('A39:F39')
ws.cell(row=40, column=1, value="Year")
ws.cell(row=40, column=2, value="TAM (m)")
ws.cell(row=40, column=3, value="SAM (m)")
ws.cell(row=40, column=4, value="Base penetration (P50)")
ws.cell(row=40, column=5, value="Named ranges")
ws.cell(row=40, column=6, value="Source")
hdr(ws, 40, 6)

market = [
    ("Y1", 110.0, 5.55,  0.0024589),
    ("Y2", 110.0, 5.55,  0.0061473),
    ("Y3", 185.0, 8.95,  0.0214639),
    ("Y4", 235.0, 11.26, 0.0269024),
    ("Y5", 270.0, 12.80, 0.0332830),
    ("Y6", 300.0, 13.89, 0.0435000),
]
for i, (yr, tam, sam, pen) in enumerate(market, start=41):
    idx = i - 40
    ws.cell(row=i, column=1, value=yr).alignment = CENTER
    c = ws.cell(row=i, column=2, value=tam); c.number_format = '#,##0.0'; c.fill = INPUT_FILL
    c = ws.cell(row=i, column=3, value=sam); c.number_format = '#,##0.00'; c.fill = INPUT_FILL
    c = ws.cell(row=i, column=4, value=pen); c.number_format = '0.00000%'; c.fill = INPUT_FILL
    ws.cell(row=i, column=5, value=f"tam_Y{idx}, sam_Y{idx}, base_pen_Y{idx}").font = ITALIC
    ws.cell(row=i, column=6, value="TAM=79% fans; SAM=CFA+163.com; Pen=8.7/2 scaled").font = ITALIC
    for col in range(1, 7):
        ws.cell(row=i, column=col).border = BORDER
        if col != 1: ws.cell(row=i, column=col).alignment = CENTER
    set_name(wb, f"tam_Y{idx}", "Assumptions", f"$B${i}")
    set_name(wb, f"sam_Y{idx}", "Assumptions", f"$C${i}")
    set_name(wb, f"base_pen_Y{idx}", "Assumptions", f"$D${i}")

# Regional weights — rows 48-56
ws.cell(row=48, column=1, value="REGIONAL SAM WEIGHTS").font = SUB_FONT
ws.cell(row=48, column=1).fill = SUB_FILL
ws.merge_cells('A48:F48')
ws.cell(row=49, column=1, value="Region")
ws.cell(row=49, column=2, value="Weight")
ws.cell(row=49, column=3, value="Regional SAM (m)")
ws.cell(row=49, column=4, value="Rationale")
ws.merge_cells('D49:F49')
hdr(ws, 49, 6)

regions = [
    ("HK & South China", 0.400, 5.55, "FIFA Talent Academy; PSG HK girls' academy; Shenzhen youth centre"),
    ("East China",       0.244, 3.40, "Suzhou regional youth centre; dense school network"),
    ("Southwest",        0.167, 2.32, "Chengdu regional youth centre"),
    ("North",            0.111, 1.54, "Beijing national youth training centre"),
    ("Northwest",        0.044, 0.61, "Xi'an regional youth centre"),
    ("Northeast",        0.033, 0.46, "Dalian regional youth centre"),
]
for i, (reg, w, sam, note) in enumerate(regions, start=50):
    ws.cell(row=i, column=1, value=reg).alignment = LEFT
    c = ws.cell(row=i, column=2, value=w); c.number_format = '0.0%'; c.fill = INPUT_FILL
    c = ws.cell(row=i, column=3, value=sam); c.number_format = '#,##0.00'; c.fill = INPUT_FILL
    ws.cell(row=i, column=4, value=note).font = ITALIC
    ws.merge_cells(start_row=i, start_column=4, end_row=i, end_column=6)
    for col in range(1, 7):
        ws.cell(row=i, column=col).border = BORDER
        ws.cell(row=i, column=col).alignment = CENTER if col in (2,3) else LEFT

# Data validations
dv_mkt = DataValidation(type="list", formula1="=$A$35:$A$37", allow_blank=False,
                        promptTitle="Marketing %", prompt="30%, 35%, or 40%")
ws.add_data_validation(dv_mkt); dv_mkt.add("C17")

dv_sel = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False,
                        promptTitle="Scenario", prompt="1=P10, 2=P50, 3=P90")
ws.add_data_validation(dv_sel); dv_sel.add("C28")

widths(ws, [14, 38, 16, 22, 14, 55])


# =====================================================================
# 3. MARKET SIZING
# =====================================================================
ws = wb.create_sheet("Market Sizing")
ws['A1'] = "Market Sizing — TAM / SAM / SOM (all formulas)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:J1')
ws['A2'] = "SOM = SAM × 1,000,000 × base penetration × scenario multiplier"
ws['A2'].font = ITALIC
ws.merge_cells('A2:J2')

ws.append([])
ws.append(["Year", "Regions Active", "TAM (m)", "SAM (m)",
           "SOM P10", "SOM P50", "SOM P90", "Active SOM", "P50/SAM", "P50/TAM"])
hdr(ws, 4, 10)

regions_active = ["HK & South China", "HK & South China", "+ East China",
                  "+ Southwest", "+ North", "+ Northwest & Northeast"]

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=regions_active[i-1]).alignment = LEFT
    ws.cell(row=r, column=3, value=f"=tam_Y{i}").number_format = '#,##0.0'
    ws.cell(row=r, column=4, value=f"=sam_Y{i}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"=D{r}*1000000*base_pen_Y{i}*mult_p10").number_format = '#,##0'
    ws.cell(row=r, column=6, value=f"=D{r}*1000000*base_pen_Y{i}*mult_p50").number_format = '#,##0'
    ws.cell(row=r, column=7, value=f"=D{r}*1000000*base_pen_Y{i}*mult_p90").number_format = '#,##0'
    ws.cell(row=r, column=8, value=f"=D{r}*1000000*base_pen_Y{i}*scenario_multiplier").number_format = '#,##0'
    ws.cell(row=r, column=9, value=f"=F{r}/(D{r}*1000000)").number_format = '0.0000%'
    ws.cell(row=r, column=10, value=f"=F{r}/(C{r}*1000000)").number_format = '0.0000%'
    for c in range(1, 11):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 3:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

ws.cell(row=12, column=1, value="Active SOM (col H) = P10/P50/P90 based on scenario_selector on Assumptions").font = ITALIC
ws.merge_cells('A12:J12')
ws.cell(row=13, column=1, value="P10 = 8.7/4 = 2.175%; P50 = 8.7/2 = 4.35%; P90 = 8.7/1 = 8.7%").font = ITALIC
ws.merge_cells('A13:J13')
widths(ws, [6, 24, 11, 11, 12, 12, 12, 13, 11, 11])


# =====================================================================
# 4. B2B P&L
# =====================================================================
ws = wb.create_sheet("B2B P&L")
ws['A1'] = "B2B Technology Enabler — P&L (live formulas)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws['A2'] = '="Active: marketing = "&TEXT(marketing_pct,"0%")&" | scenario = "&CHOOSE(scenario_selector,"P10","P50","P90")'
ws['A2'].font = Font(bold=True, size=10, color="1F4E78")
ws.merge_cells('A2:P2')

ws.append([])
ws.append(["Year", "SOM", "Sub Rev", "Ball Rev", "Custom", "License", "ITF Grant",
           "Total Rev", "Platform", "User Enh", "Ball Enh", "Ball Cost",
           "Manpower", "Total Cost", "Profit", "Cumulative"])
hdr(ws, 4, 16)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!H{r}").number_format = '#,##0'
    ws.cell(row=r, column=3, value=f"=B{r}*arpu_annual*split_b2b/1000000").number_format = '#,##0.00'
    if i == 1:
        ws.cell(row=r, column=4, value=f"=B{r}*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=5, value="=customization_y1").number_format = '#,##0.000'
        ws.cell(row=r, column=7, value="=manpower_annual*grant_pct*grant_lag_frac").number_format = '#,##0.00'
        ws.cell(row=r, column=11, value=f"=B{r}*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=12, value=f"=B{r}*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=4, value=f"=(B{r}-B{r-1})*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=5, value="=customization_annual").number_format = '#,##0.000'
        ws.cell(row=r, column=7, value="=manpower_annual*grant_pct").number_format = '#,##0.00'
        ws.cell(row=r, column=11, value=f"=(B{r}-B{r-1})*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=12, value=f"=(B{r}-B{r-1})*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=6, value="=license_fee").number_format = '#,##0.000'
    ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=9, value=f"=platform_base+platform_per_user*B{r}/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=10, value=f"=user_enhancement*B{r}/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=13, value="=manpower_annual").number_format = '#,##0.00'
    ws.cell(row=r, column=14, value=f"=SUM(I{r}:M{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=15, value=f"=H{r}-N{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=16, value=f"=O{r}" if i == 1 else f"=P{r-1}+O{r}").number_format = '#,##0.00'
    for c in range(1, 17):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

# Y7 Q2 grant catch-up
r = 11
ws.cell(row=r, column=1, value="Y7 Q2").alignment = CENTER
for c in range(2, 7): ws.cell(row=r, column=c, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=7, value="=grant_y7q2").number_format = '#,##0.00'
ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
for c in range(9, 14): ws.cell(row=r, column=c, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=14, value=f"=SUM(I{r}:M{r})").number_format = '#,##0.00'
ws.cell(row=r, column=15, value=f"=H{r}-N{r}").number_format = '#,##0.00'
ws.cell(row=r, column=16, value=f"=P{r-1}+O{r}").number_format = '#,##0.00'
for c in range(1, 17):
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).font = NUM_FONT
    if c >= 2:
        ws.cell(row=r, column=c).fill = CALC_FILL
        ws.cell(row=r, column=c).alignment = CENTER

widths(ws, [8, 10, 9, 9, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 10, 11])


# =====================================================================
# 5. B2C P&L
# =====================================================================
ws = wb.create_sheet("B2C P&L")
ws['A1'] = "B2C Distribution Partners — P&L (live formulas)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = '="Active: marketing = "&TEXT(marketing_pct,"0%")&" | scenario = "&CHOOSE(scenario_selector,"P10","P50","P90")'
ws['A2'].font = Font(bold=True, size=10, color="1F4E78")
ws.merge_cells('A2:K2')

ws.append([])
ws.append(["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev",
           "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"])
hdr(ws, 4, 11)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!H{r}").number_format = '#,##0'
    ws.cell(row=r, column=3, value=f"=B{r}*arpu_annual*split_b2c/1000000").number_format = '#,##0.00'
    if i == 1:
        ws.cell(row=r, column=4, value=f"=B{r}*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=8, value="=customization_y1").number_format = '#,##0.000'
    else:
        ws.cell(row=r, column=4, value=f"=(B{r}-B{r-1})*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=8, value="=customization_annual").number_format = '#,##0.000'
    ws.cell(row=r, column=5, value=f"=C{r}+D{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=6, value=f"=B{r}*arpu_annual*marketing_pct/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=7, value="=license_fee").number_format = '#,##0.000'
    ws.cell(row=r, column=9, value=f"=SUM(F{r}:H{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=10, value=f"=E{r}-I{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=11, value=f"=J{r}" if i == 1 else f"=K{r-1}+J{r}").number_format = '#,##0.00'
    for c in range(1, 12):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

widths(ws, [8, 10, 10, 10, 10, 10, 10, 14, 10, 10, 11])


# =====================================================================
# 6. SUMMARY
# =====================================================================
ws = wb.create_sheet("Summary")
ws['A1'] = "Summary — Fairness Check + Invariants"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws['A2'] = "References B2B and B2C P&L (driven by dropdowns)."
ws['A2'].font = ITALIC
ws.merge_cells('A2:G2')

ws.append([])
ws.append(["Year", "B2B Profit", "B2C Profit", "Total Profit",
           "B2B Cumulative", "B2C Cumulative", "B2C / B2B"])
hdr(ws, 4, 7)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='B2B P&L'!O{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=3, value=f"='B2C P&L'!J{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=4, value=f"=B{r}+C{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"='B2B P&L'!P{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=6, value=f"='B2C P&L'!K{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=7, value=f'=IF(B{r}=0,"n/a",F{r}/B{r})').number_format = '0.0000'
    for c in range(1, 8):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

# Full-cycle row
r = 11
ws.cell(row=r, column=1, value="Full-cycle").alignment = CENTER
ws.cell(row=r, column=5, value="='B2B P&L'!P11").number_format = '#,##0.00'
ws.cell(row=r, column=6, value="='B2C P&L'!K10").number_format = '#,##0.00'
ws.cell(row=r, column=7, value=f"=F{r}/E{r}").number_format = '0.0000'
for c in range(1, 8):
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).font = BOLD_NUM
    if c >= 5:
        ws.cell(row=r, column=c).fill = CALC_FILL
        ws.cell(row=r, column=c).alignment = CENTER

# Checks
ws.cell(row=13, column=1, value="CHECKS").font = SUB_FONT
ws.cell(row=13, column=1).fill = SUB_FILL
ws.merge_cells('A13:G13')

checks = [
    ("B2B cumulative = sum of annual", '=IF(ROUND(\'B2B P&L\'!P10-SUM(\'B2B P&L\'!O5:O10),2)=0,"PASS","FAIL")'),
    ("B2C cumulative = sum of annual", '=IF(ROUND(\'B2C P&L\'!K10-SUM(\'B2C P&L\'!J5:J10),2)=0,"PASS","FAIL")'),
    ("Full-cycle fairness <= 1.10", '=IF(G11<=1.1,"PASS","FAIL")'),
    ("Full-cycle fairness >= 1.09", '=IF(G11>=1.09,"PASS","REVIEW")'),
    ("All years B2B non-negative", '=IF(MIN(\'B2B P&L\'!O5:O10)>=0,"PASS","FAIL")'),
    ("All years B2C non-negative", '=IF(MIN(\'B2C P&L\'!J5:J10)>=0,"PASS","FAIL")'),
    ("Subscription split sums to 100%", '=IF(ROUND(split_coach+split_b2b+split_b2c,4)=1,"PASS","FAIL")'),
    ("Marketing rate is calibrated (30/35/40)", '=IF(OR(marketing_pct=0.30,marketing_pct=0.35,marketing_pct=0.40),"PASS","FAIL")'),
    ("Scenario selector is valid (1/2/3)", '=IF(OR(scenario_selector=1,scenario_selector=2,scenario_selector=3),"PASS","FAIL")'),
]
for i, (label, formula) in enumerate(checks, start=14):
    ws.cell(row=i, column=1, value=label).font = NUM_FONT
    ws.cell(row=i, column=1).border = BORDER
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    c = ws.cell(row=i, column=4, value=formula)
    c.font = BOLD_NUM; c.border = BORDER; c.alignment = CENTER; c.fill = CALC_FILL
    ws.cell(row=i, column=4).number_format = 'General'

widths(ws, [10, 12, 12, 12, 14, 14, 12])


# =====================================================================
# 7. MATRIX (3x3 grids + helper arrays)
# =====================================================================
ws = wb.create_sheet("Matrix")
ws['A1'] = "Matrix — 3 Marketing Levels × 3 Scenarios (all 9 computed in parallel)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "Every cell is a self-contained SUMPRODUCT formula referencing the Helper Array section below."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

# Helper: B2B formula generator
def b2b_f(som_row, dsom_row, split_name):
    return (
        f"=SUMPRODUCT($B${som_row}:$G${som_row},arpu_annual*{split_name}/1000000)"
        f"+SUMPRODUCT($B${dsom_row}:$G${dsom_row},ball_attach*ball_price*ball_share_b2b/1000000)"
        f"+SUM($B$40:$G$40)+6*license_fee+SUM($B$41:$G$41)+grant_y7q2"
        f"-6*platform_base"
        f"-SUMPRODUCT($B${som_row}:$G${som_row},platform_per_user/1000000)"
        f"-SUMPRODUCT($B${som_row}:$G${som_row},user_enhancement/1000000)"
        f"-SUMPRODUCT($B${dsom_row}:$G${dsom_row},ball_attach*(ball_enhancement+ball_cost)/1000000)"
        f"-6*manpower_annual"
    )

# Helper: B2C formula generator
def b2c_f(som_row, dsom_row, split_name, mkt_name):
    return (
        f"=SUMPRODUCT($B${som_row}:$G${som_row},arpu_annual*{split_name}/1000000)"
        f"+SUMPRODUCT($B${dsom_row}:$G${dsom_row},ball_attach*ball_price*(1-ball_share_b2b)/1000000)"
        f"-SUMPRODUCT($B${som_row}:$G${som_row},arpu_annual*{mkt_name}/1000000)"
        f"-6*license_fee"
        f"-SUM($B$40:$G$40)"
    )

# B2B grid — rows 4-8
ws.cell(row=4, column=1, value="B2B FULL-CYCLE CUMULATIVE (Y6 + Y7 Q2 grant), HK$m").font = SUB_FONT
ws.cell(row=4, column=1).fill = SUB_FILL
ws.merge_cells('A4:E4')
ws.cell(row=5, column=1, value="Marketing \\ Scenario")
ws.cell(row=5, column=2, value="P10")
ws.cell(row=5, column=3, value="P50")
ws.cell(row=5, column=4, value="P90")
hdr(ws, 5, 4)

b2b_grid = [
    ("30%", "calib_b2b_30"),
    ("35%", "calib_b2b_35"),
    ("40%", "calib_b2b_40"),
]
for i, (label, split_name) in enumerate(b2b_grid, start=6):
    ws.cell(row=i, column=1, value=label).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    # P10 (som row 34, dsom row 37)
    ws.cell(row=i, column=2, value=b2b_f(34, 37, split_name)).number_format = '#,##0.00'
    # P50 (som row 35, dsom row 38)
    ws.cell(row=i, column=3, value=b2b_f(35, 38, split_name)).number_format = '#,##0.00'
    # P90 (som row 36, dsom row 39)
    ws.cell(row=i, column=4, value=b2b_f(36, 39, split_name)).number_format = '#,##0.00'
    for c in range(2, 5):
        ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=c).fill = CALC_FILL
        ws.cell(row=i, column=c).alignment = CENTER
        ws.cell(row=i, column=c).font = NUM_FONT

# B2C grid — rows 10-14
ws.cell(row=10, column=1, value="B2C Y6 CUMULATIVE, HK$m").font = SUB_FONT
ws.cell(row=10, column=1).fill = SUB_FILL
ws.merge_cells('A10:E10')
ws.cell(row=11, column=1, value="Marketing \\ Scenario")
ws.cell(row=11, column=2, value="P10")
ws.cell(row=11, column=3, value="P50")
ws.cell(row=11, column=4, value="P90")
hdr(ws, 11, 4)

b2c_grid = [
    ("30%", "calib_b2c_30", "calib_mkt_30"),
    ("35%", "calib_b2c_35", "calib_mkt_35"),
    ("40%", "calib_b2c_40", "calib_mkt_40"),
]
for i, (label, split_name, mkt_name) in enumerate(b2c_grid, start=12):
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

# Fairness grid — rows 16-20
ws.cell(row=16, column=1, value="FAIRNESS RATIO (B2C / B2B)").font = SUB_FONT
ws.cell(row=16, column=1).fill = SUB_FILL
ws.merge_cells('A16:E16')
ws.cell(row=17, column=1, value="Marketing \\ Scenario")
ws.cell(row=17, column=2, value="P10")
ws.cell(row=17, column=3, value="P50")
ws.cell(row=17, column=4, value="P90")
hdr(ws, 17, 4)

for i in range(18, 21):
    b2c_row = i - 6  # rows 12-14
    b2b_row = i - 12  # rows 6-8
    ws.cell(row=i, column=1, value=["30%", "35%", "40%"][i - 18]).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    for j, col in enumerate(['B', 'C', 'D'], 2):
        ws.cell(row=i, column=j, value=f"={col}{b2c_row}/{col}{b2b_row}").number_format = '0.0000'
        ws.cell(row=i, column=j).border = BORDER
        ws.cell(row=i, column=j).fill = CALC_FILL
        ws.cell(row=i, column=j).alignment = CENTER
        ws.cell(row=i, column=j).font = NUM_FONT

# Active cell note — row 22
ws.cell(row=22, column=1, value='="ACTIVE on P&L sheets: marketing = "&TEXT(marketing_pct,"0%")&", scenario = "&CHOOSE(scenario_selector,"P10","P50","P90")').font = Font(bold=True, size=10, color="1F4E78")
ws.merge_cells('A22:E22')

# Helper arrays — rows 30-41
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

# Rows 34-36: som P10/P50/P90
for i, (label, mult) in enumerate([("som_p10","mult_p10"),("som_p50","mult_p50"),("som_p90","mult_p90")], start=34):
    ws.cell(row=i, column=1, value=label).font = ITALIC
    for j in range(1, 7):
        col = get_column_letter(j+1)
        ws.cell(row=i, column=j+1, value=f"={col}32*1000000*{col}33*{mult}").number_format = '#,##0'

# Rows 37-39: dsom
for i, (label, som_row) in enumerate([("dsom_p10",34),("dsom_p50",35),("dsom_p90",36)], start=37):
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

ws.cell(row=43, column=1, value="Rows 34-39 feed the 3×3 grids above. Formulas use SUMPRODUCT over 6-element row arrays.").font = ITALIC
ws.merge_cells('A43:G43')

widths(ws, [22, 14, 14, 14, 14, 14, 14])


# =====================================================================
# 8. SOURCES
# =====================================================================
ws = wb.create_sheet("Sources")
ws['A1'] = "Sources — Reference URL Register"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:D1')
ws['A2'] = "Every key data point has a specific article-level URL (no home pages)."
ws['A2'].font = ITALIC
ws.merge_cells('A2:D2')

ws.append([])
ws.append(["Reference", "URL", "Data used", "Status"])
hdr(ws, 4, 4)

sources = [
    ("Hurun Quality of Life Report",
     "https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4",
     "Health spending No.2; +12% service experience; >2 hrs exercise",
     "Primary"),
    ("People's Daily Shanghai — Youth Sports Training",
     "http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html",
     "44.67% pay more for quality; RMB 253.1bn market; 57.07% renewal",
     "Primary"),
    ("163.com — 3200亿 football industry",
     "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html",
     "3.8亿 fans; 3200亿 industry; 487.6亿 training; 4.23万 schools; 1280万 students",
     "Primary"),
    ("Football Training Institution Report",
     "https://www.renrendoc.com/paper/415852129.html",
     "92% high-end renewal; 8% of students; 76% mass base; 55% mass renewal; 68% Gen-Z data; 75% coach credentials",
     "Industry research"),
    ("adidas Leagues Cup 26 Pro Ball",
     "https://www.adidas.com/us/leagues-cup-26-pro-ball/KE5154.html",
     "US$170 retail price",
     "Primary"),
    ("INSAIT JOY Smart Football",
     "https://insaitjoy.com/products/insait-joy-smart-football",
     "US$139, no subscription",
     "Primary"),
    ("INSAIT government procurement",
     "https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/",
     "RMB 1,745/unit smart ball; RMB 4,798/unit sensor",
     "Primary"),
    ("Catapult One review",
     "https://soccerstories.net/catapult-one-review/",
     "≈US$179/year, membership only, one pod per player",
     "Secondary review"),
    ("注册球员2740人 (163.com)",
     "https://www.163.com/dy/article/L8E6C26F0556BPXI.html",
     "8.7% pathway commitment; 10.92万 elite youth",
     "Primary"),
    ("足协数字年报 (china.com.cn)",
     "http://news.china.com.cn/2026-01/30/content_118308779.shtml",
     "109,220 elite youth; 979,544 amateur; 125,976 coaches",
     "Primary"),
]
for i, row in enumerate(sources, start=5):
    for j, val in enumerate(row, 1):
        cell = ws.cell(row=i, column=j, value=val)
        cell.font = NUM_FONT
        cell.border = BORDER
        cell.alignment = LEFT if j in (1, 3) else CENTER

widths(ws, [35, 55, 55, 18])


# ═══════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════
OUT = "KungFu_AI_Football_Financial_Model_v3.1.xlsx"
wb.save(OUT)
print(f"✅ Saved: {OUT}")
print(f"   Sheets: {wb.sheetnames}")
print(f"   Named ranges: {len(wb.defined_names)}")