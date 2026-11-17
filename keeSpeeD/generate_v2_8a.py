"""
KungFu AI Football v2.8a — Live Workbook with Auto-Switching Split
Output: KungFu_AI_Football_v2.8a.xlsx
Key change: marketing_pct drives split_b2b and split_b2c automatically via VLOOKUP
Run: python generate_v2_8a.py
Requires: openpyxl
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ── STYLES ──
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
SUB_FONT = Font(bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
NUM_FONT = Font(size=9)
BOLD_NUM = Font(size=9, bold=True)
ITALIC = Font(italic=True, size=8)
TITLE_FONT = Font(bold=True, size=14)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)


def add_named(wb, name, sheet, cell):
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!{cell}")


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ═══════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ═══════════════════════════════════════════════════════════
# 1. ASSUMPTIONS
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Assumptions")
ws['A1'] = "Assumptions — INPUT SHEET (yellow = editable)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:F1')
ws['A2'] = "Change marketing_pct (30%/35%/40%) — the split auto-adjusts to preserve ~1.10 fairness."
ws['A2'].font = ITALIC
ws.merge_cells('A2:F2')

ws.append([])
ws.append(["Category", "Parameter", "Value", "Named range", "Unit", "Source / Note"])
style_header(ws, 4, 6)

# Main params — rows 5 to 28
main_params = [
    ("Subscription", "Coach/Player share", 0.50, "split_coach", "%", "Contract term", '0.00%'),
    ("Subscription", "B2B subscription share", "=VLOOKUP(marketing_pct,$A$34:$C$37,2,FALSE)", "split_b2b", "%", "AUTO from split calibration", '0.00%'),
    ("Subscription", "B2C subscription share", "=VLOOKUP(marketing_pct,$A$34:$C$37,3,FALSE)", "split_b2c", "%", "AUTO from split calibration", '0.00%'),
    ("Pricing", "Weighted ARPU (annual)", 1592, "arpu_annual", "HK$/year", "Tier mix 60/25/12/3", '#,##0'),
    ("Ball", "Selling price", 980, "ball_price", "HK$", "Price anchor", '#,##0'),
    ("Ball", "Manufacturing cost", 250, "ball_cost", "HK$", "Confirmed", '#,##0'),
    ("Ball", "B2B revenue share", 0.35, "ball_share_b2b", "%", "Minimum to keep B2B ball-positive", '0.00%'),
    ("Ball", "Attach rate (new users)", 0.50, "ball_attach", "%", "One ball per two new users", '0.00%'),
    ("Platform", "Base (fixed)", 0.20, "platform_base", "HK$m/yr", "Model D Hybrid", '#,##0.00'),
    ("Platform", "Per-user compute", 2.00, "platform_per_user", "HK$/user", "Model D Hybrid", '#,##0.00'),
    ("Platform", "User enhancement", 5.00, "user_enhancement", "HK$/user", "Proportional to SOM", '#,##0.00'),
    ("Platform", "Ball enhancement", 10.00, "ball_enhancement", "HK$/ball", "Proportional to balls sold", '#,##0.00'),
    ("B2C", "Marketing % of gross sub", 0.30, "marketing_pct", "%", "SELECT 30% / 35% / 40%", '0.00%'),
    ("B2C", "Exclusive license fee", 0.20, "license_fee", "HK$m/yr", "Paid to B2B", '#,##0.00'),
    ("Customization", "Y1 fee", 0.10, "customization_y1", "HK$m", "Both sides", '#,##0.000'),
    ("Customization", "Y2-Y6 annual fee", 0.025, "customization_annual", "HK$m", "Hybrid structure", '#,##0.000'),
    ("ITF", "Manpower annual (gross)", 0.48, "manpower_annual", "HK$m/yr", "2 BSc × HK$20k × 12", '#,##0.00'),
    ("ITF", "Grant coverage", 1.00, "grant_pct", "%", "100% eligible", '0.00%'),
    ("ITF", "Reimbursement lag (fraction)", 0.5, "grant_lag_frac", "fraction", "6 months", '0.00'),
    ("ITF", "Y7 Q2 catch-up", 0.24, "grant_y7q2", "HK$m", "2 × 20k × 6 months", '#,##0.00'),
    ("Scenario", "P10 multiplier", 0.50, "mult_p10", "×", "Lower = 8.7/4", '0.00'),
    ("Scenario", "P50 multiplier", 1.00, "mult_p50", "×", "Mean = 8.7/2", '0.00'),
    ("Scenario", "P90 multiplier", 2.00, "mult_p90", "×", "Upper = 8.7/1", '0.00'),
    ("Scenario", "Selector (1=P10, 2=P50, 3=P90)", 2, "scenario_selector", "1-3", "Drives main P&L", '#,##0'),
]

r = 5
for cat, param, val, name, unit, src, fmt in main_params:
    ws.cell(row=r, column=1, value=cat).font = NUM_FONT
    ws.cell(row=r, column=2, value=param).font = NUM_FONT
    c = ws.cell(row=r, column=3, value=val)
    c.number_format = fmt
    c.font = BOLD_NUM if isinstance(val, (int, float)) else NUM_FONT
    c.alignment = CENTER
    if isinstance(val, str) and val.startswith('='):
        c.fill = CALC_FILL
    else:
        c.fill = INPUT_FILL
    ws.cell(row=r, column=4, value=name).font = ITALIC
    ws.cell(row=r, column=5, value=unit).font = ITALIC
    ws.cell(row=r, column=6, value=src).font = ITALIC
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
        if col != 3:
            ws.cell(row=r, column=col).alignment = LEFT if col in (2, 6) else CENTER
    add_named(wb, name, "Assumptions", f"$C${r}")
    r += 1

# Active scenario multiplier — row 31
ws.cell(row=30, column=1, value="ACTIVE SCENARIO").font = SUB_FONT
ws.cell(row=30, column=1).fill = SUB_FILL
ws.merge_cells('A30:F30')
ws.cell(row=31, column=2, value="Active scenario multiplier").font = NUM_FONT
c = ws.cell(row=31, column=3, value="=CHOOSE(scenario_selector,mult_p10,mult_p50,mult_p90)")
c.fill = CALC_FILL; c.font = BOLD_NUM; c.alignment = CENTER
c.number_format = '0.00'
ws.cell(row=31, column=4, value="scenario_multiplier").font = ITALIC
ws.cell(row=31, column=6, value="= 0.50 / 1.00 / 2.00").font = ITALIC
for col in range(1, 7):
    ws.cell(row=31, column=col).border = BORDER
add_named(wb, "scenario_multiplier", "Assumptions", "$C$31")

# SPLIT CALIBRATION — rows 33 to 37
ws.cell(row=33, column=1, value="SPLIT CALIBRATION — auto-switches with marketing_pct").font = SUB_FONT
ws.cell(row=33, column=1).fill = SUB_FILL
ws.merge_cells('A33:F33')
for j, h in enumerate(["Marketing", "B2B Sub Share", "B2C Sub Share", "Fairness (P50)", "Sum (must = 50%)", "Status"], 1):
    c = ws.cell(row=34, column=j, value=h)
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

split_calib = [
    (0.30, 0.1286, 0.3714, 1.0997),
    (0.35, 0.1048, 0.3952, 1.0990),
    (0.40, 0.0809, 0.4191, 1.0990),
]
r = 35
for mkt, b2b, b2c, fair in split_calib:
    c = ws.cell(row=r, column=1, value=mkt); c.number_format = '0.0%'; c.fill = INPUT_FILL
    c.font = BOLD_NUM; c.alignment = CENTER
    c = ws.cell(row=r, column=2, value=b2b); c.number_format = '0.0000%'; c.fill = INPUT_FILL
    c.font = BOLD_NUM; c.alignment = CENTER
    c = ws.cell(row=r, column=3, value=b2c); c.number_format = '0.0000%'; c.fill = INPUT_FILL
    c.font = BOLD_NUM; c.alignment = CENTER
    c = ws.cell(row=r, column=4, value=fair); c.number_format = '0.0000'; c.fill = CALC_FILL
    c.alignment = CENTER
    c = ws.cell(row=r, column=5, value=f"=B{r}+C{r}"); c.number_format = '0.0000%'
    c.fill = CALC_FILL; c.alignment = CENTER
    c = ws.cell(row=r, column=6, value=f'=IF(ROUND(E{r},4)=0.5,"OK","ERROR")')
    c.font = BOLD_NUM; c.alignment = CENTER
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
    r += 1

# MARKET SIZING — rows 39 to 46
ws.cell(row=39, column=1, value="MARKET SIZING INPUTS").font = SUB_FONT
ws.cell(row=39, column=1).fill = SUB_FILL
ws.merge_cells('A39:F39')
for j, h in enumerate(["Year", "TAM (m)", "SAM (m)", "Base penetration (P50)", "Named ranges", "Source"], 1):
    c = ws.cell(row=40, column=j, value=h)
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

market = [
    ("Y1", 110.0, 5.59, 0.0024589),
    ("Y2", 110.0, 5.59, 0.0061473),
    ("Y3", 185.0, 9.01, 0.0214639),
    ("Y4", 235.0, 11.34, 0.0269024),
    ("Y5", 270.0, 12.89, 0.0332830),
    ("Y6", 300.0, 13.98, 0.0435000),
]
r = 41
for i, (yr, tam, sam, pen) in enumerate(market, start=1):
    ws.cell(row=r, column=1, value=yr).alignment = CENTER
    c = ws.cell(row=r, column=2, value=tam); c.fill = INPUT_FILL
    c.number_format = '#,##0.0'; c.alignment = CENTER
    c = ws.cell(row=r, column=3, value=sam); c.fill = INPUT_FILL
    c.number_format = '#,##0.00'; c.alignment = CENTER
    c = ws.cell(row=r, column=4, value=pen); c.fill = INPUT_FILL
    c.number_format = '0.00000%'; c.alignment = CENTER
    ws.cell(row=r, column=5, value=f"tam_Y{i}, sam_Y{i}, base_pen_Y{i}").font = ITALIC
    ws.cell(row=r, column=6, value="TAM=79% fans; SAM=CFA+163.com; Pen=8.7/2 scaled").font = ITALIC
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
    add_named(wb, f"tam_Y{i}", "Assumptions", f"$B${r}")
    add_named(wb, f"sam_Y{i}", "Assumptions", f"$C${r}")
    add_named(wb, f"base_pen_Y{i}", "Assumptions", f"$D${r}")
    r += 1

# REGIONAL WEIGHTS — rows 48 to 56
ws.cell(row=48, column=1, value="REGIONAL SAM WEIGHTS").font = SUB_FONT
ws.cell(row=48, column=1).fill = SUB_FILL
ws.merge_cells('A48:F48')
for j, h in enumerate(["Region", "Weight", "Regional SAM (m)", "Rationale", "", ""], 1):
    c = ws.cell(row=49, column=j, value=h)
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

regions = [
    ("HK & South China", 0.400, 5.59, "FIFA Talent Academy; PSG HK girls; Shenzhen youth centre"),
    ("East China", 0.244, 3.42, "Suzhou regional youth centre; dense school network"),
    ("Southwest", 0.167, 2.33, "Chengdu regional youth centre"),
    ("North", 0.111, 1.55, "Beijing national youth training centre"),
    ("Northwest", 0.044, 0.62, "Xi'an regional youth centre"),
    ("Northeast", 0.033, 0.47, "Dalian regional youth centre"),
]
r = 50
for reg, w, sam, note in regions:
    ws.cell(row=r, column=1, value=reg).alignment = LEFT
    c = ws.cell(row=r, column=2, value=w); c.fill = INPUT_FILL
    c.number_format = '0.0%'; c.alignment = CENTER
    ws.cell(row=r, column=3, value=sam).number_format = '#,##0.00'
    ws.cell(row=r, column=3).alignment = CENTER
    ws.cell(row=r, column=4, value=note).font = ITALIC
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
    r += 1

# Data validations
# marketing_pct dropdown — C17
dv_mkt = DataValidation(type="list", formula1="$A$35:$A$37", allow_blank=False,
                        promptTitle="Marketing %",
                        prompt="Select 30%, 35%, or 40%. Split auto-adjusts.")
ws.add_data_validation(dv_mkt)
dv_mkt.add("C17")

# scenario_selector dropdown — C28
dv_sel = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False,
                        promptTitle="Scenario",
                        prompt="1=P10 (Lower), 2=P50 (Mean), 3=P90 (Upper)")
ws.add_data_validation(dv_sel)
dv_sel.add("C28")

set_widths(ws, [14, 38, 14, 22, 14, 60])

# ═══════════════════════════════════════════════════════════
# 2. MARKET SIZING
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Market Sizing")
ws['A1'] = "Market Sizing — TAM / SAM / SOM (calculated)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:J1')
ws['A2'] = "All cells are formulas. Edit inputs on Assumptions."
ws['A2'].font = ITALIC
ws.merge_cells('A2:J2')

ws.append([])
ws.append(["Year", "Regions Active", "TAM (m)", "SAM (m)",
           "SOM P10", "SOM P50", "SOM P90", "Active SOM", "P50/SAM", "P50/TAM"])
style_header(ws, 4, 10)

regions_active = ["HK & South China", "HK & South China", "+ East China",
                  "+ Southwest", "+ North", "+ Northwest & Northeast"]

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=regions_active[i-1]).alignment = LEFT
    ws.cell(row=r, column=3, value=f"=tam_Y{i}").number_format = '#,##0.0'
    ws.cell(row=r, column=4, value=f"=sam_Y{i}").number_format = '#,##0.00'
    ws.cell(row=r, column=5,
            value=f"=C{r}*1000000*base_pen_Y{i}*mult_p10").number_format = '#,##0'
    ws.cell(row=r, column=6,
            value=f"=D{r}*1000000*base_pen_Y{i}*mult_p50").number_format = '#,##0'
    ws.cell(row=r, column=7,
            value=f"=D{r}*1000000*base_pen_Y{i}*mult_p90").number_format = '#,##0'
    # Active SOM — driven by scenario_selector
    ws.cell(row=r, column=8,
            value=f"=IF(scenario_selector=1,E{r},IF(scenario_selector=2,F{r},G{r}))").number_format = '#,##0'
    ws.cell(row=r, column=9, value=f"=F{r}/(D{r}*1000000)").number_format = '0.0000%'
    ws.cell(row=r, column=10, value=f"=F{r}/(C{r}*1000000)").number_format = '0.0000%'
    for c in range(1, 11):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 3:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

set_widths(ws, [8, 24, 12, 12, 14, 14, 14, 14, 12, 12])

ws.cell(row=12, column=1, value="SOM = SAM × 1,000,000 × base_penetration × scenario_multiplier").font = ITALIC
ws.merge_cells('A12:J12')
ws.cell(row=13, column=1, value="Active SOM (col H) = P10/P50/P90 based on scenario_selector on Assumptions").font = ITALIC
ws.merge_cells('A13:J13')
ws.cell(row=14, column=1, value="P10 = 8.7/4 = 2.175%; P50 = 8.7/2 = 4.35%; P90 = 8.7/1 = 8.7%").font = ITALIC
ws.merge_cells('A14:J14')

# ═══════════════════════════════════════════════════════════
# 3. B2B P&L
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("B2B P&L")
ws['A1'] = "B2B Technology Enabler — P&L (driven by scenario_selector AND marketing_pct)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws['A2'] = "Split and SOM both auto-update. Change marketing_pct or scenario_selector on Assumptions."
ws['A2'].font = ITALIC
ws.merge_cells('A2:P2')

ws.append([])
ws.append(["Year", "SOM", "Sub Rev", "Ball Rev", "Custom", "License", "ITF Grant",
           "Total Rev", "Platform", "User Enh", "Ball Enh", "Ball Cost",
           "Manpower", "Total Cost", "Profit", "Cumulative"])
style_header(ws, 4, 16)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!H{4+i}").number_format = '#,##0'
    ws.cell(row=r, column=3,
            value=f"=B{r}*arpu_annual*split_b2b/1000000").number_format = '#,##0.00'
    if i == 1:
        ws.cell(row=r, column=4,
                value=f"=B{r}*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=4,
                value=f"=(B{r}-B{r-1})*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=5,
            value="=customization_y1" if i == 1 else "=customization_annual").number_format = '#,##0.000'
    ws.cell(row=r, column=6, value="=license_fee").number_format = '#,##0.000'
    if i == 1:
        ws.cell(row=r, column=7,
                value="=manpower_annual*grant_pct*grant_lag_frac").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=7,
                value="=manpower_annual*grant_pct").number_format = '#,##0.00'
    ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=9,
            value=f"=platform_base+platform_per_user*B{r}/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=10,
            value=f"=user_enhancement*B{r}/1000000").number_format = '#,##0.00'
    if i == 1:
        ws.cell(row=r, column=11,
                value=f"=B{r}*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=12,
                value=f"=B{r}*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=11,
                value=f"=(B{r}-B{r-1})*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=12,
                value=f"=(B{r}-B{r-1})*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=13, value="=manpower_annual").number_format = '#,##0.00'
    ws.cell(row=r, column=14, value=f"=SUM(I{r}:M{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=15, value=f"=H{r}-N{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=16,
            value=f"=O{r}" if i == 1 else f"=P{r-1}+O{r}").number_format = '#,##0.00'
    for c in range(1, 17):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

# Y7 Q2
r = 11
ws.cell(row=r, column=1, value="Y7 Q2").alignment = CENTER
for c in range(2, 7):
    ws.cell(row=r, column=c, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=7, value="=grant_y7q2").number_format = '#,##0.00'
ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
for c in range(9, 14):
    ws.cell(row=r, column=c, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=14, value=f"=SUM(I{r}:M{r})").number_format = '#,##0.00'
ws.cell(row=r, column=15, value=f"=H{r}-N{r}").number_format = '#,##0.00'
ws.cell(row=r, column=16, value=f"=P{r-1}+O{r}").number_format = '#,##0.00'
for c in range(1, 17):
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).font = NUM_FONT
    if c >= 2:
        ws.cell(row=r, column=c).fill = CALC_FILL
        ws.cell(row=r, column=c).alignment = CENTER

set_widths(ws, [8, 10, 9, 9, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 10, 11])

# ═══════════════════════════════════════════════════════════
# 4. B2C P&L
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("B2C P&L")
ws['A1'] = "B2C Distribution Partners — P&L (driven by scenario_selector AND marketing_pct)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "Marketing cost is B2C-only. Split auto-updates to preserve fairness."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

ws.append([])
ws.append(["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev",
           "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"])
style_header(ws, 4, 11)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!H{4+i}").number_format = '#,##0'
    ws.cell(row=r, column=3,
            value=f"=B{r}*arpu_annual*split_b2c/1000000").number_format = '#,##0.00'
    if i == 1:
        ws.cell(row=r, column=4,
                value=f"=B{r}*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=4,
                value=f"=(B{r}-B{r-1})*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"=C{r}+D{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=6,
            value=f"=B{r}*arpu_annual*marketing_pct/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=7, value="=license_fee").number_format = '#,##0.000'
    ws.cell(row=r, column=8,
            value="=customization_y1" if i == 1 else "=customization_annual").number_format = '#,##0.000'
    ws.cell(row=r, column=9, value=f"=SUM(F{r}:H{r})").number_format = '#,##0.00'
    ws.cell(row=r, column=10, value=f"=E{r}-I{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=11,
            value=f"=J{r}" if i == 1 else f"=K{r-1}+J{r}").number_format = '#,##0.00'
    for c in range(1, 12):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

set_widths(ws, [8, 10, 10, 10, 10, 10, 10, 14, 10, 10, 11])

# ═══════════════════════════════════════════════════════════
# 5. SCENARIOS (three parallel blocks)
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Scenarios")
ws['A1'] = "Scenario Comparison — P10 / P50 / P90 computed in parallel"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws['A2'] = "All three scenarios always live. Marketing rate and split apply to all three."
ws['A2'].font = ITALIC
ws.merge_cells('A2:P2')

# Summary at rows 4-8
ws.cell(row=4, column=1, value="SCENARIO SUMMARY").font = SUB_FONT
ws.cell(row=4, column=1).fill = SUB_FILL
ws.merge_cells('A4:P4')
sum_headers = ["Scenario", "Conversion (Y6)", "SOM (Y6)", "B2B Cum Y6",
               "B2C Cum Y6", "B2B Full-Cycle", "Fairness", "Within 10%?"]
for j, h in enumerate(sum_headers, 1):
    c = ws.cell(row=5, column=j, value=h)
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

# We'll fill summary rows 6-8 after building blocks.
# Pre-declare scenario metadata
scenarios_meta = [
    ("P10 (Lower)", "mult_p10", 6),
    ("P50 (Mean)", "mult_p50", 7),
    ("P90 (Upper)", "mult_p90", 8),
]

# Build B2B blocks starting at row 10
b2b_y6_rows = {}
r_b2b = 10
for name, mult, sum_row in scenarios_meta:
    # Title
    ws.cell(row=r_b2b, column=1, value=f"{name} — B2B Detail").font = SUB_FONT
    ws.cell(row=r_b2b, column=1).fill = SUB_FILL
    ws.merge_cells(start_row=r_b2b, start_column=1, end_row=r_b2b, end_column=16)
    hdr = r_b2b + 1
    headers = ["Year", "SOM", "Sub Rev", "Ball Rev", "Custom", "License", "ITF Grant",
               "Total Rev", "Platform", "User Enh", "Ball Enh", "Ball Cost",
               "Manpower", "Total Cost", "Profit", "Cumulative"]
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=hdr, column=j, value=h)
        c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

    for i in range(1, 7):
        r = hdr + i
        ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
        ws.cell(row=r, column=2,
                value=f"=sam_Y{i}*1000000*base_pen_Y{i}*{mult}").number_format = '#,##0'
        ws.cell(row=r, column=3,
                value=f"=B{r}*arpu_annual*split_b2b/1000000").number_format = '#,##0.00'
        if i == 1:
            ws.cell(row=r, column=4,
                    value=f"=B{r}*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
        else:
            ws.cell(row=r, column=4,
                    value=f"=(B{r}-B{r-1})*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=5,
                value="=customization_y1" if i == 1 else "=customization_annual").number_format = '#,##0.000'
        ws.cell(row=r, column=6, value="=license_fee").number_format = '#,##0.000'
        if i == 1:
            ws.cell(row=r, column=7,
                    value="=manpower_annual*grant_pct*grant_lag_frac").number_format = '#,##0.00'
        else:
            ws.cell(row=r, column=7,
                    value="=manpower_annual*grant_pct").number_format = '#,##0.00'
        ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
        ws.cell(row=r, column=9,
                value=f"=platform_base+platform_per_user*B{r}/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=10,
                value=f"=user_enhancement*B{r}/1000000").number_format = '#,##0.00'
        if i == 1:
            ws.cell(row=r, column=11,
                    value=f"=B{r}*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
            ws.cell(row=r, column=12,
                    value=f"=B{r}*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
        else:
            ws.cell(row=r, column=11,
                    value=f"=(B{r}-B{r-1})*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
            ws.cell(row=r, column=12,
                    value=f"=(B{r}-B{r-1})*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=13, value="=manpower_annual").number_format = '#,##0.00'
        ws.cell(row=r, column=14, value=f"=SUM(I{r}:M{r})").number_format = '#,##0.00'
        ws.cell(row=r, column=15, value=f"=H{r}-N{r}").number_format = '#,##0.00'
        ws.cell(row=r, column=16,
                value=f"=O{r}" if i == 1 else f"=P{r-1}+O{r}").number_format = '#,##0.00'
        for c in range(1, 17):
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).font = NUM_FONT
            if c >= 2:
                ws.cell(row=r, column=c).fill = CALC_FILL
                ws.cell(row=r, column=c).alignment = CENTER

    y6_row = hdr + 6
    b2b_y6_rows[name] = y6_row

    # Y7 Q2 catch-up
    y7q2_row = hdr + 7
    ws.cell(row=y7q2_row, column=1, value="Y7 Q2").alignment = CENTER
    for c in range(2, 7):
        ws.cell(row=y7q2_row, column=c, value=0).number_format = '#,##0.00'
    ws.cell(row=y7q2_row, column=7, value="=grant_y7q2").number_format = '#,##0.00'
    ws.cell(row=y7q2_row, column=8, value=f"=SUM(C{y7q2_row}:G{y7q2_row})").number_format = '#,##0.00'
    for c in range(9, 14):
        ws.cell(row=y7q2_row, column=c, value=0).number_format = '#,##0.00'
    ws.cell(row=y7q2_row, column=14, value=f"=SUM(I{y7q2_row}:M{y7q2_row})").number_format = '#,##0.00'
    ws.cell(row=y7q2_row, column=15, value=f"=H{y7q2_row}-N{y7q2_row}").number_format = '#,##0.00'
    ws.cell(row=y7q2_row, column=16, value=f"=P{y7q2_row-1}+O{y7q2_row}").number_format = '#,##0.00'
    for c in range(1, 17):
        ws.cell(row=y7q2_row, column=c).border = BORDER
        ws.cell(row=y7q2_row, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=y7q2_row, column=c).fill = CALC_FILL
            ws.cell(row=y7q2_row, column=c).alignment = CENTER

    r_b2b = y7q2_row + 2  # skip a row

# Build B2C blocks starting after B2B
b2c_y6_rows = {}
r_b2c = r_b2b + 1
for name, mult, sum_row in scenarios_meta:
    ws.cell(row=r_b2c, column=1, value=f"{name} — B2C Detail").font = SUB_FONT
    ws.cell(row=r_b2c, column=1).fill = SUB_FILL
    ws.merge_cells(start_row=r_b2c, start_column=1, end_row=r_b2c, end_column=11)
    hdr = r_b2c + 1
    headers = ["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev",
               "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"]
    for j, h in enumerate(headers, 1):
        c = ws.cell(row=hdr, column=j, value=h)
        c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

    for i in range(1, 7):
        r = hdr + i
        ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
        ws.cell(row=r, column=2,
                value=f"=sam_Y{i}*1000000*base_pen_Y{i}*{mult}").number_format = '#,##0'
        ws.cell(row=r, column=3,
                value=f"=B{r}*arpu_annual*split_b2c/1000000").number_format = '#,##0.00'
        if i == 1:
            ws.cell(row=r, column=4,
                    value=f"=B{r}*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
        else:
            ws.cell(row=r, column=4,
                    value=f"=(B{r}-B{r-1})*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=5, value=f"=C{r}+D{r}").number_format = '#,##0.00'
        ws.cell(row=r, column=6,
                value=f"=B{r}*arpu_annual*marketing_pct/1000000").number_format = '#,##0.00'
        ws.cell(row=r, column=7, value="=license_fee").number_format = '#,##0.000'
        ws.cell(row=r, column=8,
                value="=customization_y1" if i == 1 else "=customization_annual").number_format = '#,##0.000'
        ws.cell(row=r, column=9, value=f"=SUM(F{r}:H{r})").number_format = '#,##0.00'
        ws.cell(row=r, column=10, value=f"=E{r}-I{r}").number_format = '#,##0.00'
        ws.cell(row=r, column=11,
                value=f"=J{r}" if i == 1 else f"=K{r-1}+J{r}").number_format = '#,##0.00'
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = BORDER
            ws.cell(row=r, column=c).font = NUM_FONT
            if c >= 2:
                ws.cell(row=r, column=c).fill = CALC_FILL
                ws.cell(row=r, column=c).alignment = CENTER

    b2c_y6_rows[name] = hdr + 6
    r_b2c = hdr + 6 + 2

# Fill summary rows 6-8
for name, mult, sum_row in scenarios_meta:
    b2b_y6 = b2b_y6_rows[name]
    b2c_y6 = b2c_y6_rows[name]
    b2b_y7q2 = b2b_y6 + 1
    ws.cell(row=sum_row, column=1, value=name).font = BOLD_NUM
    ws.cell(row=sum_row, column=2,
            value=f"=base_pen_Y6*{mult}").number_format = '0.000%'
    ws.cell(row=sum_row, column=3,
            value=f"=sam_Y6*1000000*base_pen_Y6*{mult}").number_format = '#,##0'
    ws.cell(row=sum_row, column=4,
            value=f"=P{b2b_y6}").number_format = '#,##0.00'
    ws.cell(row=sum_row, column=5,
            value=f"=K{b2c_y6}").number_format = '#,##0.00'
    ws.cell(row=sum_row, column=6,
            value=f"=P{b2b_y7q2}").number_format = '#,##0.00'
    ws.cell(row=sum_row, column=7,
            value=f"=E{sum_row}/F{sum_row}").number_format = '0.0000'
    ws.cell(row=sum_row, column=8,
            value=f'=IF(G{sum_row}<=1.1,"PASS","FAIL")').font = BOLD_NUM
    for c in range(1, 9):
        ws.cell(row=sum_row, column=c).border = BORDER
        if c >= 2 and c <= 7:
            ws.cell(row=sum_row, column=c).fill = CALC_FILL
            ws.cell(row=sum_row, column=c).alignment = CENTER

set_widths(ws, [8, 12, 9, 9, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 10, 11])

# ═══════════════════════════════════════════════════════════
# 6. SUMMARY
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Summary")
ws['A1'] = "Summary — Main P&L (driven by scenario_selector) + 9 checks"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws['A2'] = "Change marketing_pct or scenario_selector on Assumptions to update."
ws['A2'].font = ITALIC
ws.merge_cells('A2:G2')

ws.append([])
ws.append(["Year", "B2B Profit", "B2C Profit", "Total Profit",
           "B2B Cumulative", "B2C Cumulative", "B2C / B2B"])
style_header(ws, 4, 7)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='B2B P&L'!O{4+i}").number_format = '#,##0.00'
    ws.cell(row=r, column=3, value=f"='B2C P&L'!J{4+i}").number_format = '#,##0.00'
    ws.cell(row=r, column=4, value=f"=B{r}+C{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"='B2B P&L'!P{4+i}").number_format = '#,##0.00'
    ws.cell(row=r, column=6, value=f"='B2C P&L'!K{4+i}").number_format = '#,##0.00'
    ws.cell(row=r, column=7,
            value=f'=IF(B{r}=0,"n/a",F{r}/B{r})').number_format = '0.0000'
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

set_widths(ws, [10, 12, 12, 12, 14, 14, 12])

# CHECKS
ws.cell(row=13, column=1, value="CHECKS").font = SUB_FONT
ws.cell(row=13, column=1).fill = SUB_FILL
ws.merge_cells('A13:G13')

checks = [
    ("B2B cumulative = sum of annual", "=IF(ROUND('B2B P&L'!P10-SUM('B2B P&L'!O5:O10),2)=0,\"PASS\",\"FAIL\")"),
    ("B2C cumulative = sum of annual", "=IF(ROUND('B2C P&L'!K10-SUM('B2C P&L'!J5:J10),2)=0,\"PASS\",\"FAIL\")"),
    ("Full-cycle fairness <= 1.10", "=IF(G11<=1.1,\"PASS\",\"FAIL\")"),
    ("Full-cycle fairness >= 1.09 (both partners ~equal)", "=IF(G11>=1.09,\"PASS\",\"REVIEW\")"),
    ("All years B2B non-negative", "=IF(MIN('B2B P&L'!O5:O10)>=0,\"PASS\",\"FAIL\")"),
    ("All years B2C non-negative", "=IF(MIN('B2C P&L'!J5:J10)>=0,\"PASS\",\"FAIL\")"),
    ("Subscription split = 100%", "=IF(ROUND(split_coach+split_b2b+split_b2c,4)=1,\"PASS\",\"FAIL\")"),
    ("Marketing rate calibrated (30/35/40)", "=IF(OR(marketing_pct=0.30,marketing_pct=0.35,marketing_pct=0.40),\"PASS\",\"FAIL — recalibrate\")"),
    ("All 3 scenarios fairness <= 1.10", "=IF(AND(Scenarios!G6<=1.1,Scenarios!G7<=1.1,Scenarios!G8<=1.1),\"PASS\",\"FAIL\")"),
]
for i, (label, formula) in enumerate(checks, start=14):
    ws.cell(row=i, column=1, value=label).font = NUM_FONT
    ws.cell(row=i, column=1).border = BORDER
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    c = ws.cell(row=i, column=4, value=formula)
    c.font = BOLD_NUM; c.border = BORDER; c.alignment = CENTER
    c.fill = CALC_FILL

# ═══════════════════════════════════════════════════════════
# 7. SOURCES
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Sources")
ws['A1'] = "Sources — Reference URL Register"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:D1')
ws['A2'] = "Every TAM / SAM / SOM parameter, the 8.7% conversion benchmark, and the 68% Gen-Z demand signal are proofed below."
ws['A2'].font = ITALIC
ws.merge_cells('A2:D2')

ws.append([])
ws.append(["Reference", "URL", "Parameter(s) proven", "What it shows"])
style_header(ws, 4, 4)

sources = [
    ("2025 industry data (163.com)", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html",
     "TAM 300m; SAM campus 12.8m; market capital; 8.7% ceiling; 68% Gen-Z demand",
     "3.8亿球迷; 1,280万校园注册学生; 3,200亿产业规模; 487.6亿训练市场; 8.7% 家长长期支持率; 68% Z世代家长要求数据化反馈; 90%+ 家长接受培训花费; 63% 青训机构智能设备渗透; 92% 高端层复购率"),
    ("CFA 2025 Digital Annual Report (cited at 163.com)", "https://www.163.com/dy/article/L8E6C26F0556BPXI.html",
     "SAM amateur 979,500; SAM elite youth 109,220",
     "中国足协2025数字年报：社会业余球员约97.95万；精英青少年注册球员约10.92万"),
    ("NDRC Medium- and Long-Term Football Development Plan", "https://www.ndrc.gov.cn/xxgk/zcfb/ghwb/201604/t20160411_962179.html",
     "Cross-check: 50m participation target; 100m+ fan floor",
     "全社会经常参加足球运动人数目标5,000万；球迷人数过亿"),
    ("CFA youth-system reform (2026)", "https://www.thecfa.cn/mtzl/20260309/37401.html",
     "Policy trajectory",
     "青少年足球联赛扩张；扩大可测量球员池"),
    ("CEIBS analysis (FIFA Big Count)", "https://cn.ceibs.edu/new-papers-columns/29518",
     "Cross-check: 26.16m broad participants; 0.078% player density",
     "FIFA 2006广义足球人口2,616万；注册球员密度仅0.078%"),
    ("iResearch (middle-class sports spending)", "https://www.iresearch.com.cn",
     "Spending power",
     "中产家庭体育培训年支出1–5万元（63%）"),
    ("36Kr (youth academy fees)", "https://36kr.com",
     "Spending power",
     "一线城市青训班年费8,000–20,000元；精英俱乐部年费8–12万元"),
    ("Lanxiong Sports (Descente user profile)", "https://www.lanxiongsports.com",
     "Spending power; new gear fans",
     "一个孩子全年足球综合开销普遍突破3万元；迪桑特核心用户'少而忠诚、原价买单'"),
    ("Amer Sports investor relations (Arc'teryx)", "https://www.amersports.com/investors",
     "New gear fans",
     "始祖鸟35–55岁中产男性人均年消费约3万元；复购率>38%；大中华区营收增长47%"),
    ("Keep investor relations (AI coach)", "https://ir.keep.com",
     "Subscription benchmark",
     "Keep AI教练付费会员使付费率从8%提升至15%；续订率82%"),
    ("Shenxianqiu (GPS football vest)", "https://www.shenxianqiu.com",
     "Smart hardware price anchor",
     "国产GPS足球背心'神仙球'定价1,599元 — HK$980 AI足球低于此价格"),
    ("Adidas (World Cup ball retail)", "https://www.adidas.com",
     "Ball price anchor",
     "世界杯用球零售版约160美元（约1,150元人民币）"),
    ("JD.com (smart football consumer pricing)", "https://www.jd.com",
     "Smart ball price range",
     "带传感器智能足球消费级版本从数百元到数千元不等"),
]
for ref, url, param, shows in sources:
    ws.append([ref, url, param, shows])

for r in range(5, 5 + len(sources)):
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        ws.cell(row=r, column=c).alignment = LEFT if c in (1, 3, 4) else CENTER

set_widths(ws, [35, 50, 40, 80])

# ═══════════════════════════════════════════════════════════
# 8. COVER
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Cover", 0)
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 70
ws['B2'] = "Kung Fu AI Football — Live Model v2.8a"
ws['B2'].font = Font(bold=True, size=18)
ws.merge_cells('B2:C2')

cover = [
    ("File", "KungFu_AI_Football_v2.8a.xlsx"),
    ("Type", "Live model with auto-switching split + three parallel scenarios"),
    ("Base case", "P50 (mean conversion 4.35% → SOM 608,130 by Y6)"),
    ("Three scenarios", "P10 = 2.175% (Lower); P50 = 4.35% (Mean); P90 = 8.7% (Upper)"),
    ("Scenario switch", "Assumptions!scenario_selector: 1=P10, 2=P50, 3=P90"),
    ("Marketing switch", "Assumptions!marketing_pct: 30% / 35% / 40%"),
    ("Auto-split logic", "Split calibration table (rows 34-37) drives split_b2b and split_b2c via VLOOKUP"),
    ("Fairness preservation", "Each marketing rate is paired with a calibrated split that yields ~1.10 fairness"),
    ("Key sheets", "Assumptions · Market Sizing · B2B P&L · B2C P&L · Scenarios · Summary · Sources"),
    ("Yellow cells", "INPUT — editable"),
    ("Green cells", "CALCULATED — formulas, do not overwrite"),
]
for i, (k, v) in enumerate(cover, start=4):
    ws.cell(row=i, column=2, value=k).font = SUB_FONT
    ws.cell(row=i, column=2).fill = SUB_FILL
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=3, value=v).font = NUM_FONT
    ws.cell(row=i, column=3).border = BORDER

ws['B16'] = "HOW TO USE:"
ws['B16'].font = SUB_FONT
notes = [
    "1. Assumptions!marketing_pct → pick 30%, 35%, or 40%. Split auto-adjusts.",
    "2. Assumptions!scenario_selector → 1/2/3 for P10/P50/P90. Main P&L follows.",
    "3. Scenarios sheet → all three scenarios always live side by side.",
    "4. Summary → 9 PASS/FAIL checks including fairness and calibration checks.",
    "5. Sources → every parameter proofed with URL.",
]
for i, n in enumerate(notes, start=17):
    ws.cell(row=i, column=2, value=n).font = NUM_FONT
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)

# ── SAVE ──
OUT = "KungFu_AI_Football_v2.8a.xlsx"
wb.save(OUT)
print(f"✅ Workbook saved: {OUT}")
print(f"   Sheets: {wb.sheetnames}")
print(f"   Named ranges: {len(wb.defined_names)}")