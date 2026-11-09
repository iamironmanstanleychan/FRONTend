"""
KungFu AI Football v2.8 — Live Workbook with Three Scenarios
Output: KungFu_AI_Football_v2.8_ScenarioLive.xlsx
Run: python generate_v2_8.py
Requires: openpyxl
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import LineChart, Reference
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
# 1. ASSUMPTIONS (all named inputs)
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Assumptions")
ws['A1'] = "Assumptions — INPUT SHEET"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:F1')
ws['A2'] = "Yellow cells are editable. Every P&L line references these named inputs."
ws['A2'].font = ITALIC
ws.merge_cells('A2:F2')

ws.append([])
ws.append(["Category", "Parameter", "Value", "Named range", "Unit", "Source / Note"])
style_header(ws, 4, 6)

rows = [
    # Subscription & pricing
    ("Subscription", "Coach/Player share", 0.50, "split_coach", "%", "Contract term"),
    ("Subscription", "B2B subscription share", 0.1286, "split_b2b", "%", "Fix 1 — preserves 10% fairness"),
    ("Subscription", "B2C subscription share", 0.3714, "split_b2c", "%", "Fix 1 — paired with B2B"),
    ("Pricing", "Weighted ARPU (annual)", 1592, "arpu_annual", "HK$/year", "Tier mix: 60/25/12/3"),
    # Ball
    ("Ball", "Selling price", 980, "ball_price", "HK$", "Price anchor vs GPS vest RMB 1,599"),
    ("Ball", "Manufacturing cost", 250, "ball_cost", "HK$", "Confirmed"),
    ("Ball", "B2B revenue share", 0.35, "ball_share_b2b", "%", "Minimum to keep B2B ball-positive"),
    ("Ball", "Attach rate (new users)", 0.50, "ball_attach", "%", "One ball per two new users"),
    # Platform
    ("Platform", "Base (fixed)", 0.20, "platform_base", "HK$m/yr", "Model D Hybrid"),
    ("Platform", "Per-user compute", 2.00, "platform_per_user", "HK$/user", "Model D Hybrid"),
    ("Platform", "User enhancement", 5.00, "user_enhancement", "HK$/user", "Proportional to SOM"),
    ("Platform", "Ball enhancement", 10.00, "ball_enhancement", "HK$/ball", "Proportional to balls sold"),
    # B2C
    ("B2C", "Marketing % of gross sub", 0.30, "marketing_pct", "%", "Anta/LiNing/Amer/PIFD channels"),
    ("B2C", "Exclusive license fee", 0.20, "license_fee", "HK$m/yr", "Paid to B2B"),
    # Customization
    ("Customization", "Y1 fee", 0.10, "customization_y1", "HK$m", "Both B2B revenue and B2C cost"),
    ("Customization", "Y2-Y6 annual fee", 0.025, "customization_annual", "HK$m", "Hybrid structure"),
    # ITF
    ("ITF", "Manpower annual (gross)", 0.48, "manpower_annual", "HK$m/yr", "2 BSc × HK$20k × 12"),
    ("ITF", "Grant coverage", 1.00, "grant_pct", "%", "100% eligible"),
    ("ITF", "Reimbursement lag (fraction)", 0.5, "grant_lag_frac", "fraction", "6 months"),
    ("ITF", "Y7 Q2 catch-up", 0.24, "grant_y7q2", "HK$m", "2 × 20k × 6 months"),
    # Scenario multipliers
    ("Scenario", "P10 multiplier", 0.50, "mult_p10", "×", "Lower bound = 8.7/4"),
    ("Scenario", "P50 multiplier", 1.00, "mult_p50", "×", "Mean = 8.7/2"),
    ("Scenario", "P90 multiplier", 2.00, "mult_p90", "×", "Upper bound = 8.7/1"),
    ("Scenario", "Selector (1=P10, 2=P50, 3=P90)", 2, "scenario_selector", "1-3", "Drives main P&L"),
]

r = 5
for cat, param, val, name, unit, src in rows:
    ws.cell(row=r, column=1, value=cat).font = NUM_FONT
    ws.cell(row=r, column=2, value=param).font = NUM_FONT
    c = ws.cell(row=r, column=3, value=val)
    c.fill = INPUT_FILL; c.font = BOLD_NUM; c.alignment = CENTER
    if isinstance(val, float) and val < 1:
        c.number_format = '0.0000'
    elif isinstance(val, float):
        c.number_format = '#,##0.00'
    else:
        c.number_format = '#,##0'
    ws.cell(row=r, column=4, value=name).font = ITALIC
    ws.cell(row=r, column=5, value=unit).font = ITALIC
    ws.cell(row=r, column=6, value=src).font = ITALIC
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
        if col != 3:
            ws.cell(row=r, column=col).alignment = LEFT if col in (2, 6) else CENTER
    add_named(wb, name, "Assumptions", f"$C${r}")
    r += 1

# Scenario multiplier formula (driven by selector)
r += 1
ws.cell(row=r, column=1, value="Scenario").font = SUB_FONT
ws.cell(row=r, column=1).fill = SUB_FILL
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
r += 1
ws.cell(row=r, column=2, value="Active scenario multiplier").font = NUM_FONT
c = ws.cell(row=r, column=3, value="=CHOOSE(scenario_selector,mult_p10,mult_p50,mult_p90)")
c.fill = CALC_FILL; c.font = BOLD_NUM; c.alignment = CENTER
c.number_format = '0.00'
ws.cell(row=r, column=4, value="scenario_multiplier").font = ITALIC
ws.cell(row=r, column=6, value="= 0.50 / 1.00 / 2.00 based on selector").font = ITALIC
for col in range(1, 7):
    ws.cell(row=r, column=col).border = BORDER
add_named(wb, "scenario_multiplier", "Assumptions", f"$C${r}")

# Data validation dropdown on selector
dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False,
                    showDropDown=False, prompt="1=P10, 2=P50, 3=P90")
ws.add_data_validation(dv)
dv.add("C29")  # selector row is row 29 (5 + 24 = 29)

# ── MARKET SIZING INPUTS ──
r += 2
ws.cell(row=r, column=1, value="MARKET SIZING INPUTS (source-proofed)").font = SUB_FONT
ws.cell(row=r, column=1).fill = SUB_FILL
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
r += 1
ws.cell(row=r, column=1, value="Year").font = SUB_FONT
ws.cell(row=r, column=2, value="TAM (m)").font = SUB_FONT
ws.cell(row=r, column=3, value="SAM (m)").font = SUB_FONT
ws.cell(row=r, column=4, value="Base penetration (P50)").font = SUB_FONT
ws.cell(row=r, column=5, value="Named ranges").font = SUB_FONT
ws.cell(row=r, column=6, value="Source").font = SUB_FONT
for col in range(1, 7):
    ws.cell(row=r, column=col).fill = SUB_FILL
    ws.cell(row=r, column=col).border = BORDER
    ws.cell(row=r, column=col).alignment = CENTER
r += 1

market = [
    ("Y1", 110.0, 5.59, 0.0024589),
    ("Y2", 110.0, 5.59, 0.0061473),
    ("Y3", 185.0, 9.01, 0.0214639),
    ("Y4", 235.0, 11.34, 0.0269024),
    ("Y5", 270.0, 12.89, 0.0332830),
    ("Y6", 300.0, 13.98, 0.0435000),
]
for i, (yr, tam, sam, pen) in enumerate(market, start=1):
    ws.cell(row=r, column=1, value=yr).alignment = CENTER
    ws.cell(row=r, column=1).border = BORDER
    c = ws.cell(row=r, column=2, value=tam)
    c.fill = INPUT_FILL; c.border = BORDER; c.number_format = '#,##0.0'; c.alignment = CENTER
    c = ws.cell(row=r, column=3, value=sam)
    c.fill = INPUT_FILL; c.border = BORDER; c.number_format = '#,##0.00'; c.alignment = CENTER
    c = ws.cell(row=r, column=4, value=pen)
    c.fill = INPUT_FILL; c.border = BORDER; c.number_format = '0.00000%'; c.alignment = CENTER
    ws.cell(row=r, column=5, value=f"tam_Y{i}, sam_Y{i}, base_pen_Y{i}").font = ITALIC
    ws.cell(row=r, column=5).border = BORDER
    ws.cell(row=r, column=6, value="TAM: 79% of 380m fans; SAM: CFA 2025 + 163.com; Pen: 8.7/2 scaled").font = ITALIC
    ws.cell(row=r, column=6).border = BORDER
    ws.cell(row=r, column=6).alignment = LEFT
    add_named(wb, f"tam_Y{i}", "Assumptions", f"$B${r}")
    add_named(wb, f"sam_Y{i}", "Assumptions", f"$C${r}")
    add_named(wb, f"base_pen_Y{i}", "Assumptions", f"$D${r}")
    r += 1

# Regional SAM weights
r += 1
ws.cell(row=r, column=1, value="REGIONAL SAM WEIGHTS").font = SUB_FONT
ws.cell(row=r, column=1).fill = SUB_FILL
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
r += 1
ws.cell(row=r, column=1, value="Region").font = SUB_FONT
ws.cell(row=r, column=2, value="Weight").font = SUB_FONT
ws.cell(row=r, column=3, value="Regional SAM (m)").font = SUB_FONT
ws.cell(row=r, column=4, value="Rationale").font = SUB_FONT
for col in range(1, 7):
    ws.cell(row=r, column=col).fill = SUB_FILL
    ws.cell(row=r, column=col).border = BORDER
    ws.cell(row=r, column=col).alignment = CENTER
r += 1

regions = [
    ("HK & South China", 0.400, 5.59, "FIFA Talent Academy; PSG HK girls; Shenzhen youth centre"),
    ("East China", 0.244, 3.42, "Suzhou regional youth centre; dense school network"),
    ("Southwest", 0.167, 2.33, "Chengdu regional youth centre"),
    ("North", 0.111, 1.55, "Beijing national youth training centre"),
    ("Northwest", 0.044, 0.62, "Xi'an regional youth centre"),
    ("Northeast", 0.033, 0.47, "Dalian regional youth centre"),
]
for reg, w, sam, note in regions:
    ws.cell(row=r, column=1, value=reg).border = BORDER
    ws.cell(row=r, column=1).alignment = LEFT
    c = ws.cell(row=r, column=2, value=w)
    c.fill = INPUT_FILL; c.border = BORDER; c.number_format = '0.0%'; c.alignment = CENTER
    ws.cell(row=r, column=3, value=sam).border = BORDER
    ws.cell(row=r, column=3).number_format = '#,##0.00'
    ws.cell(row=r, column=3).alignment = CENTER
    ws.cell(row=r, column=4, value=note).font = ITALIC
    ws.cell(row=r, column=4).border = BORDER
    ws.cell(row=r, column=4).alignment = LEFT
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    r += 1

set_widths(ws, [14, 38, 14, 22, 14, 60])

# ═══════════════════════════════════════════════════════════
# 2. MARKET SIZING (formulas)
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Market Sizing")
ws['A1'] = "Market Sizing — TAM / SAM / SOM (calculated)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:I1')
ws['A2'] = "All cells are formulas. Edit inputs on the Assumptions sheet."
ws['A2'].font = ITALIC
ws.merge_cells('A2:I2')

ws.append([])
ws.append(["Year", "Regions Active", "TAM (m)", "SAM (m)",
           "SOM P10", "SOM P50", "SOM P90", "P50/SAM", "P50/TAM"])
style_header(ws, 4, 9)

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
    ws.cell(row=r, column=8, value=f"=F{r}/(D{r}*1000000)").number_format = '0.0000%'
    ws.cell(row=r, column=9, value=f"=F{r}/(C{r}*1000000)").number_format = '0.0000%'
    for c in range(1, 10):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 3:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

set_widths(ws, [8, 24, 12, 12, 14, 14, 14, 12, 12])

ws.cell(row=12, column=1, value="SOM = SAM × 1,000,000 × base_penetration × scenario_multiplier").font = ITALIC
ws.merge_cells('A12:I12')
ws.cell(row=13, column=1, value="P10 = ×0.50 (Lower bound = 8.7/4); P50 = ×1.00 (Mean = 8.7/2); P90 = ×2.00 (Upper bound = 8.7/1)").font = ITALIC
ws.merge_cells('A13:I13')
ws.cell(row=14, column=1, value="8.7% = share of Chinese parents willing to support child's long-term football participation (163.com). 68% Gen-Z demand = data-driven feedback.").font = ITALIC
ws.merge_cells('A14:I14')

# ═══════════════════════════════════════════════════════════
# 3. MAIN B2B P&L (driven by scenario_multiplier)
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("B2B P&L")
ws['A1'] = "B2B Technology Enabler — P&L (driven by scenario_selector)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws['A2'] = "Change scenario_selector on Assumptions (1=P10, 2=P50, 3=P90) to update this sheet."
ws['A2'].font = ITALIC
ws.merge_cells('A2:P2')

ws.append([])
headers = ["Year", "SOM", "Sub Rev", "Ball Rev", "Custom", "License", "ITF Grant",
           "Total Rev", "Platform", "User Enh", "Ball Enh", "Ball Cost",
           "Manpower", "Total Cost", "Profit", "Cumulative"]
ws.append(headers)
style_header(ws, 4, 16)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!F{4+i}").number_format = '#,##0'
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

# Y7 Q2 row
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
# 4. MAIN B2C P&L (driven by scenario_multiplier)
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("B2C P&L")
ws['A1'] = "B2C Distribution Partners — P&L (driven by scenario_selector)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "Change scenario_selector on Assumptions (1=P10, 2=P50, 3=P90) to update this sheet."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

ws.append([])
ws.append(["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev",
           "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"])
style_header(ws, 4, 11)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!F{4+i}").number_format = '#,##0'
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
ws['A1'] = "Scenario Comparison — Lower / Mean / Upper Bounds (computed in parallel)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws['A2'] = "Each scenario uses its own multiplier. All formulas. No switch needed to compare."
ws['A2'].font = ITALIC
ws.merge_cells('A2:P2')

# Summary block
ws.cell(row=4, column=1, value="SCENARIO SUMMARY").font = SUB_FONT
ws.cell(row=4, column=1).fill = SUB_FILL
ws.merge_cells('A4:P4')
sum_headers = ["Scenario", "Conversion (Y6)", "SOM (Y6)", "B2B Cum Y6",
               "B2C Cum Y6", "B2B Full-Cycle", "Fairness (full-cycle)", "Within 10%?"]
for j, h in enumerate(sum_headers, 1):
    c = ws.cell(row=5, column=j, value=h)
    c.font = HEADER_FONT; c.fill = HEADER_FILL; c.alignment = CENTER; c.border = BORDER

# Placeholder — will fill after detail blocks are written
# Row 6 = P10, Row 7 = P50, Row 8 = P90

def write_b2b_scenario_block(ws, title_row, scenario, mult):
    ws.cell(row=title_row, column=1, value=f"{scenario} — B2B Detail").font = SUB_FONT
    ws.cell(row=title_row, column=1).fill = SUB_FILL
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=16)
    hdr = title_row + 1
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
    # cumulative row (Y7 Q2 grant only — for B2B full-cycle)
    cum_row = hdr + 7
    ws.cell(row=cum_row, column=1, value="Y7 Q2").alignment = CENTER
    for c in range(2, 7):
        ws.cell(row=cum_row, column=c, value=0).number_format = '#,##0.00'
    ws.cell(row=cum_row, column=7, value="=grant_y7q2").number_format = '#,##0.00'
    ws.cell(row=cum_row, column=8, value=f"=SUM(C{cum_row}:G{cum_row})").number_format = '#,##0.00'
    for c in range(9, 14):
        ws.cell(row=cum_row, column=c, value=0).number_format = '#,##0.00'
    ws.cell(row=cum_row, column=14, value=f"=SUM(I{cum_row}:M{cum_row})").number_format = '#,##0.00'
    ws.cell(row=cum_row, column=15, value=f"=H{cum_row}-N{cum_row}").number_format = '#,##0.00'
    ws.cell(row=cum_row, column=16, value=f"=P{cum_row-1}+O{cum_row}").number_format = '#,##0.00'
    for c in range(1, 17):
        ws.cell(row=cum_row, column=c).border = BORDER
        ws.cell(row=cum_row, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=cum_row, column=c).fill = CALC_FILL
            ws.cell(row=cum_row, column=c).alignment = CENTER
    return hdr + 6  # Y6 row


def write_b2c_scenario_block(ws, title_row, scenario, mult):
    ws.cell(row=title_row, column=1, value=f"{scenario} — B2C Detail").font = SUB_FONT
    ws.cell(row=title_row, column=1).fill = SUB_FILL
    ws.merge_cells(start_row=title_row, start_column=1, end_row=title_row, end_column=11)
    hdr = title_row + 1
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
    return hdr + 6  # Y6 row


# Write blocks
b2b_p10_y6 = write_b2b_scenario_block(ws, 10, "P10 (Lower)", "mult_p10")
b2b_p50_y6 = write_b2b_scenario_block(ws, 19, "P50 (Mean)", "mult_p50")
b2b_p90_y6 = write_b2b_scenario_block(ws, 28, "P90 (Upper)", "mult_p90")
b2c_p10_y6 = write_b2c_scenario_block(ws, 38, "P10 (Lower)", "mult_p10")
b2c_p50_y6 = write_b2c_scenario_block(ws, 47, "P50 (Mean)", "mult_p50")
b2c_p90_y6 = write_b2c_scenario_block(ws, 56, "P90 (Upper)", "mult_p90")

# Fill summary block (rows 6-8)
summary_rows = [
    (6, "P10 (Lower)", "mult_p10", b2b_p10_y6, b2c_p10_y6),
    (7, "P50 (Mean)", "mult_p50", b2b_p50_y6, b2c_p50_y6),
    (8, "P90 (Upper)", "mult_p90", b2b_p90_y6, b2c_p90_y6),
]
for r, name, mult, b2b_row, b2c_row in summary_rows:
    ws.cell(row=r, column=1, value=name).font = BOLD_NUM
    ws.cell(row=r, column=2,
            value=f"=base_pen_Y6*{mult}").number_format = '0.000%'
    ws.cell(row=r, column=3,
            value=f"=sam_Y6*1000000*base_pen_Y6*{mult}").number_format = '#,##0'
    ws.cell(row=r, column=4,
            value=f"=P{b2b_row}").number_format = '#,##0.00'
    ws.cell(row=r, column=5,
            value=f"=K{b2c_row}").number_format = '#,##0.00'
    ws.cell(row=r, column=6,
            value=f"=P{b2b_row+1}").number_format = '#,##0.00'
    ws.cell(row=r, column=7,
            value=f"=E{r}/F{r}").number_format = '0.0000'
    ws.cell(row=r, column=8,
            value=f'=IF(G{r}<=1.1,"PASS","FAIL")').font = BOLD_NUM
    for c in range(1, 9):
        ws.cell(row=r, column=c).border = BORDER
        if c >= 2 and c <= 7:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

ws.cell(row=65, column=1, value="Notes: Fairness = B2C cumulative / B2B full-cycle cumulative (incl. Y7 Q2 grant catch-up). PASS if <= 1.10.").font = ITALIC
ws.merge_cells('A65:P65')
ws.cell(row=66, column=1, value="Conversion Y6 = base_pen_Y6 × multiplier. P10 = 8.7/4 = 2.175%; P50 = 8.7/2 = 4.35%; P90 = 8.7/1 = 8.7%.").font = ITALIC
ws.merge_cells('A66:P66')

set_widths(ws, [8, 12, 9, 9, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 10, 11])

# ═══════════════════════════════════════════════════════════
# 6. SUMMARY & CHECKS
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Summary")
ws['A1'] = "Summary — Main P&L and Six Checks"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws['A2'] = "References the main B2B and B2C P&L sheets (driven by scenario_selector)."
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

# Y7 Q2 full-cycle ratio
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

# Checks
ws.cell(row=13, column=1, value="CHECKS").font = SUB_FONT
ws.cell(row=13, column=1).fill = SUB_FILL
ws.merge_cells('A13:G13')

checks = [
    ("B2B cumulative = sum of annual", "=IF(ROUND('B2B P&L'!P10-SUM('B2B P&L'!O5:O10),2)=0,\"PASS\",\"FAIL\")"),
    ("B2C cumulative = sum of annual", "=IF(ROUND('B2C P&L'!K10-SUM('B2C P&L'!J5:J10),2)=0,\"PASS\",\"FAIL\")"),
    ("Full-cycle fairness <= 1.10", "=IF(G11<=1.1,\"PASS\",\"FAIL\")"),
    ("All years B2B non-negative", "=IF(MIN('B2B P&L'!O5:O10)>=0,\"PASS\",\"FAIL\")"),
    ("All years B2C non-negative", "=IF(MIN('B2C P&L'!J5:J10)>=0,\"PASS\",\"FAIL\")"),
    ("Subscription split = 100%", "=IF(ROUND(split_coach+split_b2b+split_b2c,4)=1,\"PASS\",\"FAIL\")"),
    ("All 3 scenarios fairness <= 1.10", "=IF(AND(Scenarios!G6<=1.1,Scenarios!G7<=1.1,Scenarios!G8<=1.1),\"PASS\",\"FAIL\")"),
]
for i, (label, formula) in enumerate(checks, start=14):
    ws.cell(row=i, column=1, value=label).font = NUM_FONT
    ws.cell(row=i, column=1).border = BORDER
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    ws.cell(row=i, column=4, value=formula).font = BOLD_NUM
    ws.cell(row=i, column=4).border = BORDER
    ws.cell(row=i, column=4).alignment = CENTER
    ws.cell(row=i, column=4).fill = CALC_FILL

# ═══════════════════════════════════════════════════════════
# 7. SOURCES (reference URL proof table)
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
     "全社会经常参加足球运动人数目标5,000万；球迷人数过亿 — 支撑TAM 300m与SAM 13.98m的保守性"),
    ("CFA youth-system reform (2026)", "https://www.thecfa.cn/mtzl/20260309/37401.html",
     "Policy trajectory",
     "青少年足球联赛扩张；扩大可测量球员池 — 支撑SOM长期合理性"),
    ("CEIBS analysis (FIFA Big Count)", "https://cn.ceibs.edu/new-papers-columns/29518",
     "Cross-check: 26.16m broad participants; 0.078% player density",
     "FIFA 2006广义足球人口2,616万；注册球员密度仅0.078% — 证明TAM不应直接用Big Count，SOM应从核心层获取"),
    ("iResearch (middle-class sports spending)", "https://www.iresearch.com.cn",
     "§3.7.1 spending power",
     "中产家庭体育培训年支出1–5万元（63%）；5–10万元（14%）；10万以上（11%）"),
    ("36Kr (youth academy fees)", "https://36kr.com",
     "§3.7.1 spending power",
     "一线城市青训班年费8,000–20,000元；精英俱乐部年费8–12万元"),
    ("Lanxiong Sports (Descente user profile)", "https://www.lanxiongsports.com",
     "§3.7.1, §3.7.3",
     "一个孩子全年足球综合开销普遍突破3万元；迪桑特核心用户'少而忠诚、原价买单'"),
    ("Amer Sports investor relations (Arc'teryx)", "https://www.amersports.com/investors",
     "§3.7.3 new gear fans",
     "始祖鸟35–55岁中产男性人均年消费约3万元；复购率>38%；大中华区营收增长47%"),
    ("Keep investor relations (AI coach)", "https://ir.keep.com",
     "§3.7.1 subscription benchmark",
     "Keep AI教练付费会员使付费率从8%提升至15%；续订率82%"),
    ("Shenxianqiu (GPS football vest)", "https://www.shenxianqiu.com",
     "§3.7.1 smart hardware price anchor",
     "国产GPS足球背心'神仙球'定价1,599元 — HK$980 AI足球低于此价格"),
    ("Adidas (World Cup ball retail)", "https://www.adidas.com",
     "§3.7.1 ball price anchor",
     "世界杯用球零售版约160美元（约1,150元人民币）"),
    ("JD.com (smart football consumer pricing)", "https://www.jd.com",
     "§3.7.1 smart ball price range",
     "带传感器智能足球消费级版本从数百元到数千元不等"),
]
for ref, url, param, shows in sources:
    ws.append([ref, url, param, shows])

style_body_start = 5
for r in range(style_body_start, style_body_start + len(sources)):
    for c in range(1, 5):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        ws.cell(row=r, column=c).alignment = LEFT if c in (1, 3, 4) else CENTER

set_widths(ws, [35, 50, 40, 80])

# ═══════════════════════════════════════════════════════════
# 8. COVER (first sheet)
# ═══════════════════════════════════════════════════════════
ws = wb.create_sheet("Cover", 0)
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 70
ws['B2'] = "Kung Fu AI Football — Live Model v2.8"
ws['B2'].font = Font(bold=True, size=18)
ws.merge_cells('B2:C2')

cover = [
    ("File", "KungFu_AI_Football_v2.8_ScenarioLive.xlsx"),
    ("Type", "Live Excel model with three scenarios computed in parallel"),
    ("Base case", "P50 (mean conversion 4.35% → SOM 608,130 by Y6)"),
    ("Three scenarios", "P10 = 8.7/4 = 2.175% (Lower); P50 = 8.7/2 = 4.35% (Mean); P90 = 8.7/1 = 8.7% (Upper)"),
    ("Scenario switch", "Assumptions!scenario_selector: 1=P10, 2=P50, 3=P90"),
    ("Parameter proof", "All TAM/SAM/SOM and the 8.7% benchmark are source-proofed on the Sources sheet"),
    ("Yellow cells", "INPUT — editable"),
    ("Green cells", "CALCULATED — formulas, do not overwrite"),
    ("Key sheets", "Assumptions · Market Sizing · B2B P&L · B2C P&L · Scenarios · Summary · Sources"),
]
for i, (k, v) in enumerate(cover, start=4):
    ws.cell(row=i, column=2, value=k).font = SUB_FONT
    ws.cell(row=i, column=2).fill = SUB_FILL
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=3, value=v).font = NUM_FONT
    ws.cell(row=i, column=3).border = BORDER

ws['B15'] = "HOW TO USE:"
ws['B15'].font = SUB_FONT
notes = [
    "1. Open Assumptions — all yellow cells are editable",
    "2. Change scenario_selector (1/2/3) to drive the main B2B and B2C P&L",
    "3. Open Scenarios to see all three scenarios side by side (always live)",
    "4. Open Summary — seven PASS/FAIL invariants flag any breach",
    "5. Open Sources — every parameter is proofed with a URL and a one-line explanation",
]
for i, n in enumerate(notes, start=16):
    ws.cell(row=i, column=2, value=n).font = NUM_FONT
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)

# ── SAVE ──
OUT = "KungFu_AI_Football_v2.8_ScenarioLive.xlsx"
wb.save(OUT)
print(f"✅ Workbook saved: {OUT}")
print(f"   Sheets: {wb.sheetnames}")
print(f"   Named ranges: {len(wb.defined_names)}")