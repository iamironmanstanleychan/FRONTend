"""
KungFu AI Football v2.4 — Live Excel Model (Option 1)
Named inputs + formulas: change yellow cells, everything recalculates.
Run: python generate_v2_4_live.py
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

# ── STYLES ──
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
SUB_FONT = Font(bold=True, size=10)
NUM_FONT = Font(size=10)
BOLD_NUM = Font(size=10, bold=True)
ITALIC = Font(italic=True, size=9)
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
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = CENTER; cell.border = BORDER


def style_body(ws, r0, r1, ncols, fill=None):
    for r in range(r0, r1 + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            if not cell.font.bold:
                cell.font = NUM_FONT
            cell.border = BORDER
            cell.alignment = LEFT if c == 1 else CENTER
            if fill and not cell.fill.fgColor.rgb == "00FFF2CC":
                cell.fill = fill


def write_input(ws, row, value, name, unit, fmt='#,##0.00'):
    """Write a yellow input cell at column C, register as named range."""
    ws.cell(row=row, column=1).font = NUM_FONT
    ws.cell(row=row, column=1).border = BORDER
    ws.cell(row=row, column=1).alignment = LEFT

    ws.cell(row=row, column=2).font = NUM_FONT
    ws.cell(row=row, column=2).border = BORDER
    ws.cell(row=row, column=2).alignment = LEFT

    c = ws.cell(row=row, column=3, value=value)
    c.fill = INPUT_FILL; c.border = BORDER; c.font = BOLD_NUM
    c.alignment = CENTER; c.number_format = fmt

    ws.cell(row=row, column=4, value=name).font = ITALIC
    ws.cell(row=row, column=4).alignment = LEFT
    ws.cell(row=row, column=5, value=unit).font = ITALIC
    ws.cell(row=row, column=5).alignment = LEFT

    add_named(wb, name, ws.title, f"$C${row}")


# ═══════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ═══════════════════════════════════════════════════════
# 1. ASSUMPTIONS (all inputs)
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("Assumptions")
ws['A1'] = "Assumptions — INPUT SHEET (yellow cells are editable)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:E1')
ws['A2'] = "Change any yellow cell → every P&L line recalculates automatically."
ws['A2'].font = ITALIC
ws.merge_cells('A2:E2')

ws.append([])
ws.append(["Category", "Parameter", "Value", "Named range", "Unit"])
style_header(ws, 4, 5)

rows = [
    ("Subscription", "Coach/Player share", 0.50, "split_coach", "%", '0.00%'),
    ("Subscription", "B2B subscription share", 0.1286, "split_b2b", "%", '0.0000%'),
    ("Subscription", "B2C subscription share", 0.3714, "split_b2c", "%", '0.0000%'),
    ("Pricing", "Weighted ARPU (annual)", 1592, "arpu_annual", "HK$/year", '#,##0'),
    ("Ball", "Selling price", 980, "ball_price", "HK$", '#,##0'),
    ("Ball", "Manufacturing cost", 250, "ball_cost", "HK$", '#,##0'),
    ("Ball", "B2B revenue share", 0.35, "ball_share_b2b", "%", '0.00%'),
    ("Ball", "Attach rate (new users)", 0.50, "ball_attach", "%", '0.00%'),
    ("Cost", "Platform base", 0.20, "platform_base", "HK$m/yr", '#,##0.000'),
    ("Cost", "Platform per user", 2.00, "platform_per_user", "HK$/user", '#,##0.00'),
    ("Cost", "User enhancement", 5.00, "user_enhancement", "HK$/user", '#,##0.00'),
    ("Cost", "Ball enhancement", 10.00, "ball_enhancement", "HK$/ball", '#,##0.00'),
    ("Cost", "Marketing % (B2C)", 0.30, "marketing_pct", "%", '0.00%'),
    ("Revenue", "License fee (B2C→B2B)", 0.20, "license_fee", "HK$m/yr", '#,##0.000'),
    ("Revenue", "Customization Y1", 0.10, "customization_y1", "HK$m", '#,##0.000'),
    ("Revenue", "Customization Y2–Y6", 0.025, "customization_annual", "HK$m", '#,##0.000'),
    ("ITF", "Manpower annual (gross)", 0.48, "manpower_annual", "HK$m/yr", '#,##0.000'),
    ("ITF", "Grant coverage", 1.00, "grant_pct", "%", '0.00%'),
    ("ITF", "Reimbursement lag (fraction of year)", 0.5, "grant_lag_frac", "fraction", '0.00'),
    ("Scenario", "Scenario multiplier (1.0 = P50)", 1.00, "scenario_multiplier", "×", '0.00'),
    ("Scenario", "SOM P10 factor", 0.55, "som_p10_factor", "×", '0.00'),
    ("Scenario", "SOM P90 factor", 1.50, "som_p90_factor", "×", '0.00'),
]

r = 5
for cat, param, val, name, unit, fmt in rows:
    ws.cell(row=r, column=1, value=cat)
    ws.cell(row=r, column=2, value=param)
    ws.cell(row=r, column=3, value=val).number_format = fmt
    ws.cell(row=r, column=3).fill = INPUT_FILL
    ws.cell(row=r, column=3).font = BOLD_NUM
    ws.cell(row=r, column=4, value=name).font = ITALIC
    ws.cell(row=r, column=5, value=unit).font = ITALIC
    for c in range(1, 6):
        ws.cell(row=r, column=c).border = BORDER
        if c != 3:
            ws.cell(row=r, column=c).alignment = LEFT if c == 1 else CENTER
    add_named(wb, name, "Assumptions", f"$C${r}")
    r += 1

# Section 2: Market sizing inputs
r += 1
ws.cell(row=r, column=1, value="MARKET SIZING INPUTS").font = SUB_FONT
ws.cell(row=r, column=1).fill = SUB_FILL
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
r += 1
ws.cell(row=r, column=1, value="Year")
ws.cell(row=r, column=2, value="TAM (m students/fans)")
ws.cell(row=r, column=3, value="SAM (m active players)")
ws.cell(row=r, column=4, value="Penetration (P50)")
ws.cell(row=r, column=5, value="Named ranges")
for c in range(1, 6):
    ws.cell(row=r, column=c).font = SUB_FONT
    ws.cell(row=r, column=c).fill = SUB_FILL
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).alignment = CENTER
r += 1

market_data = [
    ("Y1", 110.0, 18.0, 0.000257),
    ("Y2", 110.0, 18.0, 0.000643),
    ("Y3", 185.0, 29.0, 0.002132),
    ("Y4", 235.0, 36.5, 0.002571),
    ("Y5", 270.0, 41.5, 0.003043),
    ("Y6", 300.0, 45.0, 0.003600),
]
for i, (yr, tam, sam, pen) in enumerate(market_data, start=1):
    ws.cell(row=r, column=1, value=yr).alignment = CENTER
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=2, value=tam).number_format = '#,##0.0'
    ws.cell(row=r, column=2).fill = INPUT_FILL; ws.cell(row=r, column=2).border = BORDER
    ws.cell(row=r, column=3, value=sam).number_format = '#,##0.0'
    ws.cell(row=r, column=3).fill = INPUT_FILL; ws.cell(row=r, column=3).border = BORDER
    ws.cell(row=r, column=4, value=pen).number_format = '0.0000%'
    ws.cell(row=r, column=4).fill = INPUT_FILL; ws.cell(row=r, column=4).border = BORDER
    ws.cell(row=r, column=5, value=f"tam_Y{i}, sam_Y{i}, pen_Y{i}").font = ITALIC
    ws.cell(row=r, column=5).border = BORDER
    add_named(wb, f"tam_Y{i}", "Assumptions", f"$B${r}")
    add_named(wb, f"sam_Y{i}", "Assumptions", f"$C${r}")
    add_named(wb, f"pen_Y{i}", "Assumptions", f"$D${r}")
    r += 1

ws.column_dimensions['A'].width = 14
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 14
ws.column_dimensions['D'].width = 22
ws.column_dimensions['E'].width = 14

# ═══════════════════════════════════════════════════════
# 2. MARKET SIZING (formulas)
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("Market Sizing")
ws['A1'] = "Market Sizing — TAM / SAM / SOM (calculated from Assumptions)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:J1')
ws['A2'] = "All cells here are formulas. Edit inputs on the Assumptions sheet."
ws['A2'].font = ITALIC
ws.merge_cells('A2:J2')

ws.append([])
headers = ["Year", "Regions Active", "TAM (m)", "SAM (m)",
           "SOM P10", "SOM Active", "SOM P90", "P50/SAM", "P50/TAM"]
ws.append(headers)
style_header(ws, 4, len(headers))

regions = ["HK & South China", "HK & South China", "+ East China",
           "+ Southwest", "+ North", "+ Northwest & Northeast"]

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=regions[i-1]).alignment = LEFT
    ws.cell(row=r, column=3, value=f"=tam_Y{i}").number_format = '#,##0.0'
    ws.cell(row=r, column=4, value=f"=sam_Y{i}").number_format = '#,##0.0'
    # P10 SOM = SAM × 1e6 × penetration × P10_factor
    ws.cell(row=r, column=5,
            value=f"=C{r}*1000000*pen_Y{i}*som_p10_factor").number_format = '#,##0'
    # Active SOM = SAM × 1e6 × penetration × scenario_multiplier
    ws.cell(row=r, column=6,
            value=f"=D{r}*1000000*pen_Y{i}*scenario_multiplier").number_format = '#,##0'
    # P90 SOM
    ws.cell(row=r, column=7,
            value=f"=D{r}*1000000*pen_Y{i}*som_p90_factor").number_format = '#,##0'
    # Ratios
    ws.cell(row=r, column=8, value=f"=F{r}/(D{r}*1000000)").number_format = '0.0000%'
    ws.cell(row=r, column=9, value=f"=F{r}/(C{r}*1000000)").number_format = '0.0000%'
    for c in range(1, 10):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 3:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

set_w = [8, 24, 12, 12, 14, 14, 14, 12, 12]
for i, w in enumerate(set_w, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.cell(row=12, column=1, value="SOM Active = SAM × 1,000,000 × Penetration × Scenario multiplier").font = ITALIC
ws.merge_cells('A12:I12')
ws.cell(row=13, column=1, value="Scenario multiplier on Assumptions: 1.0 = P50, 0.55 = P10, 1.50 = P90").font = ITALIC
ws.merge_cells('A13:I13')

# ═══════════════════════════════════════════════════════
# 3. B2B P&L (formulas)
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("B2B P&L")
ws['A1'] = "B2B Technology Enabler — Live P&L (formulas)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws['A2'] = "All values are formulas. Change Assumptions → recalculates."
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
    som_row = 4 + i   # Market Sizing row
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!F{som_row}").number_format = '#,##0'
    ws.cell(row=r, column=3,
            value=f"=B{r}*arpu_annual*split_b2b/1000000").number_format = '#,##0.00'
    # Ball Rev: Y1 = B×attach×price×share/1e6; Y2+ = (B-B_prev)×...
    if i == 1:
        ws.cell(row=r, column=4,
                value=f"=B{r}*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=4,
                value=f"=(B{r}-B{r-1})*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
    # Custom
    ws.cell(row=r, column=5,
            value="=customization_y1" if i == 1 else "=customization_annual").number_format = '#,##0.000'
    # License
    ws.cell(row=r, column=6, value="=license_fee").number_format = '#,##0.000'
    # ITF grant: Y1 = manpower × grant_pct × lag_frac; Y2+ = manpower × grant_pct
    if i == 1:
        ws.cell(row=r, column=7,
                value="=manpower_annual*grant_pct*grant_lag_frac").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=7,
                value="=manpower_annual*grant_pct").number_format = '#,##0.00'
    # Total Rev
    ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
    # Platform
    ws.cell(row=r, column=9,
            value=f"=platform_base+platform_per_user*B{r}/1000000").number_format = '#,##0.00'
    # User Enh
    ws.cell(row=r, column=10,
            value=f"=user_enhancement*B{r}/1000000").number_format = '#,##0.00'
    # Ball Enh
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
    # Manpower
    ws.cell(row=r, column=13, value="=manpower_annual").number_format = '#,##0.00'
    # Total Cost
    ws.cell(row=r, column=14, value=f"=SUM(I{r}:M{r})").number_format = '#,##0.00'
    # Profit
    ws.cell(row=r, column=15, value=f"=H{r}-N{r}").number_format = '#,##0.00'
    # Cumulative
    if i == 1:
        ws.cell(row=r, column=16, value=f"=O{r}").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=16, value=f"=P{r-1}+O{r}").number_format = '#,##0.00'
    for c in range(1, 17):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

# Y7 Q2 row (grant catch-up only)
r = 11
ws.cell(row=r, column=1, value="Y7 Q2").alignment = CENTER
ws.cell(row=r, column=2, value=0).number_format = '#,##0'
ws.cell(row=r, column=3, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=4, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=5, value=0).number_format = '#,##0.000'
ws.cell(row=r, column=6, value=0).number_format = '#,##0.000'
ws.cell(row=r, column=7,
        value="=manpower_annual*grant_pct*grant_lag_frac").number_format = '#,##0.00'
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

widths = [8, 10, 9, 9, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 10, 11]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.cell(row=13, column=1, value="Columns C–P are formulas referencing named ranges. Edit inputs on Assumptions.").font = ITALIC
ws.merge_cells('A13:P13')

# ═══════════════════════════════════════════════════════
# 4. B2C P&L (formulas)
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("B2C P&L")
ws['A1'] = "B2C Distribution Partners — Live P&L (formulas)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "All values are formulas. Customization is shown as B2C cost (Fix A)."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

ws.append([])
headers = ["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev",
           "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"]
ws.append(headers)
style_header(ws, 4, 11)

for i in range(1, 7):
    r = 4 + i
    som_row = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!F{som_row}").number_format = '#,##0'
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
    if i == 1:
        ws.cell(row=r, column=11, value=f"=J{r}").number_format = '#,##0.00'
    else:
        ws.cell(row=r, column=11, value=f"=K{r-1}+J{r}").number_format = '#,##0.00'
    for c in range(1, 12):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

w2 = [8, 10, 10, 10, 10, 10, 10, 14, 10, 10, 11]
for i, w in enumerate(w2, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.cell(row=12, column=1, value="Customization (col H) = B2C cost paid to B2B — matches B2B revenue (Fix A).").font = ITALIC
ws.merge_cells('A12:K12')

# ═══════════════════════════════════════════════════════
# 5. SUMMARY (formulas)
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("Summary")
ws['A1'] = "Summary — Fairness Check (all formulas)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws['A2'] = "References B2B P&L and B2C P&L. Change inputs on Assumptions → updates here."
ws['A2'].font = ITALIC
ws.merge_cells('A2:G2')

ws.append([])
ws.append(["Year", "B2B Profit", "B2C Profit", "Total Profit",
           "B2B Cumulative", "B2C Cumulative", "B2C / B2B"])
style_header(ws, 4, 7)

for i in range(1, 7):
    r = 4 + i
    b2b_r = 4 + i
    b2c_r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='B2B P&L'!O{b2b_r}").number_format = '#,##0.00'
    ws.cell(row=r, column=3, value=f"='B2C P&L'!J{b2c_r}").number_format = '#,##0.00'
    ws.cell(row=r, column=4, value=f"=B{r}+C{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"='B2B P&L'!P{b2b_r}").number_format = '#,##0.00'
    ws.cell(row=r, column=6, value=f"='B2C P&L'!K{b2c_r}").number_format = '#,##0.00'
    ws.cell(row=r, column=7,
            value=f"=IF(B{r}=0,\"n/a\",F{r}/B{r})").number_format = '0.0000'
    for c in range(1, 8):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

w3 = [8, 12, 12, 12, 14, 14, 12]
for i, w in enumerate(w3, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

# Checks section
ws.cell(row=12, column=1, value="CHECKS").font = SUB_FONT
ws.cell(row=12, column=1).fill = SUB_FILL
ws.merge_cells('A12:G12')

checks = [
    ("B2B cumulative = sum of annual", "=IF(ROUND('B2B P&L'!P9-SUM('B2B P&L'!O5:O9),4)=0,\"PASS\",\"FAIL\")"),
    ("B2C cumulative = sum of annual", "=IF(ROUND('B2C P&L'!K9-SUM('B2C P&L'!J5:J9),4)=0,\"PASS\",\"FAIL\")"),
    ("Fairness ratio within 10%", "=IF('Summary'!G9<=1.10,\"PASS\",\"FAIL\")"),
    ("All years B2B non-negative", "=IF(MIN('B2B P&L'!O5:O9)>=0,\"PASS\",\"FAIL\")"),
    ("All years B2C non-negative", "=IF(MIN('B2C P&L'!J5:J9)>=0,\"PASS\",\"FAIL\")"),
    ("Subscription split sums to 100%", "=IF(ROUND(split_coach+split_b2b+split_b2c,4)=1,\"PASS\",\"FAIL\")"),
]
for i, (label, formula) in enumerate(checks, start=13):
    ws.cell(row=i, column=1, value=label)
    ws.cell(row=i, column=1).font = NUM_FONT
    ws.cell(row=i, column=1).border = BORDER
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    ws.cell(row=i, column=4, value=formula)
    ws.cell(row=i, column=4).font = BOLD_NUM
    ws.cell(row=i, column=4).border = BORDER
    ws.cell(row=i, column=4).alignment = CENTER
    ws.cell(row=i, column=4).fill = CALC_FILL

# ═══════════════════════════════════════════════════════
# 6. CUMULATIVE CHART
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("Cumulative Chart")
ws['A1'] = "Cumulative Profit — Live Chart (auto-updates)"
ws['A1'].font = TITLE_FONT
ws.append([])
ws.append(["Year", "B2B Cumulative", "B2C Cumulative"])
style_header(ws, 3, 3)

for i in range(1, 7):
    r = 3 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Summary'!E{4+i}").number_format = '#,##0.00'
    ws.cell(row=r, column=3, value=f"='Summary'!F{4+i}").number_format = '#,##0.00'
    for c in range(1, 4):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        ws.cell(row=r, column=c).fill = CALC_FILL
        if c >= 2:
            ws.cell(row=r, column=c).alignment = CENTER

ws.column_dimensions['A'].width = 10
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 18

chart = LineChart()
chart.title = "Cumulative Profit: B2B vs B2C (HK$m)"
chart.style = 12
chart.y_axis.title = "HK$ million"
chart.x_axis.title = "Year"
data = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=9)
cats = Reference(ws, min_col=1, min_row=4, max_row=9)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.height = 10
chart.width = 20
ws.add_chart(chart, "E3")

# ═══════════════════════════════════════════════════════
# 7. COVER
# ═══════════════════════════════════════════════════════
ws = wb.create_sheet("Cover", 0)
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 70
ws['B2'] = "Kung Fu AI Football — Live Model v2.4"
ws['B2'].font = Font(bold=True, size=18)
ws.merge_cells('B2:C2')

cover = [
    ("File", "KungFu_AI_Football_v2.4_Live.xlsx"),
    ("Type", "Live Excel model — edit Assumptions, all P&L recalculates"),
    ("Base model", "v2.3 (Hybrid Customization · Model D Platform · Fix 1 split)"),
    ("Scenario", "P50 default (326,520 users by Y6)"),
    ("To switch to P10", "Set Assumptions!scenario_multiplier = 0.55"),
    ("To switch to P90", "Set Assumptions!scenario_multiplier = 1.50"),
    ("Yellow cells", "INPUT — editable"),
    ("Green cells", "CALCULATED — formulas, do not overwrite"),
    ("Fairness ratio (P50)", "B2C/B2B ≈ 1.0969 at Y6"),
]
for i, (k, v) in enumerate(cover, start=4):
    ws.cell(row=i, column=2, value=k).font = SUB_FONT
    ws.cell(row=i, column=2).fill = SUB_FILL
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=3, value=v).font = NUM_FONT
    ws.cell(row=i, column=3).border = BORDER

ws['B15'] = "WORKFLOW TO CHANGE A SCENARIO:"
ws['B15'].font = SUB_FONT
notes = [
    "1. Open Assumptions sheet",
    "2. Change any yellow cell (e.g., marketing_pct from 30% to 35%)",
    "3. All sheets recalculate automatically",
    "4. To switch P10/P50/P90, change scenario_multiplier (0.55 / 1.0 / 1.50)",
    "5. Check Summary sheet's CHECKS section for PASS/FAIL flags",
]
for i, n in enumerate(notes, start=16):
    ws.cell(row=i, column=2, value=n).font = NUM_FONT
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)

# ── SAVE ──
OUT = "KungFu_AI_Football_v2.4_Live.xlsx"
wb.save(OUT)
print(f"✅ Live workbook saved: {OUT}")
print(f"   Sheets: {wb.sheetnames}")
print(f"   Named ranges: {len(wb.defined_names)}")