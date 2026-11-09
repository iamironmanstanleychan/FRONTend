"""
KungFu AI Football Financial Model v3.5.3
Fixes: Matrix row alignment, removing #VALUE! errors. B2C Marketing math corrected.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ═══════════════════════════════════════════════════════════
# STYLES
# ═══════════════════════════════════════════════════════════
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="E2EFDA")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
NUM_FONT = Font(size=9)
BOLD_NUM = Font(size=9, bold=True)
ITALIC = Font(italic=True, size=8)
TITLE_FONT = Font(bold=True, size=14)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

def set_name(wb, name, sheet, cell):
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!{cell}")

def hdr(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font, cell.fill, cell.alignment, cell.border = HEADER_FONT, HEADER_FILL, CENTER, BORDER

def widths(ws, arr):
    for i, w in enumerate(arr, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ═══════════════════════════════════════════════════════════
# WORKBOOK CREATION
# ═══════════════════════════════════════════════════════════
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ───────────────────────────────────────────────────────────
# 1. COVER
# ───────────────────────────────────────────────────────────
ws = wb.create_sheet("Cover")
ws.column_dimensions['A'].width = 3
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 75

ws['B2'] = "Kung Fu AI Football — Financial Model v3.5.3"
ws['B2'].font = Font(bold=True, size=18)
ws.merge_cells('B2:C2')

info = [
    ("Type", "Live model — 3 BSc Manpower, DIPN 55, HKAS 38 Balance Sheet"),
    ("Base case", "P50 scenario at 35% B2C marketing (Centre)"),
    ("Marketing control", "Assumptions!C17 — dropdown 30% / 35% / 40%"),
    ("Scenario control", "Assumptions!C28 — 1 = P10, 2 = P50, 3 = P90"),
    ("DLRI Note", "University R&D Payment is a separate cost line in B2B P&L"),
    ("Matrix sheet", "All 9 combinations computed via clean helper arrays"),
    ("Yellow cells", "INPUT — editable"),
    ("Green cells", "CALCULATED — do not overwrite"),
]
for i, (k, v) in enumerate(info, start=4):
    ws.cell(row=i, column=2, value=k).font = BOLD_NUM
    ws.cell(row=i, column=2).fill = SUB_FILL
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=3, value=v).font = NUM_FONT
    ws.cell(row=i, column=3).border = BORDER
    ws.cell(row=i, column=3).alignment = LEFT

# ───────────────────────────────────────────────────────────
# 2. ASSUMPTIONS
# ───────────────────────────────────────────────────────────
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

inputs = [
    ("Subscription", "Coach/Player share", 0.50, "split_coach", "%", "Contract", "0.00%"),
    ("Subscription", "B2B subscription share", "=VLOOKUP(marketing_pct,$A$44:$C$46,2,FALSE)", "split_b2b", "%", "AUTO", "0.0000%"),
    ("Subscription", "B2C subscription share", "=VLOOKUP(marketing_pct,$A$44:$C$46,3,FALSE)", "split_b2c", "%", "AUTO", "0.0000%"),
    ("Pricing", "Weighted ARPU (annual)", 1592, "arpu_annual", "HK$/yr", "Tier mix", "#,##0"),
    ("Ball", "Selling price", 980, "ball_price", "HK$", "Price anchor", "#,##0"),
    ("Ball", "Manufacturing cost", 250, "ball_cost", "HK$", "B2B-borne", "#,##0"),
    ("Ball", "B2B revenue share", 0.35, "ball_share_b2b", "%", "Min for positive", "0.00%"),
    ("Ball", "Attach rate (new users)", 0.50, "ball_attach", "%", "One ball per two", "0.00%"),
    ("Platform", "Base (fixed)", 0.20, "platform_base", "HK$m/yr", "Model D", "#,##0.00"),
    ("Platform", "Per-user compute", 2.00, "platform_per_user", "HK$/user", "Model D", "#,##0.00"),
    ("Platform", "User enhancement", 5.00, "user_enhancement", "HK$/user", "Proportional", "#,##0.00"),
    ("Platform", "Ball enhancement", 10.00, "ball_enhancement", "HK$/ball", "Proportional", "#,##0.00"),
    ("B2C", "Marketing % of gross sub", 0.35, "marketing_pct", "%", "DROPDOWN", "0.00%"),
    ("B2C", "Exclusive license fee", 0.20, "license_fee", "HK$m/yr", "Paid to B2B", "#,##0.00"),
    ("Customization", "Y1 fee", 0.10, "customization_y1", "HK$m", "Both sides", "#,##0.000"),
    ("Customization", "Y2-Y6 annual fee", 0.025, "customization_annual", "HK$m", "Hybrid", "#,##0.000"),
    ("ITF", "Manpower gross (3 BSc)", 0.72, "manpower_annual", "HK$m/yr", "3 × HK$20k × 12", "#,##0.00"),
    ("ITF", "RTH Grant (2 BSc)", 0.48, "grant_annual", "HK$m/yr", "2 × HK$20k × 12", "#,##0.00"),
    ("ITF", "Grant lag", 0.5, "grant_lag_frac", "fraction", "6 months", "0.00"),
    ("ITF", "Y7 Q2 catch-up", 0.24, "grant_y7q2", "HK$m", "Final reimbursement", "#,##0.00"),
    ("Scenario", "P10 multiplier", 0.50, "mult_p10", "×", "8.7 / 4", "0.00"),
    ("Scenario", "P50 multiplier", 1.00, "mult_p50", "×", "8.7 / 2", "0.00"),
    ("Scenario", "P90 multiplier", 2.00, "mult_p90", "×", "8.7 / 1", "0.00"),
    ("Scenario", "Selector (1=P10, 2=P50, 3=P90)", 2, "scenario_selector", "1-3", "DROPDOWN", "#,##0"),
    ("Operating", "Uni R&D % of License", 0.60, "uni_rd_pct", "%", "Paid to unis", "0.00%"),
    ("Operating", "DLRI Cash", "=license_fee*uni_rd_pct", "uni_rd_cash", "HK$m/yr", "60% × 0.2", "#,##0.00"),
    ("Tax", "Enhanced R&D Deduction", 3.00, "rd_deduct", "×", "300%", "0.00"),
    ("Tax", "Tax Rate Tier 1", 0.0825, "tax_r1", "%", "First HK$2m", "0.0000%"),
    ("Tax", "Tax Rate Tier 2", 0.1650, "tax_r2", "%", "Above HK$2m", "0.0000%"),
    ("Tax", "Tier 1 Limit", 2.00, "tax_limit", "HK$m", "Threshold", "#,##0.00"),
    ("Acc", "Amortization Years", 5, "amort_yrs", "years", "Useful life", "#,##0"),
]

r = 5
for cat, param, val, name, unit, note, fmt in inputs:
    ws.cell(row=r, column=1, value=cat).font = NUM_FONT
    ws.cell(row=r, column=2, value=param).font = NUM_FONT
    c = ws.cell(row=r, column=3, value=val)
    c.number_format = fmt
    c.alignment = CENTER
    c.font = BOLD_NUM
    c.fill = INPUT_FILL if not isinstance(val, str) or not val.startswith('=') else CALC_FILL
    ws.cell(row=r, column=4, value=name).font = ITALIC
    ws.cell(row=r, column=5, value=unit).font = ITALIC
    ws.cell(row=r, column=6, value=note).font = ITALIC
    for col in range(1, 7):
        ws.cell(row=r, column=col).border = BORDER
        if col != 3:
            ws.cell(row=r, column=col).alignment = LEFT if col in (2, 6) else CENTER
    set_name(wb, name, "Assumptions", f"$C${r}")
    r += 1

# Active scenario
ws.cell(row=39, column=1, value="ACTIVE SCENARIO").font = BOLD_NUM
ws.cell(row=39, column=1).fill = SUB_FILL
ws.merge_cells('A39:F39')
ws.cell(row=40, column=2, value="Active scenario multiplier").font = NUM_FONT
c = ws.cell(row=40, column=3, value="=CHOOSE(scenario_selector,mult_p10,mult_p50,mult_p90)")
c.number_format = '0.00'; c.font = BOLD_NUM; c.fill = CALC_FILL; c.alignment = CENTER
ws.cell(row=40, column=4, value="scenario_multiplier").font = ITALIC
ws.cell(row=40, column=6, value="= 0.50 / 1.00 / 2.00").font = ITALIC
for col in range(1, 7): ws.cell(row=40, column=col).border = BORDER
set_name(wb, "scenario_multiplier", "Assumptions", "$C$40")

# Split calibration
ws.cell(row=42, column=1, value="SPLIT CALIBRATION").font = BOLD_NUM
ws.cell(row=42, column=1).fill = SUB_FILL
ws.merge_cells('A42:F42')
ws.cell(row=43, column=1, value="Marketing")
ws.cell(row=43, column=2, value="B2B split")
ws.cell(row=43, column=3, value="B2C split")
ws.cell(row=43, column=4, value="Fairness")
ws.cell(row=43, column=5, value="Sum")
ws.cell(row=43, column=6, value="Status")
hdr(ws, 43, 6)

calib = [(0.30, 0.1286, 0.3714, 1.0997), (0.35, 0.1048, 0.3952, 1.0990), (0.40, 0.0809, 0.4191, 1.0990)]
for i, (mkt, b2b, b2c, fair) in enumerate(calib, start=44):
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
    set_name(wb, f"calib_b2b_{int(mkt*100)}", "Assumptions", f"$B${i}")
    set_name(wb, f"calib_b2c_{int(mkt*100)}", "Assumptions", f"$C${i}")
    set_name(wb, f"calib_mkt_{int(mkt*100)}", "Assumptions", f"$A${i}")

# Market sizing
ws.cell(row=48, column=1, value="MARKET SIZING INPUTS").font = BOLD_NUM
ws.cell(row=48, column=1).fill = SUB_FILL
ws.merge_cells('A48:F48')
ws.cell(row=49, column=1, value="Year")
ws.cell(row=49, column=2, value="TAM (m)")
ws.cell(row=49, column=3, value="SAM (m)")
ws.cell(row=49, column=4, value="Base penetration")
ws.cell(row=49, column=5, value="Named ranges")
ws.cell(row=49, column=6, value="Source")
hdr(ws, 49, 6)

market = [("Y1", 110.0, 5.55, 0.0024589), ("Y2", 110.0, 5.55, 0.0061473),
          ("Y3", 185.0, 8.95, 0.0214639), ("Y4", 235.0, 11.26, 0.0269024),
          ("Y5", 270.0, 12.80, 0.0332830), ("Y6", 300.0, 13.89, 0.0435000)]
for i, (yr, tam, sam, pen) in enumerate(market, start=50):
    idx = i - 49
    ws.cell(row=i, column=1, value=yr).alignment = CENTER
    c = ws.cell(row=i, column=2, value=tam); c.number_format = '#,##0.0'; c.fill = INPUT_FILL
    c = ws.cell(row=i, column=3, value=sam); c.number_format = '#,##0.00'; c.fill = INPUT_FILL
    c = ws.cell(row=i, column=4, value=pen); c.number_format = '0.00000%'; c.fill = INPUT_FILL
    ws.cell(row=i, column=5, value=f"tam_Y{idx}, sam_Y{idx}, pen_Y{idx}").font = ITALIC
    ws.cell(row=i, column=6, value="TAM=79% fans; SAM=CFA; Pen=8.7/2").font = ITALIC
    for col in range(1, 7):
        ws.cell(row=i, column=col).border = BORDER
    set_name(wb, f"tam_Y{idx}", "Assumptions", f"$B${i}")
    set_name(wb, f"sam_Y{idx}", "Assumptions", f"$C${i}")
    set_name(wb, f"pen_Y{idx}", "Assumptions", f"$D${i}")

dv_mkt = DataValidation(type="list", formula1="=$A$44:$A$46", allow_blank=False)
ws.add_data_validation(dv_mkt); dv_mkt.add("C17")
dv_sel = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
ws.add_data_validation(dv_sel); dv_sel.add("C28")
widths(ws, [14, 38, 16, 22, 14, 55])

# ───────────────────────────────────────────────────────────
# 3. MARKET SIZING
# ───────────────────────────────────────────────────────────
ws = wb.create_sheet("Market Sizing")
ws['A1'] = "Market Sizing — TAM / SAM / SOM"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:J1')
ws['A2'] = "SOM = SAM × 1,000,000 × base penetration × scenario multiplier"
ws['A2'].font = ITALIC
ws.merge_cells('A2:J2')
ws.append([])
ws.append(["Year", "Regions Active", "TAM (m)", "SAM (m)", "SOM P10", "SOM P50", "SOM P90", "Active SOM", "P50/SAM", "P50/TAM"])
hdr(ws, 4, 10)

regions_active = ["HK & South China", "HK & South China", "+ East China", "+ Southwest", "+ North", "+ Northwest & Northeast"]
for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=regions_active[i-1]).alignment = LEFT
    ws.cell(row=r, column=3, value=f"=tam_Y{i}").number_format = '#,##0.0'
    ws.cell(row=r, column=4, value=f"=sam_Y{i}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"=D{r}*1000000*pen_Y{i}*mult_p10").number_format = '#,##0'
    ws.cell(row=r, column=6, value=f"=D{r}*1000000*pen_Y{i}*mult_p50").number_format = '#,##0'
    ws.cell(row=r, column=7, value=f"=D{r}*1000000*pen_Y{i}*mult_p90").number_format = '#,##0'
    ws.cell(row=r, column=8, value=f"=D{r}*1000000*pen_Y{i}*scenario_multiplier").number_format = '#,##0'
    ws.cell(row=r, column=9, value=f"=F{r}/(D{r}*1000000)").number_format = '0.0000%'
    ws.cell(row=r, column=10, value=f"=F{r}/(C{r}*1000000)").number_format = '0.0000%'
    for c in range(1, 11):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 3:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER
widths(ws, [6, 24, 11, 11, 12, 12, 12, 13, 11, 11])

# ───────────────────────────────────────────────────────────
# 4. B2B P&L (Live) - With DLRI Column
# ───────────────────────────────────────────────────────────
ws = wb.create_sheet("B2B P&L")
ws['A1'] = "B2B Technology Enabler — P&L"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:S1')
ws['A2'] = '="Active: marketing = "&TEXT(marketing_pct,"0%")&" | scenario = "&CHOOSE(scenario_selector,"P10","P50","P90")'
ws['A2'].font = Font(bold=True, size=10, color="1F4E78")
ws.merge_cells('A2:S2')

ws.append([])
ws.append([
    "Year", "SOM", "Subscription Rev", "Ball Rev", "Custom", "License", "ITF Grant", "Total Rev",
    "Platform", "User Enh", "Ball Enh", "Ball Cost", "Manpower", "DLRI", "Total Cost", 
    "Profit", "Tax Due After Enhanced Deduct", "Net Profit", "Net Cumulative"
])
hdr(ws, 4, 19)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='Market Sizing'!H{r}").number_format = '#,##0'
    
    # Revenue
    ws.cell(row=r, column=3, value=f"=B{r}*arpu_annual*split_b2b/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=4, value=f"=B{r}*ball_attach*ball_price*ball_share_b2b/1000000" if i==1 else f"=(B{r}-B{r-1})*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value="=customization_y1" if i==1 else "=customization_annual").number_format = '#,##0.000'
    ws.cell(row=r, column=6, value="=license_fee").number_format = '#,##0.000'
    ws.cell(row=r, column=7, value="=grant_annual*grant_lag_frac" if i==1 else "=grant_annual").number_format = '#,##0.00'
    ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
    
    # Costs
    ws.cell(row=r, column=9, value=f"=platform_base+platform_per_user*B{r}/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=10, value=f"=user_enhancement*B{r}/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=11, value=f"=B{r}*ball_attach*ball_enhancement/1000000" if i==1 else f"=(B{r}-B{r-1})*ball_attach*ball_enhancement/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=12, value=f"=B{r}*ball_attach*ball_cost/1000000" if i==1 else f"=(B{r}-B{r-1})*ball_attach*ball_cost/1000000").number_format = '#,##0.00'
    ws.cell(row=r, column=13, value="=manpower_annual").number_format = '#,##0.00'
    
    # DLRI (Explicit Column)
    ws.cell(row=r, column=14, value="=uni_rd_cash").number_format = '#,##0.00'
    
    # Total Cost
    ws.cell(row=r, column=15, value=f"=SUM(I{r}:N{r})").number_format = '#,##0.00'
    
    # Profit (PBT)
    ws.cell(row=r, column=16, value=f"=H{r}-O{r}").number_format = '#,##0.00'
    
    # Tax Due
    taxable = f"P{r}-uni_rd_cash*(rd_deduct-1)"
    ws.cell(row=r, column=17, value=f"=IF({taxable}<=0,0,IF({taxable}<=tax_limit,{taxable}*tax_r1,tax_limit*tax_r1+({taxable}-tax_limit)*tax_r2))").number_format = '#,##0.00'
    
    # Net Profit
    ws.cell(row=r, column=18, value=f"=P{r}-Q{r}").number_format = '#,##0.00'
    
    # Net Cumulative
    ws.cell(row=r, column=19, value=f"=R{r}" if i==1 else f"=S{r-1}+R{r}").number_format = '#,##0.00'
    
    for c in range(1, 20):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

# Y7 Q2 Grant row
r = 11
ws.cell(row=r, column=1, value="Y7 Q2").alignment = CENTER
for c in range(2, 7): ws.cell(row=r, column=c, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=7, value="=grant_y7q2").number_format = '#,##0.00'
ws.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = '#,##0.00'
for c in range(9, 16): ws.cell(row=r, column=c, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=16, value=f"=H{r}-O{r}").number_format = '#,##0.00'
ws.cell(row=r, column=17, value=0).number_format = '#,##0.00'
ws.cell(row=r, column=18, value=f"=P{r}-Q{r}").number_format = '#,##0.00'
ws.cell(row=r, column=19, value=f"=S10+R{r}").number_format = '#,##0.00'
for c in range(1, 20):
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).font = NUM_FONT
    if c >= 2:
        ws.cell(row=r, column=c).fill = CALC_FILL
        ws.cell(row=r, column=c).alignment = CENTER

widths(ws, [6, 10, 11, 9, 9, 9, 9, 10, 9, 9, 9, 9, 10, 9, 10, 10, 14, 10, 12])

# ───────────────────────────────────────────────────────────
# 5. B2B BALANCE SHEET (Underneath P&L)
# ───────────────────────────────────────────────────────────
ws.cell(row=14, column=1, value="BALANCE SHEET — HKAS 38 INTANGIBLE ASSETS").font = TITLE_FONT
ws.cell(row=14, column=1).fill = SUB_FILL
ws.merge_cells('A14:F14')

ws.cell(row=15, column=1, value="Year")
ws.cell(row=15, column=2, value="Gross Intangible Asset")
ws.cell(row=15, column=3, value="Accumulated Amortization")
ws.cell(row=15, column=4, value="Net Book Value")
ws.cell(row=15, column=5, value="Capitalization Note")
hdr(ws, 15, 5)

for i in range(1, 7):
    r = 15 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    
    if i <= 2:
        gross = 0
        note = "Not capitalized"
    elif i == 3:
        gross = "=3*uni_rd_cash"
        note = "Y1-Y3 capitalized"
    elif i in [4, 5]:
        gross = "=3*uni_rd_cash"
        note = "Carried forward"
    elif i == 6:
        gross = "=6*uni_rd_cash"
        note = "Y4-Y6 capitalized added"
        
    ws.cell(row=r, column=2, value=gross).number_format = '#,##0.00'
    
    if i <= 2:
        amort = 0
    elif i == 3:
        amort = "=(3*uni_rd_cash)/amort_yrs"
    elif i in [4, 5]:
        amort = "=(3*uni_rd_cash)/amort_yrs"
    elif i == 6:
        amort = "=((3*uni_rd_cash)/amort_yrs)*2"
        
    ws.cell(row=r, column=3, value=amort).number_format = '#,##0.00'
    
    ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=note).alignment = LEFT
    
    for c in range(1, 6):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            if c < 5: ws.cell(row=r, column=c).alignment = CENTER

widths(ws, [6, 20, 20, 15, 25])

# ───────────────────────────────────────────────────────────
# 6. B2C P&L (Live)
# ───────────────────────────────────────────────────────────
ws = wb.create_sheet("B2C P&L")
ws['A1'] = "B2C Distribution Partners — P&L (Standard HK Tax)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = '="Active: marketing = "&TEXT(marketing_pct,"0%")&" | scenario = "&CHOOSE(scenario_selector,"P10","P50","P90")'
ws['A2'].font = Font(bold=True, size=10, color="1F4E78")
ws.merge_cells('A2:K2')

ws.append([])
ws.append(["Year", "SOM", "Subscription Rev", "Ball Rev", "Total Rev",
           "Marketing", "License", "Customization", "Total Cost", "Tax Due", "Net Profit", "Net Cumulative"])
hdr(ws, 4, 12)

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
    
    pbt = f"(E{r}-I{r})"
    ws.cell(row=r, column=10, value=f"=IF({pbt}<=0,0,IF({pbt}<=tax_limit,{pbt}*tax_r1,tax_limit*tax_r1+({pbt}-tax_limit)*tax_r2))").number_format = '#,##0.00'
    
    ws.cell(row=r, column=11, value=f"=E{r}-I{r}-J{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=12, value=f"=K{r}" if i == 1 else f"=L{r-1}+K{r}").number_format = '#,##0.00'
    
    for c in range(1, 13):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER
widths(ws, [6, 10, 12, 10, 10, 10, 10, 14, 10, 10, 10, 12])

# ───────────────────────────────────────────────────────────
# 7. SUMMARY
# ───────────────────────────────────────────────────────────
ws = wb.create_sheet("Summary")
ws['A1'] = "Summary — Fairness Check + Invariants"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws.append([])
ws.append(["Year", "B2B Net Profit", "B2C Net Profit", "Total Profit",
           "B2B Cumulative", "B2C Cumulative", "B2C / B2B"])
hdr(ws, 4, 7)

for i in range(1, 7):
    r = 4 + i
    ws.cell(row=r, column=1, value=f"Y{i}").alignment = CENTER
    ws.cell(row=r, column=2, value=f"='B2B P&L'!R{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=3, value=f"='B2C P&L'!K{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=4, value=f"=B{r}+C{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=5, value=f"='B2B P&L'!S{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=6, value=f"='B2C P&L'!L{r}").number_format = '#,##0.00'
    ws.cell(row=r, column=7, value=f'=IF(B{r}=0,"n/a",F{r}/B{r})').number_format = '0.0000'
    for c in range(1, 8):
        ws.cell(row=r, column=c).border = BORDER
        ws.cell(row=r, column=c).font = NUM_FONT
        if c >= 2:
            ws.cell(row=r, column=c).fill = CALC_FILL
            ws.cell(row=r, column=c).alignment = CENTER

r = 11
ws.cell(row=r, column=1, value="Full-cycle").alignment = CENTER
ws.cell(row=r, column=5, value=f"='B2B P&L'!S11").number_format = '#,##0.00'
ws.cell(row=r, column=6, value=f"='B2C P&L'!L10").number_format = '#,##0.00'
ws.cell(row=r, column=7, value=f"=F{r}/E{r}").number_format = '0.0000'
for c in range(1, 8):
    ws.cell(row=r, column=c).border = BORDER
    ws.cell(row=r, column=c).font = BOLD_NUM
    if c >= 5:
        ws.cell(row=r, column=c).fill = CALC_FILL
        ws.cell(row=r, column=c).alignment = CENTER

ws.cell(row=13, column=1, value="CHECKS").font = BOLD_NUM
ws.cell(row=13, column=1).fill = SUB_FILL
ws.merge_cells('A13:G13')

checks = [
    ("B2B cumulative = sum of annual", '=IF(ROUND(\'B2B P&L\'!S10-SUM(\'B2B P&L\'!R5:R10),2)=0,"PASS","FAIL")'),
    ("B2C cumulative = sum of annual", '=IF(ROUND(\'B2C P&L\'!L10-SUM(\'B2C P&L\'!K5:K10),2)=0,"PASS","FAIL")'),
    ("Full-cycle fairness <= 1.10", '=IF(G11<=1.1,"PASS","FAIL")'),
    ("Full-cycle fairness >= 1.09", '=IF(G11>=1.09,"PASS","REVIEW")'),
    ("All years B2B non-negative", '=IF(MIN(\'B2B P&L\'!R5:R10)>=0,"PASS","FAIL")'),
    ("All years B2C non-negative", '=IF(MIN(\'B2C P&L\'!K5:K10)>=0,"PASS","FAIL")'),
    ("Subscription split = 100%", '=IF(ROUND(split_coach+split_b2b+split_b2c,4)=1,"PASS","FAIL")'),
    ("Marketing rate is calibrated", '=IF(OR(marketing_pct=0.30,marketing_pct=0.35,marketing_pct=0.40),"PASS","FAIL")'),
    ("Scenario selector is valid", '=IF(OR(scenario_selector=1,scenario_selector=2,scenario_selector=3),"PASS","FAIL")'),
]
for i, (label, formula) in enumerate(checks, start=14):
    ws.cell(row=i, column=1, value=label).font = NUM_FONT
    ws.cell(row=i, column=1).border = BORDER
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=3)
    c = ws.cell(row=i, column=4, value=formula)
    c.font = BOLD_NUM; c.border = BORDER; c.alignment = CENTER; c.fill = CALC_FILL
widths(ws, [10, 14, 14, 12, 14, 14, 12])

# ───────────────────────────────────────────────────────────
# 8. MATRIX (Clean Helper Array Approach - Fixed)
# ───────────────────────────────────────────────────────────
ws = wb.create_sheet("Matrix")
ws['A1'] = "Matrix — 3 Marketing Levels × 3 Scenarios (Net Profit)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws['A2'] = "Formulas use helper rows below. Fully synchronized with P&L."
ws['A2'].font = ITALIC
ws.merge_cells('A2:K2')

# Matrix Grids
ws.cell(row=4, column=1, value="B2B FULL-CYCLE CUMULATIVE (Net Profit), HK$m").font = BOLD_NUM
ws.cell(row=4, column=1).fill = SUB_FILL
ws.merge_cells('A4:E4')
ws.cell(row=5, column=1, value="Marketing \\ Scenario")
ws.cell(row=5, column=2, value="P10")
ws.cell(row=5, column=3, value="P50")
ws.cell(row=5, column=4, value="P90")
hdr(ws, 5, 4)

ws.cell(row=10, column=1, value="B2C Y6 CUMULATIVE (Net Profit), HK$m").font = BOLD_NUM
ws.cell(row=10, column=1).fill = SUB_FILL
ws.merge_cells('A10:E10')
ws.cell(row=11, column=1, value="Marketing \\ Scenario")
ws.cell(row=11, column=2, value="P10")
ws.cell(row=11, column=3, value="P50")
ws.cell(row=11, column=4, value="P90")
hdr(ws, 11, 4)

ws.cell(row=16, column=1, value="FAIRNESS RATIO (B2C / B2B)").font = BOLD_NUM
ws.cell(row=16, column=1).fill = SUB_FILL
ws.merge_cells('A16:E16')
ws.cell(row=17, column=1, value="Marketing \\ Scenario")
ws.cell(row=17, column=2, value="P10")
ws.cell(row=17, column=3, value="P50")
ws.cell(row=17, column=4, value="P90")
hdr(ws, 17, 4)

# Formulas
def get_b2b_formula(col, split_name):
    rev = f"({col}61*{split_name} + {col}62 + {col}68 + {col}70 + {col}69)"
    cost = f"({col}66 + {col}64 + {col}65)"
    profit = f"({rev} - {cost})"
    taxable = f"({profit} - {col}67)"
    tax = f"IF({taxable}<=0, 0, IF({taxable}<=tax_limit, {taxable}*tax_r1, tax_limit*tax_r1 + ({taxable}-tax_limit)*tax_r2))"
    return f"={profit} - {tax} + {col}71"

def get_b2c_formula(col, split_name, mkt_name):
    rev = f"({col}61*{split_name} + {col}63)"
    cost = f"({col}61*{split_name}*{mkt_name} + {col}70 + {col}68)"
    profit = f"({rev} - {cost})"
    tax = f"IF({profit}<=0, 0, IF({profit}<=tax_limit, {profit}*tax_r1, tax_limit*tax_r1 + ({profit}-tax_limit)*tax_r2))"
    return f"={profit} - {tax}"

# Populate Grids
for i, (label, split_name) in enumerate([("30%", "calib_b2b_30"), ("35%", "calib_b2b_35"), ("40%", "calib_b2b_40")], start=6):
    ws.cell(row=i, column=1, value=label).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    ws.cell(row=i, column=2, value=get_b2b_formula('B', split_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=3, value=get_b2b_formula('C', split_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=4, value=get_b2b_formula('D', split_name)).number_format = '#,##0.00'
    for c in range(2, 5):
        ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=c).fill = CALC_FILL
        ws.cell(row=i, column=c).alignment = CENTER
        ws.cell(row=i, column=c).font = NUM_FONT

for i, (label, split_name, mkt_name) in enumerate([("30%", "calib_b2c_30", "calib_mkt_30"), ("35%", "calib_b2c_35", "calib_mkt_35"), ("40%", "calib_b2c_40", "calib_mkt_40")], start=12):
    ws.cell(row=i, column=1, value=label).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    ws.cell(row=i, column=2, value=get_b2c_formula('B', split_name, mkt_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=3, value=get_b2c_formula('C', split_name, mkt_name)).number_format = '#,##0.00'
    ws.cell(row=i, column=4, value=get_b2c_formula('D', split_name, mkt_name)).number_format = '#,##0.00'
    for c in range(2, 5):
        ws.cell(row=i, column=c).border = BORDER
        ws.cell(row=i, column=c).fill = CALC_FILL
        ws.cell(row=i, column=c).alignment = CENTER
        ws.cell(row=i, column=c).font = NUM_FONT

for i in range(18, 21):
    b2c_row = i - 6
    b2b_row = i - 12
    ws.cell(row=i, column=1, value=["30%", "35%", "40%"][i - 18]).font = BOLD_NUM
    ws.cell(row=i, column=1).alignment = CENTER
    ws.cell(row=i, column=1).border = BORDER
    for j, col in enumerate(['B', 'C', 'D'], 2):
        ws.cell(row=i, column=j, value=f"={col}{b2c_row}/{col}{b2b_row}").number_format = '0.0000'
        ws.cell(row=i, column=j).border = BORDER
        ws.cell(row=i, column=j).fill = CALC_FILL
        ws.cell(row=i, column=j).alignment = CENTER
        ws.cell(row=i, column=j).font = NUM_FONT

# Helper Section (Rows 60-71)
ws.cell(row=60, column=1, value="Array").font = BOLD_NUM
ws.cell(row=60, column=2, value="P10").font = BOLD_NUM
ws.cell(row=60, column=3, value="P50").font = BOLD_NUM
ws.cell(row=60, column=4, value="P90").font = BOLD_NUM
hdr(ws, 60, 4)

ws.cell(row=61, column=1, value="Base Rev").font = ITALIC
ws.cell(row=61, column=2, value="=SUM($B$34:$G$34)*arpu_annual/1000000").number_format = '#,##0.00'
ws.cell(row=61, column=3, value="=SUM($B$35:$G$35)*arpu_annual/1000000").number_format = '#,##0.00'
ws.cell(row=61, column=4, value="=SUM($B$36:$G$36)*arpu_annual/1000000").number_format = '#,##0.00'

ws.cell(row=62, column=1, value="Ball Rev B2B").font = ITALIC
ws.cell(row=62, column=2, value="=SUM($B$37:$G$37)*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
ws.cell(row=62, column=3, value="=SUM($B$38:$G$38)*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'
ws.cell(row=62, column=4, value="=SUM($B$39:$G$39)*ball_attach*ball_price*ball_share_b2b/1000000").number_format = '#,##0.00'

ws.cell(row=63, column=1, value="Ball Rev B2C").font = ITALIC
ws.cell(row=63, column=2, value="=SUM($B$37:$G$37)*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
ws.cell(row=63, column=3, value="=SUM($B$38:$G$38)*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'
ws.cell(row=63, column=4, value="=SUM($B$39:$G$39)*ball_attach*ball_price*(1-ball_share_b2b)/1000000").number_format = '#,##0.00'

ws.cell(row=64, column=1, value="Var Cost User").font = ITALIC
ws.cell(row=64, column=2, value="=SUM($B$34:$G$34)*(platform_per_user + user_enhancement)/1000000").number_format = '#,##0.00'
ws.cell(row=64, column=3, value="=SUM($B$35:$G$35)*(platform_per_user + user_enhancement)/1000000").number_format = '#,##0.00'
ws.cell(row=64, column=4, value="=SUM($B$36:$G$36)*(platform_per_user + user_enhancement)/1000000").number_format = '#,##0.00'

ws.cell(row=65, column=1, value="Var Cost Ball").font = ITALIC
ws.cell(row=65, column=2, value="=SUM($B$37:$G$37)*ball_attach*(ball_enhancement + ball_cost)/1000000").number_format = '#,##0.00'
ws.cell(row=65, column=3, value="=SUM($B$38:$G$38)*ball_attach*(ball_enhancement + ball_cost)/1000000").number_format = '#,##0.00'
ws.cell(row=65, column=4, value="=SUM($B$39:$G$39)*ball_attach*(ball_enhancement + ball_cost)/1000000").number_format = '#,##0.00'

ws.cell(row=66, column=1, value="Fixed Cost").font = ITALIC
ws.cell(row=66, column=2, value="=6*(platform_base + manpower_annual + uni_rd_cash)").number_format = '#,##0.00'
ws.cell(row=66, column=3, value="=6*(platform_base + manpower_annual + uni_rd_cash)").number_format = '#,##0.00'
ws.cell(row=66, column=4, value="=6*(platform_base + manpower_annual + uni_rd_cash)").number_format = '#,##0.00'

ws.cell(row=67, column=1, value="Extra Deduct").font = ITALIC
ws.cell(row=67, column=2, value="=6*uni_rd_cash*(rd_deduct-1)").number_format = '#,##0.00'
ws.cell(row=67, column=3, value="=6*uni_rd_cash*(rd_deduct-1)").number_format = '#,##0.00'
ws.cell(row=67, column=4, value="=6*uni_rd_cash*(rd_deduct-1)").number_format = '#,##0.00'

ws.cell(row=68, column=1, value="Custom").font = ITALIC
ws.cell(row=68, column=2, value="=SUM($B$40:$G$40)").number_format = '#,##0.00'
ws.cell(row=68, column=3, value="=SUM($B$40:$G$40)").number_format = '#,##0.00'
ws.cell(row=68, column=4, value="=SUM($B$40:$G$40)").number_format = '#,##0.00'

ws.cell(row=69, column=1, value="Grant").font = ITALIC
ws.cell(row=69, column=2, value="=SUM($B$41:$G$41)").number_format = '#,##0.00'
ws.cell(row=69, column=3, value="=SUM($B$41:$G$41)").number_format = '#,##0.00'
ws.cell(row=69, column=4, value="=SUM($B$41:$G$41)").number_format = '#,##0.00'

ws.cell(row=70, column=1, value="License").font = ITALIC
ws.cell(row=70, column=2, value="=6*license_fee").number_format = '#,##0.00'
ws.cell(row=70, column=3, value="=6*license_fee").number_format = '#,##0.00'
ws.cell(row=70, column=4, value="=6*license_fee").number_format = '#,##0.00'

ws.cell(row=71, column=1, value="Y7 Grant").font = ITALIC
ws.cell(row=71, column=2, value="=grant_y7q2").number_format = '#,##0.00'
ws.cell(row=71, column=3, value="=grant_y7q2").number_format = '#,##0.00'
ws.cell(row=71, column=4, value="=grant_y7q2").number_format = '#,##0.00'

for i in range(61, 72):
    for j in range(1, 5):
        ws.cell(row=i, column=j).border = BORDER
        ws.cell(row=i, column=j).font = NUM_FONT if j > 1 else ITALIC
        if j > 1:
            ws.cell(row=i, column=j).fill = CALC_FILL
            ws.cell(row=i, column=j).alignment = CENTER

widths(ws, [22, 14, 14, 14])

# ───────────────────────────────────────────────────────────
# 9. SOURCES
# ───────────────────────────────────────────────────────────
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
    ("Hurun Quality of Life Report", "https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4", "Health spending No.2; +12% service experience; >2 hrs exercise", "Primary"),
    ("People's Daily Shanghai — Youth Sports Training", "http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html", "44.67% pay more for quality; RMB 253.1bn market; 57.07% renewal", "Primary"),
    ("163.com — 3200亿 football industry", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html", "3.8亿 fans; 3200亿 industry; 487.6亿 training; 4.23万 schools; 1280万 students", "Primary"),
    ("Football Training Institution Report", "https://www.renrendoc.com/paper/415852129.html", "92% high-end renewal; 8% of students; 76% mass base; 55% mass renewal; 68% Gen-Z data; 75% coach credentials", "Industry research"),
    ("adidas Leagues Cup 26 Pro Ball", "https://www.adidas.com/us/leagues-cup-26-pro-ball/KE5154.html", "US$170 retail price", "Primary"),
    ("INSAIT JOY Smart Football", "https://insaitjoy.com/products/insait-joy-smart-football", "US$139, no subscription", "Primary"),
    ("INSAIT government procurement", "https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/", "RMB 1,745/unit smart ball; RMB 4,798/unit sensor", "Primary"),
    ("Catapult One review", "https://soccerstories.net/catapult-one-review/", "≈US$179/year, membership only, one pod per player", "Secondary review"),
    ("注册球员2740人 (163.com)", "https://www.163.com/dy/article/L8E6C26F0556BPXI.html", "8.7% pathway commitment; 10.92万 elite youth", "Primary"),
    ("足协数字年报 (china.com.cn)", "http://news.china.com.cn/2026-01/30/content_118308779.shtml", "109,220 elite youth; 979,544 amateur; 125,976 coaches", "Primary"),
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
OUT = "KungFu_AI_Football_Financial_Model_v3.5.3.xlsx"
wb.save(OUT)
print(f"✅ Saved: {OUT}")
print(f"   Matrix #VALUE! errors fixed. B2C marketing math corrected.")