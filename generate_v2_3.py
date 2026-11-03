"""
KungFu AI Football v2.3 — Workbook Generator
Run: python generate_v2_3.py
Requires: openpyxl
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter

# ── STYLES ──
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUB_FONT = Font(bold=True, size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
NUM_FONT = Font(size=10)
BOLD_NUM = Font(size=10, bold=True)
BORDER = Border(left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin'))
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
TITLE_FONT = Font(bold=True, size=14)
ITALIC = Font(italic=True, size=9)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT; cell.fill = HEADER_FILL
        cell.alignment = CENTER; cell.border = BORDER


def style_data(ws, r0, r1, ncols, bold=False):
    f = BOLD_NUM if bold else NUM_FONT
    for r in range(r0, r1 + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = f; cell.border = BORDER
            cell.alignment = LEFT if c == 1 else CENTER


def numfmt(ws, r0, r1, c0, c1, fmt='#,##0.00'):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(row=r, column=c).number_format = fmt


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = openpyxl.Workbook()
wb.remove(wb.active)

# ══════════════════════════════════════════════════════
# 1. COVER
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Cover")
ws.column_dimensions['A'].width = 4
ws.column_dimensions['B'].width = 38
ws.column_dimensions['C'].width = 70

ws['B2'] = "Kung Fu AI Football — 6-Year Forecast Model"
ws['B2'].font = Font(bold=True, size=18)
ws.merge_cells('B2:C2')

cover = [
    ("File", "KungFu_AI_Football_v2.3_ITF_RTH.xlsx"),
    ("Version", "2.3 (Fix A + Fix B + §3.6 SOM rationale)"),
    ("Scenario", "SOM P50 (326,520 users by Y6) · 30% B2C marketing · Split 12.86/37.14"),
    ("Currency", "HK$ million unless stated otherwise"),
    ("B2C Partners", "Anta · Li Ning · Amer Sports · Asia Partners IFBD Ltd (Svenska Capital)"),
    ("ITF RTH Grant", "2 BSc Y1-Y3 + 2 BSc Y4-Y6 · HK$0.48m/year gross · 6-month lag"),
    ("Private Cloud Platform", "Model D Hybrid: HK$0.20m base + HK$2.00/user/year"),
    ("Customization (v2.3)", "Hybrid: HK$0.10m Y1 + HK$0.025m/yr Y2-Y6 (both B2B rev & B2C cost)"),
    ("Fairness Ratio (Y6)", "B2C/B2B = 1.0969 (within 10% ceiling)"),
]
for i, (k, v) in enumerate(cover, start=4):
    ws.cell(row=i, column=2, value=k).font = SUB_FONT
    ws.cell(row=i, column=2).fill = SUB_FILL
    ws.cell(row=i, column=2).border = BORDER
    ws.cell(row=i, column=3, value=v).font = NUM_FONT
    ws.cell(row=i, column=3).border = BORDER

ws['B14'] = "Key fixes in v2.3:"
ws['B14'].font = SUB_FONT
notes = [
    "Fix A — Customization fee added as B2C cost line (removes double count)",
    "Fix B — Hybrid structure: HK$0.10m Y1 + HK$0.025m/yr Y2-Y6 (total HK$0.225m)",
    "§3.6 — SOM captures affordable/willing-to-pay segment, not Alibaba low-end",
]
for i, n in enumerate(notes, start=15):
    ws.cell(row=i, column=2, value=f"• {n}").font = NUM_FONT
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=3)

ws['B19'] = "Note: Static snapshot. Regenerate via Python script for alternative scenarios."
ws['B19'].font = ITALIC
ws.merge_cells('B19:C19')

# ══════════════════════════════════════════════════════
# 2. ASSUMPTIONS
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Assumptions")
ws['A1'] = "Assumptions Register"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:C1')
ws.append([])
ws.append(["Category", "Parameter", "Value"])
style_header(ws, 3, 3)

assumptions = [
    ("Subscription", "Coach/Player share", "50.00%"),
    ("Subscription", "B2B subscription share", "12.86%"),
    ("Subscription", "B2C subscription share", "37.14%"),
    ("Pricing", "Weighted ARPU", "HK$1,592/year"),
    ("Pricing", "Tier mix", "Local 60% · National 25% · World Cup 12% · Legends 3%"),
    ("Ball", "Selling price", "HK$980"),
    ("Ball", "Manufacturing cost (B2B-borne)", "HK$250"),
    ("Ball", "B2B ball revenue share", "35% (HK$343/ball)"),
    ("Ball", "B2C ball revenue share", "65% (HK$637/ball)"),
    ("Ball", "Attach rate (new users)", "50%"),
    ("Cost", "Platform Base", "HK$0.20m/year (FIXED)"),
    ("Cost", "Platform Compute", "HK$2.00/user/year (VARIABLE)"),
    ("Cost", "User Enhancement", "HK$5.00/user/year (VARIABLE)"),
    ("Cost", "Ball Enhancement", "HK$10/ball (VARIABLE)"),
    ("Cost", "Marketing & distribution (B2C)", "30% of gross subscription"),
    ("Revenue", "Exclusive license B2C->B2B", "HK$0.20m/year"),
    ("Revenue", "Customization (B2B revenue)", "HK$0.10m Y1 + HK$0.025m/yr Y2-Y6"),
    ("Cost", "Customization (B2C cost)", "HK$0.10m Y1 + HK$0.025m/yr Y2-Y6"),
    ("ITF RTH", "BSc manpower", "2 (Y1-Y3) + 2 (Y4-Y6)"),
    ("ITF RTH", "Monthly rate", "HK$20,000"),
    ("ITF RTH", "Annual gross cost", "HK$0.48m/year x 6 years"),
    ("ITF RTH", "Grant coverage", "100% eligible"),
    ("ITF RTH", "Reimbursement lag", "6 months"),
    ("Market", "Total TAM Y6 midpoint", "300m"),
    ("Market", "SAM Y6 midpoint", "45m"),
    ("Market", "SOM Y6 P50", "326,520"),
    ("Market", "SOM confidence band", "P10 ~180k; P90 ~490k"),
]
for row in assumptions:
    ws.append(list(row))
style_data(ws, 4, 3 + len(assumptions), 3)
set_widths(ws, [14, 38, 60])

# ══════════════════════════════════════════════════════
# 3. MARKET SIZING
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Market Sizing")
ws['A1'] = "Market Sizing — Total TAM / SAM / SOM"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:I1')
ws.append([])
ws.append(["Year", "Regions Active", "Total TAM (m)", "SAM (m)",
           "SOM P10", "SOM P50", "SOM P90", "P50/SAM", "P50/TAM"])
style_header(ws, 3, 9)

mkt = [
    ("Y1", "HK & South China", 110.0, 18.0, 4100, 7380, 11100, 0.00041, 0.000067),
    ("Y2", "HK & South China", 110.0, 18.0, 10200, 18450, 27700, 0.00103, 0.000168),
    ("Y3", "+ East China", 185.0, 29.0, 57100, 103835, 155800, 0.00358, 0.000561),
    ("Y4", "+ Southwest", 235.0, 36.5, 90100, 163800, 245700, 0.00449, 0.000697),
    ("Y5", "+ North", 270.0, 41.5, 126700, 230345, 345500, 0.00555, 0.000853),
    ("Y6", "+ Northwest & Northeast", 300.0, 45.0, 179600, 326520, 489800, 0.00726, 0.001088),
]
for row in mkt: ws.append(list(row))
style_data(ws, 4, 3 + len(mkt), 9)
numfmt(ws, 4, 3 + len(mkt), 3, 4, '#,##0.0')
numfmt(ws, 4, 3 + len(mkt), 5, 7, '#,##0')
numfmt(ws, 4, 3 + len(mkt), 8, 9, '0.000%')
set_widths(ws, [6, 26, 14, 10, 12, 14, 12, 10, 10])

# §3.6 note
start = 11
ws.cell(row=start, column=1, value="§3.6 — Why SOM Captures the 'Affordable and Willing-to-Pay' Segment").font = Font(bold=True, size=12)
ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=9)
notes = [
    "China's sports market is K-shaped: low-end Alibaba/Taobao (< RMB 41) = 56.3% of football units but only 23.9% of revenue.",
    "Mid-tier (RMB 41-89) = ~40% of units, 57.7% of revenue — the 'profit bullseye'.",
    "Premium (> RMB 189) = only 2.8% of units but outsized brand premium; price is just 6.4% of purchase decision.",
    "Amer Sports Greater China revenue grew 43.4% in 2025 to US$1.86bn; Anta's Descente/Kolon grew 50-65%.",
    "Li Ning gross margin ~49-50% vs Xtep 42.8%, 361 deg 41.5% — evidence of premium pricing power.",
    "SOM = middle-to-high-end segment that spends HK$3,000-5,000/yr on sports gear and values performance tech.",
    "The low-end Alibaba market is STRUCTURALLY EXCLUDED — a HK$980 ball + subscription cannot monetise on commodity purchase.",
    "SOM Y6 penetration = 0.726% of SAM — conservative claim about the slice that can afford and value the product.",
]
for i, n in enumerate(notes, start=start + 1):
    ws.cell(row=i, column=1, value=f"• {n}").font = NUM_FONT
    ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=9)

# ══════════════════════════════════════════════════════
# 4. B2B P&L
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("B2B P&L")
ws['A1'] = "B2B Technology Enabler — 6-Year P&L (P50, v2.3 Hybrid Customization)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:P1')
ws.append([])
headers = ["Year", "SOM", "Sub Rev", "Ball Rev", "Custom", "License", "ITF Grant",
           "Total Rev", "Platform", "User Enh", "Ball Enh", "Ball Cost",
           "Manpower", "Total Cost", "Profit", "Cumulative"]
ws.append(headers)
style_header(ws, 3, 16)

b2b = [
    ("Y1", 7380, 1.51, 1.27, 0.10, 0.20, 0.24, 3.32, 0.21, 0.04, 0.04, 0.92, 0.48, 1.69, 1.63, 1.63),
    ("Y2", 18450, 3.78, 1.90, 0.025, 0.20, 0.48, 6.39, 0.24, 0.09, 0.06, 1.38, 0.48, 2.25, 4.14, 5.77),
    ("Y3", 103835, 21.26, 14.64, 0.025, 0.20, 0.48, 36.61, 0.41, 0.52, 0.43, 10.67, 0.48, 12.51, 24.10, 29.87),
    ("Y4", 163800, 33.54, 10.28, 0.025, 0.20, 0.48, 44.53, 0.53, 0.82, 0.30, 7.50, 0.48, 9.63, 34.90, 64.77),
    ("Y5", 230345, 47.16, 11.41, 0.025, 0.20, 0.48, 59.28, 0.66, 1.15, 0.33, 8.32, 0.48, 10.94, 48.34, 113.11),
    ("Y6", 326520, 66.85, 16.49, 0.025, 0.20, 0.48, 84.05, 0.85, 1.63, 0.48, 12.02, 0.48, 15.46, 68.59, 181.70),
    ("Y7 Q2", "—", "—", "—", "—", "—", 0.24, 0.24, "—", "—", "—", "—", "—", 0.00, 0.24, 181.94),
]
for row in b2b: ws.append(list(row))
style_data(ws, 4, 3 + len(b2b), 16)
numfmt(ws, 4, 3 + len(b2b), 2, 2, '#,##0')
numfmt(ws, 4, 3 + len(b2b), 3, 16, '#,##0.00')
set_widths(ws, [8, 10, 9, 9, 9, 9, 9, 10, 9, 9, 9, 10, 10, 10, 10, 11])
ws.cell(row=12, column=1, value="Note: Custom column = B2B customization revenue (Hybrid: 0.10 Y1, 0.025 Y2-Y6).").font = ITALIC
ws.merge_cells('A12:P12')
ws.cell(row=13, column=1, value="ITF Grant column = cash received (Y1=0.24, Y2-Y6=0.48, Y7Q2=0.24). Manpower column = gross ITF cost (0.48/yr).").font = ITALIC
ws.merge_cells('A13:P13')

# ══════════════════════════════════════════════════════
# 5. B2C P&L
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("B2C P&L")
ws['A1'] = "B2C Distribution Partners — 6-Year P&L (P50, incl. Customization Cost — Fix A)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:K1')
ws.append([])
headers = ["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev",
           "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"]
ws.append(headers)
style_header(ws, 3, 11)

b2c = [
    ("Y1", 7380, 4.36, 2.35, 6.71, 3.52, 0.20, 0.10, 3.82, 2.89, 2.89),
    ("Y2", 18450, 10.91, 3.52, 14.43, 8.81, 0.20, 0.025, 9.035, 5.40, 8.29),
    ("Y3", 103835, 61.39, 27.20, 88.59, 49.59, 0.20, 0.025, 49.815, 38.78, 47.07),
    ("Y4", 163800, 96.87, 19.10, 115.97, 78.23, 0.20, 0.025, 78.455, 37.52, 84.59),
    ("Y5", 230345, 136.21, 21.19, 157.40, 110.01, 0.20, 0.025, 110.235, 47.17, 131.76),
    ("Y6", 326520, 193.08, 30.63, 223.71, 155.95, 0.20, 0.025, 156.175, 67.54, 199.30),
]
for row in b2c: ws.append(list(row))
style_data(ws, 4, 3 + len(b2c), 11)
numfmt(ws, 4, 3 + len(b2c), 2, 2, '#,##0')
numfmt(ws, 4, 3 + len(b2c), 3, 11, '#,##0.00')
set_widths(ws, [8, 10, 10, 10, 10, 10, 10, 14, 10, 10, 11])
ws.cell(row=11, column=1, value="v2.3 Fix A: Customization is now a B2C cost line (0.10 Y1, 0.025 Y2-Y6). Removes double count from v2.2.").font = ITALIC
ws.merge_cells('A11:K11')

# ══════════════════════════════════════════════════════
# 6. COST BEHAVIOUR
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Cost Behaviour")
ws['A1'] = "Fixed vs Variable Cost Decomposition"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:E1')
ws.append([])
ws.append(["Year", "Fixed (Platform Base + ITF Gross)", "Variable (Platform Compute + User Enh + Ball Enh + Ball Cost)", "Total B2B Cost", "Fixed %"])
style_header(ws, 3, 5)

costbeh = [
    ("Y1", 0.68, 1.01, 1.69, "40.2%"),
    ("Y2", 0.68, 1.57, 2.25, "30.2%"),
    ("Y3", 0.68, 11.83, 12.51, "5.4%"),
    ("Y4", 0.68, 8.95, 9.63, "7.1%"),
    ("Y5", 0.68, 10.26, 10.94, "6.2%"),
    ("Y6", 0.68, 14.78, 15.46, "4.4%"),
]
for row in costbeh: ws.append(list(row))
style_data(ws, 4, 3 + len(costbeh), 5)
numfmt(ws, 4, 3 + len(costbeh), 2, 4, '#,##0.00')
set_widths(ws, [8, 35, 65, 16, 10])
ws.cell(row=12, column=1, value="Fixed per user Y1: HK$0.68m / 7,380 = HK$92.14").font = ITALIC
ws.merge_cells('A12:E12')
ws.cell(row=13, column=1, value="Fixed per user Y6: HK$0.68m / 326,520 = HK$2.08 → operating leverage 44x").font = ITALIC
ws.merge_cells('A13:E13')
ws.cell(row=14, column=1, value="Break-even on fixed base: ~HK$0.68m / ~HK$240 contribution/user ≈ 2,835 users. Y1 SOM = 7,380 (2.6x break-even).").font = ITALIC
ws.merge_cells('A14:E14')
ws.cell(row=15, column=1, value="B2C customization (0.10 Y1 + 0.025 x 5) is SEMI-FIXED — contractual schedule, not SOM-linked.").font = ITALIC
ws.merge_cells('A15:E15')

# ══════════════════════════════════════════════════════
# 7. ITF RTH GRANT
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("ITF RTH Grant")
ws['A1'] = "ITF RTH Manpower Grant — Cash-Timing Schedule"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:F1')
ws.append([])
ws.append(["Year", "Gross Manpower", "Grant Received", "Net Impact", "Cumulative Net Impact", "Notes"])
style_header(ws, 3, 6)

grant = [
    ("Y1", 0.48, 0.24, -0.24, -0.24, "Only H1 Y1 reimbursed within Y1"),
    ("Y2", 0.48, 0.48, 0.00, -0.24, "H2 Y1 + H1 Y2 received"),
    ("Y3", 0.48, 0.48, 0.00, -0.24, "Steady state"),
    ("Y4", 0.48, 0.48, 0.00, -0.24, "Steady state"),
    ("Y5", 0.48, 0.48, 0.00, -0.24, "Steady state"),
    ("Y6", 0.48, 0.48, 0.00, -0.24, "H2 Y6 not yet received"),
    ("Y7 Q2", 0.00, 0.24, 0.24, 0.00, "Final 2 x 20k x 6 months"),
    ("Total", 2.88, 2.88, 0.00, "—", "Fully reimbursed by Y7 Q2"),
]
for row in grant: ws.append(list(row))
style_data(ws, 4, 3 + len(grant), 6)
numfmt(ws, 4, 3 + len(grant), 2, 5, '#,##0.00')
set_widths(ws, [10, 16, 16, 14, 20, 36])

# ══════════════════════════════════════════════════════
# 8. CONFIDENCE ANALYSIS
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Confidence Analysis")
ws['A1'] = "Confidence Analysis — P10 / P50 / P90"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws.append([])
ws.append(["Scenario", "Y6 SOM", "SOM / SAM", "B2B Cum Y6", "B2C Cum Y6", "Fairness Ratio", "Interpretation"])
style_header(ws, 3, 7)

conf = [
    ("P10 Conservative", 179600, 0.00399, 100.0, 110.0, 1.10, "Slow activation, high CAC, weak uptake"),
    ("P50 Base (v2.3)", 326520, 0.00726, 181.70, 199.30, 1.0969, "Base case — this workbook"),
    ("P90 Optimistic", 489800, 0.01089, 273.0, 300.0, 1.10, "Fast activation, falling CAC, strong uptake"),
]
for row in conf: ws.append(list(row))
style_data(ws, 4, 3 + len(conf), 7)
numfmt(ws, 4, 3 + len(conf), 2, 2, '#,##0')
numfmt(ws, 4, 3 + len(conf), 3, 3, '0.000%')
numfmt(ws, 4, 3 + len(conf), 4, 5, '#,##0.00')
numfmt(ws, 4, 3 + len(conf), 6, 6, '0.0000')
set_widths(ws, [20, 12, 12, 14, 14, 14, 45])
ws.cell(row=8, column=1, value="Combined uncertainty: ±45% (partner activation ±30%, CAC ±20%, uptake ±25%).").font = ITALIC
ws.merge_cells('A8:G8')
ws.cell(row=9, column=1, value="P10 ≈ 55% of P50; P90 ≈ 150% of P50. Fairness ratio preserved because subscription split is fixed.").font = ITALIC
ws.merge_cells('A9:G9')

# ══════════════════════════════════════════════════════
# 9. SUMMARY
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Summary")
ws['A1'] = "Summary — Annual and Cumulative Profit (v2.3)"
ws['A1'].font = TITLE_FONT
ws.merge_cells('A1:G1')
ws.append([])
ws.append(["Year", "B2B Profit", "B2C Profit", "Total Profit",
           "B2B Cumulative", "B2C Cumulative", "B2C / B2B"])
style_header(ws, 3, 7)

summary = [
    ("Y1", 1.63, 2.89, 4.52, 1.63, 2.89, 1.77),
    ("Y2", 4.14, 5.40, 9.54, 5.77, 8.29, 1.44),
    ("Y3", 24.10, 38.78, 62.88, 29.87, 47.07, 1.58),
    ("Y4", 34.90, 37.52, 72.42, 64.77, 84.59, 1.31),
    ("Y5", 48.34, 47.17, 95.51, 113.11, 131.76, 1.16),
    ("Y6", 68.59, 67.54, 136.13, 181.70, 199.30, 1.0969),
]
for row in summary: ws.append(list(row))
style_data(ws, 4, 3 + len(summary), 7)
numfmt(ws, 4, 3 + len(summary), 2, 6, '#,##0.00')
numfmt(ws, 4, 3 + len(summary), 7, 7, '0.0000')
set_widths(ws, [8, 12, 12, 12, 14, 14, 12])
ws.cell(row=11, column=1, value="Y7 Q2 grant catch-up: +HK$0.24m → B2B cumulative = HK$181.94m. B2C unchanged at HK$199.30m. Final ratio = 1.0955.").font = ITALIC
ws.merge_cells('A11:G11')
ws.cell(row=12, column=1, value="v2.3 Fix A: B2C profit reduced by HK$0.225m vs v2.2 (customization cost now shown). Fairness improved from 1.0974 to 1.0969.").font = ITALIC
ws.merge_cells('A12:G12')

# ══════════════════════════════════════════════════════
# 10. CUMULATIVE CHART
# ══════════════════════════════════════════════════════
ws = wb.create_sheet("Cumulative Chart")
ws['A1'] = "Cumulative Profit — B2B vs B2C (v2.3)"
ws['A1'].font = TITLE_FONT
ws.append([])
ws.append(["Year", "B2B Cumulative", "B2C Cumulative"])
style_header(ws, 3, 3)

chart_data = [
    ("Y1", 1.63, 2.89),
    ("Y2", 5.77, 8.29),
    ("Y3", 29.87, 47.07),
    ("Y4", 64.77, 84.59),
    ("Y5", 113.11, 131.76),
    ("Y6", 181.70, 199.30),
]
for row in chart_data: ws.append(list(row))
style_data(ws, 4, 3 + len(chart_data), 3)
numfmt(ws, 4, 3 + len(chart_data), 2, 3, '#,##0.00')
set_widths(ws, [10, 18, 18])

chart = LineChart()
chart.title = "Cumulative Profit: B2B vs B2C (HK$ million)"
chart.style = 12
chart.y_axis.title = "HK$ million"
chart.x_axis.title = "Year"
data = Reference(ws, min_col=2, min_row=3, max_col=3, max_row=3 + len(chart_data))
cats = Reference(ws, min_col=1, min_row=4, max_row=3 + len(chart_data))
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.height = 10
chart.width = 20
ws.add_chart(chart, "E3")

# ── SAVE ──
OUT = "KungFu_AI_Football_v2.3_ITF_RTH.xlsx"
wb.save(OUT)
print(f"✅ Workbook saved: {OUT}")
print(f"   Sheets: {wb.sheetnames}")