import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import LineChart, Reference

# ============================================================
# 1. Create workbook
# ============================================================
wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

# Styles
header_font = Font(bold=True, color="FFFFFF", size=11)
header_fill = PatternFill("solid", fgColor="1F4E78")
subheader_font = Font(bold=True, size=10)
subheader_fill = PatternFill("solid", fgColor="D9E1F2")
number_font = Font(size=10)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

def style_header(ws, row, cols):
    for col in range(1, cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

def style_data(ws, start_row, end_row, start_col, end_col):
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = number_font
            cell.border = thin_border
            if c == 1:
                cell.alignment = left_align
            else:
                cell.alignment = center_align

# ============================================================
# 2. Assumptions sheet
# ============================================================
ws_assump = wb.create_sheet("Assumptions")
ws_assump.title = "Assumptions"
ws_assump.column_dimensions['A'].width = 45
ws_assump.column_dimensions['B'].width = 25
ws_assump.column_dimensions['C'].width = 60

ws_assump['A1'] = "Kung Fu AI Football – 6-Year Forecast Assumptions"
ws_assump['A1'].font = Font(bold=True, size=14)
ws_assump.merge_cells('A1:C1')

assumptions = [
    ("Parameter", "Value", "Notes"),
    ("Target SAM by Year 6 (3%)", "304,260 users", "3% of 10.142m active football youth"),
    ("Penetration path", "Y1 0.2×, Y2 0.5×, Y3 0.7×, Y4 0.8×, Y5 0.9×, Y6 1.0×", "Applied to activated SAM each year"),
    ("Weighted ARPU", "HK$1,592 / year", "Tier mix: Local 60%, National 25%, World Cup 12%, Legends 3%"),
    ("Coach/player share", "50%", "Taken first from gross subscription"),
    ("B2B subscription share", "10.43%", "Of gross subscription"),
    ("B2C subscription share", "39.57%", "Of gross subscription"),
    ("Ball attach rate", "50% of new paying users", "One ball per two new users"),
    ("Ball selling price", "HK$980", ""),
    ("Ball manufacturing cost", "HK$250", "Borne entirely by B2B"),
    ("B2B ball revenue share", "35%", "Of ball selling price"),
    ("B2C ball revenue share", "65%", "Of ball selling price"),
    ("B2B enhancement cost", "HK$5 / user / year + HK$10 / ball", "Proportional to active users and balls sold"),
    ("B2B R&D", "HK$0.67m / year for Y1–Y3", "Total HK$2.0m amortised over 3 years"),
    ("License fee (B2C → B2B)", "HK$0.2m / year", "Exclusive annual license"),
    ("B2C marketing cost", "35% of gross subscription", "Distribution and marketing"),
    ("Taxes, payment fees, support, warranty", "Excluded", "Planning model only"),
]

for i, row in enumerate(assumptions, start=2):
    for j, val in enumerate(row, start=1):
        cell = ws_assump.cell(row=i, column=j, value=val)
        cell.border = thin_border
        if i == 2:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
        else:
            cell.font = number_font
            cell.alignment = left_align if j == 1 else center_align

# ============================================================
# 3. B2B P&L sheet
# ============================================================
ws_b2b = wb.create_sheet("B2B P&L")
ws_b2b.title = "B2B P&L"

headers_b2b = [
    "Year", "Paying Users", "Gross Subscription (HK$m)",
    "B2B Sub Share", "B2B Sub Revenue (HK$m)",
    "Ball Units", "Ball Revenue (HK$m)", "B2B Ball Share",
    "B2B Ball Revenue (HK$m)", "License Fee (HK$m)",
    "Total B2B Revenue (HK$m)", "R&D (HK$m)",
    "User Enhancement (HK$m)", "Ball Enhancement (HK$m)",
    "Ball Manufacturing Cost (HK$m)", "Total B2B Costs (HK$m)",
    "B2B Profit (HK$m)", "Cumulative B2B Profit (HK$m)"
]
ws_b2b.append(headers_b2b)
style_header(ws_b2b, 1, len(headers_b2b))

# Data (computed)
b2b_data = [
    [1, 6930, 11.03, 0.1043, 1.15, 3465, 3.40, 0.35, 1.19, 0.20, 2.54, 0.67, 0.03, 0.03, 0.87, 1.61, 0.93, 0.93],
    [2, 17325, 27.59, 0.1043, 2.88, 5198, 5.09, 0.35, 1.78, 0.20, 4.86, 0.67, 0.09, 0.05, 1.30, 2.11, 2.75, 3.69],
    [3, 97104, 154.59, 0.1043, 16.12, 39890, 39.09, 0.35, 13.68, 0.20, 30.01, 0.67, 0.49, 0.40, 9.97, 11.53, 18.48, 22.17],
    [4, 152592, 242.93, 0.1043, 25.34, 27744, 27.19, 0.35, 9.52, 0.20, 35.05, 0.00, 0.76, 0.28, 6.94, 7.98, 27.08, 49.25],
    [5, 214596, 341.64, 0.1043, 35.63, 31002, 30.38, 0.35, 10.63, 0.20, 46.47, 0.00, 1.07, 0.31, 7.75, 9.13, 37.33, 86.58],
    [6, 304260, 484.38, 0.1043, 50.52, 44832, 43.94, 0.35, 15.38, 0.20, 66.10, 0.00, 1.52, 0.45, 11.21, 13.18, 52.92, 139.50],
]

for row in b2b_data:
    ws_b2b.append(row)

style_data(ws_b2b, 2, 7, 1, len(headers_b2b))

# Number formats
for r in range(2, 8):
    for c in range(2, len(headers_b2b) + 1):
        ws_b2b.cell(row=r, column=c).number_format = '#,##0.00'
    ws_b2b.cell(row=r, column=2).number_format = '#,##0'
    ws_b2b.cell(row=r, column=6).number_format = '#,##0'

# Column widths
widths_b2b = [6, 12, 18, 12, 16, 10, 16, 12, 16, 12, 18, 10, 16, 16, 18, 16, 14, 20]
for i, w in enumerate(widths_b2b, start=1):
    ws_b2b.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

# ============================================================
# 4. B2C P&L sheet
# ============================================================
ws_b2c = wb.create_sheet("B2C P&L")
ws_b2c.title = "B2C P&L"

headers_b2c = [
    "Year", "Paying Users", "Gross Subscription (HK$m)",
    "B2C Sub Share", "B2C Sub Revenue (HK$m)",
    "Ball Units", "Ball Revenue (HK$m)", "B2C Ball Share",
    "B2C Ball Revenue (HK$m)", "Total B2C Revenue (HK$m)",
    "Marketing Cost (HK$m)", "License Fee (HK$m)",
    "Total B2C Costs (HK$m)", "B2C Profit (HK$m)",
    "Cumulative B2C Profit (HK$m)"
]
ws_b2c.append(headers_b2c)
style_header(ws_b2c, 1, len(headers_b2c))

b2c_data = [
    [1, 6930, 11.03, 0.3957, 4.37, 3465, 3.40, 0.65, 2.21, 6.57, 3.86, 0.20, 4.06, 2.51, 2.51],
    [2, 17325, 27.59, 0.3957, 10.92, 5198, 5.09, 0.65, 3.31, 14.23, 9.65, 0.20, 9.85, 4.37, 6.88],
    [3, 97104, 154.59, 0.3957, 61.17, 39890, 39.09, 0.65, 25.41, 86.58, 54.11, 0.20, 54.31, 32.27, 39.16],
    [4, 152592, 242.93, 0.3957, 96.13, 27744, 27.19, 0.65, 17.67, 113.80, 85.02, 0.20, 85.22, 28.57, 67.73],
    [5, 214596, 341.64, 0.3957, 135.19, 31002, 30.38, 0.65, 19.75, 154.93, 119.57, 0.20, 119.77, 35.16, 102.89],
    [6, 304260, 484.38, 0.3957, 191.67, 44832, 43.94, 0.65, 28.56, 220.23, 169.53, 0.20, 169.73, 50.49, 153.39],
]

for row in b2c_data:
    ws_b2c.append(row)

style_data(ws_b2c, 2, 7, 1, len(headers_b2c))

for r in range(2, 8):
    for c in range(2, len(headers_b2c) + 1):
        ws_b2c.cell(row=r, column=c).number_format = '#,##0.00'
    ws_b2c.cell(row=r, column=2).number_format = '#,##0'
    ws_b2c.cell(row=r, column=6).number_format = '#,##0'

widths_b2c = [6, 12, 18, 12, 16, 10, 16, 12, 16, 18, 16, 12, 16, 14, 20]
for i, w in enumerate(widths_b2c, start=1):
    ws_b2c.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

# ============================================================
# 5. Summary sheet
# ============================================================
ws_sum = wb.create_sheet("Summary")
ws_sum.title = "Summary"

headers_sum = [
    "Year", "B2B Profit (HK$m)", "B2C Profit (HK$m)",
    "Total Profit (HK$m)", "B2B Cumulative (HK$m)",
    "B2C Cumulative (HK$m)", "B2C / B2B Ratio"
]
ws_sum.append(headers_sum)
style_header(ws_sum, 1, len(headers_sum))

sum_data = [
    [1, 0.93, 2.51, 3.45, 0.93, 2.51, 2.69],
    [2, 2.75, 4.37, 7.12, 3.69, 6.88, 1.87],
    [3, 18.48, 32.27, 50.75, 22.17, 39.16, 1.77],
    [4, 27.08, 28.57, 55.65, 49.25, 67.73, 1.38],
    [5, 37.33, 35.16, 72.49, 86.58, 102.89, 1.19],
    [6, 52.92, 50.49, 103.41, 139.50, 153.39, 1.10],
]

for row in sum_data:
    ws_sum.append(row)

style_data(ws_sum, 2, 7, 1, len(headers_sum))

for r in range(2, 8):
    for c in range(2, len(headers_sum) + 1):
        ws_sum.cell(row=r, column=c).number_format = '#,##0.00'
    ws_sum.cell(row=r, column=7).number_format = '0.00'

widths_sum = [6, 18, 18, 18, 20, 20, 16]
for i, w in enumerate(widths_sum, start=1):
    ws_sum.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

# ============================================================
# 6. Chart sheet
# ============================================================
ws_chart = wb.create_sheet("Cumulative Chart")
ws_chart.title = "Cumulative Chart"

# Write data for chart
ws_chart.append(["Year", "B2B Cumulative (HK$m)", "B2C Cumulative (HK$m)"])
for row in sum_data:
    ws_chart.append([row[0], row[4], row[5]])

style_header(ws_chart, 1, 3)
style_data(ws_chart, 2, 7, 1, 3)

for r in range(2, 8):
    for c in range(2, 4):
        ws_chart.cell(row=r, column=c).number_format = '#,##0.00'

# Create line chart
chart = LineChart()
chart.title = "Cumulative Profit: B2B vs B2C"
chart.style = 12
chart.y_axis.title = "HK$ million"
chart.x_axis.title = "Year"

data = Reference(ws_chart, min_col=2, min_row=1, max_col=3, max_row=7)
cats = Reference(ws_chart, min_col=1, min_row=2, max_row=7)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

ws_chart.add_chart(chart, "E2")

# ============================================================
# 7. Save workbook
# ============================================================
output_file = "KungFu_AI_Football_3pct_SAM_35pct_Marketing.xlsx"
wb.save(output_file)
print(f"Workbook saved as: {output_file}")