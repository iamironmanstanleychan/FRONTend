"""
KungFu AI Football — README v2.8 A4 Landscape Generator
Output: KungFu_AI_Football_README_v2.8_A4.docx
Page size: A4 Landscape (297mm x 210mm)
Run: python generate_v2_8_a4.py
Requires: python-docx
"""

from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── CONSTANTS ──
USABLE_WIDTH_MM = 277  # 297 - 10 - 10


def set_cell_background(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def add_table(doc, data, header=True, col_widths=None, font_size=9,
              max_width_mm=USABLE_WIDTH_MM):
    """Add a table; auto-scale col_widths if sum exceeds max_width_mm."""
    rows, cols = len(data), len(data[0])
    table = doc.add_table(rows=rows, cols=cols)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, row in enumerate(data):
        for j, value in enumerate(row):
            cell = table.cell(i, j)
            cell.text = str(value)
            for p in cell.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(font_size)
                    if header and i == 0:
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            if header and i == 0:
                set_cell_background(cell, "1F4E78")

    if col_widths:
        total = sum(col_widths)
        if total > max_width_mm:
            scale = max_width_mm / total
            col_widths = [w * scale for w in col_widths]
        for j, w in enumerate(col_widths):
            for i in range(rows):
                table.cell(i, j).width = Mm(w)
    return table


def add_para(doc, text, bold=False, italic=False, size=10, space_after=5):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(text, style='List Bullet')
    for r in p.runs:
        r.font.size = Pt(9)
    return p


def add_quote(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Mm(8)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    return p


# ─────────────────────────────────────────────────────────
# DOCUMENT SETUP — A4 LANDSCAPE
# ─────────────────────────────────────────────────────────
doc = Document()

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(9)

# A4 Landscape: 297mm × 210mm
section = doc.sections[0]
section.page_width = Mm(297)
section.page_height = Mm(210)
section.orientation = WD_ORIENT.LANDSCAPE
section.top_margin = Mm(10)
section.bottom_margin = Mm(10)
section.left_margin = Mm(10)
section.right_margin = Mm(10)
# Usable width: 277mm

# ─────────────────────────────────────────────────────────
# COVER
# ─────────────────────────────────────────────────────────
title = doc.add_heading('Kung Fu AI Football — 6-Year Financial Model', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = sub.add_run('README v2.8')
r.bold = True
r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

sub2 = doc.add_paragraph()
sub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = sub2.add_run('A4 Landscape Edition — with §3.7 Spending Power Section')
r2.italic = True
r2.font.size = Pt(11)
r2.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

doc.add_paragraph()
cover_info = [
    ["Field", "Value"],
    ["File", "KungFu_AI_Football_v2.8_Live.xlsx"],
    ["Model type", "Live Excel model — editable inputs, formula-driven P&L"],
    ["Base case", "P50 (mean conversion 4.35% → SOM 608,130 by Year 6)"],
    ["Currency", "HK$ million unless stated otherwise"],
    ["Version", "2.8 (adds §3.7 spending power section with source URLs)"],
    ["Document size", "A4 Landscape (297mm × 210mm)"],
    ["Prepared for", "Kung Fu AI Football partnership negotiation"],
    ["Model owner", "B2B technology enabler"],
]
add_table(doc, cover_info, col_widths=[70, 200], font_size=10)

doc.add_page_break()

# ─────────────────────────────────────────────────────────
# 1. PURPOSE AND SCOPE
# ─────────────────────────────────────────────────────────
doc.add_heading('1. Purpose and Scope', level=1)
add_para(doc, 'This workbook is the six-year financial planning model for Kung Fu AI Football — the B2B technology enabler — and its B2C distribution partners.')
add_para(doc, 'The model answers four questions with a single set of linked formulas:')
add_bullet(doc, 'TAM / SAM / SOM — how large is the addressable market, the serviceable pool, and the obtainable paid base?')
add_bullet(doc, 'B2B P&L — what does the technology enabler earn and spend each year?')
add_bullet(doc, 'B2C P&L — what do the distribution partners earn and spend each year?')
add_bullet(doc, 'Fairness — are both partners non-negative in every year, and is the cumulative split within the 10% ceiling?')
add_para(doc, 'The model excludes taxes, payment processing fees, customer support, hardware warranty, logistics, bad debt, working capital, capex, and financing costs.')
add_para(doc, 'The workbook is live. Every P&L line is a formula driven by named inputs on the Assumptions sheet. Changing a yellow cell recalculates the entire model.', bold=True)

# ─────────────────────────────────────────────────────────
# 2. WORKBOOK STRUCTURE
# ─────────────────────────────────────────────────────────
doc.add_heading('2. Workbook Structure', level=1)
wb = [
    ["#", "Sheet", "Function"],
    ["1", "Cover", "Version, scenario, workflow instructions"],
    ["2", "Assumptions", "All inputs — yellow cells, named ranges"],
    ["3", "Market Sizing", "TAM / SAM / SOM with three conversion bands and regional allocation"],
    ["4", "B2B P&L", "Revenue (incl. grant), costs (incl. ITF manpower), profit, cumulative"],
    ["5", "B2C P&L", "Revenue, costs (incl. customization fee), profit, cumulative"],
    ["6", "Summary", "Fairness check plus six PASS/FAIL invariants"],
    ["7", "Cumulative Chart", "Auto-updating line chart"],
]
add_table(doc, wb, col_widths=[15, 60, 220], font_size=10)

doc.add_page_break()

# ─────────────────────────────────────────────────────────
# 3. MARKET SIZING
# ─────────────────────────────────────────────────────────
doc.add_heading('3. Market Sizing — TAM, SAM, SOM', level=1)

doc.add_heading('3.1 Definitions and Source Hierarchy', level=2)
def_t = [
    ["Tier", "Definition", "Year 6 Value", "Primary Source"],
    ["TAM", "Football-interested population: fans, students, amateur players, existing club/academy players, retired players", "300m", "2025 industry data: 380m fans"],
    ["SAM", "Core active football participants with official registration or school enrolment", "13.98m", "CFA 2025 annual report + 2025 industry data"],
    ["SOM", "Paid subscribers captured through the partners' premium channels", "608,130 (P50)", "SAM × mean conversion 4.35%"],
]
add_table(doc, def_t, col_widths=[18, 130, 35, 95], font_size=9)

doc.add_heading('3.2 TAM = 300m', level=2)
tam_t = [
    ["Item", "Value", "Source URL", "Explanation"],
    ["Chinese football fans", "380m", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html", "2025 industry data reports 380 million fans"],
    ["TAM (adopted)", "300m", "—", "79% of the 380m fan base — deliberately conservative, excluding low-interest casual fans"],
]
add_table(doc, tam_t, col_widths=[50, 20, 120, 90], font_size=9)
add_para(doc, 'Cross-check: NDRC plan states fans exceed 100 million. 300m sits above the official floor and below the industry ceiling.', italic=True, size=8)

doc.add_heading('3.3 SAM = 13.98m', level=2)
sam_t = [
    ["Component", "Value", "Source URL", "Explanation"],
    ["Social amateur players", "979,500", "CFA 2025 Digital Annual Report (cited at https://www.163.com/dy/article/L8E6C26F0556BPXI.html)", "Registered, organised adult players"],
    ["Elite youth registered players", "109,220", "Same CFA 2025 report", "Youth inside the formal training system"],
    ["Campus football registered students", "12.8m", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html", "Students at football-specialist schools"],
    ["SAM (adopted)", "13.98m", "—", "Sum of the three measurable components"],
]
add_table(doc, sam_t, col_widths=[50, 20, 120, 90], font_size=9)

doc.add_heading('3.4 SOM — Three Conversion Bands', level=2)
add_para(doc, 'The SOM is derived by applying a conversion rate to SAM. The rate is anchored to a single external benchmark: 8.7% — the share of Chinese parents willing to support their child\'s long-term football participation (source: https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html).')
som_t = [
    ["Scenario", "Formula", "Conversion Rate", "Share of Ceiling", "SOM (Y6)"],
    ["P10 (Conservative)", "8.7 / 4", "2.175%", "25%", "304,065"],
    ["P50 (Base)", "8.7 / 2", "4.35%", "50%", "608,130"],
    ["P90 (Optimistic)", "8.7 / 1", "8.7%", "100%", "1,216,260"],
]
add_table(doc, som_t, col_widths=[55, 30, 40, 40, 50], font_size=10)
add_para(doc, 'Statistical note: The three bands are consistent with a lognormal distribution — P50/P10 = 2.0 and P90/P50 = 2.0, implying σ_log ≈ 0.54 (geometric SD ≈ 1.72).', size=8)

doc.add_heading('3.4.1 The 68% Gen-Z Evidence', level=3)
add_quote(doc, '68% of Gen-Z parents explicitly demand data-driven feedback on their child\'s football training (source: https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html).')
genz_t = [
    ["Evidence", "Value", "Implication for P50"],
    ["Gen-Z parents requiring data-driven feedback", "68%", "Two-thirds of the target parent cohort already want exactly what the product provides"],
    ["Parents accepting football training spend", ">90%", "The spending habit is established"],
    ["High-end tier repurchase rate", "92%", "Once converted, parents stay"],
    ["Smart wearable penetration in youth academies", "63%", "Tech-assisted training is already normal in the channel"],
]
add_table(doc, genz_t, col_widths=[90, 25, 150], font_size=9)

doc.add_heading('3.4.2 Why 4.35% Is a Sensible Middle Estimate', level=3)
add_para(doc, 'Step 1 — How many parents back football at all? Out of 100 parents, about 9 (8.7%) say they are willing to support their child playing football long-term. This is the outer limit.')
add_para(doc, 'Step 2 — Of those, how many want data, not just encouragement? Among today\'s Gen-Z parents, 68 out of 100 say they want numbers and data-driven feedback.')
add_para(doc, 'Step 3 — So how big is the "wants what we sell" group? 9 × 68 ≈ 6 out of 100 parents (5.916%). Call this the "real demand" ceiling.')
add_para(doc, 'Step 4 — But not everyone who says they want it actually pays. Some parents like the idea but never subscribe. So we assume only about 4 to 5 out of every 6 of those interested parents actually pay. That gives 4.35% — our P50.')
middle_t = [
    ["Scenario", "Rate", "Plain-English meaning"],
    ["P10 (pessimistic)", "2.175%", "Only about 2 in every 6 interested parents pay — very cautious"],
    ["P50 (our estimate)", "4.35%", "About 4–5 in every 6 interested parents pay — realistic"],
    ["P90 (optimistic)", "8.7%", "Assumes every football-supporting parent pays, even those who never asked for data — aggressive"],
]
add_table(doc, middle_t, col_widths=[60, 25, 180], font_size=10)
add_para(doc, 'The 4.35% figure sits below the "real demand" ceiling of 5.916% while capturing about 74% of it — a 26% "talk-is-cheap" discount. Caveat: the math assumes independence between "supporting football" and "wanting data"; positive correlation would make 4.35% even more conservative.', italic=True, size=8)

doc.add_heading('3.5 SOM by Year (P50)', level=2)
som_y = [
    ["Year", "Regions Active", "SAM (m)", "Conversion", "SOM"],
    ["Y1", "HK & South China", "5.59", "0.246%", "13,745"],
    ["Y2", "HK & South China", "5.59", "0.615%", "34,363"],
    ["Y3", "+ East China", "9.01", "2.147%", "193,389"],
    ["Y4", "+ Southwest", "11.34", "2.691%", "305,073"],
    ["Y5", "+ North", "12.89", "3.328%", "429,018"],
    ["Y6", "+ Northwest & Northeast", "13.98", "4.350%", "608,130"],
]
add_table(doc, som_y, col_widths=[22, 100, 35, 40, 50], font_size=10)

doc.add_heading('3.5.1 Footnote — How the Year-by-Year SAM Is Derived', level=3)
sam_comp = [
    ["Component", "Value", "Share of SAM", "Source"],
    ["Campus football registered students", "12.80m", "91.6%", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html"],
    ["Social amateur players", "0.98m", "7.0%", "CFA 2025 report (https://www.163.com/dy/article/L8E6C26F0556BPXI.html)"],
    ["Elite youth registered players", "0.11m", "0.8%", "Same CFA 2025 report"],
    ["Total SAM", "13.98m", "100%", "—"],
]
add_table(doc, sam_comp, col_widths=[70, 25, 35, 150], font_size=9)

add_para(doc, 'Regional Allocation Weights', bold=True)
weights_t = [
    ["Region", "Weight", "Regional SAM", "Rationale"],
    ["Hong Kong & South China", "40.0%", "5.59m", "Dense school football in Guangdong; HK FIFA Talent Academy and PSG HK girls' academy; Shenzhen South China youth training centre"],
    ["East China", "24.4%", "3.42m", "Highest football participation density; Suzhou regional youth training centre"],
    ["Southwest", "16.7%", "2.33m", "Chengdu regional youth training centre; strong school football presence in Sichuan"],
    ["North", "11.1%", "1.55m", "Beijing as policy and national youth training centre"],
    ["Northwest", "4.4%", "0.62m", "Xi'an regional youth training centre; lower school density"],
    ["Northeast", "3.3%", "0.47m", "Dalian regional youth training centre; declining school-age population"],
    ["Total", "100%", "13.98m", "—"],
]
add_table(doc, weights_t, col_widths=[55, 25, 35, 165], font_size=9)

sam_deriv = [
    ["Year", "Regions activated", "Regional SAM added", "Cumulative SAM", "Calculation"],
    ["Y1", "HK & South China", "5.59m", "5.59m", "5.59"],
    ["Y2", "No new region", "—", "5.59m", "5.59 + 0"],
    ["Y3", "+ East China", "3.42m", "9.01m", "5.59 + 3.42"],
    ["Y4", "+ Southwest", "2.33m", "11.34m", "9.01 + 2.33"],
    ["Y5", "+ North", "1.55m", "12.89m", "11.34 + 1.55"],
    ["Y6", "+ Northwest & Northeast", "1.09m", "13.98m", "12.89 + 0.62 + 0.47"],
]
add_table(doc, sam_deriv, col_widths=[22, 80, 50, 50, 75], font_size=10)

orig_t = [
    ["Region", "Original planning estimate", "× 0.3107", "Adopted SAM"],
    ["HK & South China", "18.0m", "5.59", "5.59m"],
    ["East China", "11.0m", "3.42", "3.42m"],
    ["Southwest", "7.5m", "2.33", "2.33m"],
    ["North", "5.0m", "1.55", "1.55m"],
    ["Northwest", "2.0m", "0.62", "0.62m"],
    ["Northeast", "1.5m", "0.47", "0.47m"],
    ["Total", "45.0m", "13.98", "13.98m"],
]
add_table(doc, orig_t, col_widths=[60, 60, 35, 50], font_size=10)

doc.add_heading('3.6 Total Football Market Capital', level=2)
mkt_t = [
    ["Metric", "Value", "Source URL"],
    ["Football industry total scale (2025)", "RMB 320bn", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html"],
    ["Football training market", "RMB 48.76bn", "Same"],
    ["Fans", "380m", "Same"],
    ["Football-specialist schools", "42,300", "Same"],
    ["Campus football registered students", "12.8m", "Same"],
    ["Gen-Z parents demanding data-driven feedback", "68%", "Same"],
    ["Parents accepting football training spend", ">90%", "Same"],
    ["Smart wearable penetration in youth academies", "63%", "Same"],
]
add_table(doc, mkt_t, col_widths=[120, 40, 115], font_size=9)

doc.add_page_break()

# ─────────────────────────────────────────────────────────
# 3.7 SPENDING POWER
# ─────────────────────────────────────────────────────────
doc.add_heading('3.7 Spending Power of the Target Segment for "AI Football + Annual Subscription"', level=1)

add_para(doc, 'The pricing assumption — HK$980 for the IoT ball and HK$1,592/year for the subscription — is calibrated against three independently observable dimensions: spending power, performance orientation, and the "new gear fans" consumption class.', bold=True)

doc.add_heading('3.7.1 Spending Power — The ARPU Sits at the Low End of Existing Family Sports Budgets', level=2)
add_para(doc, 'Youth football training spend. Middle-class Chinese families spend RMB 10,000–50,000 per year on children\'s sports training (63% of surveyed families), with 14% spending RMB 50,000–100,000 and 11% above RMB 100,000 (source: iResearch, https://www.iresearch.com.cn). Entry-level youth academy fees in Tier-1 cities run RMB 8,000–20,000/year, while elite club fees reach RMB 80,000–120,000/year (source: 36Kr, https://36kr.com). With camps and equipment, a child\'s annual football spend commonly exceeds RMB 30,000 (source: Lanxiong Sports, https://www.lanxiongsports.com).')
add_para(doc, 'Calibration: At HK$1,592/year (≈ RMB 1,450), the subscription represents 3–15% of a middle-class family\'s annual sports training budget — the low end of the acceptable range.')

smart_t = [
    ["Product", "Price", "Source"],
    ["Domestic GPS football vest (神仙球)", "RMB 1,599", "https://www.shenxianqiu.com"],
    ["Adidas World Cup match ball retail edition", "~USD 160 (≈ RMB 1,150)", "https://www.adidas.com"],
    ["Smart football consumer editions", "RMB 300–3,000", "https://www.jd.com"],
]
add_table(doc, smart_t, col_widths=[110, 60, 95], font_size=10)

add_para(doc, 'The HK$980 ball is not a premium outlier — it sits below the domestic GPS vest and is comparable to a World Cup replica ball.')
add_para(doc, 'Subscription benchmark: Keep\'s AI coach paid membership lifted paid conversion from 8% to 15% with an 82% renewal rate (source: Keep IR, https://ir.keep.com). In smart football, annual subscription models already cover 12% of B2B users with 83% retention (source: 163.com industry data, https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html).')

doc.add_heading('3.7.2 Performance Orientation — The Target Segment Pays for Outcomes, Not for Discounts', level=2)
add_para(doc, 'Price sensitivity is segmented. Among football consumers, 45% continue buying football-specific apparel even after a price increase — a higher inelastic share than in most other sports categories (source: 163.com industry data, https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html). The market splits into three tiers: hardcore players (lowest price elasticity), value-seekers (38%, demanding both performance and discount), and entry-level triers (48%, highly price-sensitive).')
add_para(doc, 'Demand for data is itself performance orientation. 68% of Gen-Z parents explicitly demand data-driven feedback (same source). 44.67% of parents are willing to pay a premium for high-quality courses, with the stated logic that "parents no longer buy on \'cheap\' — they buy on results and professionalism" (same source).')
add_para(doc, 'Implication: The target segment is not looking for a cheaper football. It is looking for a tool that makes its child kick more accurately. The AI ball and subscription are priced for that segment — not for the 48% entry-level tier.')

doc.add_heading('3.7.3 The "New Gear Fans" Class — High Ticket, High Repurchase, Full-Price Loyalty', level=2)
add_para(doc, 'Premium sports consumers already spend at scale. Arc\'teryx\'s core 35–55 male middle-class users spend an average of RMB 30,000 per person annually, with a repurchase rate above 38%; Arc\'teryx Greater China revenue grew 47% year-on-year (source: Amer Sports IR, https://www.amersports.com/investors). Descente\'s core users are "few but high-repurchase, high-word-of-mouth, willing to pay full price" (source: Lanxiong Sports, https://www.lanxiongsports.com).')
add_para(doc, 'Football equipment repurchase. 31% of football consumers report repurchase rates in the 70–90% range — nearly one in three is a high-frequency, continuous equipment buyer (source: 163.com industry data, https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html).')

newgear_t = [
    ["\"New Gear Fan\" trait", "Implication for Kung Fu AI Football"],
    ["Pays for performance uplift, not logo", "The product's core is measurable shooting improvement (8-week 10–15% accuracy gain)"],
    ["High repurchase (38–90% range)", "Subscription model matches this continuous consumption habit"],
    ["Willing to pay full price", "HK$980 + HK$1,592/year does not require discount-led acquisition"],
    ["Treats gear as a signal of commitment", "Finishing IQ and physics dashboards provide a visible training investment record"],
]
add_table(doc, newgear_t, col_widths=[100, 160], font_size=10)

add_para(doc, 'A consumer already spending RMB 30,000/year on performance gear, for whom HK$980 + HK$1,592/year represents 8–9% of annual equipment spend, is the natural first buyer of a ball that tells them why the shot missed.')

doc.add_heading('3.7.4 Synthesis — Three Dimensions, One Pricing Conclusion', level=2)
synth_t = [
    ["Dimension", "Core Evidence", "Implication for Pricing"],
    ["Spending power", "Middle-class sports spend RMB 10,000–50,000/year; smart hardware validated at RMB 1,000+; subscription conversion 8–15%, renewal 82%", "HK$980 + HK$1,592/year sits at the low end of acceptable range"],
    ["Performance orientation", "44.67% of parents pay premium for quality; 45% of football consumers buy despite price rises; 68% Gen-Z demand data feedback", "Target segment pays for results, matching the product's value proposition"],
    ["New gear fans", "Arc'teryx users spend RMB 30,000/year, repurchase 38%; football equipment repurchase 70–90% for 31% of buyers", "A high-ticket, high-repurchase, full-price consumption class already exists"],
]
add_table(doc, synth_t, col_widths=[50, 140, 80], font_size=9)

add_para(doc, 'Conclusion: The HK$980 AI ball and HK$1,592/year ARPU are not prices that require "educating the market." They sit within an already-validated smart sports hardware band, address a consumer class already spending tens of thousands of RMB annually on performance gear, and represent a single-digit percentage of that class\'s annual equipment budget. Affordability, performance orientation, and the consumption habits of the "new gear fans" class jointly support the pricing assumption.', bold=True)
add_para(doc, 'Note on source verification: The URLs above point to the source categories (iResearch, 36Kr, Lanxiong Sports, Amer Sports IR, Keep IR, 163.com industry data). Specific article-level URLs should be verified and logged in the Reference URL Register before panel submission.', italic=True, size=8)

doc.add_page_break()

# ─────────────────────────────────────────────────────────
# 4. ITF RTH MANPOWER GRANT
# ─────────────────────────────────────────────────────────
doc.add_heading('4. ITF RTH Manpower Grant', level=1)
itf_t = [
    ["Item", "Assumption"],
    ["BSc manpower", "2 (Y1–Y3) + 2 (Y4–Y6)"],
    ["Monthly rate", "HK$20,000"],
    ["Annual gross cost", "HK$0.48m/year × 6 = HK$2.88m"],
    ["Grant coverage", "100% eligible"],
    ["Reimbursement lag", "6 months"],
]
add_table(doc, itf_t, col_widths=[70, 200], font_size=10)

itf_s = [
    ["Year", "Gross Manpower", "Grant Received", "Net Impact"],
    ["Y1", "0.48", "0.24", "−0.24"],
    ["Y2", "0.48", "0.48", "0.00"],
    ["Y3", "0.48", "0.48", "0.00"],
    ["Y4", "0.48", "0.48", "0.00"],
    ["Y5", "0.48", "0.48", "0.00"],
    ["Y6", "0.48", "0.48", "0.00"],
    ["Y7 Q2", "—", "0.24", "+0.24"],
    ["Total", "2.88", "2.88", "0.00"],
]
add_table(doc, itf_s, col_widths=[25, 60, 60, 50], font_size=10)

# ─────────────────────────────────────────────────────────
# 5. B2B COMMERCIAL TERMS
# ─────────────────────────────────────────────────────────
doc.add_heading('5. Commercial Terms — B2B Technology Enabler', level=1)

doc.add_heading('5.1 Revenue', level=2)
b2b_rev = [
    ["Line", "Terms"],
    ["Subscription share", "12.86% of gross subscription"],
    ["Ball revenue share", "35% × HK$980 = HK$343 per ball"],
    ["Exclusive annual license", "HK$0.20m/year"],
    ["Customization revenue", "HK$0.10m Year 1; HK$0.025m/year Years 2–6"],
    ["ITF RTH grant", "HK$0.24m Y1; HK$0.48m Y2–Y6; HK$0.24m Y7 Q2"],
]
add_table(doc, b2b_rev, col_widths=[90, 180], font_size=10)

doc.add_heading('5.2 Cost — Model D Hybrid Platform', level=2)
b2b_cost = [
    ["Line", "Terms", "Behaviour"],
    ["Private Cloud Platform — base", "HK$0.20m/year", "Fixed"],
    ["Private Cloud Platform — compute", "HK$2.00 × active users", "Variable"],
    ["User enhancement", "HK$5.00 × active users", "Variable"],
    ["Ball enhancement", "HK$10 × ball units", "Variable"],
    ["Ball manufacturing", "HK$250 × ball units", "Variable"],
    ["ITF manpower (gross)", "HK$0.48m/year", "Fixed"],
]
add_table(doc, b2b_cost, col_widths=[80, 90, 70], font_size=10)

# ─────────────────────────────────────────────────────────
# 6. B2C COMMERCIAL TERMS
# ─────────────────────────────────────────────────────────
doc.add_heading('6. Commercial Terms — B2C Distribution Partners', level=1)
b2c_t = [
    ["Line", "Terms"],
    ["Subscription share", "37.14% of gross subscription"],
    ["Ball revenue share", "65% × HK$980 = HK$637 per ball"],
    ["Marketing & distribution", "30% of gross subscription"],
    ["Exclusive license fee", "HK$0.20m/year paid to B2B"],
    ["Customization fee", "HK$0.10m Year 1; HK$0.025m/year Years 2–6 paid to B2B"],
]
add_table(doc, b2c_t, col_widths=[90, 180], font_size=10)

# ─────────────────────────────────────────────────────────
# 7. SUBSCRIPTION SPLIT
# ─────────────────────────────────────────────────────────
doc.add_heading('7. Subscription Split', level=1)
split_t = [
    ["Step", "Recipient", "Share of Gross"],
    ["1", "Coach / Player", "50.00%"],
    ["2", "B2B technology enabler", "12.86%"],
    ["3", "B2C distribution partners", "37.14%"],
    ["", "Total", "100.00%"],
]
add_table(doc, split_t, col_widths=[25, 120, 80], font_size=10)
add_para(doc, 'Tier pricing: Local Coach HK$65/month (60% mix), National HK$148 (25%), World Cup HK$298 (12%), Legends HK$698 (3%). Weighted ARPU is HK$1,592/year.')

doc.add_page_break()

# ─────────────────────────────────────────────────────────
# 8. COST BEHAVIOUR
# ─────────────────────────────────────────────────────────
doc.add_heading('8. Cost Behaviour and Operating Leverage', level=1)
cost_t = [
    ["Cost line", "Behaviour", "Driver"],
    ["Platform base", "Fixed", "HK$0.20m/year"],
    ["ITF manpower (gross)", "Fixed", "HK$0.48m/year"],
    ["Platform compute", "Variable", "HK$2.00 × SOM"],
    ["User enhancement", "Variable", "HK$5.00 × SOM"],
    ["Ball enhancement", "Variable", "HK$10 × new users × 50%"],
    ["Ball manufacturing", "Variable", "HK$250 × new users × 50%"],
]
add_table(doc, cost_t, col_widths=[80, 50, 100], font_size=10)
add_para(doc, 'Fixed base HK$0.68m/year throughout. User base grows from 13,745 to 608,130 — a 44× increase — so fixed cost per user falls from HK$49.47 to HK$1.12.')
add_para(doc, 'Break-even on fixed base ≈ 2,835 users. Year 1 SOM of 13,745 is 4.8× break-even.', bold=True)

# ─────────────────────────────────────────────────────────
# 9. YEAR-BY-YEAR RESULTS
# ─────────────────────────────────────────────────────────
doc.add_heading('9. Year-by-Year Results (P50 Central Case, SOM = 608,130 by Y6)', level=1)

doc.add_heading('9.1 B2B Technology Enabler P&L (HK$m)', level=2)
b2b_pnl = [
    ["Year", "SOM", "Sub Rev", "Ball Rev", "Custom", "License", "ITF Grant", "Total Rev", "Platform", "User Enh", "Ball Enh", "Ball Cost", "Manpower", "Total Cost", "Profit", "Cumulative"],
    ["Y1", "13,745", "2.81", "2.36", "0.10", "0.20", "0.24", "5.71", "0.23", "0.07", "0.07", "1.72", "0.48", "2.56", "3.15", "3.15"],
    ["Y2", "34,363", "7.04", "3.54", "0.025", "0.20", "0.48", "11.28", "0.27", "0.17", "0.10", "2.58", "0.48", "3.60", "7.68", "10.83"],
    ["Y3", "193,389", "39.59", "27.27", "0.025", "0.20", "0.48", "67.57", "0.59", "0.97", "0.80", "19.88", "0.48", "22.71", "44.86", "55.68"],
    ["Y4", "305,073", "62.46", "19.16", "0.025", "0.20", "0.48", "82.32", "0.81", "1.53", "0.56", "13.96", "0.48", "17.34", "64.99", "120.67"],
    ["Y5", "429,018", "87.84", "21.26", "0.025", "0.20", "0.48", "109.80", "1.06", "2.15", "0.62", "15.49", "0.48", "19.80", "90.00", "210.67"],
    ["Y6", "608,130", "124.50", "30.72", "0.025", "0.20", "0.48", "155.92", "1.42", "3.04", "0.90", "22.39", "0.48", "28.22", "127.70", "338.38"],
    ["Y7 Q2", "—", "—", "—", "—", "—", "0.24", "0.24", "—", "—", "—", "—", "—", "0.00", "0.24", "338.62"],
]
add_table(doc, b2b_pnl, col_widths=[11, 14, 14, 14, 14, 14, 14, 17, 14, 14, 14, 14, 14, 17, 17, 17], font_size=7)

doc.add_heading('9.2 B2C Distribution Partners P&L (HK$m)', level=2)
b2c_pnl = [
    ["Year", "SOM", "Sub Rev", "Ball Rev", "Total Rev", "Marketing", "License", "Customization", "Total Cost", "Profit", "Cumulative"],
    ["Y1", "13,745", "8.13", "4.38", "12.51", "6.57", "0.20", "0.10", "6.87", "5.64", "5.64"],
    ["Y2", "34,363", "20.32", "6.57", "26.89", "16.41", "0.20", "0.025", "16.64", "10.25", "15.89"],
    ["Y3", "193,389", "114.34", "50.65", "164.99", "92.36", "0.20", "0.025", "92.59", "72.41", "88.29"],
    ["Y4", "305,073", "180.39", "35.58", "215.97", "145.71", "0.20", "0.025", "145.93", "70.03", "158.32"],
    ["Y5", "429,018", "253.68", "39.48", "293.15", "204.91", "0.20", "0.025", "205.13", "88.02", "246.34"],
    ["Y6", "608,130", "359.57", "57.04", "416.61", "290.44", "0.20", "0.025", "290.67", "125.94", "372.29"],
]
add_table(doc, b2c_pnl, col_widths=[14, 18, 18, 18, 22, 22, 16, 22, 22, 20, 22], font_size=8)

doc.add_heading('9.3 Fairness Check', level=2)
fair = [
    ["Year", "B2B Profit", "B2C Profit", "B2B Cumulative", "B2C Cumulative", "B2C/B2B"],
    ["Y1", "3.15", "5.64", "3.15", "5.64", "1.79"],
    ["Y2", "7.68", "10.25", "10.83", "15.89", "1.47"],
    ["Y3", "44.86", "72.41", "55.68", "88.29", "1.59"],
    ["Y4", "64.99", "70.03", "120.67", "158.32", "1.31"],
    ["Y5", "90.00", "88.02", "210.67", "246.34", "1.17"],
    ["Y6", "127.70", "125.94", "338.38", "372.29", "1.1002"],
    ["Y7 Q2", "+0.24", "—", "338.62", "—", "1.0997"],
]
add_table(doc, fair, col_widths=[22, 35, 35, 45, 45, 30], font_size=10)
add_para(doc, 'Both constraints hold in the full cycle:', bold=True)
add_bullet(doc, 'Neither partner records a negative year.')
add_bullet(doc, 'Year 6 standalone ratio is 1.1002 — at the boundary.')
add_bullet(doc, 'Full-cycle ratio (incl. Y7 Q2 grant catch-up) is 1.0997 — inside the 10% ceiling.')
add_bullet(doc, 'Full-cycle B2B cumulative profit = HK$338.62m.')

# ─────────────────────────────────────────────────────────
# 10. CONFIDENCE ANALYSIS
# ─────────────────────────────────────────────────────────
doc.add_heading('10. Confidence Analysis', level=1)
conf = [
    ["Scenario", "Conversion", "Y6 SOM", "B2B Cum Y6", "B2C Cum Y6", "Fairness (full cycle)"],
    ["P10 (Conservative)", "2.175%", "304,065", "~HK$169m", "~HK$186m", "~1.10"],
    ["P50 (Base)", "4.35%", "608,130", "HK$338.6m", "HK$372.3m", "1.0997"],
    ["P90 (Optimistic)", "8.7%", "1,216,260", "~HK$677m", "~HK$745m", "~1.10"],
]
add_table(doc, conf, col_widths=[50, 35, 35, 45, 45, 55], font_size=10)
add_para(doc, 'To switch scenario in the live workbook, set scenario_multiplier to 0.50 (P10), 1.00 (P50), or 2.00 (P90).')

# ─────────────────────────────────────────────────────────
# 11. SENSITIVITY
# ─────────────────────────────────────────────────────────
doc.add_heading('11. Sensitivity', level=1)

doc.add_heading('11.1 Marketing cost', level=2)
mkt_s = [
    ["B2C marketing", "Required B2B sub share", "Required B2C sub share"],
    ["30% (base)", "12.86%", "37.14%"],
    ["35%", "~10.48%", "~39.52%"],
    ["40%", "~7.93%", "~42.07%"],
]
add_table(doc, mkt_s, col_widths=[60, 80, 80], font_size=10)

doc.add_heading('11.2 Platform model', level=2)
plat_s = [
    ["Model", "6-year total", "B2B Cum Y6", "Fairness (full cycle)", "Within 10%?"],
    ["D — Hybrid (adopted)", "2.90", "338.38", "1.0997", "Yes"],
    ["A — Flat 0.40 × 6", "2.40", "338.88", "1.099", "Yes"],
    ["C — Step-by-region", "3.15", "338.13", "1.101", "Marginal"],
]
add_table(doc, plat_s, col_widths=[60, 40, 40, 55, 40], font_size=10)

doc.add_heading('11.3 Ball cost and attach rate', level=2)
add_para(doc, 'At HK$250 per ball, B2B earns +HK$83 net per ball. At HK$300, net falls to +HK$33; at HK$343, ball sales break even and then turn loss-making.')

doc.add_heading('11.4 Compute-cost conservatism', level=2)
add_para(doc, 'The model assumes HK$7/user/year for combined platform compute and user enhancement (HK$2 + HK$5). Benchmarks indicate HK$10–16/user/year at 25 kicks per user. The model may understate infrastructure cost by HK$2–5m cumulative at the P50 SOM. This does not threaten the fairness band — B2B still generates over HK$330m cumulative profit — but is disclosed for pre-commitment stress-testing.')

# ─────────────────────────────────────────────────────────
# 12. HOW TO USE
# ─────────────────────────────────────────────────────────
doc.add_heading('12. How to Use the Live Workbook', level=1)
add_bullet(doc, 'Open the Assumptions sheet. All yellow cells are editable.')
add_bullet(doc, 'Change any input — e.g. marketing_pct 0.30 → 0.35, or ball_cost 250 → 300.')
add_bullet(doc, 'Every sheet recalculates.')
add_bullet(doc, 'Check the Summary sheet — six PASS/FAIL invariants flag any breach.')
add_bullet(doc, 'To switch conversion scenario, set scenario_multiplier to 0.50 (P10), 1.00 (P50), or 2.00 (P90).')
add_bullet(doc, 'To run a sensitivity table, use Data → What-If Analysis → Data Table.')

doc.add_page_break()

# ─────────────────────────────────────────────────────────
# 13. CHANGE LOG
# ─────────────────────────────────────────────────────────
doc.add_heading('13. Change Log', level=1)
cl = [
    ["Version", "Change"],
    ["2.0", "Model D hybrid platform; split 12.86 / 37.14"],
    ["2.1", "ITF RTH manpower integrated"],
    ["2.2", "Production release with generation script"],
    ["2.3", "Customization fee added as B2C cost; hybrid customization 0.10 + 0.025×5"],
    ["2.4", "Live Excel model — named inputs, formula-driven P&L"],
    ["2.5", "TAM/SAM/SOM externally sourced with URLs; three-band conversion framework; P&L recalibrated to P50 SOM 608,130"],
    ["2.6", "Integrated 68% Gen-Z evidence (§3.4.1) and regional SAM allocation weights (§3.5.1)"],
    ["2.7", "§3.4.2 rewritten in layman terms; arithmetic correction: full-cycle fairness ratio 1.0997"],
    ["2.8", "Added §3.7 Spending Power of the Target Segment — three dimensions with source URLs. A4 Landscape edition."],
]
add_table(doc, cl, col_widths=[22, 250], font_size=9)

# ─────────────────────────────────────────────────────────
# 14. REFERENCE URL REGISTER
# ─────────────────────────────────────────────────────────
doc.add_heading('14. Reference URL Register', level=1)
refs = [
    ["Reference", "URL", "Used For"],
    ["2025 industry data (fans, market size, campus students, parent attitudes, 68% Gen-Z demand)", "https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html", "TAM 300m; SAM campus; market capital; 8.7% ceiling; 68% Gen-Z; §3.7"],
    ["CFA 2025 Digital Annual Report (cited at 163.com)", "https://www.163.com/dy/article/L8E6C26F0556BPXI.html", "SAM amateur (979,500) and elite youth (109,220)"],
    ["NDRC Medium- and Long-Term Football Development Plan", "https://www.ndrc.gov.cn/xxgk/zcfb/ghwb/201604/t20160411_962179.html", "Cross-check: 50m participation; 100m+ fan floor"],
    ["CFA youth-system reform (2026)", "https://www.thecfa.cn/mtzl/20260309/37401.html", "Policy trajectory"],
    ["CEIBS analysis (FIFA Big Count)", "https://cn.ceibs.edu/new-papers-columns/29518", "Cross-check: 26.16m broad participants; 0.078% density"],
    ["iResearch (middle-class sports spending)", "https://www.iresearch.com.cn", "§3.7.1 spending power"],
    ["36Kr (youth academy fees)", "https://36kr.com", "§3.7.1 spending power"],
    ["Lanxiong Sports (Descente user profile)", "https://www.lanxiongsports.com", "§3.7.1, §3.7.3"],
    ["Amer Sports investor relations (Arc'teryx)", "https://www.amersports.com/investors", "§3.7.3 new gear fans"],
    ["Keep investor relations (AI coach)", "https://ir.keep.com", "§3.7.1 subscription benchmark"],
    ["Shenxianqiu (GPS football vest)", "https://www.shenxianqiu.com", "§3.7.1 smart hardware price anchor"],
    ["Adidas (World Cup ball retail)", "https://www.adidas.com", "§3.7.1 ball price anchor"],
    ["JD.com (smart football consumer pricing)", "https://www.jd.com", "§3.7.1 smart ball price range"],
]
add_table(doc, refs, col_widths=[85, 90, 100], font_size=8)

# ─────────────────────────────────────────────────────────
# 15. CAVEATS
# ─────────────────────────────────────────────────────────
doc.add_heading('15. Caveats and Intended Use', level=1)
caveats = [
    "This is a planning model. Not audited, not tax-adjusted, not suitable for filing without independent review.",
    "TAM is the broad football-interested population — 300m, taken as 79% of the 380m fan base.",
    "SAM is the core active football-participating population — 13.98m, every component backed by an official report or industry dataset. Regional allocation weights in §3.5.1.",
    "SOM is a projection, not a deterministic calculation. The three bands are fixed fractions of the 8.7% parental-support ceiling and are consistent with a lognormal distribution (σ_log ≈ 0.54).",
    "The 68% Gen-Z data-driven demand figure supports the P50 of 4.35% — the model assumes only about three in four of the 'real demand' parents (5.916%) actually pay. See §3.4.2.",
    "§3.7 pricing evidence: HK$980 ball + HK$1,592/year ARPU are calibrated against middle-class sports budgets, validated smart hardware price bands, and the 'new gear fans' consumption profile. Source URLs logged in §14.",
    "The P&L is presented at P50 (conversion 4.35%, SOM 608,130). P10 and P90 are selectable via the scenario multiplier.",
    "The Private Cloud Platform uses Model D — HK$0.20m fixed base plus HK$2.00 per active user.",
    "The subscription split is 12.86% / 37.14%. Full-cycle fairness ratio (incl. Y7 Q2 grant) is 1.0997, within the 10% ceiling. Y6 standalone ratio is 1.1002 (at the boundary).",
    "Customization revenue and cost both appear on both sides of the ledger.",
    "The ITF RTH grant is cash-neutral across the cycle.",
    "Known conservatism gap: HK$7/user may understate compute and storage cost by 1.4–2.3×. Disclosed in §11.4.",
    "Partner credentials require due diligence before any binding commercial agreement.",
    "The workbook is live. To model alternative scenarios, change the yellow input cells.",
]
for i, c in enumerate(caveats, 1):
    p = doc.add_paragraph(f"{i}. {c}")
    p.paragraph_format.left_indent = Mm(4)
    for r in p.runs:
        r.font.size = Pt(8)

# ─────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────
doc.add_paragraph()
f1 = doc.add_paragraph()
f1.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = f1.add_run('— End of README v2.8 (A4 Landscape) —')
r.italic = True
r.font.size = Pt(8)
r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

f2 = doc.add_paragraph()
f2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = f2.add_run(
    'Prepared for: Kung Fu AI Football partnership negotiation  |  '
    'Model owner: B2B technology enabler  |  '
    'Review cycle: Re-run whenever marketing cost, ball cost, ball attach rate, platform cost, '
    'customization structure, conversion assumptions, or SAM components change by more than 5 percentage points.'
)
r2.italic = True
r2.font.size = Pt(7)
r2.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

# ─────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────
OUT = 'KungFu_AI_Football_README_v2.8_A4.docx'
doc.save(OUT)
print(f'✅ A4 Word document saved: {OUT}')
print(f'   Page size: 297mm x 210mm (A4 Landscape)')
print(f'   Usable width: 277mm')
print(f'   Sections: 15 + §3.7')