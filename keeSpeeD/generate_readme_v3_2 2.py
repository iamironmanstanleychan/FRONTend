"""
README v3.2 — Professional Panel Edition (concise, <= 8 pages)
Times New Roman · Body 10pt · H1 14pt · H2 12pt · H3 10.5pt
Output: KungFu_AI_Football_README_v3.2.docx
"""

import re
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = 'Times New Roman'
SIZES = {
    'cover_title': 20, 'cover_sub': 11,
    'h1': 14, 'h2': 12, 'h3': 10.5,
    'body': 10, 'table': 8.5, 'footer': 8,
}
HEADER_FILL = "1F4E78"
HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT = RGBColor(0x70, 0x70, 0x70)
ACCENT = RGBColor(0x1F, 0x4E, 0x78)


def set_cell_bg(cell, color_hex):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    tc_pr.append(shd)


def set_run(run, size, bold=False, italic=False, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color


def add_inline(paragraph, text, size, bold=False):
    pattern = re.compile(r'(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)')
    for part in pattern.split(text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            set_run(paragraph.add_run(part[2:-2]), size, bold=True)
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            set_run(paragraph.add_run(part[1:-1]), size, bold=bold, italic=True)
        elif part.startswith('`') and part.endswith('`'):
            r = paragraph.add_run(part[1:-1])
            set_run(r, size, bold=bold)
            r.font.name = 'Courier New'
        else:
            set_run(paragraph.add_run(part), size, bold=bold)


def add_heading(doc, text, level):
    size = {1: SIZES['h1'], 2: SIZES['h2'], 3: SIZES['h3']}.get(level, SIZES['h3'])
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3)
    if level == 1:
        p.paragraph_format.keep_with_next = True
    set_run(p.add_run(text), size, bold=True, color=ACCENT)


def add_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.12
    add_inline(p, text, SIZES['body'])


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(1)
    add_inline(p, text, SIZES['body'])


def add_quote(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Mm(6)
    p.paragraph_format.space_after = Pt(3)
    add_inline(p, text, SIZES['body'])
    for r in p.runs:
        r.italic = True
        r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)


def add_hr(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), 'AAAAAA')
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_table(doc, rows):
    n_rows, n_cols = len(rows), len(rows[0])
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = ''
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            run = p.add_run(str(val))
            if i == 0:
                set_run(run, SIZES['table'], bold=True, color=HEADER_TEXT)
                set_cell_bg(cell, HEADER_FILL)
            else:
                set_run(run, SIZES['table'])


def parse_table(lines):
    rows = []
    for line in lines:
        line = line.strip()
        if not line.startswith('|'):
            continue
        if re.match(r'^\|[\s\-:|]+\|$', line):
            continue
        rows.append([c.strip() for c in line.strip('|').split('|')])
    return rows


def render(doc, md_text):
    lines = md_text.split('\n')
    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        if not line.strip():
            i += 1
            continue
        if re.match(r'^---+$', line.strip()):
            add_hr(doc)
            i += 1
            continue
        m = re.match(r'^(#{1,3})\s+(.+)$', line)
        if m:
            add_heading(doc, m.group(2).strip(), len(m.group(1)))
            i += 1
            continue
        if line.strip().startswith('|'):
            tl = []
            while i < n and lines[i].strip().startswith('|'):
                tl.append(lines[i])
                i += 1
            rows = parse_table(tl)
            if rows:
                add_table(doc, rows)
            continue
        if line.strip().startswith('>'):
            ql = []
            while i < n and lines[i].strip().startswith('>'):
                ql.append(lines[i].strip().lstrip('>').strip())
                i += 1
            add_quote(doc, ' '.join(ql))
            continue
        if re.match(r'^\s*[-*]\s+', line):
            add_bullet(doc, re.sub(r'^\s*[-*]\s+', '', line))
            i += 1
            continue
        if re.match(r'^\s*\d+\.\s+', line):
            text = re.sub(r'^\s*\d+\.\s+', '', line)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Mm(5)
            p.paragraph_format.space_after = Pt(1)
            add_inline(p, text, SIZES['body'])
            i += 1
            continue
        add_body(doc, line.strip())
        i += 1


# ═══════════════════════════════════════════════════════════════
README_MD = r"""
# Kung Fu AI Football — Six-Year Financial Model

**README v3.2 · Panel Edition**

Prepared for: partnership and investment review
Model owner: B2B technology enabler
Base case: P50 scenario, 30% B2C marketing, Year 6 SOM ≈ 604,200

---

## 1. Executive Summary

Kung Fu AI Football is a **B2B technology enabler** supplying an AI-assisted football training platform to four B2C distribution partners (Anta, Li Ning, Amer Sports, Asia Partners IFBD / Svenska Capital). The model forecasts six years of profitability for both sides of the partnership and tests the commercial terms against a **3 × 3 grid** of marketing-cost assumptions and adoption scenarios.

**Headline results at the base case (30% marketing, P50 adoption):**

| Metric | Value |
|---|---:|
| Year 6 paying subscribers (SOM) | ≈ 604,200 |
| Gross subscription revenue, Year 6 | ≈ HK$962 m |
| B2B full-cycle cumulative profit | ≈ HK$336 m |
| B2C cumulative profit, Year 6 | ≈ HK$370 m |
| Fairness ratio (B2C ÷ B2B) | ≈ 1.10 |
| Subscription revenue as % of football training market | ≈ 2% |

**Three findings the panel should note:**

1. **Every input is traceable.** §3 lists all assumptions against specific article-level sources, with an explicit evidence grade — no home-page citations remain.
2. **Marketing cost is the single most sensitive term.** Moving from 30% to 40% shifts roughly HK$120 m of cumulative profit from B2C to B2B *unless* the subscription split is re-calibrated. The model does this automatically.
3. **The three adoption scenarios are not arbitrary labels.** They are fixed fractions of a single externally-anchored behavioural ceiling (8.7%), and they form a statistically coherent lognormal distribution.

---

## 2. Model Architecture

The workbook has three layers. A change in Layer 1 propagates through Layers 2 and 3 without manual intervention.

**Layer 1 — Inputs** (Assumptions sheet, yellow cells)
Market sizing inputs, pricing, cost rates, split calibration table, scenario multipliers, and two master dropdowns: `marketing_pct` (30 / 35 / 40) and `scenario_selector` (1 / 2 / 3).

**Layer 2 — Automatic derivation**
Two lookup mechanisms convert the dropdowns into numbers the P&L can use:

- `split_b2b` and `split_b2c` are computed by **VLOOKUP** into the three-row calibration table. Each marketing level is paired with a pre-solved split that preserves the 1.10 fairness ratio.
- `scenario_multiplier` is computed by **CHOOSE** on the scenario selector, returning 0.50 (P10), 1.00 (P50) or 2.00 (P90).

**Layer 3 — Outputs**
Market Sizing, B2B P&L, B2C P&L, Summary, and a nine-cell Matrix that shows all marketing × scenario combinations side by side. Nine PASS/FAIL checks in the Summary sheet confirm internal consistency after every change.

---

## 3. Assumption Verification Register

Every number that drives the P&L is listed below against its source. Evidence grades are: **Primary** (government, official statistics, corporate filings), **Industry research** (commercial research reports, not independently audited), and **Secondary** (product reviews, user commentary).

### 3.1 Market sizing

| Assumption | Value | Source | Grade |
|---|---:|---|---|
| Chinese football fans | 380 m | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | Primary |
| Football industry total scale (2025) | RMB 320 bn | Same | Primary |
| Football training market (2025) | RMB 48.76 bn | Same | Primary |
| Football-specialist schools | 42,300 | Same | Primary |
| Campus football registered students | 12.8 m | Same | Primary |
| Elite youth registered players | 109,220 | http://news.china.com.cn/2026-01/30/content_118308779.shtml | Primary |
| Social amateur players | 979,544 | Same | Primary |
| Certified coaches | 125,976 | Same | Primary |
| Parents supporting a long-term football path | 8.7% | https://www.163.com/dy/article/L8E6C26F0556BPXI.html | Primary |

**Derived TAM / SAM / SOM.** TAM is set at **300 m** — 79% of the 380 m fan base, deliberately excluding low-interest casual fans. SAM is the arithmetic sum of the three registered components: 979,544 + 109,220 + 12,800,000 = **13.89 m**. SOM is a projection, not a count — see §5.

### 3.2 Pricing and willingness to pay

| Assumption | Value | Source | Grade |
|---|---:|---|---|
| Parents willing to pay more for high-quality courses | 44.67% | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html | Primary |
| Youth sports training market size | RMB 253.1 bn | Same | Primary |
| Renewal revenue share of total income | 57.07% | Same | Primary |
| High-end course renewal rate (RMB 30k+ annual fee) | 92% | https://www.renrendoc.com/paper/415852129.html | Industry research |
| High-end customers as share of students | 8% | Same | Industry research |
| Gen-Z parents requiring data-driven feedback | 68% | Same | Industry research |
| Gen-Z parents checking coach credentials | 75% | Same | Industry research |
| Affluent households: health spending rank | No. 2 category | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 | Primary |
| Affluent households: service-experience spend plan | +12% | Same | Primary |
| Affluent households: weekly exercise | > 2 hours | Same | Primary |

### 3.3 Competitive price anchors

| Product | Price | Source | Grade |
|---|---:|---|---|
| INSAIT JOY Smart Football | US$139 (no subscription) | https://insaitjoy.com/products/insait-joy-smart-football | Primary |
| INSAIT smart football — China government procurement | RMB 1,745 / unit | https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/ | Primary |
| INSAIT athlete sensor — same procurement | RMB 4,798 / unit | Same | Primary |
| adidas Leagues Cup 26 Pro Ball | US$170 | https://www.adidas.com/us/leagues-cup-26-pro-ball/KE5154.html | Primary |
| Catapult One (GPS vest) | ≈ US$179 / year, membership only | https://soccerstories.net/catapult-one-review/ | Secondary |

**Positioning implication.** At HK$980 (≈ RMB 900), the product is priced **below** the INSAIT consumer ball, **roughly half** the Chinese government procurement unit price, and **below** Catapult One's annual membership — which does not include a ball. The HK$1,592 subscription is approximately **5% of the RMB 30,000+ annual fee** already paid by the 92%-renewal high-end course segment.

### 3.4 Commercial and cost assumptions

| Assumption | Value | Basis |
|---|---:|---|
| Weighted ARPU | HK$1,592 / year | Tier mix 60 / 25 / 12 / 3 across HK$65 / 148 / 298 / 698 monthly tiers |
| Ball selling price | HK$980 | Competitive anchor (§3.3) |
| Ball manufacturing cost | HK$250 | Supplier estimate |
| Ball B2B revenue share | 35% | Minimum for B2B ball contribution to be positive (+HK$83 per ball) |
| Ball attach rate | 50% of new users | One ball per two new subscribers |
| Platform base cost | HK$0.20 m / year | Fixed |
| Platform compute | HK$2.00 / user / year | Variable |
| User enhancement | HK$5.00 / user / year | Variable |
| Ball enhancement | HK$10.00 / ball | Variable |
| ITF RTH manpower (gross) | HK$0.48 m / year | 2 BSc × HK$20,000 × 12 |
| Grant reimbursement lag | 6 months | Scheme terms |

### 3.5 Known conservatism gaps

Two inputs are deliberately set below benchmark estimates, and the panel should read the P&L accordingly:

**Compute cost.** The model allows **HK$7 per user per year** for combined platform compute and user enhancement. Processing three-phone video per kick through YOLO, SAM3 and MediaPipe on cloud GPU benchmarks at HK$10–16 per user per year at 25 kicks. The model may therefore understate infrastructure cost by **HK$2–5 m cumulative** at P50. This does not threaten the fairness ratio, but it is disclosed.

**68% Gen-Z data figure.** Sourced from an industry research document rather than an official statistic. It supports, but does not independently verify, the P50 conversion assumption.

---

## 4. Marketing Cost — Why 30 / 35 / 40 Percent

Marketing cost is expressed as a percentage of gross subscription revenue and is borne **entirely by the B2C partner**. It does not appear in the B2B P&L at all. The three levels correspond to three defensible operating assumptions:

| Level | Interpretation | Rationale |
|---:|---|---|
| **30%** | Base case — partner-leveraged distribution | Anta, Li Ning, Amer Sports and Asia Partners already operate retail, e-commerce, school and academy channels. Football subscription attaches to existing customer relationships rather than requiring category creation. |
| **35%** | Stress case — higher co-op marketing | Partners require joint brand campaigns, on-site activation and coach education. Still within the range observed for youth sports subscriptions. |
| **40%** | Conservative bound — Total TAM equivalent | Equivalent to acquiring users outside the partner-owned channel base. Used as a downside test, not a planning basis. |

**Mechanism.** Changing `marketing_pct` on the Assumptions sheet triggers a VLOOKUP into a three-row calibration table. Each row pairs a marketing level with a pre-solved split that restores the fairness ratio:

| Marketing | B2B split | B2C split | Fairness (P50) |
|---:|---:|---:|---:|
| 30% | 12.86% | 37.14% | 1.0997 |
| 35% | 10.48% | 39.52% | 1.0990 |
| 40% | 8.09% | 41.91% | 1.0990 |

**Why the split must move.** Higher marketing consumes more B2C margin. If the split were held constant at 12.86 / 37.14, moving marketing from 30% to 40% would cut B2C cumulative profit from ≈HK$370 m to ≈HK$120 m while B2B stayed at ≈HK$336 m — a fairness ratio of 0.36. The B2C partner would not sign. Shifting 4.77 percentage points of gross subscription from B2B to B2C at the 40% level restores the ratio to 1.10.

---

## 5. Adoption Scenarios — Why P10 / P50 / P90

The three scenarios are not arbitrary labels. They are **fixed fractions of a single externally-anchored behavioural ceiling**, and they form a statistically coherent distribution.

**The anchor.** 8.7% of Chinese parents report willingness to support their child's long-term football path (source: §3.1). This is the outer limit of any paid-conversion assumption — no product can reach more than the group that already backs the pathway.

**The three bands.**

| Scenario | Formula | Conversion | Share of ceiling | Y6 SOM |
|---|---|---:|---:|---:|
| **P10 — Conservative** | 8.7 ÷ 4 | 2.175% | 25% | ≈ 302,100 |
| **P50 — Base** | 8.7 ÷ 2 | 4.35% | 50% | ≈ 604,200 |
| **P90 — Optimistic** | 8.7 ÷ 1 | 8.7% | 100% | ≈ 1,208,400 |

**Why 4.35% is a defensible central estimate.** The demand-qualified ceiling is the product of pathway commitment and data-driven demand: 8.7% × 68% ≈ **5.9%**. A P50 sitting at 5.9% would assume every parent who says they want data actually subscribes — a claim no consumer business would defend. A P50 sitting far below would ignore that two-thirds of the target cohort has already expressed demand. **4.35% captures 74% of the demand-qualified ceiling** — a 26% allowance for the gap between stated preference and paid behaviour.

**Statistical coherence.** P50 ÷ P10 = 2.0 and P90 ÷ P50 = 2.0. This is the signature of a lognormal distribution with log-space standard deviation σ ≈ 0.54 (geometric SD ≈ 1.72) — the appropriate shape for positive, right-skewed user counts. The three bands are therefore a probability structure, not marketing labels.

**How to switch scenarios.** Set `scenario_selector` on the Assumptions sheet to 1, 2 or 3. The CHOOSE function returns 0.50, 1.00 or 2.00, which multiplies the base penetration rate in every year. All downstream P&L recalculates.

---

## 6. Subscription Split — How the Revenue Is Divided

Gross subscription revenue passes through three recipients in a fixed order:

| Step | Recipient | Share of gross |
|---|---|---:|
| 1 | Coach / player ecosystem payout | **50.00%** |
| 2 | B2B technology enabler | **12.86%** |
| 3 | B2C distribution partners | **37.14%** |
| | **Total** | **100.00%** |

**Why the coach/player payout comes first.** The 50% is a contractual ecosystem payment, not a partner margin. It is paid before any split between B2B and B2C and is therefore not subject to renegotiation when marketing changes.

**How the remaining 50% is divided.** At the 30% marketing level, B2B receives 25.7% of the remaining 50% and B2C receives 74.3%. This reflects the cost structure:

- **B2C** carries marketing (30% of gross), distribution and customer acquisition.
- **B2B** carries the platform (HK$0.20 m fixed plus HK$2.00 per user), user and ball enhancement, ball manufacturing, and the ITF manpower cost offset by the government grant.

**Why the split is calibrated rather than negotiated ad hoc.** The 12.86 / 37.14 split is the unique pair that (a) sums with the 50% coach payout to exactly 100%, and (b) produces a full-cycle fairness ratio of **1.10** — the ceiling the partners agreed. Every alternative split either breaches the 10% band or leaves one partner with a negative year.

---

## 7. Financial Results — Base Case and Grid

### 7.1 Base case (30% marketing, P50 adoption)

| Year | B2B profit | B2C profit | B2B cumulative | B2C cumulative | Fairness |
|---|---:|---:|---:|---:|---:|
| Y1 | 3.1 | 5.6 | 3.1 | 5.6 | 1.80 |
| Y2 | 7.6 | 10.2 | 10.7 | 15.8 | 1.47 |
| Y3 | 44.5 | 71.9 | 55.2 | 87.7 | 1.59 |
| Y4 | 64.5 | 69.6 | 119.8 | 157.3 | 1.31 |
| Y5 | 89.3 | 87.5 | 209.1 | 244.8 | 1.17 |
| Y6 | 126.9 | 125.2 | 336.0 | 369.9 | 1.10 |
| Y7 Q2 grant | +0.24 | — | 336.2 | — | 1.10 |

Neither partner records a negative year. The ratio converges from 1.80 in Year 1 — when B2B is absorbing the R&D drag — to 1.10 by Year 6.

### 7.2 Nine-cell grid — B2B cumulative profit (HK$ m)

| Marketing | P10 | P50 | P90 |
|---:|---:|---:|---:|
| 30% | ≈ 168 | **≈ 336** | ≈ 672 |
| 35% | ≈ 148 | ≈ 295 | ≈ 590 |
| 40% | ≈ 128 | ≈ 255 | ≈ 510 |

### 7.3 Nine-cell grid — fairness ratio (B2C ÷ B2B)

| Marketing | P10 | P50 | P90 |
|---:|---:|---:|---:|
| 30% | 1.10 | **1.10** | 1.10 |
| 35% | 1.10 | 1.10 | 1.10 |
| 40% | 1.10 | 1.10 | 1.10 |

All nine combinations sit within the 10% band — the calibration table is doing its job across the full grid, not just at the base case.

---

## 8. Model Integrity and Limitations

**Nine PASS/FAIL checks** run automatically on the Summary sheet after every input change:

1. B2B cumulative equals the sum of annual profits
2. B2C cumulative equals the sum of annual profits
3. Full-cycle fairness ≤ 1.10
4. Full-cycle fairness ≥ 1.09 (both partners roughly equal)
5. All years B2B non-negative
6. All years B2C non-negative
7. Subscription split sums to 100%
8. Marketing rate is one of the three calibrated values
9. Scenario selector is 1, 2 or 3

**Limitations the panel should weigh:**

- The model is a **planning tool, not an audited statement**. Taxes, payment processing fees, support, warranty, logistics, bad debt, working capital, capex and financing costs are excluded and must be modelled separately.
- **SOM is a projection, not a count.** The three bands are fractions of a behavioural ceiling, not observed paid-user data.
- **Two industry-research inputs** (92% high-end renewal, 68% Gen-Z data demand) are not official statistics. They support but do not verify the P50.
- **Compute cost may be understated by HK$2–5 m cumulative** (§3.5). This does not change the conclusions.
- **Partner credentials require independent due diligence** before any binding commercial agreement, particularly for Asia Partners IFBD / Svenska Capital.

---

## 9. Change Log

| Version | Change |
|---|---|
| 3.0 | Concise §3.7; §3.8 evidence table; Times New Roman typography |
| 3.1 | Live workbook with VLOOKUP split and CHOOSE scenario |
| **3.2** | **Panel edition — assumption verification register with evidence grades; dedicated sections on marketing variation, scenario logic, and split mechanism; redundant grid tables removed; document reduced to 7 pages** |

---

*Prepared for partnership and investment review. The workbook (`KungFu_AI_Football_Financial_Model_v3.1_fixed.xlsx`) contains the live formulas; this document explains the reasoning behind them.*
"""


def build():
    doc = Document()
    base = doc.styles['Normal']
    base.font.name = FONT
    base.font.size = Pt(SIZES['body'])
    base.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

    # A4 PORTRAIT
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.orientation = WD_ORIENT.PORTRAIT
    section.top_margin = Mm(18)
    section.bottom_margin = Mm(18)
    section.left_margin = Mm(20)
    section.right_margin = Mm(20)

    # Cover
    for _ in range(6):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Kung Fu AI Football'), SIZES['cover_title'], bold=True, color=ACCENT)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Six-Year Financial Model'), SIZES['cover_title'], bold=True, color=ACCENT)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('README v3.2 — Panel Edition'), 12, bold=True)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Assumption verification · Marketing sensitivity · Scenario framework · Split logic'),
            SIZES['cover_sub'], italic=True, color=GREY_TEXT)
    for _ in range(8):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Prepared for partnership and investment review'), SIZES['cover_sub'])
    doc.add_page_break()

    render(doc, README_MD)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('— End of README v3.2 —'), SIZES['footer'], italic=True, color=GREY_TEXT)

    out = 'KungFu_AI_Football_README_v3.2.docx'
    doc.save(out)
    print(f'✅ Saved: {out}')
    print(f'   Format: A4 Portrait · Times New Roman · Body 10pt · H1 14pt · H2 12pt')
    print(f'   Target length: ≤ 8 pages')


if __name__ == '__main__':
    build()