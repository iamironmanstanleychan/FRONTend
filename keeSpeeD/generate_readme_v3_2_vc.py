"""
README v3.2 — VC Edition
Times New Roman · A4 Portrait · Body 10pt · H1 14pt · H2 12pt
Output: KungFu_AI_Football_README_v3.2_VC.docx
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
MONO = 'Consolas'
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


def set_run(run, size, bold=False, italic=False, color=None, font=None):
    f = font or FONT
    run.font.name = f
    run._element.rPr.rFonts.set(qn('w:eastAsia'), f)
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
            set_run(paragraph.add_run(part[1:-1]), size, bold=bold, font=MONO)
        else:
            set_run(paragraph.add_run(part), size, bold=bold)


def add_heading(doc, text, level):
    size = {1: SIZES['h1'], 2: SIZES['h2'], 3: SIZES['h3']}.get(level, SIZES['h3'])
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
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


def add_table(doc, rows, bar_col=None):
    """bar_col: 0-indexed column that should use monospace font for block characters."""
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
            is_bar = (bar_col is not None and j == bar_col)
            font = MONO if is_bar else None
            run = p.add_run(str(val))
            if i == 0:
                set_run(run, SIZES['table'], bold=True, color=HEADER_TEXT, font=font)
                set_cell_bg(cell, HEADER_FILL)
            else:
                set_run(run, SIZES['table'], font=font,
                        color=ACCENT if is_bar else None)


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
                # Detect bar column by header
                bar_col = None
                if len(rows[0]) > 0:
                    for idx, h in enumerate(rows[0]):
                        if 'bar' in h.lower() or 'visual' in h.lower() or '■' in h:
                            bar_col = idx
                            break
                add_table(doc, rows, bar_col=bar_col)
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

**README v3.2 · Venture Capital Edition**

Prepared for: partnership and investment review
Model owner: B2B technology enabler
Anchor case: **35% marketing · P50 adoption** — the centre cell of the 3 × 3 grid

---

## 1. Executive Summary

Kung Fu AI Football is a **B2B technology enabler** supplying an AI-assisted football training platform to four B2C distribution partners — Anta, Li Ning, Amer Sports, and Asia Partners IFBD (Svenska Capital). The model forecasts six years of profitability for both sides of the partnership.

**The base case sits at the centre of the risk grid, not at an optimistic edge.** At 35% B2C marketing cost and P50 adoption, the model returns **≈HK$277 m** in cumulative B2B profit over six years, with a full-cycle fairness ratio of **1.10** between the two partners.

**The risk/reward spread is the core of the investment case:**

| Case | Marketing | Adoption | Six-year B2B cumulative |
|---|---:|---:|---:|
| **Safe floor** | 40% | P10 | **HK$108.55 m** |
| **Anchor** | 35% | P50 | **≈HK$277 m** |
| **Upside ceiling** | 30% | P90 | **HK$672.45 m** |

Even the **safe floor** — highest marketing cost, lowest adoption — returns **HK$108.55 m**. The anchor is **2.6× the floor**. The ceiling is **6.2× the floor**. This is the risk/reward profile presented in §2.

**Three strengths the panel should note:**

1. **Every input traces to a specific source.** §4 lists all assumptions against article-level URLs with an explicit evidence grade — no home-page citations remain.
2. **The three adoption scenarios form a statistical structure, not marketing labels.** They are fixed fractions of a single behavioural ceiling and form a lognormal distribution (§6).
3. **Marketing cost is the most sensitive term — and the model auto-rebalances.** §5 explains the VLOOKUP mechanism that keeps the fairness ratio at 1.10 across all three marketing levels.

---

## 2. Market Potential — The Risk/Reward Case

### 2.1 The anchor: 35% marketing at P50 adoption

The centre cell of the 3 × 3 grid is the planning basis. It sits neither at the most optimistic corner (30% / P90) nor at the most conservative (40% / P10), but at the **moderate intersection** of marketing cost and adoption.

**B2B profit by year — 35% marketing, P50 adoption (HK$ m):**

| Year | Profit | Visual bar |
|:---|---:|:---|
| Y1 | 2.6 | █ |
| Y2 | 6.3 | ███ |
| Y3 | 37.2 | ██████████████████ |
| Y4 | 53.0 | █████████████████████████ |
| Y5 | 73.2 | ███████████████████████████████████ |
| Y6 | 104.0 | ██████████████████████████████████████████████████ |

The shape is **convex** — profit accelerates as the fixed cost base (HK$0.68 m per year in platform base plus ITF manpower) is amortised over a rapidly growing subscriber count. Year 1 profit is modest because the platform is being stood up. By Year 6, profit exceeds Year 1 by **40×**.

### 2.2 Why the safe floor (HK$108.55 m) is genuinely safe

The bottom-left corner of the grid — **40% marketing, P10 adoption** — is the worst case across all nine combinations. Four structural features make it a **credible floor**, not a distress scenario:

| Safeguard | Mechanism | Effect |
|---|---|---|
| **Partner-owned distribution** | Anta, Li Ning, Amer Sports and Asia Partners reach existing school, academy and retail customers | Even at P10 (half the P50 user count), partner channels still deliver ≈302,000 subscribers by Year 6 |
| **Auto-calibrated split** | At 40% marketing, the B2B share drops from 12.86% to 8.09% | B2B profit stays positive because the split re-solves to preserve 1.10 fairness |
| **ITF RTH grant** | Government covers 100% of the HK$0.48 m/year BSc manpower cost | HK$2.88 m of cost is neutralised over six years |
| **Ball contribution** | Each ball earns B2B +HK$83 net of manufacturing and enhancement | Ball profit is unaffected by marketing percentage — it depends only on user growth |

The floor is not a "what if everything goes wrong" case — it is a **well-defended pessimistic case**. The partner channels, the calibrated split, the grant, and the ball margin all remain intact even at maximum marketing cost and minimum adoption.

### 2.3 What unlocks the upside (HK$672.45 m)

The top-right corner — **30% marketing, P90 adoption** — is the best case across the grid. Three drivers combine to produce **6.2× the floor**:

**Operating leverage.** The fixed base of HK$0.68 m per year is amortised over **1.21 million subscribers by Year 6** (double the P50 count). Fixed cost per user falls from HK$49.82 in Year 1 to **HK$0.56** in Year 6.

**Partner-leveraged marketing.** At 30% marketing cost, the B2C partners attach the subscription to existing football equipment, school sponsorships and retail flows. The 30% rate reflects their established channel advantage, not a discount to market rates.

**Ball revenue scaling.** Ball revenue grows with **new subscribers**, not with the total base. At P90, new-user flow is roughly double the P50 case, which doubles the ball business **without doubling the fixed cost base**.

### 2.4 The full 3 × 3 grid

**B2B cumulative profit, six years plus Y7 Q2 grant (HK$ m):**

| Marketing | P10 | P50 | P90 |
|---:|---:|---:|---:|
| **30%** | 168 | 336 | **672** ← ceiling |
| **35%** | 138 | **277** ← anchor | 553 |
| **40%** | **109** ← floor | 217 | 433 |

Every cell is positive. Every cell is traceable to the live formulas in the workbook. The anchor sits at the centre; the floor and ceiling bracket it at **2.6×** and **2.4×** respectively.

---

## 3. B2B Financial Forecasts — 35% / P50 Anchor Case

The full six-year P&L at the anchor case, in HK$ million.

| Year | SOM | Sub Rev | Ball Rev | Custom | License | ITF Grant | Total Rev | Platform | User Enh | Ball Enh | Ball Cost | Manpower | Total Cost | Profit | Cumulative |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Y1 | 13,650 | 2.28 | 2.34 | 0.10 | 0.20 | 0.24 | 5.16 | 0.23 | 0.07 | 0.07 | 1.71 | 0.48 | 2.56 | 2.60 | 2.60 |
| Y2 | 34,130 | 5.69 | 3.51 | 0.025 | 0.20 | 0.48 | 9.91 | 0.27 | 0.17 | 0.10 | 2.57 | 0.48 | 3.59 | 6.32 | 8.92 |
| Y3 | 192,160 | 32.06 | 27.11 | 0.025 | 0.20 | 0.48 | 59.88 | 0.58 | 0.96 | 0.79 | 19.85 | 0.48 | 22.66 | 37.22 | 46.14 |
| Y4 | 303,000 | 50.55 | 19.01 | 0.025 | 0.20 | 0.48 | 70.29 | 0.81 | 1.52 | 0.56 | 13.93 | 0.48 | 17.30 | 52.99 | 99.13 |
| Y5 | 425,980 | 71.07 | 21.09 | 0.025 | 0.20 | 0.48 | 92.87 | 1.05 | 2.13 | 0.62 | 15.44 | 0.48 | 19.72 | 73.15 | 172.28 |
| Y6 | 604,200 | 100.84 | 30.52 | 0.025 | 0.20 | 0.48 | 132.07 | 1.41 | 3.02 | 0.89 | 22.29 | 0.48 | 28.09 | 103.98 | **276.26** |
| Y7 Q2 | — | — | — | — | — | 0.24 | 0.24 | — | — | — | — | — | 0.00 | 0.24 | **276.50** |

**Reading the table.** Subscription revenue is **36% of total revenue by Year 6** — the largest single line. Ball revenue is the second-largest and is **bimodal**: strong in Year 3 (when the mid-tier regions activate) and again in Year 6 (when the final regions come online). The cost base is dominated by ball manufacturing (HK$22.29 m in Year 6), which is fully variable and therefore not a fixed-risk item.

---

## 4. Assumption Verification Register

Every number that drives the P&L, with its source and evidence grade.

### 4.1 Market sizing

| Assumption | Value | Source | Grade |
|---|---:|---|---|
| Chinese football fans | 380 m | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | Primary |
| Football industry total scale (2025) | RMB 320 bn | Same | Primary |
| Football training market (2025) | RMB 48.76 bn | Same | Primary |
| Campus football registered students | 12.8 m | Same | Primary |
| Elite youth registered players | 109,220 | http://news.china.com.cn/2026-01/30/content_118308779.shtml | Primary |
| Social amateur players | 979,544 | Same | Primary |
| Parents supporting a long-term football path | 8.7% | https://www.163.com/dy/article/L8E6C26F0556BPXI.html | Primary |

**Derived market layers.** TAM = 300 m (79% of the 380 m fan base). SAM = 13.89 m (the arithmetic sum of the three registered components). SOM at P50 ≈ 604,200 — a projection, not a count.

### 4.2 Pricing and willingness to pay

| Assumption | Value | Source | Grade |
|---|---:|---|---|
| Parents willing to pay more for quality courses | 44.67% | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html | Primary |
| Renewal revenue share of total income | 57.07% | Same | Primary |
| High-end course renewal rate (RMB 30k+ annual fee) | 92% | https://www.renrendoc.com/paper/415852129.html | Industry research |
| Gen-Z parents requiring data-driven feedback | 68% | Same | Industry research |
| Affluent households: health spending rank | No. 2 category | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 | Primary |

### 4.3 Competitive price anchors

| Product | Price | Source | Grade |
|---|---:|---|---|
| INSAIT JOY Smart Football | US$139 (no subscription) | https://insaitjoy.com/products/insait-joy-smart-football | Primary |
| INSAIT smart football — China government procurement | RMB 1,745 / unit | https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/ | Primary |
| adidas Leagues Cup 26 Pro Ball | US$170 | https://www.adidas.com/us/leagues-cup-26-pro-ball/KE5154.html | Primary |
| Catapult One (GPS vest) | ≈ US$179 / year, membership only | https://soccerstories.net/catapult-one-review/ | Secondary |

**Positioning implication.** At HK$980, the product is priced **below** the INSAIT consumer ball, **roughly half** the Chinese government procurement price, and **below** Catapult One's annual membership — which does not include a ball. The HK$1,592 subscription is approximately **5% of the RMB 30,000+ annual fee** already paid by the 92%-renewal high-end course segment.

### 4.4 Cost assumptions

| Assumption | Value | Basis |
|---|---:|---|
| Weighted ARPU | HK$1,592 / year | Tier mix 60 / 25 / 12 / 3 across HK$65 / 148 / 298 / 698 monthly tiers |
| Ball selling price | HK$980 | Competitive anchor (§4.3) |
| Ball manufacturing cost | HK$250 | Supplier estimate |
| Ball B2B revenue share | 35% | Minimum for positive B2B ball contribution |
| Ball attach rate | 50% of new users | One ball per two new subscribers |
| Platform base | HK$0.20 m / year | Fixed |
| Platform compute | HK$2.00 / user / year | Variable |
| User enhancement | HK$5.00 / user / year | Variable |
| Ball enhancement | HK$10.00 / ball | Variable |
| ITF RTH manpower (gross) | HK$0.48 m / year | 2 BSc × HK$20,000 × 12 |

### 4.5 Known conservatism gaps

Two inputs are deliberately set below benchmark estimates:

**Compute cost.** The model allows HK$7 per user per year. Cloud GPU benchmarks for three-phone kick video processing indicate HK$10–16 per user per year at 25 kicks. The model may understate infrastructure cost by **HK$2–5 m cumulative** at P50. Does not threaten the fairness ratio.

**68% Gen-Z data figure.** Sourced from industry research, not an official statistic. It supports, but does not independently verify, the P50 conversion assumption.

---

## 5. Marketing Cost — Why 30 / 35 / 40 Percent

Marketing cost is expressed as a percentage of gross subscription revenue and is borne **entirely by the B2C partner**. It does not appear in the B2B P&L at all.

| Level | Interpretation | Rationale |
|---:|---|---|
| **30%** | Partner-leveraged distribution | Anta, Li Ning, Amer Sports and Asia Partners already operate retail, e-commerce, school and academy channels |
| **35%** | Base case — higher co-op marketing | Partners require joint campaigns, on-site activation and coach education |
| **40%** | Conservative bound | Equivalent to acquiring users outside the partner-owned channel base |

**The split must move with marketing.** Higher marketing consumes B2C margin. If the split were held at 12.86 / 37.14 while marketing rose to 40%, B2C cumulative profit would collapse from ≈HK$370 m to ≈HK$120 m while B2B stayed at ≈HK$336 m — a fairness ratio of **0.36**. The model auto-rebalances:

| Marketing | B2B split | B2C split | Fairness (P50) |
|---:|---:|---:|---:|
| 30% | 12.86% | 37.14% | 1.0997 |
| **35%** | **10.48%** | **39.52%** | **1.0990** |
| 40% | 8.09% | 41.91% | 1.0990 |

Shifting 4.77 percentage points of gross subscription from B2B to B2C at the 40% level restores the 1.10 ratio. The mechanism is a VLOOKUP into a three-row calibration table.

---

## 6. Scenario Framework — Why P10 / P50 / P90

The three scenarios are **fixed fractions of a single externally-anchored behavioural ceiling**.

**The anchor.** 8.7% of Chinese parents report willingness to support a child's long-term football path (§4.1). No paid-conversion assumption can exceed this group.

| Scenario | Formula | Conversion | Share of ceiling | Y6 SOM |
|---|---|---:|---:|---:|
| **P10 — Conservative** | 8.7 ÷ 4 | 2.175% | 25% | ≈ 302,100 |
| **P50 — Base** | 8.7 ÷ 2 | 4.35% | 50% | ≈ 604,200 |
| **P90 — Optimistic** | 8.7 ÷ 1 | 8.7% | 100% | ≈ 1,208,400 |

**Why 4.35% is defensible.** The demand-qualified ceiling is 8.7% × 68% ≈ **5.9%**. A P50 at 5.9% would assume every parent who wants data actually subscribes — a claim no consumer business would defend. **4.35% captures 74% of the demand-qualified ceiling**, allowing a 26% gap between stated preference and paid behaviour.

**Statistical coherence.** P50 ÷ P10 = 2.0 and P90 ÷ P50 = 2.0 — the signature of a lognormal distribution with log-space standard deviation σ ≈ 0.54 (geometric SD ≈ 1.72). The three bands form a probability structure, not marketing labels.

---

## 7. Subscription Split — How the Revenue Is Divided

Gross subscription revenue passes through three recipients in a fixed order:

| Step | Recipient | Share of gross |
|---|---|---:|
| 1 | Coach / player ecosystem payout | **50.00%** |
| 2 | B2B technology enabler | **10.48%** (at 35% marketing) |
| 3 | B2C distribution partners | **39.52%** (at 35% marketing) |
| | **Total** | **100.00%** |

**Why the coach/player payout comes first.** The 50% is a contractual ecosystem payment, not a partner margin. It is paid before any B2B/B2C split and is therefore **not subject to renegotiation** when marketing changes.

**Why the split is calibrated rather than negotiated ad hoc.** The 10.48 / 39.52 split is the unique pair at 35% marketing that (a) sums with the 50% coach payout to exactly 100%, and (b) produces a full-cycle fairness ratio of **1.10**. Every alternative either breaches the 10% band or leaves one partner with a negative year.

---

## 8. Risk, Integrity, and Limitations

**Nine PASS/FAIL checks** run automatically after every input change:

1. B2B cumulative equals the sum of annual profits
2. B2C cumulative equals the sum of annual profits
3. Full-cycle fairness ≤ 1.10
4. Full-cycle fairness ≥ 1.09
5. All years B2B non-negative
6. All years B2C non-negative
7. Subscription split sums to 100%
8. Marketing rate is one of the three calibrated values
9. Scenario selector is 1, 2 or 3

**Limitations the panel should weigh:**

- The model is a **planning tool, not an audited statement**. Taxes, payment processing fees, support, warranty, logistics, bad debt, working capital, capex and financing costs are excluded.
- **SOM is a projection, not a count.** The three bands are fractions of a behavioural ceiling, not observed paid-user data.
- **Two industry-research inputs** (92% high-end renewal, 68% Gen-Z data demand) are not official statistics.
- **Compute cost may be understated by HK$2–5 m cumulative** (§4.5). Does not change the conclusions.
- **Partner credentials require independent due diligence** before any binding commercial agreement, particularly for Asia Partners IFBD / Svenska Capital.

---

## 9. Change Log

| Version | Change |
|---|---|
| 3.0 | Concise §3.7; §3.8 evidence table; Times New Roman typography |
| 3.1 | Live workbook with VLOOKUP split and CHOOSE scenario |
| **3.2 (VC)** | **Market potential leads after executive summary; 35% / P50 middle-centre anchor with visual bar chart; HK$108.55 m floor and HK$672.45 m ceiling framed as risk/reward; §2 Model Architecture deleted and replaced by B2B Financial Forecasts; remaining sections reduced to verification and cross-check** |

---

*Prepared for venture capital and partnership review. The workbook (`KungFu_AI_Football_Financial_Model_v3.1_fixed.xlsx`) contains the live formulas; this document explains the reasoning and the investment case.*
"""


def build():
    doc = Document()
    base = doc.styles['Normal']
    base.font.name = FONT
    base.font.size = Pt(SIZES['body'])
    base.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

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
    set_run(p.add_run('README v3.2 — Venture Capital Edition'), 12, bold=True)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Market potential · Risk/reward · Anchor case · Verification'), SIZES['cover_sub'], italic=True, color=GREY_TEXT)
    for _ in range(8):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Prepared for venture capital and partnership review'), SIZES['cover_sub'])
    doc.add_page_break()

    render(doc, README_MD)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('— End of README v3.2 —'), SIZES['footer'], italic=True, color=GREY_TEXT)

    out = 'KungFu_AI_Football_README_v3.2_VC.docx'
    doc.save(out)
    print(f'✅ Saved: {out}')
    print(f'   Format: A4 Portrait · Times New Roman · Body 10pt · H1 14pt · H2 12pt')
    print(f'   Anchor case: 35% / P50 · Floor: HK$108.55 m · Ceiling: HK$672.45 m')


if __name__ == '__main__':
    build()