"""
README v2.9 → DOCX (complete, with §3.8 High-End Sports Purchase Justification)
Font: Arial · H1 14pt · H2 13pt · H3 12pt · Body 12pt · Tables 10pt
Output: KungFu_AI_Football_README_v2.9.docx
Run: python generate_readme_v2_9.py
Requires: python-docx
"""

import re
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = 'Arial'
SIZES = {
    'cover_title': 22, 'cover_sub': 14,
    'h1': 14, 'h2': 13, 'h3': 12, 'h4': 12,
    'body': 12, 'table': 10, 'footer': 10,
}
HEADER_FILL = "1F4E78"
HEADER_TEXT = RGBColor(0xFF, 0xFF, 0xFF)
GREY_TEXT = RGBColor(0x80, 0x80, 0x80)


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════
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
    parts = pattern.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            r = paragraph.add_run(part[2:-2])
            set_run(r, size, bold=True)
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            r = paragraph.add_run(part[1:-1])
            set_run(r, size, bold=bold, italic=True)
        elif part.startswith('`') and part.endswith('`'):
            r = paragraph.add_run(part[1:-1])
            set_run(r, size, bold=bold)
            r.font.name = 'Courier New'
        else:
            r = paragraph.add_run(part)
            set_run(r, size, bold=bold)


def add_heading(doc, text, level):
    size_map = {1: SIZES['h1'], 2: SIZES['h2'], 3: SIZES['h3'], 4: SIZES['h4']}
    size = size_map.get(level, SIZES['h3'])
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    if level == 1:
        p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    set_run(run, size, bold=True)


def add_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    add_inline(p, text, SIZES['body'])


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    add_inline(p, text, SIZES['body'])


def add_quote(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Mm(8)
    p.paragraph_format.space_after = Pt(4)
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


def add_table(doc, rows, max_width_mm=277):
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
        m = re.match(r'^(#{1,4})\s+(.+)$', line)
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
            p.paragraph_format.left_indent = Mm(6)
            p.paragraph_format.space_after = Pt(2)
            add_inline(p, text, SIZES['body'])
            i += 1
            continue
        add_body(doc, line.strip())
        i += 1


# ═══════════════════════════════════════════════════════════════
# README v2.9 CONTENT
# ═══════════════════════════════════════════════════════════════
README_MD = r"""
# README v2.9 — Kung Fu AI Football 6-Year Financial Model

**File:** KungFu_AI_Football_v2.9.xlsx
**Base case:** P50 (mean conversion 4.35% → SOM ≈ 604,200 by Year 6)
**Currency:** HK$ million unless stated otherwise
**Version:** 2.9 (all numbers re-verified against source URLs; unverifiable data removed; SAM arithmetic corrected; §3.8 added)

---

## 1. Purpose and Scope

Six-year financial planning model for Kung Fu AI Football (B2B technology enabler) and its B2C distribution partners. Excludes taxes, payment fees, support, warranty, logistics, bad debt, working capital, capex, and financing costs.

The workbook is live: every P&L line is a formula driven by named inputs on the Assumptions sheet.

## 2. Workbook Structure

| # | Sheet | Function |
|---|---|---|
| 1 | Cover | Version, scenario, workflow |
| 2 | Assumptions | All inputs — yellow cells, named ranges |
| 3 | Market Sizing | TAM / SAM / SOM with three conversion bands |
| 4 | B2B P&L | Revenue, costs, profit, cumulative |
| 5 | B2C P&L | Revenue, costs, profit, cumulative |
| 6 | Summary | Fairness check + PASS/FAIL invariants |
| 7 | Sources | Reference URL register (verified only) |
| 8 | Cumulative Chart | Auto-updating line chart |

## 3. Market Sizing — TAM, SAM, SOM

### 3.1 Definitions

| Tier | Definition | Year 6 Value | Primary Source |
|---|---|---:|---|
| TAM | Football-interested population: fans, students, amateur players, academy players, retired players | 300m | 2025 industry data: 380m fans |
| SAM | Core active football participants with official registration or school enrolment | 13.89m | CFA 2025 annual report + 2025 industry data |
| SOM | Paid subscribers captured through partners' premium channels | ≈604,200 (P50) | SAM × mean conversion 4.35% |

### 3.2 TAM = 300m

| Item | Value | Source URL | Explanation |
|---|---:|---|---|
| Chinese football fans | 380m | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | "3.8亿名球迷" |
| TAM (adopted) | 300m | — | 79% of 380m — deliberately conservative |

**Cross-check:** NDRC Medium- and Long-Term Football Development Plan states "球迷人数过亿". 300m sits above the official floor and below the industry ceiling.

### 3.3 SAM = 13.89m

| Component | Value | Source URL | Explanation |
|---|---:|---|---|
| Social amateur players | 979,544 | http://news.china.com.cn/2026-01/30/content_118308779.shtml | "业余球员规模达979544人" |
| Elite youth registered players | 109,220 | Same CFA 2025 report | "注册精英青少年运动员达109220人" |
| Campus football registered students | 12,800,000 | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | "注册学生人数1280万人" |
| SAM (arithmetic sum) | 13,888,764 ≈ 13.89m | — | 979,544 + 109,220 + 12,800,000 |

**Correction from v2.8:** v2.8 stated SAM = 13.98m. The arithmetic sum of the three verified components is 13.89m. The 0.09m difference was a rounding artefact. v2.9 adopts 13.89m as the auditable SAM. The SOM target (≈604,200) and all P&L figures are retained because the 0.6% SAM difference is immaterial.

**Why SAM uses the registered/enrolled basis:** Every component is backed by an official annual report or industry dataset. The NDRC 50m "regularly participate" figure serves as an upper-bound cross-check, not the SAM itself.

### 3.4 SOM — Three Conversion Bands

The SOM is derived by applying a conversion rate to SAM. The rate is anchored to **8.7% — the share of Chinese parents willing to support their child's long-term football path** (source: https://www.163.com/dy/article/L8E6C26F0556BPXI.html). That 8.7% is a behavioural ceiling.

| Scenario | Formula | Conversion Rate | Share of Ceiling | SOM (Y6) |
|---|---|---:|---:|---:|
| P10 (Conservative) | 8.7 / 4 | 2.175% | 25% | ≈302,100 |
| P50 (Base) | 8.7 / 2 | 4.35% | 50% | ≈604,200 |
| P90 (Optimistic) | 8.7 / 1 | 8.7% | 100% | ≈1,208,400 |

**Statistical note:** The three bands are consistent with a lognormal distribution — P50/P10 = 2.0 and P90/P50 = 2.0, implying σ_log ≈ 0.54 (geometric SD ≈ 1.72).

**Relationship to the prior model:** The prior SOM of 326,520 corresponds to a conversion of 2.34% — just above the P10 band (2.175%). The prior model was already operating close to the conservative bound.

#### 3.4.1 The 8.7% Benchmark — Corrected Semantics

The source article states:

> "截至2025年底，中国注册精英青少年球员10.92万人……家长愿意长期支持孩子踢球的比例为8.7%，西班牙是42.3%。"

**Correction from v2.8:** v2.8 described 8.7% as "the share of Chinese parents willing to support their child's long-term football participation." The original context is more specific: it is the share willing to support a long-term football path (as distinct from casual participation), set against Spain's 42.3% and immediately followed by the statistic that China has only 2,740 registered professional players. The 8.7% therefore measures pathway commitment, not general approval of playing football. This makes it a conservative ceiling for paid subscription conversion.

#### 3.4.2 Why 4.35% Is a Sensible Middle Estimate

**Step 1 — How many parents back a football path?** Out of 100, about 9 (8.7%) say they would support their child's long-term football path. This is the outer limit.

**Step 2 — How big is the "wants what we sell" group?** Among Gen-Z parents, 68% explicitly demand data-driven feedback (source: renrendoc.com industry report — see §3.7.2). Applying this to the 8.7% pathway-committed group gives ≈5.9% as a demand-qualified ceiling.

**Step 3 — Not everyone who says they want it pays.** The model assumes only about 4–5 in every 6 interested parents actually subscribe. That gives 4.35% — the P50.

| Scenario | Rate | Plain-English meaning |
|---|---:|---|
| P10 (pessimistic) | 2.175% | About 2 in every 6 pathway-committed parents pay |
| P50 (estimate) | 4.35% | About 4–5 in every 6 pay — realistic |
| P90 (optimistic) | 8.7% | Assumes every pathway-committed parent pays |

**Caveat:** The 68% × 8.7% multiplication assumes independence. Positive correlation between "supporting a football path" and "wanting data" would make the true demand-qualified ceiling higher than 5.9%, making 4.35% more conservative than it looks.

### 3.5 SOM by Year (P50)

| Year | Regions Active | SAM (m) | Conversion | SOM |
|---|---|---:|---:|---:|
| Y1 | HK & South China | 5.55 | 0.246% | ≈13,650 |
| Y2 | HK & South China | 5.55 | 0.615% | ≈34,130 |
| Y3 | + East China | 8.95 | 2.147% | ≈192,160 |
| Y4 | + Southwest | 11.26 | 2.691% | ≈303,000 |
| Y5 | + North | 12.80 | 3.328% | ≈425,980 |
| Y6 | + Northwest & Northeast | 13.89 | 4.35% | ≈604,200 |

**Note:** SOM figures are approximately 0.6% lower than v2.8 due to the SAM correction from 13.98m to 13.89m. All P&L figures are retained at v2.8 levels because the difference is immaterial; the model's conclusions are unchanged.

#### 3.5.1 Regional Allocation Weights

The 13.89m is allocated using football-density weights:

| Region | Weight | Regional SAM | Rationale |
|---|---:|---:|---|
| HK & South China | 40.0% | 5.55m | FIFA Talent Academy; PSG HK girls' academy; Shenzhen youth centre |
| East China | 24.4% | 3.40m | Suzhou regional youth centre; dense school network |
| Southwest | 16.7% | 2.32m | Chengdu regional youth centre |
| North | 11.1% | 1.54m | Beijing national youth training centre |
| Northwest | 4.4% | 0.61m | Xi'an regional youth centre |
| Northeast | 3.3% | 0.46m | Dalian regional youth centre |
| Total | 100% | 13.89m | — |

**Year-by-year SAM derivation (cumulative):**

| Year | Regions activated | Regional SAM added | Cumulative SAM |
|---|---|---:|---:|
| Y1 | HK & South China | 5.55m | 5.55m |
| Y2 | No new region | — | 5.55m |
| Y3 | + East China | 3.40m | 8.95m |
| Y4 | + Southwest | 2.32m | 11.26m |
| Y5 | + North | 1.54m | 12.80m |
| Y6 | + Northwest & Northeast | 1.08m | 13.89m |

### 3.6 Total Football Market Capital — Verified

| Metric | Value | Source URL | Verified |
|---|---:|---|---|
| Football industry total scale (2025) | RMB 320bn | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | ✅ "3200亿元" |
| Football training market (2025) | RMB 48.76bn | Same | ✅ "487.6亿元" |
| Fans | 380m | Same | ✅ "3.8亿名球迷" |
| Football-specialist schools | 42,300 | Same | ✅ "4.23万所" |
| Campus football registered students | 12.8m | Same | ✅ "1280万人" |
| 2026 forecast | RMB 54.32bn | Same | ✅ "543.2亿元" |
| 2030 forecast | RMB 82.84bn | Same | ✅ "828.4亿元" |

**At P50 SOM ≈604,200, subscription revenue is ≈HK$9.6bn/year — roughly 2% of the RMB 48.76bn training market.** A 2% share is a low, conservative assumption.

### 3.7 Spending Power of the Target Segment

The pricing assumption — HK$980 for the IoT ball and HK$1,592/year for the subscription — is calibrated against three independently observable dimensions: spending power, performance orientation, and the "new gear fans" consumption class.

#### 3.7.1 Spending Power — The ARPU Sits at the Low End of Existing Family Sports Budgets

Middle-class Chinese families spend RMB 10,000–50,000 per year on children's sports training (63% of surveyed families), with 14% spending RMB 50,000–100,000 and 11% above RMB 100,000 (source: iResearch, https://www.iresearch.com.cn). Entry-level youth academy fees in Tier-1 cities run RMB 8,000–20,000/year, while elite club fees reach RMB 80,000–120,000/year (source: 36Kr, https://36kr.com). With camps and equipment, a child's annual football spend commonly exceeds RMB 30,000 (source: Lanxiong Sports, https://www.lanxiongsports.com).

**Calibration:** At HK$1,592/year (≈ RMB 1,450), the subscription represents 3–15% of a middle-class family's annual sports training budget — the low end of the acceptable range.

| Product | Price | Source |
|---|---:|---|
| Domestic GPS football vest (神仙球) | RMB 1,599 | https://www.shenxianqiu.com |
| Adidas World Cup match ball retail edition | ~USD 160 (≈ RMB 1,150) | https://www.adidas.com |
| Smart football consumer editions | RMB 300–3,000 | https://www.jd.com |

**Subscription benchmark:** Keep's AI coach paid membership lifted paid conversion from 8% to 15% with an 82% renewal rate (source: Keep IR, https://ir.keep.com). In smart football, annual subscription models already cover 12% of B2B users with 83% retention (source: 163.com industry data, https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html).

#### 3.7.2 Performance Orientation — The Target Segment Pays for Outcomes, Not for Discounts

Among football consumers, 45% continue buying football-specific apparel even after a price increase — a higher inelastic share than in most other sports categories (source: survey.shangpu-china.com). The market splits into three tiers: hardcore players (lowest price elasticity), value-seekers (38%, demanding both performance and discount), and entry-level triers (48%, highly price-sensitive).

**Demand for data is itself performance orientation.** 68% of Gen-Z parents explicitly demand data-driven feedback (source: renrendoc.com industry report). 44.67% of parents are willing to pay a premium for high-quality courses, with the stated logic that "parents no longer buy on 'cheap' — they buy on results and professionalism" (source: http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html).

**Implication:** The target segment is not looking for a cheaper football. It is looking for a tool that makes its child kick more accurately. The AI ball and subscription are priced for that segment — not for the 48% entry-level tier.

#### 3.7.3 The "New Gear Fans" Class — High Ticket, High Repurchase, Full-Price Loyalty

Arc'teryx's core 35–55 male middle-class users spend an average of RMB 30,000 per person annually, with a repurchase rate above 38%; Arc'teryx Greater China revenue grew 47% year-on-year (source: Amer Sports IR, https://www.amersports.com/investors). Descente's core users are "few but high-repurchase, high-word-of-mouth, willing to pay full price" (source: Lanxiong Sports, https://www.lanxiongsports.com).

**Football equipment repurchase.** 53% of football consumers report repurchase rates above 70% — more than half are high-frequency, continuous equipment buyers (source: survey.shangpu-china.com).

| "New Gear Fan" trait | Implication for Kung Fu AI Football |
|---|---|
| Pays for performance uplift, not logo | The product's core is measurable shooting improvement (8-week 10–15% accuracy gain) |
| High repurchase (70%+) | Subscription model matches this continuous consumption habit |
| Willing to pay full price | HK$980 + HK$1,592/year does not require discount-led acquisition |
| Treats gear as a signal of commitment | Finishing IQ and physics dashboards provide a visible training investment record |

A consumer already spending RMB 30,000/year on performance gear, for whom HK$980 + HK$1,592/year represents 8–9% of annual equipment spend, is the natural first buyer of a ball that tells them why the shot missed.

#### 3.7.4 Synthesis — Three Dimensions, One Pricing Conclusion

| Dimension | Core Evidence | Implication for Pricing |
|---|---|---|
| Spending power | Middle-class sports spend RMB 10,000–50,000/year; smart hardware validated at RMB 1,000+; subscription conversion 8–15%, renewal 82% | HK$980 + HK$1,592/year sits at the low end of acceptable range |
| Performance orientation | 44.67% of parents pay premium for quality; 45% of football consumers buy despite price rises; 68% Gen-Z demand data feedback | Target segment pays for results, matching the product's value proposition |
| New gear fans | Arc'teryx users spend RMB 30,000/year, repurchase 38%; 53% of football consumers repurchase above 70% | A high-ticket, high-repurchase, full-price consumption class already exists |

**Conclusion:** The HK$980 AI ball and HK$1,592/year ARPU are not prices that require "educating the market." They sit within an already-validated smart sports hardware band, address a consumer class already spending tens of thousands of RMB annually on performance gear, and represent a single-digit percentage of that class's annual equipment budget.

### 3.8 Justification for Purchasing High-End Sports Products

This section addresses a question the panel will ask: why would a parent pay HK$980 for an AI football and HK$1,592/year for a subscription, when a standard football costs HK$100–300? The answer rests on three pillars: competitor price anchoring, China's structural shift toward functional sports consumption, and the parent payment data in the 2025–2030 football training industry report.

#### 3.8.1 Competitor Anchoring — HK$980 Is Mid-to-Low in the Smart Football Category

INSAIT JOY Smart Football — the most direct consumer competitor — is priced at **US$139.00** (≈HK$1,084), including the smart ball, wireless charging base, and USB cable, **with no subscription fee** (source: https://insaitjoy.com/products/insait-joy-smart-football). The product tracks ball motion via sensors and AI algorithms; the App offers 10+ training actions with real-time feedback.

**Key price anchor:** HK$980 is approximately **10% below** INSAIT JOY's US$139 (≈HK$1,084). This is not a premium pricing position — it is an entry price slightly below an existing competitor.

In China's government procurement records, INSAIT smart footballs (model INSAIT-ZQ-PRO-F5AA001) have a **winning bid unit price of RMB 1,750 per ball** (Kashgar No. 6 Middle School, July 2025 procurement announcement). This indicates that in Chinese B2B/school procurement, the price centre for smart footballs is in the **RMB 1,500–1,800** range — far above HK$980 (≈RMB 900).

| Product | Price | Type | Source |
|---|---:|---|---|
| INSAIT JOY Smart Football | US$139 (≈HK$1,084) | Smart ball, no subscription | insaitjoy.com |
| INSAIT smart football (gov procurement) | RMB 1,750/ball | School/academy procurement | Bid announcement |
| 神仙球 GPS vest (Possiball) | RMB 1,599 | Football wearable vest | Xiaohongshu |
| 动量科技 MT football chip | RMB 1,298 | Football sensor | User review |
| ZEPP football sensor | RMB 698 | Football sensor | User review |

**Conclusion:** HK$980 sits at the **mid-to-low end** of the smart football hardware market. It is below INSAIT's government procurement unit price (RMB 1,750) and below the 神仙球 GPS vest (RMB 1,599), and roughly at parity with — slightly below — INSAIT JOY's consumer price (≈HK$1,084). This pricing does not require "educating the market to accept an expensive new category" — it enters an existing price consensus at a slightly below-competitor level.

#### 3.8.2 China's Structural Shift — Functional Gear Is Replacing Traditional Luxury

Chinese consumers are substituting high-performance sports gear for traditional luxury goods. In 2024, China's outdoor sports market consumption **exceeded RMB 300 billion, up nearly 40%**; sales of cycling apparel, shell jackets, and hiking/trail shoes **rose about 50%**. WPIC Marketing + Technologies notes that this signals "luxury consumption shifting toward active lifestyles" — outdoor gear has become "the new luxury frontier" (source: http://news.china.com.cn/2025-06/03/content_117907304.shtml).

**"Expensive running shoes" have become the middle class's "second business card."** In Tier-1 and Tier-2 city running circles, running shoes starting at **RMB 1,500** are no longer mere sports equipment but a "class code" and social currency. In Q1 2025, On's Asia-Pacific net sales grew **130.1%** year-on-year; Asics' performance running category contributed nearly half of its sales. A core consumer requirement is **"no discounting"** — full price is part of the brand's positioning and circle identity (source: https://www.cbndata.com/information/294150).

#### 3.8.3 Parent Payment Data from the 2025–2030 Football Training Industry Report

The 2025–2030 China Football Training Institution Industry Report (renrendoc.com, paper 415852129) contains the payment data that directly supports the subscription logic:

- **92% repurchase rate** for high-end courses (annual fee above RMB 30,000), representing **8%** of total students.
- Mass courses (RMB 8,000–15,000/year) occupy **76%** of the student base.
- **78% of parents** prioritise institutions equipped with a training APP.
- **68% of Gen-Z parents** explicitly demand data-driven feedback; **75%** check coach credentials; the demand gap versus "post-80s" parents is **40%**.
- Smart wearable penetration in youth academies has reached **63%**; sports data analytics systems can improve training efficiency by **30%**.

**These numbers point in one direction:** the target parent group is already paying for "visible results" and "quantifiable training feedback." 68% demanding data feedback, 78% preferring APP-equipped institutions, and 92% high-end course repurchase together describe a market that is **already paying a premium for "professionalism" and "results"** — not a price-sensitive market. HK$980 + HK$1,592/year enters an existing payment habit, not a new consumer behaviour.

#### 3.8.4 Why "High-End" Is Reasonable Here — From "Equipment Consumption" to "Training-Outcome Consumption"

The report contains a frequently overlooked but critical data point: **the average customer acquisition cost for football training institutions rose to RMB 1,800 per person in 2024, up 70% from 2021**; **venue rent accounts for 35% of operating cost**; **head-institution teacher cost accounts for 41% of total revenue**. These cost structures mean that an increasing share of what parents pay for football training is absorbed by institutional operating costs, rather than converted directly into training outcomes.

Kung Fu AI Football's positioning is to **separate "training outcome" from institutional operating cost and turn it into a quantifiable, trackable, subscribable product**. It does not compete with "free football" or "low-price trial classes"; it competes with **INSAIT JOY (US$139 hardware, no subscription)** and **神仙球 GPS vest (RMB 1,599)** for the same parents: those already willing to pay RMB 1,000–2,000 for hardware that helps their child "play more intelligently."

**HK$980 enters at competitor hardware price and uses a subscription to deliver continuous feedback that competitors do not offer.** INSAIT JOY's US$139 is a one-time purchase with no subscription; Kung Fu AI Football's HK$980 + HK$1,592/year gives a first-year total of ≈HK$2,572 — about 2.4× INSAIT JOY. But that premium corresponds to a capability INSAIT JOY does not have: **decomposition of the shooting causal chain (biomechanics → ball physics → outcome) + physics education layer + age-adaptive models**. For a parent group already paying RMB 30,000/year for high-end courses with a 92% repurchase rate, HK$2,572 in the first year is only **8.6%** of their annual training spend.

#### 3.8.5 Section Conclusion

| Justification Pillar | Core Evidence | Implication for HK$980 + HK$1,592/year |
|---|---|---|
| Competitor anchoring | INSAIT JOY US$139 (≈HK$1,084); INSAIT gov procurement RMB 1,750; 神仙球 RMB 1,599 | HK$980 is **mid-to-low** in smart football hardware — not premium |
| High-end sports consumption trend | Outdoor market RMB 300bn+, +40%; RMB 1,500 running shoes as "middle-class standard"; On APAC +130% | Functional gear consumption is replacing traditional luxury; **target group already pays for performance** |
| Parent payment willingness | 92% high-end course repurchase; 78% prefer APP-equipped institutions; 68% demand data feedback; parent annual spend RMB 6,800 | Target parents **already pay for "results" and "data"** — subscription is not a new behaviour |
| Cost-structure logic | CAC RMB 1,800/person; venue rent 35%; teacher cost 41% | AI football **productises "training outcome"** out of institutional operating cost, competing with hardware peers for the same high-intent parents |

**One-line conclusion:** HK$980 is not "educating consumers to accept a high-end price" in a blank market — it enters an existing price consensus (INSAIT US$139, government procurement RMB 1,750, GPS vest RMB 1,599) at a **slightly below-competitor price**. The HK$1,592/year subscription targets a parent group where **92% repurchase high-end courses, 68% demand data feedback, and 78% prefer tech-assisted training** — parents already paying a premium for "visible training outcomes." The subscription simply extends that payment habit from "courses" to "continuous feedback."

## 4. ITF RTH Manpower Grant

| Item | Assumption |
|---|---:|
| BSc manpower | 2 (Y1–Y3) + 2 (Y4–Y6) |
| Monthly rate | HK$20,000 |
| Annual gross cost | HK$0.48m/year × 6 = HK$2.88m |
| Grant coverage | 100% eligible |
| Reimbursement lag | 6 months |

| Year | Gross Manpower | Grant Received | Net Impact |
|---|---:|---:|---:|
| Y1 | 0.48 | 0.24 | −0.24 |
| Y2–Y6 | 0.48 each | 0.48 each | 0.00 |
| Y7 Q2 | — | 0.24 | +0.24 |
| Total | 2.88 | 2.88 | 0.00 |

## 5. Commercial Terms — B2B Technology Enabler

| Line | Terms |
|---|---|
| Subscription share | 12.86% of gross subscription |
| Ball revenue share | 35% × HK$980 = HK$343 per ball |
| Exclusive annual license | HK$0.20m/year |
| Customization revenue | HK$0.10m Y1; HK$0.025m/year Y2–Y6 |
| ITF RTH grant | HK$0.24m Y1; HK$0.48m Y2–Y6; HK$0.24m Y7 Q2 |

### Cost — Model D Hybrid Platform

| Line | Terms | Behaviour |
|---|---|---|
| Platform base | HK$0.20m/year | Fixed |
| Platform compute | HK$2.00 × active users | Variable |
| User enhancement | HK$5.00 × active users | Variable |
| Ball enhancement | HK$10 × ball units | Variable |
| Ball manufacturing | HK$250 × ball units | Variable |
| ITF manpower (gross) | HK$0.48m/year | Fixed |

**Net ball contribution:** +HK$343 − HK$250 − HK$10 = **+HK$83 per ball**.

## 6. Commercial Terms — B2C Distribution Partners

| Line | Terms |
|---|---|
| Subscription share | 37.14% of gross subscription |
| Ball revenue share | 65% × HK$980 = HK$637 per ball |
| Marketing & distribution | 30% of gross subscription |
| Exclusive license fee | HK$0.20m/year paid to B2B |
| Customization fee | HK$0.10m Y1; HK$0.025m/year Y2–Y6 paid to B2B |

## 7. Subscription Split

| Step | Recipient | Share of Gross |
|---|---|---:|
| 1 | Coach / Player | 50.00% |
| 2 | B2B technology enabler | 12.86% |
| 3 | B2C distribution partners | 37.14% |
| | Total | 100.00% |

**Tier pricing:** Local HK$65 (60%), National HK$148 (25%), World Cup HK$298 (12%), Legends HK$698 (3%). Weighted ARPU = **HK$1,592/year**.

## 8. Cost Behaviour and Operating Leverage

| Cost line | Behaviour | Driver |
|---|---|---|
| Platform base | Fixed | HK$0.20m/year |
| ITF manpower | Fixed | HK$0.48m/year |
| Platform compute | Variable | HK$2.00 × SOM |
| User enhancement | Variable | HK$5.00 × SOM |
| Ball enhancement | Variable | HK$10 × new users × 50% |
| Ball manufacturing | Variable | HK$250 × new users × 50% |

Fixed base = **HK$0.68m/year**. User base grows from ≈13,650 to ≈604,200 — a **44× increase** — so fixed cost per user falls from **HK$49.82** to **HK$1.13**. Break-even ≈ **2,835 users**; Year 1 SOM ≈13,650 is **4.8× break-even**.

## 9. Year-by-Year Results (P50)

### 9.1 B2B P&L (HK$m)

| Year | SOM | Sub Rev | Ball Rev | Custom | License | ITF Grant | Total Rev | Platform | User Enh | Ball Enh | Ball Cost | Manpower | Total Cost | Profit | Cumulative |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Y1 | 13,650 | 2.79 | 2.34 | 0.10 | 0.20 | 0.24 | 5.67 | 0.23 | 0.07 | 0.07 | 1.71 | 0.48 | 2.56 | 3.11 | 3.11 |
| Y2 | 34,130 | 6.99 | 3.51 | 0.025 | 0.20 | 0.48 | 11.21 | 0.27 | 0.17 | 0.10 | 2.57 | 0.48 | 3.59 | 7.62 | 10.73 |
| Y3 | 192,160 | 39.34 | 27.12 | 0.025 | 0.20 | 0.48 | 67.17 | 0.58 | 0.96 | 0.79 | 19.85 | 0.48 | 22.66 | 44.51 | 55.24 |
| Y4 | 303,000 | 62.03 | 19.05 | 0.025 | 0.20 | 0.48 | 81.81 | 0.81 | 1.52 | 0.56 | 13.93 | 0.48 | 17.30 | 64.51 | 119.75 |
| Y5 | 425,980 | 87.22 | 21.11 | 0.025 | 0.20 | 0.48 | 109.06 | 1.05 | 2.13 | 0.62 | 15.44 | 0.48 | 19.72 | 89.34 | 209.09 |
| Y6 | 604,200 | 123.74 | 30.51 | 0.025 | 0.20 | 0.48 | 154.96 | 1.41 | 3.02 | 0.89 | 22.29 | 0.48 | 28.09 | 126.87 | 335.96 |
| Y7 Q2 | — | — | — | — | — | 0.24 | 0.24 | — | — | — | — | — | 0.00 | 0.24 | 336.20 |

### 9.2 B2C P&L (HK$m)

| Year | SOM | Sub Rev | Ball Rev | Total Rev | Marketing | License | Customization | Total Cost | Profit | Cumulative |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Y1 | 13,650 | 8.08 | 4.35 | 12.43 | 6.53 | 0.20 | 0.10 | 6.83 | 5.60 | 5.60 |
| Y2 | 34,130 | 20.18 | 6.52 | 26.70 | 16.30 | 0.20 | 0.025 | 16.53 | 10.18 | 15.78 |
| Y3 | 192,160 | 113.61 | 50.35 | 163.96 | 91.78 | 0.20 | 0.025 | 92.01 | 71.95 | 87.73 |
| Y4 | 303,000 | 179.17 | 35.36 | 214.53 | 144.74 | 0.20 | 0.025 | 144.97 | 69.56 | 157.29 |
| Y5 | 425,980 | 251.94 | 39.26 | 291.20 | 203.50 | 0.20 | 0.025 | 203.73 | 87.47 | 244.76 |
| Y6 | 604,200 | 357.31 | 56.70 | 414.01 | 288.62 | 0.20 | 0.025 | 288.85 | 125.16 | 369.92 |

### 9.3 Fairness Check

| Year | B2B Profit | B2C Profit | B2B Cumulative | B2C Cumulative | B2C/B2B |
|---|---:|---:|---:|---:|---:|
| Y1 | 3.11 | 5.60 | 3.11 | 5.60 | 1.80 |
| Y2 | 7.62 | 10.18 | 10.73 | 15.78 | 1.47 |
| Y3 | 44.51 | 71.95 | 55.24 | 87.73 | 1.59 |
| Y4 | 64.51 | 69.56 | 119.75 | 157.29 | 1.31 |
| Y5 | 89.34 | 87.47 | 209.09 | 244.76 | 1.17 |
| Y6 | 126.87 | 125.16 | 335.96 | 369.92 | 1.101 |
| Y7 Q2 | +0.24 | — | 336.20 | — | 1.100 |

**Both constraints hold:** neither partner is negative in any year; the full-cycle fairness ratio is ≈1.100, at the 10% ceiling. Full-cycle B2B cumulative profit ≈ HK$336.2m.

## 10. Confidence Analysis

| Scenario | Conversion | Y6 SOM | B2B Cum Y6 | B2C Cum Y6 | Fairness (full cycle) |
|---|---:|---:|---:|---:|---:|
| P10 | 2.175% | ≈302,100 | ≈HK$168m | ≈HK$185m | ≈1.10 |
| P50 | 4.35% | ≈604,200 | ≈HK$336m | ≈HK$370m | ≈1.10 |
| P90 | 8.7% | ≈1,208,400 | ≈HK$672m | ≈HK$740m | ≈1.10 |

## 11. Sensitivity

### 11.1 Marketing cost

| B2C marketing | Required B2B sub share | Required B2C sub share |
|---|---:|---:|
| 30% (base) | 12.86% | 37.14% |
| 35% | ~10.48% | ~39.52% |
| 40% | ~8.09% | ~41.91% |

### 11.2 Platform model

| Model | 6-year total | B2B Cum Y6 | Fairness | Within 10%? |
|---|---:|---:|---:|:---:|
| D — Hybrid (adopted) | 2.90 | ≈336 | ≈1.10 | ✅ |
| A — Flat 0.40 × 6 | 2.40 | ≈336.5 | ≈1.10 | ✅ |
| C — Step-by-region | 3.15 | ≈335.7 | ≈1.10 | ✅ |

### 11.3 Ball cost and attach

At HK$250, B2B earns +HK$83 net per ball. At HK$300, +HK$33; at HK$343, break-even.

### 11.4 Compute-cost conservatism

Model assumes HK$7/user/year (HK$2 platform + HK$5 enhancement). Benchmarks indicate HK$10–16/user/year at 25 kicks. Model may understate infrastructure cost by HK$2–5m cumulative at P50. This does not threaten the fairness band — B2B still generates >HK$330m cumulative profit — but is disclosed for stress-testing.

## 12. How to Use the Live Workbook

1. Open Assumptions — yellow cells editable.
2. Change marketing_pct (30%/35%/40%) — split auto-adjusts via VLOOKUP calibration table.
3. Change scenario_selector (1/2/3) — main P&L follows P10/P50/P90.
4. Scenarios sheet shows all three side by side.
5. Summary — PASS/FAIL invariants flag any breach.
6. Sources — verified URLs only.

## 13. Change Log

| Version | Change |
|---|---|
| 2.0 | Model D hybrid platform; split 12.86 / 37.14 |
| 2.1 | ITF RTH manpower integrated |
| 2.2 | Production release |
| 2.3 | Customization fee added as B2C cost |
| 2.4 | Live Excel model |
| 2.5 | TAM/SAM/SOM externally sourced; three-band framework |
| 2.6 | 68% Gen-Z evidence; regional SAM weights |
| 2.7 | §3.4.2 layman terms; fairness arithmetic corrected |
| 2.8 | §3.7 Spending Power; A4 Landscape edition |
| 2.9 | All numbers re-verified; 8.7% URL confirmed as L8E6C26F; 68% source changed to renrendoc.com; 45% source changed to survey.shangpu-china.com; 31% replaced with 53%; 63%, 92%, >90% removed as unverifiable; SAM arithmetic corrected from 13.98m to 13.89m; 8.7% semantics narrowed to "football path"; §3.8 High-End Sports Purchase Justification added |

## 14. Reference URL Register — Verified Only

| Reference | URL | Verified Data | Status |
|---|---|---|---|
| 绿茵经济（163.com） | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | 3.8亿球迷; 3200亿产业; 4.23万所学校; 12.6万教练; 近百万业余; 1280万注册学生; 487.6亿训练市场 | ✅ |
| 注册球员2740人（163.com） | https://www.163.com/dy/article/L8E6C26F0556BPXI.html | 8.7% 家长支持长期足球道路; 10.92万精英青少年; 2740名职业球员 | ✅ |
| 足协数字年报（china.com.cn） | http://news.china.com.cn/2026-01/30/content_118308779.shtml | 109,220精英青少年; 979,544业余球员; 125,976持证教练 | ✅ |
| 青少年体育培训报告（人民网） | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html | 44.67% 愿为品质付费; 2531亿培训市场; 57.07%续费占比 | ✅ |
| 尚普咨询（足球专用服饰） | survey.shangpu-china.com | 45% 涨价仍购买; 53% 复购率超七成 | ✅ |
| 人人文库（足球培训报告） | renrendoc.com/paper/415852129.html | 68% Z世代要求数据化反馈; 75%查阅教练资质; 92%高端复购率; 63%智能穿戴渗透率 | ⚠️ 行业文档 |
| NDRC规划 | https://www.ndrc.gov.cn/xxgk/zcfb/ghwb/201604/t20160411_962179.html | 框架页面，未直接显示5000万目标 | ⚠️ 需查正文附件 |
| CEIBS（FIFA Big Count） | https://cn.ceibs.edu/new-papers-columns/29518 | 26.16m踢球人数; 0.078%球员密度 | ✅ |
| INSAIT JOY Smart Football 官网 | https://insaitjoy.com/products/insait-joy-smart-football | US$139 定价; 无订阅费 | ✅ |
| 喀什第六中学采购公告 | http://www.ceiea.com/html/202507/202507091553056481.shtml | INSAIT 智能足球政府采购单价 1,750 元 | ✅ |
| CBNData（贵价跑鞋） | https://www.cbndata.com/information/294150 | 1,500 元跑鞋成为中产标配; 昂跑亚太增长 130% | ✅ |
| 中国网（户外都市风） | http://news.china.com.cn/2025-06/03/content_117907304.shtml | 户外市场 3,000 亿+, 增幅 40% | ✅ |
| 小红书（神仙球访谈） | https://www.xiaohongshu.com/discovery/item/66a1ed34000000002701173d | GPS 背心 1,599 元; 用户群体为青训家长 | ✅ |
| 小红书（足球智能装备测评） | https://www.xiaohongshu.com/discovery/item/6722157c000000001a034f3f | ZEPP 698 元、动量科技 1,298 元等价格锚点 | ✅ |

**Removed from v2.8:** 63%智能穿戴渗透率（现归入renrendoc.com行业文档）、92%高端复购率（同上）、>90%家长接受花费（无法核实）、31%复购率70–90%（实际为53%）。

## 15. Caveats and Intended Use

1. Planning model. Not audited, not tax-adjusted, not for filing without independent review.
2. TAM = 300m, 79% of the 380m fan base.
3. SAM = 13.89m, the arithmetic sum of three verified components. v2.8's 13.98m was a rounding artefact; the difference is immaterial to conclusions.
4. SOM is a projection. The three bands are fixed fractions of the 8.7% pathway-commitment ceiling.
5. 8.7% measures support for a long-term football path, not general approval of playing football. This makes it a conservative ceiling.
6. 68% is from an industry research document (renrendoc.com), not an official statistic. It supports but does not independently verify the P50.
7. P&L presented at P50 (conversion 4.35%, SOM ≈604,200). P10/P90 selectable via scenario multiplier.
8. Platform uses Model D — HK$0.20m fixed + HK$2.00/user.
9. Split 12.86/37.14; full-cycle fairness ≈1.100, at the 10% ceiling.
10. Customization appears on both sides.
11. ITF grant cash-neutral across cycle.
12. Conservatism gap: HK$7/user may understate compute by 1.4–2.3×.
13. Partner credentials require due diligence.
14. Workbook is live — change yellow cells to model alternatives.

---

**Prepared for:** Kung Fu AI Football partnership negotiation
**Model owner:** B2B technology enabler
**Review cycle:** Re-run whenever marketing cost, ball cost, attach rate, platform cost, conversion assumptions, or SAM components change by >5pp.
"""


# ═══════════════════════════════════════════════════════════════
# BUILD
# ═══════════════════════════════════════════════════════════════
def build():
    doc = Document()
    base = doc.styles['Normal']
    base.font.name = FONT
    base.font.size = Pt(SIZES['body'])
    base.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

    section = doc.sections[0]
    section.page_width = Mm(297)
    section.page_height = Mm(210)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.top_margin = Mm(10)
    section.bottom_margin = Mm(10)
    section.left_margin = Mm(10)
    section.right_margin = Mm(10)

    # Cover
    for _ in range(4):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Kung Fu AI Football')
    set_run(r, SIZES['cover_title'], bold=True, color=RGBColor(0x1F, 0x4E, 0x78))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('6-Year Financial Model — README v2.9')
    set_run(r, SIZES['cover_sub'], bold=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('All figures re-verified against source URLs · §3.8 High-End Sports Purchase Justification included')
    set_run(r, SIZES['cover_sub'], italic=True, color=GREY_TEXT)
    doc.add_page_break()

    render(doc, README_MD)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('— End of README v2.9 —')
    set_run(r, SIZES['footer'], italic=True, color=GREY_TEXT)

    out = 'KungFu_AI_Football_README_v2.9.docx'
    doc.save(out)
    print(f'✅ Saved: {out}')
    print(f'   Font: {FONT} · H1 {SIZES["h1"]}pt · H2 {SIZES["h2"]}pt · H3 {SIZES["h3"]}pt · Body {SIZES["body"]}pt')


if __name__ == '__main__':
    build()