"""
README v3.0 → DOCX
Font: Times New Roman · Body 10pt · H1 14pt · H2 12pt · H3 11pt · Tables 9pt
Output: KungFu_AI_Football_README_v3.0.docx
Run: python generate_readme_v3_0.py
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

FONT = 'Times New Roman'
SIZES = {
    'cover_title': 20, 'cover_sub': 12,
    'h1': 14, 'h2': 12, 'h3': 11, 'h4': 10,
    'body': 10, 'table': 9, 'footer': 8,
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
    for part in pattern.split(text):
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
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    if level == 1:
        p.paragraph_format.keep_with_next = True
    set_run(p.add_run(text), size, bold=True)


def add_body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.1
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
            p.paragraph_format.left_indent = Mm(5)
            p.paragraph_format.space_after = Pt(1)
            add_inline(p, text, SIZES['body'])
            i += 1
            continue
        add_body(doc, line.strip())
        i += 1


# ═══════════════════════════════════════════════════════════════
# README v3.0 CONTENT
# ═══════════════════════════════════════════════════════════════
README_MD = r"""
# README v3.0 — Kung Fu AI Football 6-Year Financial Model

**Base case:** P50 (mean conversion 4.35% → SOM ≈ 604,200 by Year 6)
**Currency:** HK$ million unless stated otherwise
**Version:** 3.0 (concise §3.7 justification + §3.8 evidence table; Times New Roman 10pt)

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

### 3.3 SAM = 13.89m

| Component | Value | Source URL | Explanation |
|---|---:|---|---|
| Social amateur players | 979,544 | http://news.china.com.cn/2026-01/30/content_118308779.shtml | "业余球员规模达979544人" |
| Elite youth registered players | 109,220 | Same CFA 2025 report | "注册精英青少年运动员达109220人" |
| Campus football registered students | 12,800,000 | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | "注册学生人数1280万人" |
| SAM (arithmetic sum) | 13,888,764 ≈ 13.89m | — | 979,544 + 109,220 + 12,800,000 |

**Correction from v2.8:** v2.8 stated SAM = 13.98m. The arithmetic sum of the three verified components is 13.89m. The 0.09m difference was a rounding artefact. v3.0 adopts 13.89m as the auditable SAM.

### 3.4 SOM — Three Conversion Bands

The rate is anchored to **8.7% — the share of Chinese parents willing to support their child's long-term football path** (source: https://www.163.com/dy/article/L8E6C26F0556BPXI.html).

| Scenario | Formula | Conversion Rate | Share of Ceiling | SOM (Y6) |
|---|---|---:|---:|---:|
| P10 (Conservative) | 8.7 / 4 | 2.175% | 25% | ≈302,100 |
| P50 (Base) | 8.7 / 2 | 4.35% | 50% | ≈604,200 |
| P90 (Optimistic) | 8.7 / 1 | 8.7% | 100% | ≈1,208,400 |

**Statistical note:** Consistent with a lognormal distribution — P50/P10 = 2.0, P90/P50 = 2.0, σ_log ≈ 0.54.

**Relationship to prior model:** The prior SOM of 326,520 corresponds to a conversion of 2.34% — just above the P10 band (2.175%). The prior model was already operating close to the conservative bound.

#### 3.4.1 The 8.7% Benchmark — Corrected Semantics

> "截至2025年底，中国注册精英青少年球员10.92万人……家长愿意长期支持孩子踢球的比例为8.7%，西班牙是42.3%。"

The 8.7% measures **pathway commitment** (support for a long-term football path), not general approval of playing football. This makes it a conservative ceiling for paid subscription conversion.

#### 3.4.2 Why 4.35% Is a Sensible Middle Estimate

**Step 1 —** Out of 100, about 9 (8.7%) say they would support their child's long-term football path.
**Step 2 —** Among Gen-Z parents, 68% explicitly demand data-driven feedback. Applying this to the 8.7% pathway-committed group gives ≈5.9% as a demand-qualified ceiling.
**Step 3 —** Not everyone who says they want it pays. The model assumes only about 4–5 in every 6 interested parents actually subscribe. That gives **4.35% — the P50**.

**Caveat:** The 68% × 8.7% multiplication assumes independence. Positive correlation would make the true demand-qualified ceiling higher, making 4.35% more conservative than it looks.

### 3.5 SOM by Year (P50)

| Year | Regions Active | SAM (m) | Conversion | SOM |
|---|---|---:|---:|---:|
| Y1 | HK & South China | 5.55 | 0.246% | ≈13,650 |
| Y2 | HK & South China | 5.55 | 0.615% | ≈34,130 |
| Y3 | + East China | 8.95 | 2.147% | ≈192,160 |
| Y4 | + Southwest | 11.26 | 2.691% | ≈303,000 |
| Y5 | + North | 12.80 | 3.328% | ≈425,980 |
| Y6 | + Northwest & Northeast | 13.89 | 4.35% | ≈604,200 |

#### 3.5.1 Regional Allocation Weights

| Region | Weight | Regional SAM | Rationale |
|---|---:|---:|---|
| HK & South China | 40.0% | 5.55m | FIFA Talent Academy; PSG HK girls' academy; Shenzhen youth centre |
| East China | 24.4% | 3.40m | Suzhou regional youth centre; dense school network |
| Southwest | 16.7% | 2.32m | Chengdu regional youth centre |
| North | 11.1% | 1.54m | Beijing national youth training centre |
| Northwest | 4.4% | 0.61m | Xi'an regional youth centre |
| Northeast | 3.3% | 0.46m | Dalian regional youth centre |
| Total | 100% | 13.89m | — |

### 3.6 Total Football Market Capital — Verified

| Metric | Value | Source URL | Verified |
|---|---:|---|---|
| Football industry total scale (2025) | RMB 320bn | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | ✅ |
| Football training market (2025) | RMB 48.76bn | Same | ✅ |
| Fans | 380m | Same | ✅ |
| Football-specialist schools | 42,300 | Same | ✅ |
| Campus football registered students | 12.8m | Same | ✅ |
| 2026 forecast | RMB 54.32bn | Same | ✅ |
| 2030 forecast | RMB 82.84bn | Same | ✅ |

**At P50 SOM ≈604,200, subscription revenue is ≈HK$9.6bn/year — roughly 2% of the RMB 48.76bn training market.**

## §3.7 Justification for Purchasing High-End Sports Products

**Core question:** Why would a parent pay HK$980 for an AI football plus HK$1,592/year for a subscription, when a standard football costs a fraction of that?

### 3.7.1 Spending Power Is Migrating Toward Health and Performance — Not Away from Premium Spending

China's affluent households are reallocating consumption from material display to health, sport and experience. The Hurun 2026 Quality of Life Report — surveying 470 million-RMB-plus individuals with average household assets of RMB 61 million — finds that **"health and wellness" has risen to the second-largest planned spending category**, while service experiences (travel, health, education) are set to rise **12%**, against a **10% cut in material luxury** such as watches and jewellery. Respondents already exercise **more than 2 hours per week**, with hiking/trail and running the preferred activities.

The same shift is visible in youth sport. The 2025 Youth Sports Training Industry Report puts the market at **RMB 253.1 billion**, with renewal revenue rising from 48% to **57.07%** of total income — evidence that parents are consolidating spend around providers they trust rather than chasing discounts. Crucially, **44.67% of parents explicitly say they will pay more for high-quality courses**, and the report's own conclusion is that "parents no longer buy on 'cheap' — they buy on **results and professionalism**".

Within football specifically, the 2025–2030 Football Training Institution Report segments the market by price sensitivity: **92% of high-end course customers (annual fee above RMB 30,000) renew**, although they represent only **8% of students**; mass courses (RMB 8,000–15,000) hold **76% of the base but renew at only 55%**. And **68% of Gen-Z parents demand data-driven feedback**, while **75% check coach credentials**.

### 3.7.2 Price Anchoring: HK$980 Sits at the Lower End of Smart-Football Hardware

The relevant comparison set is not a standard ball — it is connected football hardware:

| Product | Price | Model |
|---|---:|---|
| INSAIT JOY Smart Football | US$139 (≈HK$1,084), no subscription | Consumer, one-off |
| INSAIT smart football — China government procurement | RMB 1,745/unit | B2B/school procurement |
| adidas Leagues Cup 26 Pro Ball | US$170 (standard match ball, no sensors) | Retail benchmark |
| Catapult One (GPS vest) | ≈US$179/year, membership only | Consumer subscription; hardware tied to membership |

HK$980 (≈RMB 900) is therefore **below** the INSAIT consumer ball, **roughly half** the Chinese government procurement price for an INSAIT smart football, and **below Catapult One's annual membership** — which does not even include a ball. Catapult One is also instructive as a business model: it is **membership-only, cannot be purchased outright, and one pod serves one player under one membership**. A subscription-based connected-training product is therefore already an accepted consumer format in this category.

### 3.7.3 Why HK$1,592/Year Is a Small Share of the Target Household's Existing Football Spend

The football training market alone is **RMB 48.76 billion in 2025**, inside a **RMB 320 billion** football industry, rising to RMB 54.32 billion in 2026 and RMB 82.84 billion by 2030. High-end football course families already spend **above RMB 30,000/year**, and **92% of them renew**. Against that base, HK$1,592/year is roughly **5% of an existing high-end course budget** — a marginal add-on for continuous, quantified feedback, not a new category of expenditure.

### 3.7.4 Section Conclusion

| Justification pillar | Evidence | Implication for HK$980 + HK$1,592/year |
|---|---|---|
| Spending power migration | Health spending now No.2 category; service experience +12%; >2 hrs exercise/week | Target households are **increasing** health/performance spend |
| Parent willingness to pay | 44.67% pay more for quality; 57.07% renewal share; 68% demand data | Parents already pay for **results and data**, not price |
| High-end football spend | 92% renewal on RMB 30k+ courses; only 8% of students | Subscription ≈ **5%** of existing high-end budget |
| Hardware anchoring | HK$980 below INSAIT US$139, half of RMB 1,745 procurement, below Catapult One US$179/yr | HK$980 is **mid-to-low** in connected football hardware |
| Subscription precedent | Catapult One is membership-only, one pod per player | Recurring payment is an **accepted format** in this category |

**In one sentence:** HK$980 enters an existing connected-football price band **below** INSAIT's consumer and procurement prices and **below** Catapult One's annual membership, while HK$1,592/year asks parents who already renew RMB 30,000+ courses at 92% and 44.67% of whom will pay more for quality to extend roughly **5% of that budget** from "courses" to "continuous, quantified feedback."

## §3.8 Evidence Table — Every Data Point and Its Source

This table lists every fact used in §3.7 against its exact source URL. It exists so that a panel reviewer can verify any individual claim without searching.

| Data / Fact | URL Source |
|---|---|
| Health & wellness is now the **No.2 planned spending category** for China's affluent households | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 |
| Service experience (travel, health, education) spending planned to **rise 12%**; material luxury (watches, jewellery) planned to **fall 10%** | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 |
| Affluent respondents exercise **more than 2 hours per week**; hiking/trail and running are the preferred activities | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 |
| Hurun survey base: **470 individuals with RMB 1 million+ assets**, average household assets **RMB 61 million** | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 |
| Youth sports training market size: **RMB 253.1 billion** | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html |
| Renewal revenue share rose from 48% to **57.07%** of total income | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html |
| **44.67% of parents** explicitly say they will pay more for high-quality courses | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html |
| Report conclusion: parents buy on **"results and professionalism"**, not on "cheap" | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html |
| Football industry total scale (2025): **RMB 320 billion** | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html |
| Football training market (2025): **RMB 48.76 billion** | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html |
| Football training market forecast: **RMB 54.32 billion (2026)**, **RMB 82.84 billion (2030)** | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html |
| Chinese football fans: **380 million** | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html |
| Football-specialist schools: **42,300**; registered students: **12.8 million** | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html |
| **92% renewal** among high-end course customers (annual fee above RMB 30,000) | https://www.renrendoc.com/paper/415852129.html |
| High-end customers represent **8%** of total students | https://www.renrendoc.com/paper/415852129.html |
| Mass courses (RMB 8,000–15,000) hold **76%** of the student base; renewal rate **55%** | https://www.renrendoc.com/paper/415852129.html |
| **68% of Gen-Z parents** demand data-driven feedback; **75%** check coach credentials | https://www.renrendoc.com/paper/415852129.html |
| adidas Leagues Cup 26 Pro Ball retail price: **US$170** | https://www.adidas.com/us/leagues-cup-26-pro-ball/KE5154.html |
| INSAIT JOY Smart Football: **US$139**, no subscription | https://insaitjoy.com/products/insait-joy-smart-football |
| INSAIT smart football (INSAIT-ZQ-PROF5AA001) government procurement unit price: **RMB 1,745/unit** | https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/ |
| INSAIT athlete sensor (INSAIT-CGQx10-001) same procurement: **RMB 4,798/unit** | https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/ |
| Catapult One: **≈US$179/year, membership only**, cannot be purchased outright | https://soccerstories.net/catapult-one-review/ |
| Catapult One model: **one pod per player per membership** | https://soccerstories.net/catapult-one-review/ |

**Status note:** The Hurun, People's Daily Shanghai, 163.com, adidas, INSAIT and xianlanzhaobiao sources are official/primary. The renrendoc.com report is an **industry research document**, not an official statistic — its 68%, 75%, 92%, 76% and 55% figures should be cited as "industry estimate" in panel materials. The Catapult One review is a **secondary product review**, used only as a price and business-model reference.

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

**Both constraints hold:** neither partner is negative in any year; the full-cycle fairness ratio is ≈1.100, at the 10% ceiling.

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

### 11.2 Ball cost and attach

At HK$250, B2B earns +HK$83 net per ball. At HK$300, +HK$33; at HK$343, break-even.

### 11.3 Compute-cost conservatism

Model assumes HK$7/user/year (HK$2 platform + HK$5 enhancement). Benchmarks indicate HK$10–16/user/year at 25 kicks. Model may understate infrastructure cost by HK$2–5m cumulative at P50. Disclosed for stress-testing.

## 12. How to Use the Live Workbook

1. Open Assumptions — yellow cells editable.
2. Change marketing_pct (30%/35%/40%) — split auto-adjusts via VLOOKUP calibration table.
3. Change scenario_selector (1/2/3) — main P&L follows P10/P50/P90.
4. Scenarios sheet shows all three side by side.
5. Summary — PASS/FAIL invariants flag any breach.

## 13. Change Log

| Version | Change |
|---|---|
| 2.8 | §3.7 Spending Power; A4 Landscape edition |
| 2.9 | All numbers re-verified; 8.7% URL confirmed; 68% source changed to renrendoc; 45% source changed to shangpu; 31% replaced with 53%; 63%, 92%, >90% removed as unverifiable; SAM corrected from 13.98m to 13.89m; §3.8 High-End Sports Purchase Justification added |
| 3.0 | §3.7 condensed to one page using only verifiable primary/secondary sources; §3.8 converted to a Data/Fact + URL evidence table; fonts changed to Times New Roman 10pt body / H1 14pt / H2 12pt; unverifiable home-page citations removed |

## 14. Reference URL Register

| Reference | URL | Data used | Status |
|---|---|---|---|
| Hurun Quality of Life Report | https://www.hurun.net/zh-cn/info/detail?num=CSX6BTOMEWY4 | Health spending No.2; +12% service experience; >2 hrs exercise | ✅ Primary |
| People's Daily Shanghai — Youth Sports Training Report | http://sh.people.com.cn/n2/2026/0127/c176739-41484959.html | 44.67% pay more for quality; RMB 253.1bn market; 57.07% renewal | ✅ Primary |
| 163.com — 3200亿足球产业 | https://www.163.com/dy/article/KKOGPQ4C0534AAOK.html | 3.8亿 fans; 3200亿 industry; 487.6亿 training; 4.23万 schools; 1280万 students | ✅ Primary |
| Football Training Institution Report | https://www.renrendoc.com/paper/415852129.html | 92% high-end renewal; 8% of students; 76% mass base; 55% mass renewal; 68% Gen-Z data; 75% coach credentials | ⚠️ Industry research |
| adidas Leagues Cup 26 Pro Ball | https://www.adidas.com/us/leagues-cup-26-pro-ball/KE5154.html | US$170 retail price | ✅ Primary |
| INSAIT JOY Smart Football | https://insaitjoy.com/products/insait-joy-smart-football | US$139, no subscription | ✅ Primary |
| INSAIT government procurement | https://www.xianlanzhaobiao.com/news-e9385871eeaa543d921d893a2a66a7f5/ | RMB 1,745/unit smart ball; RMB 4,798/unit sensor | ✅ Primary |
| Catapult One review | https://soccerstories.net/catapult-one-review/ | ≈US$179/year, membership only, one pod per player | ⚠️ Secondary review |
| 注册球员2740人（163.com） | https://www.163.com/dy/article/L8E6C26F0556BPXI.html | 8.7% pathway commitment; 10.92万 elite youth | ✅ Primary |
| 足协数字年报（china.com.cn） | http://news.china.com.cn/2026-01/30/content_118308779.shtml | 109,220 elite youth; 979,544 amateur; 125,976 coaches | ✅ Primary |

**Removed from earlier versions:** cbndata.com, lanxiongsports.com, 36kr.com, iresearch.com.cn, shenxianqiu.com, amersports.com, ir.keep.com, jd.com — all were cited as home pages without a specific article URL, and could not be verified.

## 15. Caveats and Intended Use

1. Planning model. Not audited, not tax-adjusted, not for filing without independent review.
2. TAM = 300m, 79% of the 380m fan base.
3. SAM = 13.89m, the arithmetic sum of three verified components.
4. SOM is a projection. The three bands are fixed fractions of the 8.7% pathway-commitment ceiling.
5. 8.7% measures support for a long-term football path, not general approval of playing football.
6. 68%, 92%, 76% and 55% are from an industry research document (renrendoc.com), not an official statistic. They support but do not independently verify the P50.
7. P&L presented at P50 (conversion 4.35%, SOM ≈604,200). P10/P90 selectable via scenario multiplier.
8. Platform uses Model D — HK$0.20m fixed + HK$2.00/user.
9. Split 12.86/37.14; full-cycle fairness ≈1.100, at the 10% ceiling.
10. ITF grant cash-neutral across cycle.
11. Conservatism gap: HK$7/user may understate compute by 1.4–2.3×.
12. Partner credentials require due diligence.
13. Workbook is live — change yellow cells to model alternatives.

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
    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Kung Fu AI Football'), SIZES['cover_title'], bold=True,
            color=RGBColor(0x1F, 0x4E, 0x78))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('6-Year Financial Model — README v3.0'), SIZES['cover_sub'], bold=True)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('Times New Roman · Body 10pt · H1 14pt · H2 12pt'),
            SIZES['cover_sub'], italic=True, color=GREY_TEXT)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('§3.7 justification + §3.8 evidence table'),
            SIZES['cover_sub'], italic=True, color=GREY_TEXT)
    doc.add_page_break()

    render(doc, README_MD)

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run('— End of README v3.0 —'), SIZES['footer'], italic=True, color=GREY_TEXT)

    out = 'KungFu_AI_Football_README_v3.0.docx'
    doc.save(out)
    print(f'✅ Saved: {out}')
    print(f'   Font: {FONT} · Body {SIZES["body"]}pt · H1 {SIZES["h1"]}pt · H2 {SIZES["h2"]}pt')


if __name__ == '__main__':
    build()