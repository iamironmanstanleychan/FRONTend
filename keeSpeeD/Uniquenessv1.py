"""
Generate Kung Fu AI Football — Executive Summary & Investment Thesis (.docx)
Requires: pip install python-docx

v2 — Section 2 (Product) condensed; Section 4 replaced with the rewritten
"Competitiveness & Uniqueness" (INSAIT free-app objection addressed head-on,
92% renewal denominator stated, pricing logic split into background vs. proof).
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def set_cell_background(cell, hex_color: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def add_bold_paragraph(doc, segments, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    for text, bold in segments:
        run = p.add_run(text)
        run.bold = bold
    return p


def add_bullet(doc, segments, level=0):
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    for text, bold in segments:
        run = p.add_run(text)
        run.bold = bold
    return p


def add_table_with_header(doc, headers, rows, col_widths=None, header_fill="1F3864"):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(10)
        set_cell_background(hdr[i], header_fill)

    for row_data in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row_data):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(val))
            run.font.size = Pt(10)

    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)
    return table


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

doc = Document()

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")


# ============================ COVER =======================================
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_before = Pt(120)
run = title.add_run("Kung Fu AI Football")
run.bold = True
run.font.size = Pt(32)
run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run("Executive Summary & Investment Thesis")
run.bold = True
run.font.size = Pt(18)
run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

tagline = doc.add_paragraph()
tagline.alignment = WD_ALIGN_PARAGRAPH.CENTER
tagline.paragraph_format.space_before = Pt(24)
run = tagline.add_run(
    "A sports AI platform that diagnoses why a kick failed — "
    "and prescribes the correction."
)
run.italic = True
run.font.size = Pt(12)

page_break(doc)


# ============================ 1. PROBLEM ==================================
doc.add_heading("1. The Problem", level=1)

doc.add_paragraph(
    "Football is the world's largest sport, but youth and amateur training "
    "remains feedback-scarce. Hardware trackers report how hard a ball was "
    "struck. They do not explain why it missed, which biomechanical flaw caused "
    "the error, or which drill corrects it."
)
add_bold_paragraph(doc, [
    ("Parents are voting with their wallets. ", False),
    ("China's youth sports training market reached RMB 253.1 billion in 2025", True),
    (", with renewal revenue rising from 48% to ", False),
    ("57.07%", True),
    (" of total income, and ", False),
    ("44.67%", True),
    (" of parents explicitly stating they will pay more for high-quality courses. "
     "The industry's own conclusion: parents no longer buy on \"cheap\" — they "
     "buy on results and professionalism.", False),
])


# ============================ 2. PRODUCT (CONDENSED) ======================
doc.add_heading("2. The Product", level=1)

add_bold_paragraph(doc, [
    ("A hybrid hardware-plus-subscription platform: ", True),
    ("Smart Ball with IoT sensors (ball kinetics, spin, speed) at HK$980, plus a "
     "computer-vision biomechanics layer (foot angle, hip rotation, kick style), "
     "plus a recurring subscription at HK$1,592/year.", False),
])
add_bold_paragraph(doc, [
    ("The output is prescriptive, not descriptive. ", True),
    ("The AI Kernel does not report how hard the ball was struck; it diagnoses "
     "why it missed and prescribes the drill that corrects it.", False),
])
add_bold_paragraph(doc, [
    ("The moat is the paired dataset. ", True),
    ("Ball mechanics and player biomechanics, labelled against coaching ground "
     "truth, plus the drill-correction logic that maps flaw to fix. A sensor ball "
     "can be copied; the paired dataset cannot be rebuilt without both streams.", False),
])


# ============================ 3. MARKET ===================================
doc.add_heading("3. Market Sizing — A Verifiable Funnel", level=1)

doc.add_paragraph(
    "The addressable base is the sum of officially reported cohorts:"
)

add_table_with_header(
    doc,
    headers=["Cohort", "Participants"],
    rows=[
        ["Amateur players (社会足球业余球员)", "979,544"],
        ["Registered elite youth athletes (注册精英青少年运动员)", "109,220"],
        ["Registered students, football-specialist schools (校园足球特色学校)", "12,800,000"],
        ["SAM", "≈ 13.89 million"],
    ],
    col_widths=[11, 5],
)

doc.add_paragraph(
    "Sources: China Football Association 2025 Digital Annual Report; "
    "2025 national campus football data."
).runs[0].italic = True

add_bold_paragraph(doc, [
    ("The conversion anchor is 8.7%", True),
    (" — the share of Chinese parents who state willingness to support a child's "
     "long-term football pathway. A stated-preference floor, not a top-down "
     "adoption guess.", False),
])

add_bold_paragraph(doc, [("SOM = SAM × 8.7% anchor × adoption penetration:", True)])

add_table_with_header(
    doc,
    headers=["Scenario", "Penetration of Anchor Cohort", "SOM (Year 6)"],
    rows=[
        ["P10 — Conservative", "25%", "≈ 302,100"],
        ["P50 — Centre", "50%", "≈ 604,200"],
        ["P90 — Ceiling", "100%", "≈ 1,208,400"],
    ],
    col_widths=[5, 6, 5],
)

doc.add_paragraph(
    "Each figure is reproducible: 13.89M × (8.7% / 4), (8.7% / 2), 8.7%. "
    "No hidden conversion assumptions."
)


# ============================ 4. COMPETITIVENESS ==========================
doc.add_heading("4. Competitiveness & Uniqueness", level=1)

# --- A. Moat
doc.add_heading("A. The Moat: A Proprietary Paired Dataset", level=2)
doc.add_paragraph(
    "The moat is not the hardware. It is the pairing of ball mechanics with "
    "player biomechanics, labelled against coaching ground truth, and the logic "
    "that maps a flaw to a corrective drill. A hardware-only competitor cannot "
    "replicate it without rebuilding the dataset from both streams."
)

# --- B. Price
doc.add_heading("B. Price Positioning", level=2)

add_table_with_header(
    doc,
    headers=["Product", "Price", "Model"],
    rows=[
        ["adidas standard match ball (no sensors)", "US$170", "One-time; no feedback"],
        ["INSAIT JOY Smart Football", "US$139 (~HK$1,084)", "One-time; free app, no subscription"],
        ["INSAIT smart football — China gov. procurement", "RMB 1,745/unit", "Institutional purchase"],
        ["Catapult One (GPS vest)", "≈US$179/year", "Membership only; one pod per player; no ball"],
        ["Kung Fu AI Football", "HK$980 + HK$1,592/year", "Hybrid hardware + subscription"],
    ],
    col_widths=[6.5, 4, 6.5],
)

add_bold_paragraph(doc, [
    ("HK$980 (~RMB 900) sits ", False),
    ("below", True),
    (" the INSAIT JOY consumer ball at its US$139 list price, ", False),
    ("roughly half", True),
    (" the Chinese government procurement price for a comparable smart ball "
     "(RMB 1,745/unit), and below Catapult One's annual membership — which does "
     "not include a ball.", False),
])

# --- C. The objection
doc.add_heading("C. Why the Subscription Survives INSAIT's Free-App Model", level=2)

doc.add_paragraph(
    "INSAIT JOY is a well-executed smart ball: US$139, free app, no subscription, "
    "real-time drill feedback. It is our closest direct competitor, and its "
    "no-subscription model is deliberate. It is also ball-only. INSAIT sees the "
    "ball; it does not see the player."
)
add_bold_paragraph(doc, [
    ("Our subscription is not a paywall on app access. ", True),
    ("It pays for the computer-vision biomechanics layer and the diagnostic "
     "engine — the capability to say \"your hip rotation opens 15° early, which "
     "is why the ball slices right; here is the drill that closes it.\" "
     "INSAIT's free app cannot say that, because it never sees the hip.", False),
])
doc.add_paragraph(
    "Catapult One validates subscription as an accepted format in connected "
    "sports training — membership-only, per-player, no outright purchase. We "
    "cite it for the format, not the category: Catapult applies it to GPS load "
    "tracking; we apply it to skill coaching, where diagnostic depth justifies "
    "recurrence."
)

# --- D. Pricing rationale
doc.add_heading("D. Why HK$1,592/Year Is a Marginal Decision", level=2)

add_bold_paragraph(doc, [
    ("Background condition. ", True),
    ("China's youth sports training market reached RMB 253.1 billion in 2025, "
     "with renewal revenue rising from 48% to 57.07%. 44.67% of parents say they "
     "will pay more for high-quality courses. Demand-side conditions are in place.", False),
])
add_bold_paragraph(doc, [
    ("Pricing proof. ", True),
    ("Within football training, high-end courses (annual fee above RMB 30,000) "
     "achieve a 92% renewal rate — once a family commits to a structured pathway, "
     "it stays. Against a RMB 30,000 budget, HK$1,592/year is roughly 5%: a "
     "marginal add-on for quantified feedback, not a new category of expenditure.", False),
])
add_bold_paragraph(doc, [
    ("The denominator, stated plainly. ", True),
    ("That high-end segment is currently only 8% of total students. That is the "
     "point: willingness to pay exists at the top of the market and is "
     "underserved by product. The mass segment (RMB 8,000–15,000, 76% of students, "
     "55% renewal) is a later product line, not the beachhead. Our pricing and "
     "our SAM/SOM funnel anchor to the same 8–8.7% cohort.", False),
])

# --- E. Demand alignment
doc.add_heading("E. Demand Alignment", level=2)
doc.add_paragraph(
    "The 2025–2030 China football training industry report finds that 68% of "
    "Gen-Z parents require data-driven feedback and 75% verify coach credentials "
    "before enrolling. Our product is built on exactly those two axes: quantified "
    "feedback on every touch, and recommendations traceable to biomechanical data "
    "rather than opinion."
)


# ============================ 5. GTM ======================================
doc.add_heading("5. Go-to-Market — B2B Technology Enabler", level=1)

add_bold_paragraph(doc, [
    ("Kung Fu AI Football does not build a consumer brand. ", True),
    ("It monetises the AI engine and smart-ball hardware through two channels:", False),
])
add_bullet(doc, [
    ("Institutional (pilot phase): ", True),
    ("direct sales to football academies and school programmes. The 5–10 academy "
     "pilot is the validation motion; it carries a contained direct-sales cost, "
     "by design and only during validation.", False),
])
add_bullet(doc, [
    ("Brand partnership (scale phase): ", True),
    ("licensing the engine and hardware to large B2C sports platforms that "
     "already own consumer distribution, retail, and marketing — target partners "
     "include Anta (and its subsidiary Amer Sports) and Asia Partners.", False),
])
add_bold_paragraph(doc, [
    ("The brand-partnership tier removes channel, marketing, and retail cost from "
     "the P&L at scale. Institutional sales are the proof layer, not the "
     "long-term cost structure.", True),
])


# ============================ 6. INCENTIVES ===============================
doc.add_heading("6. Hong Kong R&D Incentives", level=1)

doc.add_paragraph(
    "As a pure innovative R&D company, the build phase is materially subsidised:"
)
add_bullet(doc, [
    ("DIPN 55 (Enhanced Tax Reduction): ", True),
    ("Qualifying R&D expenditure with local DLRIs qualifies for a 300% tax "
     "deduction on the first HK$2 million.", False),
])
add_bullet(doc, [
    ("HKAS 38 (Intangible Assets): ", True),
    ("Development-phase expenditure (HK$360k per 3-year cycle) is capitalised as "
     "an intangible asset, strengthening the balance sheet while the tax "
     "deduction still applies to cash spent.", False),
])
add_bullet(doc, [
    ("Research Talent Hub (RTH): ", True),
    ("Eligible for ITF-project research talent support.", False),
])
add_bullet(doc, [
    ("PolyU Entrepreneurship Investment Fund: ", True),
    ("Eligible for up to HK$4 million in co-investment capital, subject to PolyU "
     "member equity or IP-licensing criteria.", False),
])


# ============================ 7. FINANCIALS ===============================
doc.add_heading("7. Financial Projections", level=1)

doc.add_paragraph(
    "Six-year cumulative net profits, after all operating costs, HKAS 38 "
    "amortisation, and HKSAR taxes (internal model outputs, available for diligence):"
)

add_table_with_header(
    doc,
    headers=["Case", "Assumptions", "Cumulative Net Profit"],
    rows=[
        ["Conservative", "40% marketing cost + P10 adoption", "HK$89.97 million"],
        ["Middle", "35% marketing cost + P50 adoption", "HK$230.54 million"],
        ["Ceiling", "30% marketing cost + P90 adoption", "HK$560.96 million"],
    ],
    col_widths=[3.5, 7.5, 5],
)

doc.add_paragraph(
    "Designed to be cash-generative and highly scalable — transitioning from "
    "hardware-plus-subscription into a high-margin data and IP-licensing business."
)


# ============================ 8. ASK ======================================
doc.add_heading("8. The Ask", level=1)

add_bold_paragraph(doc, [("Seeking pilot partners and seed capital to:", True)])
doc.add_paragraph(
    "1. Execute a 5–10 academy pilot validating the AI kernel against "
    "ground-truth coaching assessments", style="List Number")
doc.add_paragraph(
    "2. Build the proprietary paired dataset (ball + biomechanics) that "
    "constitutes the defensible moat", style="List Number")
doc.add_paragraph(
    "3. Convert the pilot into a repeatable B2B licensing motion with 2–3 "
    "design partners", style="List Number")


# ============================ SUMMARY =====================================
doc.add_paragraph()
add_bold_paragraph(doc, [
    ("Summary for VC: ", True),
    ("Kung Fu AI Football is not another smart ball. It is a ", False),
    ("data-fusion coaching engine", True),
    (" sold through B2B channels, with Hong Kong R&D incentives subsidising the "
     "build and a verifiable 1.21M-participant anchor cohort funding the early "
     "subscription base. ", False),
    ("The pilot is the de-risking event.", True),
])


# ============================ SAVE ========================================
output = "KungFu_AI_Football_Executive_Summary.docx"
doc.save(output)
print(f"Saved: {output}")