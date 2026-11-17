"""
Generate Kung Fu AI Football — Executive Summary & Investment Thesis (.docx)
Requires: pip install python-docx
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
    """Set table cell background colour."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def add_bold_paragraph(doc, segments, style=None, space_after=6):
    """
    Add a paragraph composed of (text, bold) tuples.
    Example: add_bold_paragraph(doc, [("The moat is ", False), ("the fusion dataset.", True)])
    """
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_after = Pt(space_after)
    for text, bold in segments:
        run = p.add_run(text)
        run.bold = bold
    return p


def add_bullet(doc, segments, level=0):
    """Add a bullet point that can contain bold segments."""
    p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
    for text, bold in segments:
        run = p.add_run(text)
        run.bold = bold
    return p


def add_table_with_header(doc, headers, rows, col_widths=None, header_fill="1F3864"):
    """Add a styled table with a coloured header row."""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        run = hdr[i].paragraphs[0].add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(10)
        set_cell_background(hdr[i], header_fill)

    # Data rows
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
# Build document
# ---------------------------------------------------------------------------

doc = Document()

# Base style
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
    "struck. They do not explain why it missed the target, which biomechanical "
    "flaw caused the error, or which drill corrects it."
)
add_bold_paragraph(doc, [
    ("Parents are voting with their wallets. ", False),
    ("China's youth sports training market reached RMB 253.1 billion in 2025", True),
    (", with renewal revenue rising from 48% to ", False),
    ("57.07%", True),
    (" of total income — and ", False),
    ("44.67%", True),
    (" of parents explicitly stating they will pay more for high-quality courses. "
     "The industry's own conclusion: parents no longer buy on \"cheap\"; they "
     "buy on results and professionalism.", False),
])


# ============================ 2. PRODUCT ==================================
doc.add_heading("2. The Product", level=1)

add_bold_paragraph(doc, [("A hybrid hardware-plus-subscription platform:", True)])

add_bullet(doc, [("Smart Ball", True),
                 (" — IoT sensors capturing ball kinetics, spin, and speed (HK$980)", False)])
add_bullet(doc, [("Computer Vision layer", True),
                 (" — player biomechanics: foot angle, hip rotation, kick style", False)])
add_bullet(doc, [("AI Kernel", True),
                 (" — fuses both streams into diagnostic and prescriptive coaching feedback", False)])
add_bullet(doc, [("Recurring subscription", True), (" — HK$1,592/year", False)])

add_bold_paragraph(doc, [
    ("The moat is the fusion dataset. ", True),
    ("Ball mechanics plus biomechanics plus drill-correction logic. A hardware "
     "competitor can replicate a sensor ball; it cannot replicate the diagnostic "
     "layer without the paired dataset. We do not tell a player ", False),
    ("how hard", True),
    (" they kicked — we tell them ", False),
    ("why", True),
    (" it missed and ", False),
    ("how", True),
    (" to fix it.", False),
])


# ============================ 3. MARKET ===================================
doc.add_heading("3. Market Sizing — A Verifiable Funnel", level=1)

doc.add_paragraph(
    "The addressable base is not an estimate. It is the sum of officially "
    "reported cohorts:"
)

add_table_with_header(
    doc,
    headers=["Cohort", "Participants"],
    rows=[
        ["Amateur players (社会足球业余球员)", "979,544"],
        ["Registered elite youth athletes (注册精英青少年运动员)", "109,220"],
        ["Registered students at football-specialist schools (校园足球特色学校)", "12,800,000"],
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
    (" — the share of Chinese parents who explicitly state willingness to "
     "support a child's long-term football pathway. This is a stated-preference "
     "floor, not a top-down adoption guess.", False),
])

add_bold_paragraph(doc, [
    ("SOM = SAM × 8.7% anchor × adoption penetration:", True),
])

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
    "Each figure is reproducible: 13.89M × (8.7% / 4), 13.89M × (8.7% / 2), "
    "13.89M × 8.7%. No hidden conversion assumptions."
)


# ============================ 4. COMPETITION ==============================
doc.add_heading("4. Competitive Positioning", level=1)

add_table_with_header(
    doc,
    headers=["", "INSAIT JOY", "Kung Fu AI Football"],
    rows=[
        ["Hardware", "Smart ball (US$139)", "Smart ball (HK$980)"],
        ["Data captured", "Ball mechanics", "Ball mechanics + biomechanics via CV"],
        ["Feedback", "Real-time drill feedback via app",
         "Diagnostic (\"why it missed\") + prescriptive correction"],
        ["Revenue model", "One-time hardware", "Hardware + recurring subscription"],
    ],
    col_widths=[3.5, 6, 7.5],
)

doc.add_paragraph(
    "INSAIT JOY is a well-executed hardware product with an app-based feedback "
    "layer. Our differentiation is the computer-vision biomechanics layer, which "
    "converts performance tracking into prescriptive coaching. That is a "
    "data-driven moat hardware-only competitors cannot replicate without the "
    "paired dataset."
)


# ============================ 5. GTM ======================================
doc.add_heading("5. Go-to-Market — B2B Technology Enabler", level=1)

doc.add_paragraph(
    "Kung Fu AI Football does not build a D2C brand. We supply the AI engine and "
    "smart-ball hardware to football academies, school programmes, and sports "
    "brands with existing distribution."
)
add_bold_paragraph(doc, [
    ("Target channel partners", True),
    (" include large B2C sports platforms (e.g., Anta, Amer Sports, Asia Partners) "
     "— the structural model is to avoid bearing channel, marketing, and retail "
     "costs. ", False),
    ("(These are target partners, not signed agreements; pilot-stage discussions "
     "are the near-term objective.)", False),
])


# ============================ 6. INCENTIVES ===============================
doc.add_heading("6. Hong Kong R&D Incentives", level=1)

doc.add_paragraph(
    "As a pure innovative R&D company, the build phase is materially subsidised:"
)
add_bullet(doc, [
    ("DIPN 55 (Enhanced Tax Reduction):", True),
    (" Qualifying R&D expenditure with local DLRIs qualifies for a 300% tax "
     "deduction on the first HK$2 million — a significant reduction in annual "
     "tax burden.", False),
])
add_bullet(doc, [
    ("HKAS 38 (Intangible Assets):", True),
    (" Development-phase R&D expenditure (HK$360k per 3-year cycle) is "
     "capitalised as an intangible asset, strengthening the balance sheet while "
     "the tax deduction still applies to cash spent.", False),
])
add_bullet(doc, [
    ("Research Talent Hub (RTH):", True),
    (" Eligible for ITF-project research talent support.", False),
])
add_bullet(doc, [
    ("PolyU Entrepreneurship Investment Fund:", True),
    (" Eligible for up to HK$4 million in co-investment capital, subject to "
     "PolyU member equity or IP-licensing criteria.", False),
])


# ============================ 7. FINANCIALS ===============================
doc.add_heading("7. Financial Projections", level=1)

doc.add_paragraph(
    "Six-year cumulative net profits, after all operating costs, HKAS 38 "
    "amortisation, and HKSAR taxes (internal model outputs, available for "
    "diligence):"
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
    "The business is designed to be cash-generative and highly scalable — "
    "transitioning from a hardware-plus-subscription model into a high-margin "
    "data and IP-licensing business."
)


# ============================ 8. ASK ======================================
doc.add_heading("8. The Ask", level=1)

add_bold_paragraph(doc, [
    ("Seeking pilot partners and seed capital", True),
    (" to:", False),
])
doc.add_paragraph(
    "1. Execute a 5–10 academy pilot validating the AI kernel against "
    "ground-truth coaching assessments",
    style="List Number",
)
doc.add_paragraph(
    "2. Build the proprietary paired dataset (ball + biomechanics) that "
    "constitutes the defensible moat",
    style="List Number",
)
doc.add_paragraph(
    "3. Convert the pilot into a repeatable B2B licensing motion with 2–3 "
    "design partners",
    style="List Number",
)


# ============================ SUMMARY BOX =================================
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