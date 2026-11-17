"""
Kung Fu AI Football — A3 Negotiation Card (.docx)
Requires: pip install python-docx
Standalone. Outputs Negotiation_Card_A3_INTERNAL.docx
Single A3 landscape page.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===========================================================================
# NUMBERS — single source of truth
# ===========================================================================

MIDPOINT_CENTRE = 134.23      # v13 Step D
REQUIRED_MULT = 20.0
POST_TARGET = MIDPOINT_CENTRE / REQUIRED_MULT     # 6.71

LADDER = [
    ("A", 5.71,  "20x at midpoint — full fund return, max dilution"),
    ("B", 7.80,  "20x if Series A clears HK$176M"),
    ("C", 9.00,  "20x if Series A clears HK$200M — round-number ask"),
    ("D", 13.42, "10x at midpoint — founder's opening"),
]

TRANCHES = [(0.5, 6.21), (0.5, 12.00)]           # (cheque, pre-money)
stake1 = 0.5 / (6.21 + 0.5)
stake2 = 0.5 / (12.00 + 0.5)
founder_after = (1 - stake1) * (1 - stake2)
blended = 1 - founder_after

SCENARIOS = [
    ("S0 Base",        "134–200", "5%",  "20x", "10.0–14.9%"),
    ("S1 Winter",      "70–100",  "4%",  "25x", "25.0–35.7%"),
    ("S2 Tailwind",    "200–300", "6%",  "15x", "5.0–7.5%"),
    ("S3 Contraction", "50–70",   "3%",  "30x", "42.9–60.0%"),
]

# ===========================================================================
# HELPERS
# ===========================================================================

def set_cell_bg(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def no_borders(tbl):
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)


def thin_borders(tbl, color="BFBFBF"):
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def cell_text(cell, text, size=9, bold=False, italic=False, color=None,
              align=None, space_after=2):
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    r.bold = bold; r.italic = italic; r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def cell_add_para(cell, text, size=9, bold=False, italic=False, color=None,
                  space_after=2):
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run(text)
    r.bold = bold; r.italic = italic; r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def heading(cell, text):
    p = cell.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text)
    r.bold = True; r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return p


def section_box(doc, parent_cell, title, fill="1F3864"):
    """Create a nested single-cell table as a titled box; return the inner cell."""
    box = parent_cell.add_table(rows=2, cols=1)
    thin_borders(box)
    box.alignment = WD_TABLE_ALIGNMENT.CENTER
    box.columns[0].width = Cm(12.8)
    hdr, body = box.rows[0].cells[0], box.rows[1].cells[0]
    hdr.width = Cm(12.8); body.width = Cm(12.8)
    set_cell_bg(hdr, fill)
    cell_text(hdr, title, size=9, bold=True,
              color=RGBColor(0xFF, 0xFF, 0xFF), space_after=2)
    return body


def inner_table(parent_cell, headers, rows, widths, font=8,
                highlight_col=None, shade_rows=None):
    n = len(headers)
    t = parent_cell.add_table(rows=1, cols=n)
    t.style = "Table Grid"
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(h); r.bold = True; r.font.size = Pt(font)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_bg(hdr[i], "1F3864")
    shade_rows = shade_rows or []
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(str(v)); r.font.size = Pt(font)
            if highlight_col is not None and i == highlight_col:
                set_cell_bg(cells[i], "EAF1F8")
        if ri in shade_rows:
            for c in cells:
                set_cell_bg(c, "F2F2F2")
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    return t


def money(x, d=2):
    return f"HK${x:,.{d}f}M"

# ===========================================================================
# DOCUMENT — A3 landscape
# ===========================================================================

doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Cm(42.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(1.1)
sec.right_margin = Cm(1.1)
sec.top_margin = Cm(0.9)
sec.bottom_margin = Cm(0.8)

st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(9)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

RED = RGBColor(0xC0, 0x00, 0x00)
NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x60, 0x60, 0x60)

# ---- Title band
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
r = p.add_run("KUNG FU AI FOOTBALL  ·  SEED NEGOTIATION CARD")
r.bold = True; r.font.size = Pt(20); r.font.color.rgb = NAVY

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(6)
r = p.add_run("INTERNAL  ·  CONFIDENTIAL  ·  NOT FOR DISTRIBUTION     |     "
              "HK$1.0M cheque     |     Series A reference HK$134.23M     |     "
              "Fund multiple 20×     |     Post-money ceiling HK$6.71M")
r.bold = True; r.font.size = Pt(8); r.font.color.rgb = RED

# ---- Outer 3-column layout
outer = doc.add_table(rows=1, cols=3)
no_borders(outer)
outer.alignment = WD_TABLE_ALIGNMENT.CENTER
col_w = Cm(13.2)
for c in outer.rows[0].cells:
    c.width = col_w
c1, c2, c3 = outer.rows[0].cells

# ===========================================================================
# COLUMN 1 — Equation, Ladder, Scenario strip
# ===========================================================================

b = section_box(doc, c1, "1 · THE GOVERNING EQUATION")
cell_text(b, "Founder dilution = (cheque × required multiple) ÷ Series A price",
          size=10, bold=True, color=NAVY, space_after=4)
cell_add_para(b, "Dilution is fixed by the Series A price and the fund's multiple — "
                 "not by the pre-money you negotiate. The pre-money is a label; the "
                 "dilution is the cost.", size=8, italic=True, color=GREY)
cell_add_para(b, "Only three levers exist:  C (cheque size)   ·   M (required "
                 "multiple)   ·   SA (Series A clearing price).", size=8, bold=True)
cell_add_para(b, f"Check: 1.0 × 20 ÷ 134.23 = 14.90%  ✓", size=8, italic=True,
              color=GREY)

b = section_box(doc, c1, "2 · THE SEED PRICE LADDER")
inner_table(b, ["", "Pre-money", "HK$1M buys", "Reaching 20× requires"],
            [[k, money(p), f"{1.0/(p+1.0)*100:.2f}%", note]
             for k, p, note in LADDER],
            widths=[0.8, 3.0, 2.8, 6.0], font=7.5, highlight_col=1,
            shade_rows=[0])
cell_add_para(b, "A is the only price that returns the fund at the midpoint. "
                 "D is a fair return (10×) but will not close. Open D, land B or C, "
                 "refuse A alone.", size=8, italic=True, color=GREY)

b = section_box(doc, c1, "3 · SCENARIO QUICK-REFERENCE (v14)")
inner_table(b, ["Scenario", "Series A", "Surv.", "Mult", "Dilution"],
            [list(s) for s in SCENARIOS],
            widths=[3.0, 2.6, 1.6, 1.4, 4.2], font=7.5, highlight_col=4,
            shade_rows=[1, 3])
cell_add_para(b, "In S1/S3 do not raise a priced round. Use PolyU co-investment "
                 "(up to HK$4M), BUD and retained earnings. A depressed anchor "
                 "persists for five years.", size=8, italic=True, color=GREY)

# ===========================================================================
# COLUMN 2 — Proposal, Positions, Do-not-say
# ===========================================================================

b = section_box(doc, c2, "4 · THE PROPOSAL — TWO TRANCHES", fill="0F6B3C")
cell_add_para(b, "HK$0.5M now at HK$6.21M pre  →  7.45%", size=9, bold=True)
cell_add_para(b, "HK$0.5M post-pilot at HK$12.0M pre  →  4.00%", size=9, bold=True)
cell_add_para(b, f"Blended: HK$1.0M for {blended*100:.2f}%  ·  founder retains "
                 f"{founder_after*100:.2f}%", size=9, bold=True, color=NAVY)
cell_add_para(b, "Versus 14.90% for HK$1.0M in one cheque — the second tranche "
                 "saves ~3.75 points of the company because it prices against "
                 "evidence, not against the plan.", size=8, italic=True, color=GREY)
cell_add_para(b, "Fallback: HK$0.5M single cheque at HK$6.21M pre — 7.45%, highest "
                 "pre-money, half the dilution, half the capital.", size=8)

b = section_box(doc, c2, "5 · NEGOTIATION POSITIONS")
inner_table(b, ["Position", "Structure", "Founder keeps", "When"],
            [["Open", "Two tranches, HK$1.0M", f"{founder_after*100:.1f}%",
              "Pilot timeline allows the gap"],
             ["Target", "HK$1.0M single at HK$7.80M pre", "88.6%",
              "Series A credible at HK$176M"],
             ["Settle", "HK$0.5M single at HK$6.21M pre", "92.6%",
              "Runway sufficient at HK$0.5M"],
             ["Walk", "No priced round", "100%",
              "S1/S3, or below HK$5.71M pre"]],
            widths=[1.8, 5.6, 2.4, 3.0], font=7.5, highlight_col=2,
            shade_rows=[3])
cell_add_para(b, "Refuse: committed dilution with optional capital. If tranche 2 is "
                 "at the investor's discretion, it must be priced after the pilot.",
              size=8, italic=True, color=RED)

b = section_box(doc, c2, "6 · DO NOT SAY", fill="7A1F1F")
for line in [
    "✗  \"Our DCF says we're worth HK$134M.\"  — they are pricing to P_A ÷ post, not to your model.",
    "✗  \"Your 10% ask is unreasonable.\"  — it is 20× arithmetic. Name it, don't dispute it.",
    "✗  \"We'll be profitable in Year 1.\"  — they have read the model. Say something new.",
    "✗  \"We need the money by Q2.\"  — urgency kills the walk and the walk is your leverage.",
    "✗  \"Let's just agree a number.\"  — you will be agreeing to their number.",
]:
    cell_add_para(b, line, size=8, space_after=3)

# ===========================================================================
# COLUMN 3 — Three sentences, two questions, outstanding
# ===========================================================================

b = section_box(doc, c3, "7 · THE THREE SENTENCES THAT MATTER", fill="0F6B3C")
for i, s in enumerate([
    "HK$9M for 10% is a dream-frame price, not a market price — it is fund "
    "arithmetic, not a valuation disagreement.",
    "Dilution is fixed by the Series A price and your multiple, not by the "
    "pre-money we negotiate. Only three levers exist: cheque size, multiple, "
    "Series A price.",
    "I cannot accept committed dilution with optional capital. If the second "
    "tranche is at your discretion, it prices after the pilot.",
], 1):
    cell_add_para(b, f"{i}.  {s}", size=8.5, space_after=4)

b = section_box(doc, c3, "8 · TWO QUESTIONS TO ASK FIRST", fill="0F6B3C")
cell_add_para(b, "1.  What is your minimum position size?", size=9, bold=True)
cell_add_para(b, "Determines whether HK$0.5M tranches are available at all. "
                 "Five minutes to ask; weeks to discover in diligence.",
              size=8, italic=True, color=GREY)
cell_add_para(b, "2.  Would you accept a HK$0.5M tranche now with a pre-agreed "
                 "price band for a second tranche post-pilot?", size=9, bold=True)
cell_add_para(b, "Determines whether the tranche structure is on the table before "
                 "you reveal that it is your preferred structure.",
              size=8, italic=True, color=GREY)

b = section_box(doc, c3, "9 · KEY LINES (verbatim)", fill="1F3864")
for line in [
    "\"The pricing has to match the risk. If the second cheque is priced today, it "
    "should be a binding commitment. If it is at your option, it should price after "
    "the pilot.\"",
    "\"Your 10% implies HK$9M. That is below every scenario we have modelled, "
    "including the contraction case. The question is which forecast it encodes.\"",
    "\"I would rather retain earnings for twelve months than set a HK$5.7M anchor "
    "that follows this company into every subsequent round.\"",
]:
    cell_add_para(b, line, size=8, italic=True, space_after=4)

b = section_box(doc, c3, "10 · OUTSTANDING BEFORE THE MEETING", fill="7A1F1F")
for line in [
    "Lower case rebuild — still Centre ÷ 2.57. In S1/S3 this becomes the base case "
    "and the company has no credible downside. Highest priority.",
    "Confirm audit engagement type — audit vs review determines whether the 12% "
    "rate argument survives.",
    "Confirm fund minimum position size — determines whether tranches exist.",
]:
    cell_add_para(b, "•  " + line, size=8, space_after=3)

# ---- Footer band
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Intrinsic value reference (v13 matrix) is NOT a market price: "
              "midpoint Centre HK$134.23M, implied rate ~12.9%, HK$1.0M buys 0.74%. "
              "Market will price at HK$9–19M. Use the matrix to plan the Series A "
              "step-up, not to fight the seed price.")
r.italic = True; r.font.size = Pt(7.5); r.font.color.rgb = GREY

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Sources: v13 Valuation Route · v14 Global Scenarios · "
              "Seed Structure Comparison. Supersedes all earlier drafts. "
              "Draft, pre-validation — conditional on Y1–Y3 audit and pilot.")
r.italic = True; r.font.size = Pt(7); r.font.color.rgb = GREY

out = "Negotiation_Card_A3_INTERNAL.docx"
doc.save(out)

print(f"Saved: {out}  (A3 landscape, 42.0 × 29.7 cm)")
print(f"Post-money ceiling: {money(POST_TARGET)}   |   "
      f"Blended tranche dilution: {blended*100:.2f}%   |   "
      f"Founder retains: {founder_after*100:.2f}%")