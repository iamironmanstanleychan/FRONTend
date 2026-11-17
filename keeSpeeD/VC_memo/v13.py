"""
Kung Fu AI Football — The Valuation Route (.docx)
Requires: pip install python-docx

A linear derivation of the midpoint valuation, seven steps:
  A  P&L projection Y1–Y7 Q2
  B  Y7 Q2 undiscounted accumulated value
  C  Y0 adjusted (9 cells)
  D  Midpoint of B and C (9 cells)
  E  Implied discount rate for each of the 9 midpoint cells
  F  Shares allocated to VC for HK$1.0M — pre-money and post-money (18 cells)
  G  Positioning and counterpart positions
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===========================================================================
# PARAMETERS
# ===========================================================================

DATA = [
    ("Y1",  0.58,   2.08,    5.24),
    ("Y2",  2.01,   5.19,   12.61),
    ("Y3", 12.43,  31.04,   74.31),
    ("Y4", 17.25,  44.20,  107.63),
    ("Y5", 23.72,  61.05,  149.13),
    ("Y6", 33.74,  86.74,  211.79),
]
GRANT = 0.24
INTANGIBLE_Y6 = 0.58
GRANT_YEARS = 6.25
INTANGIBLE_YEARS = 6
RATES = [0.25, 0.50, 0.75]
CASES = ["Ceiling", "Centre", "Lower"]
CASE_IDX = {"Lower": 1, "Centre": 2, "Ceiling": 3}

SHARE_COUNT = 100
SEED_CHEQUE = 1.0
INVESTOR_ASK = 0.0216

REPORTED_NPV = {
    "Ceiling": {0.25: 199.02, 0.50: 90.85, 0.75: 49.15},
    "Centre":  {0.25:  81.96, 0.50: 37.51, 0.75: 20.37},
    "Lower":   {0.25:  32.04, 0.50: 14.70, 0.75:  8.01},
}
REPORTED_UNDISCOUNTED = {"Ceiling": 560.96, "Centre": 230.54, "Lower": 89.97}

# ===========================================================================
# HELPERS
# ===========================================================================

def set_cell_background(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tc_pr.append(shd)


def bold_para(doc, segs, after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(after)
    for text, b in segs:
        r = p.add_run(text); r.bold = b
    return p


def bullet(doc, segs):
    p = doc.add_paragraph(style="List Bullet")
    for text, b in segs:
        r = p.add_run(text); r.bold = b
    return p


def quote_para(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text); r.italic = True; r.font.size = Pt(10)
    return p


def note(doc, text, before=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    r = p.add_run(text); r.italic = True; r.font.size = Pt(9)
    return p


def carry(doc, text):
    """The 'carry forward' line that closes each step."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(10)
    r = p.add_run("→ Carries forward: ")
    r.bold = True; r.font.size = Pt(9)
    r2 = p.add_run(text)
    r2.italic = True; r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    return p


def table(doc, headers, rows, widths=None, header_fill="1F3864",
          highlight_col=None, shade_rows=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        r = hdr[i].paragraphs[0].add_run(h)
        r.bold = True; r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        set_cell_background(hdr[i], header_fill)
    shade_rows = shade_rows or []
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            lines = v if isinstance(v, list) else [v]
            for j, line in enumerate(lines):
                p = cells[i].paragraphs[0] if j == 0 else cells[i].add_paragraph()
                r = p.add_run(str(line)); r.font.size = Pt(9)
                if j > 0:
                    r.font.size = Pt(8)
                    r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)
            if highlight_col is not None and i == highlight_col:
                set_cell_background(cells[i], "EAF1F8")
        if ri in shade_rows:
            for c in cells:
                set_cell_background(c, "F2F2F2")
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    return t


def money(x, d=2):
    return f"HK${x:,.{d}f}M"


def pshare(val_m):
    return val_m * 1_000_000 / SHARE_COUNT


def ps_fmt(val_m):
    return f"HK${pshare(val_m):,.0f}"


def stake_post(pre_m):
    """Correct ownership: cheque / (pre + cheque)."""
    return SEED_CHEQUE / (pre_m + SEED_CHEQUE)


def shares_at_pre(pre_m):
    """Trap: shares priced off pre-money only. Over-allocates."""
    return SHARE_COUNT * SEED_CHEQUE / pre_m


def shares_at_post(pre_m):
    """Correct: shares priced off post-money."""
    return SHARE_COUNT * SEED_CHEQUE / (pre_m + SEED_CHEQUE)


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

# --- Step B: Y7 Q2 undiscounted
y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE_Y6 for c in CASES}

# --- Step C: Y0 adjusted
disc_g = {r: 1 / (1 + r) ** GRANT_YEARS for r in RATES}
disc_i = {r: 1 / (1 + r) ** INTANGIBLE_YEARS for r in RATES}
pv_grant = {r: GRANT * disc_g[r] for r in RATES}
pv_intangible = {r: INTANGIBLE_Y6 * disc_i[r] for r in RATES}

y0 = {}
for c in CASES:
    for r in RATES:
        y0[(c, r)] = (
            REPORTED_NPV[c][r] - GRANT + pv_grant[r] + pv_intangible[r]
        )

# --- Step D: midpoint
mid = {(c, r): (y0[(c, r)] + y7q2[c]) / 2 for c in CASES for r in RATES}


# --- Step E: implied discount rate for each midpoint cell
def y0_value(case, r):
    pv = sum(DATA[t][CASE_IDX[case]] / (1 + r) ** (t + 1) for t in range(6))
    pv += GRANT / (1 + r) ** GRANT_YEARS
    pv += INTANGIBLE_Y6 / (1 + r) ** INTANGIBLE_YEARS
    return pv


def implied_rate(case, target):
    """Bisection: find r such that y0_value(case, r) == target."""
    lo, hi = 0.0005, 3.0
    if y0_value(case, lo) < target:
        return None  # target above achievable max
    if y0_value(case, hi) > target:
        return None  # target below achievable min
    for _ in range(120):
        m = (lo + hi) / 2
        if y0_value(case, m) > target:
            lo = m
        else:
            hi = m
    return (lo + hi) / 2


implied = {(c, r): implied_rate(c, mid[(c, r)]) for c in CASES for r in RATES}

# --- Step F: shares
shares_pre = {(c, r): shares_at_pre(mid[(c, r)]) for c in CASES for r in RATES}
shares_post = {(c, r): shares_at_post(mid[(c, r)]) for c in CASES for r in RATES}

# --- Step G: positions
POSITIONS = {
    "Opening ask": ("Ceiling", 0.50),
    "Target":      ("Centre",  0.50),
    "Settle":      ("Lower",   0.50),
    "Walk-away":   ("Lower",   0.75),
}
pos = {k: mid[v] for k, v in POSITIONS.items()}
ask_pre = SEED_CHEQUE / INVESTOR_ASK - SEED_CHEQUE

# ===========================================================================
# BUILD
# ===========================================================================

doc = Document()
st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

# Section-level header/footer (suppressed on cover page, shown on pages 2+)
sec = doc.sections[0]
sec.different_first_page_header_footer = True

hdr = sec.header.paragraphs[0]
hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
hdr_run = hdr.add_run("INTERNAL  ·  CONFIDENTIAL  ·  NOT FOR DISTRIBUTION")
hdr_run.bold = True
hdr_run.font.size = Pt(8)
hdr_run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

ftr = sec.footer.paragraphs[0]
ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
ftr_run = ftr.add_run(
    "Draft, pre-validation — all figures are model outputs conditional on "
    "Y1–Y3 audited results and pilot completion.  "
    "v13 supersedes v8 and v11; to be superseded by v14 on Lower-case rebuild."
)
ftr_run.italic = True
ftr_run.font.size = Pt(7)
ftr_run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

# Cover-page classification block (above the title)
cover_box = doc.add_paragraph()
cover_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
cover_box.paragraph_format.space_before = Pt(0)
r = cover_box.add_run("⚠   INTERNAL — CONFIDENTIAL   ⚠")
r.bold = True
r.font.size = Pt(12)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

cover_sub = doc.add_paragraph()
cover_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
cover_sub.paragraph_format.space_after = Pt(18)
r = cover_sub.add_run(
    "Not for distribution outside the founding team and named advisers.\n"
    "Contains unpublished financial projections and negotiation strategy."
)
r.italic = True
r.font.size = Pt(9)
r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("MEMORANDUM"); r.bold = True; r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()
for label, value in [
    ("To:", "Founders / Board"),
    ("From:", "Finance & Strategy"),
    ("Re:", "The Valuation Route — From P&L to the Midpoint Matrix"),
    ("Classification:", "INTERNAL — CONFIDENTIAL — NOT FOR DISTRIBUTION"),
    ("Status:", "Draft, pre-validation. All figures are model outputs "
                "conditional on Y1–Y3 audited results and pilot completion."),
    ("Version:", "v13 — seven-step derivation, A through G. "
                 "Supersedes v8 (Seed_Valuation_Y4Q1) and v11 (Crossover)."),
    ("Next version:", "v14 — pending Lower-case rebuild with a distinct cost "
                      "structure; not for external circulation until then."),
]:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " ")
    rr.bold = True
    if label in ("Classification", "Status"):
        rr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        v = p.add_run(value)
        v.bold = (label == "Classification")
        v.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    else:
        p.add_run(value)

doc.add_paragraph()

# ---- 0. The route
doc.add_heading("0. The Route", level=1)
doc.add_paragraph(
    "This document derives the midpoint valuation in seven steps, each consuming the "
    "output of the one before. Nothing is asserted that is not built from the prior "
    "step. Read it top to bottom; the nine-cell matrix at Step D is the destination, "
    "and Steps E through G convert it into price, equity and negotiating position."
)
rows = [
    ["A", "P&L projection Y1–Y7 Q2", "Three cases, six years, plus grant"],
    ["B", "Y7 Q2 undiscounted accumulated value", "A summed, plus intangible"],
    ["C", "Y0 adjusted", "Reported NPV reconciled: 9 cells"],
    ["D", "Midpoint of B and C", "The 9-cell matrix — the destination"],
    ["E", "Implied discount rate", "What rate each midpoint cell corresponds to"],
    ["F", "Shares to VC for HK$1.0M", "18 cells: pre-money and post-money priced"],
    ["G", "Positioning", "Founder and counterpart positions"],
]
table(doc, ["Step", "What it produces", "Input"], rows,
      [1.8, 7.0, 7.0])

# ===========================================================================
# STEP A
# ===========================================================================

doc.add_heading("A. P&L Projection, Y1 to Y7 Q2", level=1)
doc.add_paragraph(
    "Net profit, HK$ millions, after all operating costs, HKAS 38 amortisation and "
    "HKSAR taxes. The three cases differ by marketing cost assumption and adoption "
    "band. The grant is a separate, non-operating line at Y7 Q2."
)
rows = [[r[0], f"{r[1]:.2f}", f"{r[2]:.2f}", f"{r[3]:.2f}"] for r in DATA]
rows.append(["Y7 Q2 grant", f"{GRANT:.2f}", f"{GRANT:.2f}", f"{GRANT:.2f}"])
rows.append(["Y1–Y6 subtotal",
             f"{sum(r[1] for r in DATA):.2f}",
             f"{sum(r[2] for r in DATA):.2f}",
             f"{sum(r[3] for r in DATA):.2f}"])
table(doc, ["Year", "Lower (40% mkt / P10)", "Centre (35% mkt / P50)",
            "Ceiling (30% mkt / P90)"],
      rows, [3.5, 4.3, 4.3, 4.3], shade_rows=[7])
note(doc, "The Y1–Y6 subtotal is the operating basis. The grant is added separately "
          "and is not net profit from operations.")
carry(doc, "the three Y1–Y6 subtotals feed Step B.")

# ===========================================================================
# STEP B
# ===========================================================================

doc.add_heading("B. Y7 Q2 Undiscounted Accumulated Value", level=1)
doc.add_paragraph(
    "At Y7 Q2 every projection year has landed and the grant has been received. No "
    "discounting applies, because nothing is being valued forward. The HK$0.58M "
    "capitalised intangible at Y6 is added because it sits on the balance sheet as "
    "an asset."
)
rows = []
for c in CASES:
    rows.append([
        c,
        f"{REPORTED_UNDISCOUNTED[c]:.2f}",
        f"{INTANGIBLE_Y6:.2f}",
        money(y7q2[c]),
        ps_fmt(y7q2[c]),
    ])
table(doc, ["Case", "A: cumulative Y1–Y6 + grant", "+ intangible",
            "B: Y7 Q2 value", "Per share"],
      rows, [2.5, 4.3, 2.5, 3.2, 3.2], highlight_col=3)
note(doc, "This is book value if every year lands exactly and all profit is retained. "
          "It is the upper bound of the range — the proven end.")
carry(doc, "the three Y7 Q2 values form the upper bound in Step D.")

# ===========================================================================
# STEP C
# ===========================================================================

doc.add_heading("C. Y0 Adjusted — The Dream Price", level=1)

doc.add_heading("C.1 The reported NPVs", level=2)
doc.add_paragraph(
    "The founder's NPVs, with the grant included at face value. These are the figures "
    "on record and must be reconciled before use."
)
rows = []
for c in CASES:
    rows.append([c] + [money(REPORTED_NPV[c][r]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("C.2 The two adjustments", level=2)
doc.add_paragraph(
    "The grant is received at Y7 Q2, not today, so it must be discounted from that "
    "date. The HK$0.58M intangible is an asset and must be added at its present "
    "value. Both adjustments use the same rate as the operating cash flows."
)
rows = []
for r in RATES:
    rows.append([
        f"{int(r*100)}%",
        f"{disc_g[r]:.4f}", f"{pv_grant[r]:.3f}",
        f"{disc_i[r]:.4f}", f"{pv_intangible[r]:.3f}",
    ])
table(doc, ["Rate", "DF grant (6.25y)", "PV grant",
            "DF intangible (6y)", "PV intangible"],
      rows, [2.0, 3.4, 3.0, 3.6, 3.2])

doc.add_heading("C.3 Reconciliation to the adjusted basis", level=2)
rows = []
for c in CASES:
    for r in RATES:
        rows.append([
            c, f"{int(r*100)}%",
            f"{REPORTED_NPV[c][r]:.2f}",
            f"−{GRANT:.2f}",
            f"+{pv_grant[r]:.3f}",
            f"+{pv_intangible[r]:.3f}",
            f"{y0[(c, r)]:.2f}",
        ])
table(doc, ["Case", "Rate", "Reported", "Remove grant",
            "Add PV grant", "Add PV intangible", "C: adjusted"],
      rows, [2.2, 1.5, 2.4, 2.8, 2.4, 2.8, 2.6], highlight_col=6)

doc.add_heading("C.4 The nine cells", level=2)
rows = []
for c in CASES:
    rows.append([c] + [money(y0[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)
note(doc, "The adjustment is small: −0.03 / −0.17 / −0.21 in the Ceiling case at "
          "25/50/75%. But the schedule is now internally consistent: every future item "
          "is discounted to the same Y0 date, at the same rate, over the correct "
          "number of years.")
carry(doc, "the nine adjusted cells form the lower bound in Step D.")

# ===========================================================================
# STEP D
# ===========================================================================

doc.add_heading("D. Midpoint of C and B — The Nine-Cell Matrix", level=1)
doc.add_paragraph(
    "For each case and rate, the midpoint of the adjusted Y0 value (Step C) and the "
    "Y7 Q2 accumulated value (Step B). This is the destination of the route: the "
    "valuation at the crossover between a dream and a proven business."
)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        row.append([money(mid[(c, r)]), ps_fmt(mid[(c, r)])])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)

doc.add_heading("D.1 Worked example — Centre at 50%", level=2)
rows = [
    ["C: Y0 adjusted (50%)", money(y0[("Centre", 0.50)])],
    ["B: Y7 Q2 accumulated", money(y7q2["Centre"])],
    ["Sum", money(y0[("Centre", 0.50)] + y7q2["Centre"])],
    ["D: midpoint (sum ÷ 2)", money(mid[("Centre", 0.50)])],
    ["Per share on 100 shares", ps_fmt(mid[("Centre", 0.50)])],
]
table(doc, ["Item", "Value"], rows, [7.0, 6.0], highlight_col=1)
note(doc, "Every other cell in the matrix is computed the same way. The arithmetic is "
          "trivial; what requires justification is Step E — whether the resulting "
          "number corresponds to a defensible discount rate.")
carry(doc, "the nine midpoint values are converted to implied rates in Step E and to "
           "share allocations in Step F.")

# ===========================================================================
# STEP E
# ===========================================================================

doc.add_heading("E. Implied Discount Rate for Each Cell", level=1)
doc.add_paragraph(
    "For each midpoint cell, the discount rate that would produce that value directly "
    "from the Y1–Y7 Q2 cash flows. This is the test of whether the midpoint is "
    "defensible: if the implied rates cluster around the rate appropriate for a "
    "business with an audited track record, the matrix is coherent. If they were "
    "wildly different across cells, the midpoint would be an artefact."
)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        ir = implied[(c, r)]
        row.append(f"{ir*100:.1f}%" if ir is not None else "—")
    rows.append(row)
table(doc, ["Case", "25% column", "50% column", "75% column"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("E.1 Reading the rate matrix", level=2)
bold_para(doc, [
    ("The 50% column is the centre of gravity. ", True),
    (f"Implied rates there are {implied[('Lower', 0.50)]*100:.0f}%, "
     f"{implied[('Centre', 0.50)]*100:.0f}% and {implied[('Ceiling', 0.50)]*100:.0f}% "
     f"for Lower, Centre and Ceiling. A rate of roughly 12% is what you apply to a "
     f"profitable private business with three years of audited accounts. It is not a "
     f"seed rate (50%) and not a mature-company rate (8–10%). It is the crossover "
     f"rate.", False),
])
bold_para(doc, [
    ("The rate falls as the discount column rises. ", True),
    ("This is not a coincidence. A higher discount rate produces a lower Y0 value in "
     "Step C, which pulls the midpoint down, which requires a lower rate to reproduce. "
     "The 25% column implies the highest rates and the 75% column the lowest. The "
     "middle column is the one to defend.", False),
])
bold_para(doc, [
    ("The three cases cluster at each column. ", True),
    ("Because the three P&L cases share the same cash-flow shape, they share the same "
     "implied rate at any given column. The rate is a property of time and required "
     "return, not of which forecast is chosen. That is the strength of the method — "
     "and also its weakness, addressed in Section G.", False),
])
carry(doc, "the rate matrix establishes that the midpoint is defensible. Step F "
           "converts the midpoint values into equity.")

# ===========================================================================
# STEP F
# ===========================================================================

doc.add_heading("F. Shares Allocated to the VC for HK$1.0M", level=1)
doc.add_paragraph(
    f"At each of the nine midpoint pre-money values, a {money(SEED_CHEQUE)} cheque "
    f"converts to a share allocation. Two conventions are shown because they are "
    f"routinely confused, and the difference is material: pricing off pre-money only "
    f"over-allocates shares to the investor."
)

doc.add_heading("F.1 Priced off pre-money (the trap)", level=2)
doc.add_paragraph(
    "Shares issued = 100 × cheque ÷ pre-money. This is wrong because it ignores the "
    "dilution the cheque itself creates. The investor ends up with a larger stake "
    "than the price implies."
)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        row.append([f"{shares_pre[(c, r)]:.2f}", 
                    f"{shares_pre[(c, r)]:.2f}%"])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)

doc.add_heading("F.2 Priced off post-money (correct)", level=2)
doc.add_paragraph(
    "Shares issued = 100 × cheque ÷ (pre-money + cheque). The cheque is priced after "
    "its own dilution. This is the standard convention."
)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        row.append([f"{shares_post[(c, r)]:.2f}",
                    f"{shares_post[(c, r)]:.2f}%"])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)

doc.add_heading("F.3 The 18 cells side by side", level=2)
doc.add_paragraph(
    "Shares to the VC under each convention, across all nine midpoint cells. The gap "
    "between the two is the cost of using the pre-money convention."
)
rows = []
for c in CASES:
    for r in RATES:
        pre_sh = shares_pre[(c, r)]
        post_sh = shares_post[(c, r)]
        rows.append([
            c, f"{int(r*100)}%",
            money(mid[(c, r)]),
            f"{pre_sh:.2f}",
            f"{post_sh:.2f}",
            f"{pre_sh - post_sh:.2f}",
        ])
table(doc, ["Case", "Rate", "Midpoint pre-money",
            "VC shares (pre-money priced)",
            "VC shares (post-money priced)",
            "Over-allocation"],
      rows, [2.2, 1.5, 3.2, 3.8, 3.8, 3.0], highlight_col=4)
note(doc, "The over-allocation is largest where the pre-money is smallest — the "
          "Lower × 75% cell. In the Centre × 50% cell the gap is "
          f"{shares_pre[('Centre', 0.50)] - shares_post[('Centre', 0.50)]:.2f} "
          "shares, or "
          f"{(shares_pre[('Centre', 0.50)] - shares_post[('Centre', 0.50)]) / shares_post[('Centre', 0.50)] * 100:.1f}% "
          "more equity than the price implies.")

doc.add_heading("F.4 Founder and investor position at each midpoint cell", level=2)
doc.add_paragraph(
    "Post-issue shareholding on the correct (post-money) convention. Founder retains "
    "100 minus the VC allocation in every cell."
)
rows = []
for c in CASES:
    for r in RATES:
        vc_sh = shares_post[(c, r)]
        rows.append([
            c, f"{int(r*100)}%",
            f"{vc_sh:.2f}",
            f"{vc_sh:.2f}%",
            f"{SHARE_COUNT - vc_sh:.2f}",
            f"{SHARE_COUNT - vc_sh:.2f}%",
        ])
table(doc, ["Case", "Rate", "VC shares", "VC %",
            "Founder shares", "Founder %"],
      rows, [2.2, 1.5, 2.8, 2.5, 3.2, 2.8], highlight_col=3)
carry(doc, "the share allocations at the four negotiating positions are set out in "
           "Step G.")

# ===========================================================================
# STEP G
# ===========================================================================

# ---- Section G divider — DO NOT DISTRIBUTE ----
doc.add_page_break()

stop_bar = doc.add_paragraph()
stop_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
stop_bar.paragraph_format.space_before = Pt(0)
r = stop_bar.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_title = doc.add_paragraph()
stop_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_title.add_run("STOP  ·  SECTION G BEGINS HERE")
r.bold = True
r.font.size = Pt(14)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_warn = doc.add_paragraph()
stop_warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_warn.add_run(
    "DO NOT DISTRIBUTE THIS SECTION.\n"
    "Contains founder positions, counterpart responses, and negotiation scripts.\n"
    "Strip this section in full before any investor communication."
)
r.italic = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_bar2 = doc.add_paragraph()
stop_bar2.alignment = WD_ALIGN_PARAGRAPH.CENTER
stop_bar2.paragraph_format.space_after = Pt(14)
r = stop_bar2.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_heading("G. Positioning and Counterpart", level=1)

doc.add_heading("G.1 Founder positions", level=2)
rows = []
basis_text = {
    "Opening ask": "Midpoint Ceiling × 50% — upside case, supported by the rate matrix",
    "Target": "Midpoint Centre × 50% — base case, implied rate ≈ 12%",
    "Settle": "Midpoint Lower × 50% — conservative but coherent",
    "Walk-away": "Midpoint Lower × 75% — floor; below this, decline",
}
for k in ["Opening ask", "Target", "Settle", "Walk-away"]:
    v = pos[k]
    rows.append([
        k,
        money(v),
        ps_fmt(v),
        f"{shares_post[(POSITIONS[k][0], POSITIONS[k][1])]:.2f}",
        f"{stake_post(v)*100:.2f}%",
        basis_text[k],
    ])
table(doc, ["Position", "Pre-money", "Per share",
            "VC shares of 100", "VC %", "Basis"],
      rows, [2.6, 2.8, 2.8, 2.8, 2.2, 6.0], highlight_col=4, shade_rows=[3])
note(doc, f"Every position is a cell in the Step D matrix. The founder is not "
          f"inventing numbers in Step G — they are selecting cells that Step E has "
          f"already shown to be defensible.")

doc.add_heading("G.2 Counterpart positions and how to answer", level=2)
rows = [
    [f"{INVESTOR_ASK*100:.2f}% stake at HK$1.0M",
     money(ask_pre),
     f"{INVESTOR_ASK*100:.2f}%",
     "Below every cell in the Step D matrix, including the most conservative. The ask "
     "does not correspond to any forecast at the crossover rate. Name this first."],
    ["Y0 dream frame",
     money(y0[("Centre", 0.50)]),
     f"{stake_post(y0[('Centre', 0.50)])*100:.2f}%",
     "Prices the company as though the Y1–Y3 audits will not exist. Factually wrong "
     "for a Y4 Q1 round. Ask which forecast it encodes."],
    ["Lower × 75% (the old anchor)",
     money(mid[("Lower", 0.75)]),
     f"{stake_post(mid[('Lower', 0.75)])*100:.2f}%",
     "Below the walk-away. Decline — the company is profitable by Y3 and does not "
     "need the cheque."],
    ["50% discount rate applied to Y4 Q1",
     "—", "—",
     "The rate is a seed rate for an unproven business. At Y4 Q1 the company has "
     "three years of audits. Argue the rate, not the forecast."],
]
table(doc, ["Counterpart position", "Implied pre-money", "VC stake", "Response"],
      rows, [4.0, 3.0, 2.2, 7.0], highlight_col=2)

doc.add_heading("G.3 The three arguments, in order", level=2)

doc.add_paragraph("Argument 1 — the rate.").runs[0].bold = True
quote_para(doc, f"\"The midpoint at the Centre case implies a "
                f"{implied[('Centre', 0.50)]*100:.0f}% discount rate. That is the rate "
                f"you apply to a profitable business with three years of audited "
                f"accounts. It is not a seed rate. If you want to argue 50%, you are "
                f"arguing that the audits will not exist — and the audit is the one "
                f"fact we both can verify.\"")

doc.add_paragraph("Argument 2 — the forecast.").runs[0].bold = True
quote_para(doc, f"\"Your {INVESTOR_ASK*100:.2f}% implies {money(ask_pre)}. That is "
                f"below every cell in the matrix — including the Lower case at a 75% "
                f"rate, which still values the company at {money(mid[('Lower', 0.75)])} "
                f"and gives you {stake_post(mid[('Lower', 0.75)])*100:.2f}%. So the "
                f"question is not whether my number is too high. It is which forecast "
                f"your stake encodes. Because it encodes none of them.\"")

doc.add_paragraph("Argument 3 — the walk.").runs[0].bold = True
quote_para(doc, "\"If we cannot bridge, I would rather not raise. The company is "
                "profitable by Y3 and the accumulated cash means the HK$1.0M is for "
                "acceleration, not survival. I would rather retain earnings for "
                "twelve months and return at the proven end of the range than accept "
                "a dream-frame price after the dream has been executed.\"")

doc.add_heading("G.4 What to extract if the stake holds", level=2)
bullet(doc, [("Non-participating preference only. ", True),
             ("1× non-participating. The gap to participating is worth more than the "
              "stake concession at every midpoint cell.", False)])
bullet(doc, [("No board seat. ", True),
             (f"At {stake_post(pos['Target'])*100:.2f}% (the Target), the investor "
              f"should have information rights and an observer seat at most.", False)])
bullet(doc, [("Pro-rata, not super pro-rata. ", True),
             ("Standard pre-emptive rights only. Keeps the cap table clean.", False)])
bullet(doc, [("Follow-on price band agreed now. ", True),
             ("Lock in the next-round band before pilot data lands. Converts a price "
              "concession into a relationship.", False)])

doc.add_heading("G.5 Market Reality Check — The 10% Ask", level=2)
doc.add_paragraph(
    "The matrix in Step D is a planning tool. The market will not price against it. "
    "A realistic seed offer for HK$1.0M is 10% of the company — HK$9.0M pre-money, "
    "HK$90,000 per share. That is below every cell in the Step D matrix."
)

rows = [
    ["Y0 Lower × 75% (absolute floor)",
     money(y0[("Lower", 0.75)]),
     f"{stake_post(y0[('Lower', 0.75)])*100:.2f}%",
     "The only Y0 benchmark below the market offer"],
    ["Market offer (10%)", money(9.00), "10.0%",
     "Fund-arithmetic price, not a DCF output"],
    ["Y0 Lower × 50%", money(y0[("Lower", 0.50)]),
     f"{stake_post(y0[('Lower', 0.50)])*100:.2f}%",
     "What the Lower case implies at a standard seed rate"],
    ["Midpoint Lower × 75%", money(mid[("Lower", 0.75)]),
     f"{stake_post(mid[('Lower', 0.75)])*100:.2f}%",
     "Lowest matrix cell; implies 1.99% for HK$1M versus the market's 10%"],
    ["Midpoint Centre × 50% (target)", money(mid[("Centre", 0.50)]),
     f"{stake_post(mid[('Centre', 0.50)])*100:.2f}%",
     "Crossover target; implies 0.74% for HK$1M versus the market's 10%"],
]
table(doc, ["Basis", "Pre-money", "HK$1M buys", "Comment"], rows,
      [5.0, 2.8, 2.5, 6.5], highlight_col=1, shade_rows=[1])

note(doc, "The market offer corresponds to roughly a 70% discount rate applied to "
          "the Lower case — higher than any rate in the matrix. The matrix was "
          "never designed to contain this cell, because it is not a valuation. It "
          "is fund arithmetic: a HK$1.0M cheque in a HK$100M fund needs 10% to be "
          "worth managing.")

doc.add_heading("G.6 What this means for the negotiation", level=2)
bullet(doc, [("Do not argue the DCF against the 10% ask. ", True),
             ("The ask is not a DCF output. Argue fund structure instead: smaller "
              "cheque, tranched cheques, or a follow-on price band agreed now.",
              False)])
bullet(doc, [("Accept the 10% if the terms are clean. ", True),
             (f"A HK$9.0M pre-money with non-participating preference and no board "
              f"seat is a manageable outcome. It sets a reference price, but a "
              f"clean one.", False)])
bullet(doc, [("Reduce the cheque if HK$1.0M is more than the plan needs. ", True),
             ("HK$500K at HK$9.0M pre-money is 5.3% dilution — half the equity for "
              "half the runway. This is the single strongest lever available.",
              False)])
bullet(doc, [("Reserve the matrix for the next round. ", True),
             ("The midpoint cells are what the company should be worth at Y5 or Y6, "
              "after the pilot and the audit. Use them to plan the step-up, not to "
              "fight the current round.", False)])

# ===========================================================================
# 7. Caveats and recommendation
# ===========================================================================

# ---- End of Section G ----
end_bar = doc.add_paragraph()
end_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
end_bar.paragraph_format.space_before = Pt(18)
r = end_bar.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

end_title = doc.add_paragraph()
end_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = end_title.add_run("END OF SECTION G  ·  INTERNAL CONTENT")
r.bold = True
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

end_bar2 = doc.add_paragraph()
end_bar2.alignment = WD_ALIGN_PARAGRAPH.CENTER
end_bar2.paragraph_format.space_after = Pt(18)
r = end_bar2.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_page_break()
doc.add_heading("7. Caveats", level=1)
bullet(doc, [("The three cases are scalar multiples. ", True),
             ("From Y2 onward Centre is approximately 2.57× Lower and Ceiling "
              "approximately 6.25× Lower in every year. The cases vary size, not "
              "shape. This is visible in Step E: the implied rates cluster because "
              "the cash flows are proportional. Rebuild the Lower case with a "
              "distinct cost structure before external circulation — it is the "
              "single change that most strengthens the matrix.", False)])
bullet(doc, [("The midpoint is an average, not a method. ", True),
             ("What gives it force is Step E. The implied rates are defensible; the "
              "averaging is the shortcut used to reach them. If challenged, present "
              "the rate, not the arithmetic.", False)])
bullet(doc, [("The grant and intangible treatment must be consistent. ", True),
             ("Step C discounts the grant from Y7 Q2 and the intangible from Y6. "
              "Step B includes both at face. Confirm with the auditor that the "
              "capitalised amount is the amortised carrying value, not gross R&D "
              "spend.", False)])
bullet(doc, [("Per-share values assume 100 shares. ", True),
             ("Confirm the authorised count with counsel before quoting externally. "
              "The absolute per-share price carries no meaning; only the proportion "
              "each share represents.", False)])
bullet(doc, [("Three years of audits must actually land. ", True),
             ("The entire matrix is conditional on Y1–Y3 being audited and profitable. "
              "If they underperform, the applicable rate rises and the midpoint falls.",
              False)])

doc.add_heading("8. Recommendation", level=1)
bold_para(doc, [
    ("Walk the route in the room. ", True),
    ("Do not lead with the nine-cell matrix. Lead with Step A — the P&L — then Step B, "
     "then Step C. By the time Step D appears, the investor has watched it being "
     "built. A number that is derived is harder to dispute than a number that is "
     "announced.", False),
])
bold_para(doc, [
    ("Lead with the implied rate. ", True),
    (f"\"The midpoint implies {implied[('Centre', 0.50)]*100:.0f}%\" is a stronger "
     f"opening than any valuation figure. It engages the investor on method. Method "
     f"arguments are winnable; number arguments are not.", False),
])
bold_para(doc, [
    ("Use the midpoint cells for positions. ", True),
    (f"Open {money(pos['Opening ask'])}, target {money(pos['Target'])}, settle "
     f"{money(pos['Settle'])}, walk below {money(pos['Walk-away'])}. Every position "
     f"is a cell that Step E has already shown to be defensible.", False),
])
bold_para(doc, [
    ("Before the meeting. ", True),
    ("Confirm the audit engagement type; rebuild the Lower case with a distinct cost "
     "structure; and decide whether the company can decline HK$1.0M if the terms do "
     "not land at or above the walk-away. The third is what makes Step G's walk "
     "credible.", False),
])

out = "Valuation_Route_v13_INTERNAL.docx"
doc.save(out)

# ---- Console
print(f"Saved: {out}\n")
print("A/B — cumulative and Y7 Q2:")
for c in CASES:
    print(f"  {c:<8} cumulative {REPORTED_UNDISCOUNTED[c]:>7.2f}  "
          f"+intang {INTANGIBLE_Y6:.2f}  = {y7q2[c]:>8.2f}")
print("\nC — Y0 adjusted (9 cells):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {y0[(c, r)]:>7.2f}"
    print(line)
print("\nD — Midpoint (9 cells):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {mid[(c, r)]:>7.2f}"
    print(line)
print("\nE — Implied rate (9 cells):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        ir = implied[(c, r)]
        line += f"  {int(r*100):>3}%: {(f'{ir*100:>5.1f}%' if ir else '   n/a')}"
    print(line)
print("\nF — VC shares, post-money priced (9 cells):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {shares_post[(c, r)]:>6.2f}"
    print(line)
print("\nF — VC shares, pre-money priced (9 cells):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {shares_pre[(c, r)]:>6.2f}"
    print(line)
print("\nG — Positions:")
for k in ["Opening ask", "Target", "Settle", "Walk-away"]:
    v = pos[k]
    print(f"  {k:<12} {money(v):>12}  {ps_fmt(v):>16}  "
          f"VC {shares_post[(POSITIONS[k][0], POSITIONS[k][1])]:>6.2f} shares "
          f"({stake_post(v)*100:.2f}%)")
print(f"\nInvestor ask {INVESTOR_ASK*100:.2f}% -> pre {money(ask_pre)}")