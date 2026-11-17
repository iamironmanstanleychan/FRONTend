"""
Kung Fu AI Football — Pre-money Valuation and Negotiation Memo (.docx)
Requires: pip install python-docx

v8 — separates the Y0 pilot-funding instrument from the post-pilot seed,
     and adds a readiness diagnostic before the anchor defence.
"""

import math
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
GRANT_YEARS = 6.25
INTANGIBLE = 0.58
INTANGIBLE_YEARS = 6
RATES = [0.25, 0.50, 0.75]
CASES = ["Ceiling", "Centre", "Lower"]

SHARE_COUNT = 100
RAISE_Y0_HKD_M = 1.5
RAISE_Y1_HKD_M = 8.0
Y0_ANGEL_PRE = 30.0
Y1_SEED_PRE = 37.34
RAISE_HKD_M = RAISE_Y1_HKD_M

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


def num_bullet(doc, segs):
    p = doc.add_paragraph(style="List Number")
    for text, b in segs:
        r = p.add_run(text); r.bold = b
    return p


def quote_para(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(10)
    return p


def note(doc, text, before=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(before)
    r = p.add_run(text); r.italic = True; r.font.size = Pt(9)
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


def own(pre, raise_m):
    return raise_m / (pre + raise_m)


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

disc_grant = {r: 1 / (1 + r) ** GRANT_YEARS for r in RATES}
disc_intangible = {r: 1 / (1 + r) ** INTANGIBLE_YEARS for r in RATES}
pv_grant = {r: GRANT * disc_grant[r] for r in RATES}
pv_intangible = {r: INTANGIBLE * disc_intangible[r] for r in RATES}

adjusted_y0 = {}
for c in CASES:
    for r in RATES:
        adjusted_y0[(c, r)] = (
            REPORTED_NPV[c][r] - GRANT + pv_grant[r] + pv_intangible[r]
        )

y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE for c in CASES}

ratio = {(c, r): y7q2[c] / adjusted_y0[(c, r)] for c in CASES for r in RATES}
duration = {r: math.log(ratio[("Centre", r)]) / math.log(1 + r) for r in RATES}

POSITIONS = {
    "Opening ask":  ("Ceiling", 0.50),
    "Target":       ("Centre",  0.50),
    "Reservation":  ("Centre",  0.75),
    "Worst cell":   ("Lower",   0.75),
}
pos = {k: adjusted_y0[v] for k, v in POSITIONS.items()}

raise_for_20 = {k: 0.25 * pos[k] for k in pos}

# Series A dilution: assume 3x post-seed, HK$15M raise
SERIES_A_RAISE = 15.0
SERIES_A_MULT = 3.0
after_a = {}
for k in pos:
    post_seed = pos[k] + RAISE_HKD_M
    a_pre = post_seed * SERIES_A_MULT
    a_post = a_pre + SERIES_A_RAISE
    founder_seed = 1 - own(pos[k], RAISE_HKD_M)
    after_a[k] = founder_seed * (post_seed / a_post)

# Alternative financing path: angel HK$1.5M at HK$30M, pilot, seed HK$8M at HK$50M
ANGEL_PRE, ANGEL_RAISE = Y0_ANGEL_PRE, RAISE_Y0_HKD_M
NEXT_SEED_PRE, NEXT_SEED_RAISE = 50.0, 8.0
alt_founder = (
    (1 - ANGEL_RAISE / (ANGEL_PRE + ANGEL_RAISE))
    * (1 - NEXT_SEED_RAISE / (NEXT_SEED_PRE + NEXT_SEED_RAISE))
)

# ===========================================================================
# BUILD
# ===========================================================================

doc = Document()
st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("MEMORANDUM"); r.bold = True; r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()
for label, value in [
    ("To:", "Founders / Board"),
    ("From:", "Finance & Strategy"),
    ("Re:", "Pre-money Valuation and Negotiation Position — Kung Fu AI Football"),
    ("Status:", "INTERNAL — negotiation preparation. Not for external distribution."),
    ("Version:", "v8 — dual-date valuation, Y0/Y1 funding plan, readiness diagnostic, and anchor defence."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    p.add_run(value)

doc.add_paragraph()

# ===========================================================================
# 1. Purpose and Summary
# ===========================================================================
doc.add_heading("1. Purpose and Summary", level=1)
doc.add_paragraph(
    "This memo sets out the per-share pre-money valuation on two dates, explains how "
    "the two relate, and lays out a negotiation strategy built around the specific "
    "risk that an investor anchors on the lowest cell in the matrix — Lower case at a "
    "75% discount rate, or HK$7.80M pre-money."
)
bold_para(doc, [
    ("Pricing basis. ", True),
    (f"At the Centre case and the 50% standard seed rate, Y0 pre-money is "
     f"{money(pos['Target'])} — {ps_fmt(pos['Target'])} per share on {SHARE_COUNT} "
     f"shares. This is the number that belongs in a term sheet.", False),
])
bold_para(doc, [
    ("Reference point. ", True),
    (f"At Y7 Q2, accumulated value in the Centre case is {money(y7q2['Centre'])} — "
     f"{ps_fmt(y7q2['Centre'])} per share. This is a book value, not a valuation, and "
     f"should never be quoted as a pre-money figure.", False),
])
bold_para(doc, [
    ("The risk to manage. ", True),
    (f"If the investor anchors on Lower × 75% = {money(pos['Worst cell'])}, an "
     f"{money(RAISE_HKD_M)} raise delivers them "
     f"{own(pos['Worst cell'], RAISE_HKD_M)*100:.1f}% of the company. That is a "
     f"control transaction disguised as seed capital. Section 13 is devoted entirely "
     f"to this scenario.", False),
])

# ===========================================================================
# 2. Source P&L
# ===========================================================================
doc.add_heading("2. Source P&L Projections", level=1)
doc.add_paragraph(
    "Net profit, HK$ millions, after all operating costs, HKAS 38 amortisation and "
    "HKSAR taxes. Three cases differ by marketing cost assumption and adoption band."
)
rows = [[r[0], f"{r[1]:.2f}", f"{r[2]:.2f}", f"{r[3]:.2f}"] for r in DATA]
rows.append(["Y7 Q2 grant", f"{GRANT:.2f}", f"{GRANT:.2f}", f"{GRANT:.2f}"])
rows.append(["Cumulative (undiscounted)",
             f"{sum(r[1] for r in DATA):.2f}",
             f"{sum(r[2] for r in DATA):.2f}",
             f"{sum(r[3] for r in DATA):.2f}"])
table(doc, ["Year", "Lower (40% mkt / P10)", "Centre (35% mkt / P50)",
            "Ceiling (30% mkt / P90)"],
      rows, [4.0, 4.3, 4.3, 4.3], shade_rows=[7])
note(doc, "HK$0.58M Y6 capitalised intangible sits on the balance sheet separately "
          "and is added at the Y7 Q2 reference point; it is not a profit line.")

# ===========================================================================
# 3. The Two Valuation Dates
# ===========================================================================
doc.add_heading("3. The Two Valuation Dates", level=1)
bold_para(doc, [
    ("Y0. ", True),
    ("The seed round closes today. Every future cash flow is discounted back to this "
     "date at a rate that prices the risk the investor is taking. The 50% column is "
     "the standard seed target return; 75% prices execution risk at the maximum; 25% "
     "is an upside case.", False),
])
bold_para(doc, [
    ("Y7 Q2. ", True),
    ("The end of the forecast. Every projection year has landed, the grant has been "
     "received, the intangible is on the balance sheet. No discounting applies, "
     "because nothing is being valued forward. The figure shows what the company "
     "accumulates if the plan is executed exactly as modelled.", False),
])
bold_para(doc, [
    ("The gap between them is the value of time and risk. ", True),
    (f"Centre case: {money(y7q2['Centre'])} at Y7 Q2 discounts to "
     f"{money(pos['Target'])} at Y0 on a 50% rate — a factor of "
     f"{ratio[('Centre', 0.50)]:.2f}×. That difference is what the investor earns "
     f"for taking six years of execution risk.", False),
])

# ===========================================================================
# 4. Y0 As Calculated
# ===========================================================================
doc.add_heading("4. Y0 — As Calculated (founder's figures)", level=1)
doc.add_paragraph(
    "The reported NPVs, with the grant included at face value. Reproduced exactly "
    "for reconciliation against Section 5."
)
rows = []
for c in CASES:
    rows.append([c] + [money(REPORTED_NPV[c][r]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("Per share on 100 shares", level=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(REPORTED_NPV[c][r]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ===========================================================================
# 5. Y0 Adjusted
# ===========================================================================
doc.add_heading("5. Y0 — Adjusted to a Consistent Basis", level=1)
doc.add_paragraph(
    "The reported NPVs treat the HK$0.24M grant as if received today. It is received "
    "at Y7 Q2 and must be discounted. They also omit the HK$0.58M capitalised "
    "intangible, which is an asset and should be counted. The adjustment makes both "
    "consistent with the operating cash flows, which are already discounted year by "
    "year."
)

doc.add_heading("Discount factors and present values", level=2)
rows = []
for r in RATES:
    rows.append([
        f"{int(r*100)}%",
        f"{disc_grant[r]:.4f}",
        f"{pv_grant[r]:.3f}",
        f"{disc_intangible[r]:.4f}",
        f"{pv_intangible[r]:.3f}",
    ])
table(doc, ["Rate", "DF grant (6.25y)", "PV grant",
            "DF intangible (6y)", "PV intangible"],
      rows, [2.0, 3.2, 2.8, 3.5, 3.0])

doc.add_heading("Reconciliation", level=2)
rows = []
for c in CASES:
    for r in RATES:
        rows.append([
            c, f"{int(r*100)}%",
            f"{REPORTED_NPV[c][r]:.2f}",
            f"−{GRANT:.2f}",
            f"+{pv_grant[r]:.3f}",
            f"+{pv_intangible[r]:.3f}",
            f"{adjusted_y0[(c, r)]:.2f}",
        ])
table(doc, ["Case", "Rate", "Reported", "Less grant at face",
            "Add PV grant", "Add PV intangible", "Adjusted"],
      rows, [2.2, 1.5, 2.4, 3.2, 2.4, 2.8, 2.4], highlight_col=6)

doc.add_heading("Adjusted Y0 pre-money and per share", level=2)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        v = adjusted_y0[(c, r)]
        row.append([money(v), ps_fmt(v)])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)
note(doc, "The adjustment is small: −0.03 at 25%, −0.17 at 50%, −0.21 at 75% "
          "relative to the reported figures. No conclusion changes.")

# ===========================================================================
# 6. Y7 Q2
# ===========================================================================
doc.add_heading("6. Y7 Q2 — Undiscounted Accumulated Value", level=1)
doc.add_paragraph(
    "At Y7 Q2 everything in the forecast has been realised. No discounting is "
    "applied. The figures are cumulative undiscounted profit (grant included) plus "
    "the capitalised intangible."
)
rows = []
for c in CASES:
    rows.append([
        c,
        f"{REPORTED_UNDISCOUNTED[c]:.2f}",
        f"{INTANGIBLE:.2f}",
        f"{y7q2[c]:.2f}",
        ps_fmt(y7q2[c]),
    ])
table(doc, ["Case", "Cumulative Y1–Y6 + grant", "Intangible",
            "Total at Y7 Q2", "Per share"],
      rows, [2.5, 4.3, 2.3, 3.2, 3.5], highlight_col=3)
note(doc, "This is book value of accumulated assets, not enterprise value. No "
          "multiple, no discount, no probability. Quoting it as a pre-money figure "
          "invites the response that the company does not need external capital.")

# ===========================================================================
# 7. Per-Share Summary
# ===========================================================================
doc.add_heading("7. Per-Share Summary — All Bases", level=1)
doc.add_paragraph(
    f"Price per share on {SHARE_COUNT} authorised shares. A 100-share structure is a "
    f"seed-stage convention; the absolute figure carries no meaning, only the "
    f"proportion each share represents."
)
rows = []
for r in RATES:
    rows.append([f"Y0 adjusted — {int(r*100)}%"] +
                [ps_fmt(adjusted_y0[(c, r)]) for c in CASES])
rows.append(["Y7 Q2 — undiscounted"] + [ps_fmt(y7q2[c]) for c in CASES])
table(doc, ["Basis", "Ceiling", "Centre", "Lower"], rows,
      [5.0, 3.8, 3.8, 3.8], shade_rows=[3])

# ===========================================================================
# 8. Value Creation Multiple
# ===========================================================================
doc.add_heading("8. Value Creation Multiple — Y7 Q2 over Y0", level=1)
doc.add_paragraph(
    "The ratio of Y7 Q2 accumulated value to Y0 present value answers a simple "
    "question: for each HK$1 of pre-money paid at Y0, how much nominal cash does the "
    "business accumulate by Y7 Q2? It is a consistency check on the schedule, not a "
    "return metric."
)
rows = [[c] + [f"{ratio[(c, r)]:.2f}×" for r in RATES] for c in CASES]
table(doc, ["Case", "25% rate", "50% rate", "75% rate"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)
note(doc, "The frequently quoted figures of 1161%, 619% and 282% are arithmetically "
          "correct but mix three rates: Lower at 75%, Centre at 50%, Ceiling at 25%. "
          "Read within a rate column, the multiple is essentially identical across "
          "all three cases.")

doc.add_heading("Why the multiple is rate-determined, not case-determined", level=2)
doc.add_paragraph(
    f"The ratio is (1 + r)^D where D is cash-flow duration. Implied duration is "
    f"{duration[0.25]:.2f} years at 25%, {duration[0.50]:.2f} years at 50% and "
    f"{duration[0.75]:.2f} years at 75%. Because the three cases share nearly the "
    f"same cash-flow shape, they share the same multiple at any given rate. The "
    f"multiple is a property of time and required return, not of which scenario is "
    f"chosen."
)

doc.add_heading("What it is not", level=2)
bullet(doc, [("Not MOIC. ", True),
             ("An investor's multiple depends on the fraction of the company acquired "
              "and on exit value. At the 50% target over 7.25 years, required MOIC is "
              "approximately 1.5^7.25 ≈ 19× — achieved via a 10× earnings exit, not "
              "via liquidation at book.", False)])
bullet(doc, [("Not IRR. ", True),
             ("The discount rate is the IRR by construction. The ratio restates the "
              "same fact over the cash-flow duration rather than the full horizon.", False)])

doc.add_heading("The diagnostic finding", level=2)
doc.add_paragraph(
    "Near-identical ratios across the three cases confirm that the cases differ in "
    "scale, not in risk or timing. A genuinely distinct downside case — slower ramp, "
    "higher CAC, lost cohort — would produce a different ratio at the same rate. The "
    "current Lower case does not. Rebuild it with a different structure before "
    "diligence; an analyst running this ratio test will reach the same conclusion in "
    "five minutes."
)

# ===========================================================================
# 9. Explanation
# ===========================================================================
doc.add_heading("9. Explanation — Reading the Numbers", level=1)
bold_para(doc, [
    ("The \"0\" column is not a valuation. ", True),
    ("It is the sum of nominal cash flows. It has no time value and applies no risk. "
     "Its only use is as the Y7 Q2 base, before adding the intangible.", False),
])
bold_para(doc, [
    ("The 25/50/75% columns are present values. ", True),
    ("They answer: what is the entire future stream worth today at this required "
     "return? At 50%, HK$1 received in six years is worth HK$0.088 today.", False),
])
bold_para(doc, [
    ("Per-share values are pre-money ÷ 100. ", True),
    ("If the authorised share count is later set at 1,000 or 10,000, every per-share "
     "figure divides by 10 or 100 and nothing else changes.", False),
])
bold_para(doc, [
    ("Why Y7 Q2 is much larger. ", True),
    (f"It is undiscounted and includes six years of profit. Centre: "
     f"{money(y7q2['Centre'])} at Y7 Q2 versus {money(pos['Target'])} at Y0 on a "
     f"50% rate. The difference is what an investor earns for waiting.", False),
])

# ===========================================================================
# 10. Intangible
# ===========================================================================
doc.add_heading("10. The Intangible Asset", level=1)
bold_para(doc, [
    ("What it is. ", True),
    ("HK$0.58M of HKAS 38 capitalised R&D carried on the balance sheet at Y6. An "
     "accounting recognition of spending already incurred, not new cash.", False),
])
bold_para(doc, [
    ("How it is treated here. ", True),
    ("Added to operating NPV at present value (HK$0.152M / 0.051M / 0.020M at "
     "25/50/75%) and added at face value to the Y7 Q2 reference.", False),
])
bold_para(doc, [
    ("Why that is defensible. ", True),
    ("A DCF values the cash flows the IP generates; adding the capitalised cost of "
     "creating it risks double-counting. The adjustment is defensible because the "
     "DCF covers Y1–Y6 operating profit only and assigns no terminal value — no part "
     "of the forecast captures the balance-sheet IP separately.", False),
])
bold_para(doc, [
    ("The defence if challenged. ", True),
    ("It is a balance-sheet floor argument, not an earnings argument. HK$0.58M is "
     "capitalised R&D already spent and carried as an asset. At Y0 it floors "
     "enterprise value at accumulated IP, independent of whether the Y1–Y6 forecast "
     "is believed. It supports the reservation price, not the opening ask.", False),
])

# ===========================================================================
# 11. General Negotiation
# ===========================================================================
doc.add_heading("11. Negotiation — General Positions", level=1)

doc.add_heading("11.1 Positions", level=2)
rows = []
basis_text = {
    "Opening ask": "Ceiling × 50% — upside case at standard seed return",
    "Target": "Centre × 50% — base case at standard seed return",
    "Reservation": "Centre × 75% — base case priced for execution risk",
    "Worst cell": "Lower × 75% — do not accept (Section 13)",
}
for k in ["Opening ask", "Target", "Reservation", "Worst cell"]:
    v = pos[k]
    rows.append([k, money(v), ps_fmt(v),
                 f"{own(v, RAISE_HKD_M)*100:.1f}%", basis_text[k]])
table(doc, ["Position", "Pre-money at Y0", "Per share",
            "Investor ownership at HK$8M", "Basis"],
      rows, [3.0, 3.0, 3.2, 3.6, 5.7], highlight_col=1, shade_rows=[3])

doc.add_heading("11.2 Round size — two instruments, two dates", level=2)
bold_para(doc, [
    ("Y0. ", True),
    (f"Raise {money(RAISE_Y0_HKD_M)} at {money(Y0_ANGEL_PRE)} pre-money "
     f"(≈{own(Y0_ANGEL_PRE, RAISE_Y0_HKD_M)*100:.1f}% dilution). This funds the "
     f"5–10 academy pilot and the paired dataset. It is deliberately small: an "
     f"HK$8M cheque into a pre-pilot company is not a realistic Y0 event, and "
     f"pricing one now would set a low reference that the pilot cannot repair.", False),
])
bold_para(doc, [
    ("Y1. ", True),
    (f"Raise {money(RAISE_Y1_HKD_M)} at the Target of {money(Y1_SEED_PRE)} pre-money "
     f"once pilot data exists (≈{own(Y1_SEED_PRE, RAISE_Y1_HKD_M)*100:.1f}% "
     f"dilution). This is the round the valuation matrix in Sections 4–7 is written "
     f"for. It should not be attempted before the coach-agreement benchmarks land.", False),
])

doc.add_heading("11.3 Principles", level=2)
bullet(doc, [("Anchor on the discount rate, not the forecast. ", True),
             ("50% is the standard seed target return. If the investor argues for a "
              "higher rate, they are arguing about risk, not about the model — a "
              "narrower and more winnable debate than one about profit forecasts.", False)])
bullet(doc, [("Disclose the adjustment. ", True),
             ("Show both the reported NPVs and the adjusted figures. Volunteering the "
              "small downward reconciliation builds credibility for everything else.", False)])
bullet(doc, [("Use Y7 Q2 as narrative, never as an ask. ", True),
             (f"Centre-case {money(y7q2['Centre'])} accumulated by Y7 Q2 justifies "
              f"the Ceiling column and the exit multiple. Introduced as a pre-money "
              f"number, it invites the reply that the company does not need capital.", False)])
bullet(doc, [("Trade terms, not price. ", True),
             ("If pushed, concede valuation only for structure: milestone tranching, "
              "a smaller option pool, or a lower liquidation preference.", False)])
bullet(doc, [("Let the pilot set the price. ", True),
             ("Each month of pilot data moves the case from Centre toward Ceiling. "
              "If runway allows the wait, the worst cell becomes indefensible for the "
              "investor.", False)])

# ===========================================================================
# 12. Y0 risk and readiness
# ===========================================================================
doc.add_heading("12. The Y0 Risk Is the Pass, Not the Anchor", level=1)
doc.add_paragraph(
    "The negotiation framework in this section addresses the least likely Y0 outcome. "
    "The dominant risk at Y0 is that no priced round is offered at all, because an "
    "investor with HK$8M has alternatives that already carry proof. Planning should "
    "reflect the probability distribution, not the most colourful scenario."
)
rows = [
    ["Pass, no term sheet", "Most likely",
     "No negotiation. Pilot-first path becomes the plan."],
    ["Smaller instrument (HK$1–2M SAFE)", "Likely",
     "Sections 11.2 (Y0) applies. Cap table stays clean."],
    ["Priced seed at HK$20–30M", "Possible",
     "Between Reservation and Target. Acceptable."],
    ["Anchor at Lower × 75% (HK$7.80M)", "Least likely",
     "Section 13.4 applies. Do not concede."],
]
table(doc, ["Y0 outcome", "Likelihood", "Consequence"], rows,
      [5.0, 3.0, 9.5])

doc.add_heading("12.1 Readiness diagnostic — before seeking a priced seed", level=2)
doc.add_paragraph(
    "A priced seed at Y0 is only attemptable if all four are true. If any is false, "
    "run the angel instrument instead and revisit after the pilot."
)
num_bullet(doc, [("A named academy has agreed to pilot terms in writing. ", True),
                 ("An LOI with a timetable, not a verbal expression of interest.", False)])
num_bullet(doc, [("The pilot budget is fully specified in months of runway. ", True),
                 ("HK$1.5M buys how many months? If the answer is under 9, the raise "
                  "is too small for the milestones it must buy.", False)])
num_bullet(doc, [("At least one brand partner has engaged substantively. ", True),
                 ("A meeting is not engagement. A pilot scope discussion is.", False)])
num_bullet(doc, [("The Lower case has been rebuilt with a distinct cost structure. ", True),
                 ("If Lower is still Centre ÷ 2.57, an analyst will say so and the "
                  "priced round will not close on this schedule.", False)])

# ===========================================================================
# 13. Against the Lower x 75% Anchor
# ===========================================================================
doc.add_heading("13. Negotiating Against the Lower × 75% Anchor", level=1)
doc.add_paragraph(
    "This is the most likely attack and the most damaging if mishandled. The investor "
    "picks the intersection of the most pessimistic forecast and the most pessimistic "
    "discount rate. The result — HK$7.80M — implies a control transaction at seed. "
    "The defence is not a counter-number. It is a structural argument about why the "
    "cell is incoherent."
)

doc.add_heading("13.1 What HK$7.80M actually means", level=2)
rows = []
for k in ["Worst cell", "Reservation", "Target", "Opening ask"]:
    v = pos[k]
    o = own(v, RAISE_HKD_M)
    rows.append([k, money(v), money(RAISE_HKD_M), money(v + RAISE_HKD_M),
                 f"{o*100:.1f}%", f"{(1-o)*100:.1f}%"])
table(doc, ["Position", "Pre-money", "Raise", "Post-money",
            "Investor", "Founder"],
      rows, [3.0, 3.0, 2.8, 3.0, 2.4, 2.4], highlight_col=4, shade_rows=[0])
note(doc, "At HK$7.80M, an HK$8M raise takes 50.6% of the company. That is not a "
          "seed round. It is a control transaction. No serious seed investor asks for "
          "control; the ask is an anchor, and the correct response is to name it as "
          "one rather than counter-offer against it.")

doc.add_heading("13.2 Why the anchor is structurally incoherent", level=2)
bold_para(doc, [
    ("Two independent choices are bundled into one cell. ", True),
    ("The Lower case assumes 40% marketing cost and P10 adoption — that is a statement "
     "about the forecast. The 75% rate assumes a 75% required annual return — that is "
     "a statement about the return, not the forecast. Choosing both is choosing to be "
     "wrong twice.", False),
])
bold_para(doc, [
    ("The Lower case already prices the downside. ", True),
    ("Every year of the Lower case is profitable. A business that fails in no year "
     "does not warrant a 75% rate. 75% rates apply to binary outcomes — a drug "
     "candidate, a single-asset exploration play — not to a subscription business with "
     "a signed pilot and a named academy cohort. If the forecast is credible, the rate "
     "should be 50%. If the rate is 75%, the forecast should be Centre or Ceiling. The "
     "combination is double-counting.", False),
])
bold_para(doc, [
    ("The three cases are scalar multiples. ", True),
    ("From Y2 onward, Centre is ~2.57× Lower and Ceiling ~6.25× Lower in every year. "
     "The cases differ in size, not in structure. A conservative case that fails in no "
     "year and differs only by a scale factor is not a downside case — it is the base "
     "case divided by 2.57. The Lower × 75% cell therefore compounds two conservatisms "
     "that describe the same risk.", False),
])
bold_para(doc, [
    ("The implied duration confirms it. ", True),
    (f"The Y7 Q2 / Y0 ratio is {ratio[('Lower', 0.75)]:.2f}× at Lower × 75%, "
     f"{ratio[('Centre', 0.50)]:.2f}× at Centre × 50% and {ratio[('Ceiling', 0.25)]:.2f}× "
     f"at Ceiling × 25%. All three are essentially the same (1+r)^D with D ≈ 4.5 years. "
     f"The multiple is rate-determined. The scenario choice adds nothing except a lower "
     f"absolute number.", False),
])

doc.add_heading("13.3 The ownership red flag", level=2)
bold_para(doc, [
    (f"At HK$7.80M with an HK$8M raise, the investor takes "
     f"{own(pos['Worst cell'], RAISE_HKD_M)*100:.1f}%. ", True),
    (f"At that ownership, the founder has lost control in the first round. The next "
     f"round becomes difficult: no follow-on investor will enter above a HK$7.80M seed "
     f"price without extraordinary progress, and a Series A at 3× post-money dilutes "
     f"the founder to roughly {after_a['Worst cell']*100:.1f}%. This is not a "
     f"negotiating position; it is a structural failure of the cap table.", False),
])

doc.add_heading("13.4 The cost of conceding", level=2)
doc.add_paragraph(
    "Compare two paths. Path A accepts HK$7.80M at seed. Path B holds at the Target "
    "and raises HK$9.34M for 20%. Series A is modelled at 3× the post-seed valuation "
    "with a HK$15M raise, in both cases."
)
rows = []
for k in ["Worst cell", "Target"]:
    post_seed = pos[k] + RAISE_HKD_M
    a_pre = post_seed * SERIES_A_MULT
    a_post = a_pre + SERIES_A_RAISE
    f_seed = 1 - own(pos[k], RAISE_HKD_M)
    rows.append([
        k, money(pos[k]), money(post_seed),
        money(a_pre), money(a_post),
        f"{f_seed*100:.1f}%", f"{after_a[k]*100:.1f}%",
    ])
table(doc, ["Seed path", "Seed pre", "Post-seed",
            "Series A pre", "Series A post", "Founder after seed", "Founder after A"],
      rows, [2.6, 2.4, 2.4, 2.6, 2.6, 3.0, 3.0], highlight_col=6)
note(doc, "The gap in founder ownership after Series A is approximately "
          f"{(after_a['Target'] - after_a['Worst cell'])*100:.0f} percentage points. "
          "That is the real cost of conceding HK$7.80M — not the seed price, but the "
          "final ownership position it produces.")

doc.add_heading("13.5 The credible walk", level=2)
doc.add_paragraph(
    "The strongest response to HK$7.80M is not a counter-offer. It is a credible "
    "alternative. Before the meeting, the founder should be prepared to execute the "
    "following path if the anchor does not move."
)
bold_para(doc, [
    ("Alternative path. ", True),
    (f"Raise HK${ANGEL_PRE:.1f}M angel capital at HK${ANGEL_RAISE:.1f}M pre-money "
     f"(≈{ANGEL_RAISE/(ANGEL_PRE+ANGEL_RAISE)*100:.1f}% dilution). Complete the pilot "
     f"over 6–12 months. Raise the seed at HK${NEXT_SEED_PRE:.1f}M with an "
     f"HK${NEXT_SEED_RAISE:.1f}M round, by which point the pilot data should justify "
     f"the step-up. Founder ownership after both rounds: {alt_founder*100:.1f}%.", False),
])
bold_para(doc, [
    ("Compare. ", True),
    (f"Accepting HK$7.80M leaves the founder at "
     f"{(1-own(pos['Worst cell'], RAISE_HKD_M))*100:.1f}% immediately after seed. The "
     f"alternative path leaves the founder at {alt_founder*100:.1f}% after two rounds. "
     f"The difference is approximately {(alt_founder - (1-own(pos['Worst cell'], RAISE_HKD_M)))*100:.0f} "
     f"percentage points — the price of solving the negotiation by conceding.", False),
])

doc.add_heading("13.6 Pre-meeting preparation checklist", level=2)
doc.add_paragraph(
    "The walk is only as credible as the alternative behind it. Before the meeting, "
    "confirm each of the following. If any is unanswered, do not walk — bridge instead."
)
num_bullet(doc, [("One angel soft commitment. ", True),
                 ("HK$500K–1M at HK$25–30M pre-money. A verbal yes is sufficient to "
                  "make the alternative real.", False)])
num_bullet(doc, [("Pilot budget. ", True),
                 ("Exact cost to run the 5–10 academy pilot and the dataset build, in "
                  "months of runway. HK$1.5M buys approximately how many months?", False)])
num_bullet(doc, [("Named academy LOI. ", True),
                 ("At least one signed academy willing to start the pilot on a defined "
                  "timetable.", False)])
num_bullet(doc, [("Pilot success benchmarks. ", True),
                 ("The specific metrics — coach agreement rate, retained cohort size, "
                  "diagnostic accuracy — that convert pilot data into a valuation step-up.",
                  False)])

doc.add_heading("13.7 Fallback structures if the anchor will not move", level=2)
doc.add_paragraph(
    "If the investor will not move off the headline number but wants to do the deal, "
    "three structures bridge the gap without conceding 50%."
)
rows = [
    ["Milestone tranching",
     "HK$7.80M for tranche 1 (HK$2M); pre-agreed step-up to HK$25M for tranche 2 "
     "(HK$6M) upon pilot completion",
     "Founder keeps control; investor gets downside protection on entry"],
    ["Convertible with ratchet",
     "HK$7.80M effective cap, with anti-dilution reset to a higher valuation upon "
     "pilot data",
     "Investor's entry price adjusts upward if the pilot lands; no permanent 50% dilution"],
    ["Revenue share",
     "3% of subscription revenue until 2× invested capital returned, then converts "
     "to equity at the next round",
     "Keeps cap table clean; investor gets a defined return; no valuation argument"],
]
table(doc, ["Structure", "Mechanics", "Benefit"], rows,
      [3.5, 8.0, 6.0])
note(doc, "Milestone tranching is the strongest of the three because it prices the "
          "investor's concern (pilot uncertainty) directly rather than conceding "
          "against an unrelated number (the seed pre-money).")

doc.add_heading("13.8 Script — what to say, in order", level=2)

doc.add_paragraph("Step 1 — acknowledge without flinching.").runs[0].bold = True
quote_para(doc, "\"You've picked the most conservative cell in the model. I want to "
                "understand what you're actually assuming, because two different "
                "assumptions are bundled into that one number.\"")

doc.add_paragraph("Step 2 — separate the two assumptions.").runs[0].bold = True
quote_para(doc, "\"There are two independent choices: what you believe about the "
                "forecast, and what return you require. The Lower case assumes 40% "
                "marketing cost and P10 adoption. The 75% column assumes you require "
                "a 75% annual return. Which of those two are we actually debating?\"")

doc.add_paragraph("Step 3 — name the double-count.").runs[0].bold = True
quote_para(doc, "\"If you take the Lower case, you've already priced in the "
                "downside. Discounting that forecast at 75% prices in the downside a "
                "second time. I can defend either assumption on its own. I can't "
                "defend both together, because no other transaction in this category "
                "combines the worst forecast with the worst rate.\"")

doc.add_paragraph("Step 4 — surface the ownership.").runs[0].bold = True
quote_para(doc, "\"At HK$7.80M pre-money and an HK$8M raise, you'd own 50.6% of the "
                "company at seed. I don't think that's what you're proposing. But if "
                "it is, we should be clear that this is a control transaction, not a "
                "seed round, and it should be priced and documented as one.\"")

doc.add_paragraph("Step 5 — offer the bridge.").runs[0].bold = True
quote_para(doc, "\"What if we structure this differently? HK$7.80M for the first "
                "tranche, with a pre-agreed step-up to the Target upon pilot "
                "completion. You get your downside protection on the entry point. I "
                "get to keep enough of the company to raise the next round. We both "
                "get a defined outcome.\"")

doc.add_paragraph("Step 6 — set the walk, without threatening.").runs[0].bold = True
quote_para(doc, "\"If we can't bridge, I'd rather raise a smaller angel round at "
                "HK$25–30M, complete the pilot, and come back at HK$50M+ next year. "
                "That's not a threat — it's the math of what the pilot is worth. I "
                "want to do this with you. But I can't do it at HK$7.80M.\"")

doc.add_heading("13.9 What not to do", level=2)
bullet(doc, [("Do not counter-offer against HK$7.80M. ", True),
             ("A counter at, say, HK$20M treats the anchor as a real number. It is "
              "not. Argue the method, not the number.", False)])
bullet(doc, [("Do not say \"we'll be profitable in Year 1\". ", True),
             ("They have read the model. They know. Repeating it signals that you "
              "have nothing else.", False)])
bullet(doc, [("Do not reveal urgency. ", True),
             ("If the investor senses that you need the money in weeks, the walk is "
              "not credible and the anchor will not move. Prepare the alternative "
              "before revealing the timeline.", False)])
bullet(doc, [("Do not accept a lower number \"for now\" and hope to reprice. ", True),
             ("A HK$7.80M seed sets the reference price for every subsequent round. "
              "The cap table cannot be repaired by later performance alone.", False)])

# ===========================================================================
# 14. Caveats
# ===========================================================================
doc.add_heading("14. Caveats", level=1)
bullet(doc, [("The projections are unvalidated. ", True),
             ("No pilot has run. The discount rates price that uncertainty; the cash "
              "flows themselves are model outputs.", False)])
bullet(doc, [("Paper profitability is not a business. ", True),
             ("Even the Lower case is profitable in every year. An investor will test "
              "it against realistic CAC and a longer sales cycle. Prepare that "
              "sensitivity before the first meeting.", False)])
bullet(doc, [("The three cases are scalar multiples. ", True),
             ("Centre is approximately 2.57× Lower and Ceiling approximately 6.25× "
              "Lower in every year from Y2. The cases vary size, not shape. This is "
              "the most likely methodological challenge and it should be addressed "
              "before diligence, not in the room.", False)])
bullet(doc, [("Share count is a convention. ", True),
             (f"Per-share prices are on {SHARE_COUNT} shares and scale inversely with "
              f"the authorised count. Confirm with counsel before quoting a price per "
              f"share externally.", False)])
bullet(doc, [("Tax incentives are not earnings. ", True),
             ("DIPN 55 and HKAS 38 improve cash timing and the balance sheet. They do "
              "not belong in a profit-based valuation as income.", False)])

# ===========================================================================
# 15. Recommendation
# ===========================================================================
doc.add_heading("15. Recommendation", level=1)
bold_para(doc, [
    ("Price the round on the adjusted Y0 basis. ", True),
    (f"Target {money(pos['Target'])} pre-money ({ps_fmt(pos['Target'])}/share) at "
     f"Centre × 50%. Open at {money(pos['Opening ask'])} "
     f"({ps_fmt(pos['Opening ask'])}/share) at Ceiling × 50%. Do not close below "
     f"{money(pos['Reservation'])} ({ps_fmt(pos['Reservation'])}/share) at "
     f"Centre × 75%.", False),
])
bold_para(doc, [
    ("Do not accept the Lower × 75% cell. ", True),
    (f"At {money(pos['Worst cell'])} with an {money(RAISE_HKD_M)} raise, the investor "
     f"takes {own(pos['Worst cell'], RAISE_HKD_M)*100:.1f}% — a control transaction at "
     f"seed. The cost is not the seed price; it is the "
     f"{(after_a['Target'] - after_a['Worst cell'])*100:.0f} percentage-point gap in "
     f"founder ownership after Series A.", False),
])
bold_para(doc, [
    ("Prepare the walk before the meeting. ", True),
    ("The credible alternative — angel round at HK$25–30M, pilot completion, seed "
     "return at HK$50M+ — is the strongest card in the negotiation. It is only a card "
     "if it exists before the conversation starts.", False),
])
bold_para(doc, [
    ("Before the first meeting. ", True),
    ("Rebuild the Lower case with a different cost structure so the columns differ in "
     "kind and not only in scale. That change alone makes the Lower × 75% cell "
     "untenable, because the anchor depends on the Lower case being a plausible "
     "downside — which a scaled base case is not.", False),
])

out = "PreMoney_Valuation_v8.docx"
doc.save(out)

# ---- Console verification
print(f"Saved: {out}\n")
print("Adjusted Y0 (HK$M / per share on 100):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        v = adjusted_y0[(c, r)]
        line += f"  {int(r*100):>3}%: {v:>7.2f} ({ps_fmt(v):>12})"
    print(line)
print("\nY7 Q2 (HK$M / per share):")
for c in CASES:
    print(f"  {c:<8} {y7q2[c]:>8.2f} ({ps_fmt(y7q2[c])})")
print("\nValue creation ratio (Y7 Q2 / Y0):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {ratio[(c, r)]:>5.2f}x"
    print(line)
print(f"\nImplied duration: 25% {duration[0.25]:.2f}y  "
      f"50% {duration[0.50]:.2f}y  75% {duration[0.75]:.2f}y")
print("\nPositions (HK$8.00M raise):")
for k in ["Worst cell", "Reservation", "Target", "Opening ask"]:
    v = pos[k]
    print(f"  {k:<12} {money(v):>12}  {ps_fmt(v):>16}  "
          f"ownership {own(v, RAISE_HKD_M)*100:>5.1f}%  "
          f"founder after A {after_a[k]*100:>5.1f}%")
print(f"\nRaise for 20% at Target: {money(raise_for_20['Target'])}")
print(f"Alternative path founder ownership: {alt_founder*100:.1f}%")