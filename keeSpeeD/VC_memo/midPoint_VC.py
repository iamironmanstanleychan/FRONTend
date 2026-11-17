"""
Kung Fu AI Football — Midpoint Valuation Memo (.docx)
Requires: pip install python-docx

Anchor: midpoint between Y0 NPV (Y1-Y7 Q2, adjusted) and Y7 Q2 accumulated value.
Reflects the "trace sales" investment thesis: the company is priced as an
operating concern, not a pure projection or a completed plan.
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
INTANGIBLE_Y3 = 0.29
GRANT_YEARS = 6.25
INTANGIBLE_Y6_YEARS = 6
RATES = [0.25, 0.50, 0.75]
CASES = ["Ceiling", "Centre", "Lower"]
CASE_IDX = {"Lower": 1, "Centre": 2, "Ceiling": 3}

SHARE_COUNT = 100
SEED_CHEQUE = 1.0
INVESTOR_ASK_STAKE = 0.0216

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
    r = p.add_run(text); r.italic = True; r.font.size = Pt(10)
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


def stake_at(pre_m):
    """Stake for SEED_CHEQUE at the given pre-money (HK$M)."""
    return SEED_CHEQUE / (pre_m + SEED_CHEQUE)


def implied_rate(case, target):
    lo, hi = 0.01, 3.0
    for _ in range(80):
        mid = (lo + hi) / 2
        pv = 0.0
        for t, row in enumerate(DATA):
            pv += row[CASE_IDX[case]] / (1 + mid) ** (t + 1)
        pv += GRANT / (1 + mid) ** GRANT_YEARS
        pv += INTANGIBLE_Y6 / (1 + mid) ** INTANGIBLE_Y6_YEARS
        if pv > target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2

# ===========================================================================
# COMPUTATIONS
# ===========================================================================

# Y0 adjusted basis
disc_grant = {r: 1 / (1 + r) ** GRANT_YEARS for r in RATES}
disc_intangible = {r: 1 / (1 + r) ** INTANGIBLE_Y6_YEARS for r in RATES}
y0 = {}
for c in CASES:
    for r in RATES:
        y0[(c, r)] = (
            REPORTED_NPV[c][r] - GRANT
            + GRANT * disc_grant[r]
            + INTANGIBLE_Y6 * disc_intangible[r]
        )

# Y7 Q2 (undiscounted)
y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE_Y6 for c in CASES}

# Midpoint
mid = {(c, r): (y0[(c, r)] + y7q2[c]) / 2 for c in CASES for r in RATES}

# Implied rates
implied = {c: implied_rate(c, mid[(c, 0.50)]) for c in CASES}

# Investor ask
ask_pre = SEED_CHEQUE / INVESTOR_ASK_STAKE - SEED_CHEQUE

# Lowest midpoint cell
lowest_cell = min(mid.values())
lowest_cell_key = min(mid, key=mid.get)

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
    ("Re:", "Midpoint Valuation — The Trace-Sales Frame"),
    ("Status:", "INTERNAL — negotiation preparation."),
    ("Version:", "v10 — midpoint between Y0 NPV and Y7 Q2 accumulated value."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    p.add_run(value)

doc.add_paragraph()

# ---- 1. The Trace-Sales Thesis
doc.add_heading("1. The Trace-Sales Thesis", level=1)
doc.add_paragraph(
    "A seed or VC investor will fund a company with evidence of sales — even if "
    "revenue is minimal or negative in the first few years — because the sales record "
    "is proof that the business exists, that customers pay, and that the plan has "
    "survived contact with the market. That proof has value independent of the "
    "profitability it produces."
)
bold_para(doc, [
    ("Neither endpoint prices that proof correctly. ", True),
    ("The Y0 NPV discounts the business as if no proof exists — every future cash "
     "flow is treated as a projection. The Y7 Q2 value treats the proof as complete — "
     "every year has landed and all profit is banked. The truth lies between them: "
     "the business is worth more than its risk-adjusted PV, because the track record "
     "is itself an asset, and less than its accumulated value, because the plan has "
     "not yet fully landed.", False),
])
bold_para(doc, [
    ("The midpoint is the valuation at which a trace-sales investor should "
     "transact. ", True),
    ("It is not a DCF. It is a reconciliation: the present value of what is projected "
     "and the accumulated value of what will have been proven, averaged. The result "
     "can be read as the price a rational investor pays for the evidence itself.", False),
])

# ---- 2. The Two Endpoints
doc.add_heading("2. The Two Endpoints", level=1)
doc.add_heading("Y0 NPV — Y1 to Y7 Q2, discounted", level=2)
rows = []
for c in CASES:
    rows.append([c] + [money(y0[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)
note(doc, "Adjusted basis: grant discounted from Y7 Q2 (6.25 years); HK$0.58M "
          "intangible at Y6 discounted 6 years. The 50% column is the standard seed "
          "target return.")

doc.add_heading("Y7 Q2 — Undiscounted accumulated value", level=2)
rows = []
for c in CASES:
    rows.append([c, f"{REPORTED_UNDISCOUNTED[c]:.2f}",
                 f"{INTANGIBLE_Y6:.2f}", money(y7q2[c]), ps_fmt(y7q2[c])])
table(doc, ["Case", "Cumulative Y1–Y6 + grant", "Intangible",
            "Total", "Per share"],
      rows, [2.5, 4.3, 2.3, 3.2, 3.5], highlight_col=3)
note(doc, "No discount, no multiple, no probability. The value of the plan if every "
          "year lands exactly as modelled and all profit is retained.")

# ---- 3. Midpoint Matrix
doc.add_heading("3. The Midpoint Matrix — Primary Valuation", level=1)
doc.add_paragraph(
    "The midpoint between the two endpoints, for each case and rate. This is the "
    "valuation the trace-sales investor pays: more than the discounted present value, "
    "less than the accumulated nominal."
)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        row.append([money(mid[(c, r)]), f"{stake_at(mid[(c, r)])*100:.2f}% at HK$1M"])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)
note(doc, "Read: at the midpoint, a HK$1.0M cheque buys the stake shown. The 50% "
          "column is the benchmark; the 25% and 75% columns bracket it.")

doc.add_heading("Per share on 100 shares", level=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(mid[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ---- 4. What the midpoint implies about the rate
doc.add_heading("4. What the Midpoint Implies About the Discount Rate", level=1)
doc.add_paragraph(
    "The midpoint is equivalent to discounting the Y0 NPV at a much lower rate than "
    "50%. The implied rates below show what discount rate would produce the midpoint "
    "value directly. This is the analytical basis for using the midpoint: it is not "
    "an arbitrary average, it is the value that corresponds to a rate appropriate for "
    "a company with a track record."
)
rows = []
for c in CASES:
    rows.append([c, f"{implied[c]*100:.1f}%",
                 money(mid[(c, 0.50)]),
                 "Consistent with audited operating history, not a seed projection"])
table(doc, ["Case", "Implied rate at midpoint", "Midpoint value (50%)", "Interpretation"],
      rows, [2.5, 3.5, 3.5, 8.0], highlight_col=1)
note(doc, f"A rate of {implied['Centre']*100:.0f}% is what you apply to a mature, "
          f"audited business with three years of sales. A rate of 50% assumes the "
          f"business has no track record at all. The midpoint is the valuation at "
          f"which the rate reflects the track record the company will actually have "
          f"at Y4 Q1.")

bold_para(doc, [
    ("This is the trace-sales argument, stated technically. ", True),
    (f"At Y4 Q1 the company will have three years of audited profit. The appropriate "
     f"rate to apply to that record is approximately {implied['Centre']*100:.0f}%, "
     f"not 50%. The midpoint matrix is the valuation that results from applying that "
     f"rate, and it is the correct basis for the HK$1.0M seed.", False),
])

# ---- 5. HK$1.0M at Y4 Q1
doc.add_heading("5. The HK$1.0M Seed at Y4 Q1", level=1)
doc.add_paragraph(
    "At the midpoint Centre × 50%, the HK$1.0M cheque buys 0.74%. That is the "
    "defensible position when the company has three years of audited profit."
)
rows = []
for c in CASES:
    rows.append([
        c,
        money(mid[(c, 0.50)]),
        money(mid[(c, 0.50)] + SEED_CHEQUE),
        f"{stake_at(mid[(c, 0.50)])*100:.2f}%",
    ])
table(doc, ["Case", "Pre-money", "Post-money", "HK$1.0M buys"],
      rows, [3.0, 3.8, 3.8, 3.8], highlight_col=3)
note(doc, "The Centre case is the benchmark. The Lower case is the conservative "
          "reading; the Ceiling case is the upside.")

# ---- 6. Against the Investor's 2.16% Ask
doc.add_heading("6. Against the Investor's 2.16% Ask", level=1)
bold_para(doc, [
    ("The ask implies a pre-money of ", False),
    (money(ask_pre), True),
    (" — which is ", False),
    ("below every one of the nine midpoint cells", True),
    (f". The lowest midpoint cell is the {lowest_cell_key[0]} case at "
     f"{int(lowest_cell_key[1]*100)}%, at {money(lowest_cell)}, which still buys the "
     f"investor only {stake_at(lowest_cell)*100:.2f}%. A 2.16% ask is a valuation the "
     f"midpoint method does not support in any case.", False),
])
rows = [
    ["Investor ask", money(ask_pre), f"{INVESTOR_ASK_STAKE*100:.2f}%"],
    [f"Lowest midpoint ({lowest_cell_key[0]} × {int(lowest_cell_key[1]*100)}%)",
     money(lowest_cell), f"{stake_at(lowest_cell)*100:.2f}%"],
    ["Midpoint Centre × 75%", money(mid[("Centre", 0.75)]),
     f"{stake_at(mid[('Centre', 0.75)])*100:.2f}%"],
    ["Midpoint Centre × 50%", money(mid[("Centre", 0.50)]),
     f"{stake_at(mid[('Centre', 0.50)])*100:.2f}%"],
    ["Midpoint Centre × 25%", money(mid[("Centre", 0.25)]),
     f"{stake_at(mid[('Centre', 0.25)])*100:.2f}%"],
]
table(doc, ["Reference", "Pre-money", "HK$1.0M buys"],
      rows, [6.0, 4.5, 4.5], highlight_col=2, shade_rows=[0])
note(doc, "The investor's ask is below every midpoint cell. That is the central "
          "finding of this memo and the basis of the negotiation that follows.")

# ---- 7. Negotiation Script
doc.add_heading("7. Negotiation Script", level=1)

doc.add_paragraph("Step 1 — state the frame.").runs[0].bold = True
quote_para(doc, "\"I want to value this on a trace-sales basis. The present value of "
                "the projections is one number; the accumulated value if the plan "
                "lands is another. The honest valuation is between them, because the "
                "track record we'll have by Y4 Q1 is itself worth something. That's "
                "the midpoint method.\"")

doc.add_paragraph("Step 2 — show the midpoint matrix.").runs[0].bold = True
quote_para(doc, "\"At the midpoint, the Centre case is HK$134M at the 50% column, "
                "and the implied discount rate is about "
                f"{implied['Centre']*100:.0f}%. That's the rate you'd apply to a "
                "company with three years of audited profit, not to a seed-stage "
                "projection.\"")

doc.add_paragraph("Step 3 — convert the ask to a forecast.").runs[0].bold = True
quote_para(doc, "\"Your 2.16% ask implies HK$45.3M pre-money. That is below every "
                "cell in the midpoint matrix — including the most conservative cell, "
                "which is the Lower case at a 75% rate and still values the company "
                "at HK$49M. So the question is: which forecast do you believe? "
                "Because 2.16% doesn't correspond to any of them.\"")

doc.add_paragraph("Step 4 — offer the landing zone.").runs[0].bold = True
quote_para(doc, f"\"What I can do is accept the midpoint at the Lower case at 50% — "
                f"HK${mid[('Lower', 0.50)]:.1f}M pre-money, which is "
                f"{stake_at(mid[('Lower', 0.50)])*100:.2f}% for your HK$1.0M. That "
                f"prices the business at the conservative forecast with the track "
                f"record priced in. It's below the Centre midpoint, but it's above "
                f"the floor. It's a fair landing.\"")

doc.add_paragraph("Step 5 — if the ask holds, name the structural problem.").runs[0].bold = True
quote_para(doc, "\"If we can't agree on the basis, I'd rather not raise at this "
                "price. The company is profitable by Y3. Retaining earnings for "
                "another 12 months costs less than accepting a cap table that can't "
                "be repaired by later performance.\"")

doc.add_paragraph("Step 6 — what to extract if you concede.").runs[0].bold = True
quote_para(doc, "\"If you hold at 2.16%, I need clean terms: non-participating "
                "preference, information rights only — no board seat, standard "
                "pro-rata only, and the price band for your follow-on at the next "
                "round agreed now. The stake concession should buy a defined "
                "relationship, not just a lower number.\"")

# ---- 8. Positions
doc.add_heading("8. Positions", level=1)
rows = [
    ["Opening ask", money(mid[("Ceiling", 0.50)]),
     f"{stake_at(mid[('Ceiling', 0.50)])*100:.2f}%",
     "Midpoint Ceiling × 50% — the upside case"],
    ["Target", money(mid[("Centre", 0.50)]),
     f"{stake_at(mid[('Centre', 0.50)])*100:.2f}%",
     "Midpoint Centre × 50% — the defensible base case"],
    ["Settle", money(mid[("Lower", 0.50)]),
     f"{stake_at(mid[('Lower', 0.50)])*100:.2f}%",
     "Midpoint Lower × 50% — conservative but coherent"],
    ["Walk-away", money(mid[("Lower", 0.75)]),
     f"{stake_at(mid[('Lower', 0.75)])*100:.2f}%",
     "Midpoint Lower × 75% — the floor; below this, decline"],
]
table(doc, ["Position", "Pre-money", "HK$1.0M buys", "Basis"],
      rows, [3.0, 3.5, 3.5, 8.5], highlight_col=1, shade_rows=[3])
note(doc, f"All four positions are above the investor's 2.16% ask, which implies "
          f"{money(ask_pre)}. The founder's walk-away of "
          f"{stake_at(mid[('Lower', 0.75)])*100:.2f}% is still below the ask — "
          f"meaning even at the floor, the investor is being asked to accept a "
          f"lower stake than they have proposed. That is the central asymmetry in "
          f"the negotiation.")

# ---- 9. Caveats
doc.add_heading("9. Caveats", level=1)
bullet(doc, [("The midpoint is a heuristic, not a valuation method. ", True),
             ("It averages a discounted present value with an undiscounted "
              "accumulated value. Strictly, these are different constructs and the "
              "average has no closed-form derivation. What gives it force is the "
              "implied rate: ~"
              f"{implied['Centre']*100:.0f}% is the rate appropriate for a company "
              "with an audited track record, and that is the rate that produces the "
              "midpoint.", False)])
bullet(doc, [("The Lower case is still Centre ÷ 2.57. ", True),
             ("This flaw remains and must be fixed before any external circulation. "
              "The midpoint argument is only as strong as the Lower case behind it.",
              False)])
bullet(doc, [("Y1–Y3 must actually land. ", True),
             ("The midpoint assumes the operating history will exist by Y4 Q1. If "
              "the early years underperform, the appropriate rate rises and the "
              "midpoint falls. The matrix is conditional on execution.", False)])
bullet(doc, [("The intangible is not double-counted but is also not free. ", True),
             ("HK$0.58M at Y6 and HK$0.29M at Y3 are the amortised carrying values, "
              "not the gross R&D spend. Confirm with the auditor that the capitalised "
              "amount is consistent with the R&D schedule.", False)])
bullet(doc, [("The investor's 2.16% ask may reflect a different frame. ", True),
             ("The ask could be a deliberate anchor, or it could reflect a different "
              "valuation method (comparables, book value, or a different forecast). "
              "Ask which. The midpoint argument lands only if the investor is willing "
              "to engage on method.", False)])
bullet(doc, [("Per-share values are on 100 shares. ", True),
             ("Confirm the authorised count with counsel before quoting any per-share "
              "figure externally.", False)])

# ---- 10. Recommendation
doc.add_heading("10. Recommendation", level=1)
bold_para(doc, [
    ("Open on the midpoint, not the endpoints. ", True),
    (f"Open at {money(mid[('Ceiling', 0.50)])} (Ceiling × 50%), target "
     f"{money(mid[('Centre', 0.50)])} (Centre × 50%), settle at "
     f"{money(mid[('Lower', 0.50)])} (Lower × 50%), and do not close below "
     f"{money(mid[('Lower', 0.75)])} (Lower × 75%).", False),
])
bold_para(doc, [
    ("Lead with the implied rate, not the midpoint number. ", True),
    (f"\"The midpoint implies a {implied['Centre']*100:.0f}% rate\" is a stronger "
     f"opening than \"the midpoint is HK$134M.\" The first engages the investor on "
     f"method; the second invites a counter-number. Method arguments are winnable; "
     f"number arguments are not.", False),
])
bold_para(doc, [
    ("The 2.16% ask is below every midpoint cell. ", True),
    ("That is the single fact to carry into the meeting. The midpoint method rejects "
     "the ask in all nine cells, and the founder has the arithmetic to show it. "
     "The question to pose is not \"why is your stake so high?\" but \"which forecast "
     "does 2.16% encode?\" — because no forecast at the midpoint rate produces it.",
     False),
])
bold_para(doc, [
    ("Before the meeting. ", True),
    ("Confirm the audit engagement type; rebuild the Lower case with a distinct cost "
     "structure; and decide whether the company can decline the HK$1.0M if the terms "
     "do not land at or above the walk-away. The third is the one that determines "
     "whether the walk is credible.", False),
])

out = "Midpoint_Valuation_v10.docx"
doc.save(out)

# ---- Console
print(f"Saved: {out}\n")
print("Y0 NPV adjusted:")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {y0[(c, r)]:>7.2f}"
    print(line)
print("\nY7 Q2 (with intangible 0.58):")
for c in CASES:
    print(f"  {c:<8} {y7q2[c]:>8.2f}")
print("\nMidpoint matrix (HK$M / stake at HK$1M):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        v = mid[(c, r)]
        line += f"  {int(r*100):>3}%: {v:>7.2f} ({stake_at(v)*100:>4.2f}%)"
    print(line)
print("\nPer share on 100 shares:")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {ps_fmt(mid[(c, r)]):>12}"
    print(line)
print("\nImplied rate at midpoint (Centre column):")
for c in CASES:
    print(f"  {c:<8} {implied[c]*100:>5.1f}%")
print(f"\nInvestor 2.16% ask -> pre {money(ask_pre)}")
print(f"Lowest midpoint cell: {lowest_cell_key[0]} x {int(lowest_cell_key[1]*100)}% "
      f"= {money(lowest_cell)} ({stake_at(lowest_cell)*100:.2f}%)")
print(f"Ask is below every midpoint cell: {ask_pre < lowest_cell}")