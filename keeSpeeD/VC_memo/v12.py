"""
Kung Fu AI Football — The Crossover and the Cheque (.docx)
Requires: pip install python-docx

Two chapters, one document:
  Part I  — The strategic range (Y0 dream price to Y7 Q2 proven value)
  Part II — The operational price (Y4 Q1 DCF, the HK$1.0M cheque)
  Part III — The bridge (why the numbers differ, which to use when)

v8 is Part II. v11 is Part I. This document is both, plus the reconciliation.
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
INTANGIBLE_Y3 = 0.29
INTANGIBLE_Y6 = 0.58
GRANT_YEARS = 6.25
INTANGIBLE_Y6_YEARS = 6
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


def ps_fmt(val_m):
    return f"HK${val_m * 1_000_000 / SHARE_COUNT:,.0f}"


def stake_at(pre_m):
    return SEED_CHEQUE / (pre_m + SEED_CHEQUE)

# ===========================================================================
# COMPUTATIONS
# ===========================================================================

# --- Y0 adjusted (dream price)
disc_g = {r: 1 / (1 + r) ** GRANT_YEARS for r in RATES}
disc_i = {r: 1 / (1 + r) ** INTANGIBLE_Y6_YEARS for r in RATES}
y0 = {}
for c in CASES:
    for r in RATES:
        y0[(c, r)] = REPORTED_NPV[c][r] - GRANT + GRANT * disc_g[r] + INTANGIBLE_Y6 * disc_i[r]

# --- Y7 Q2 nominal (proven value)
y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE_Y6 for c in CASES}

# --- Midpoint (crossover)
mid = {(c, r): (y0[(c, r)] + y7q2[c]) / 2 for c in CASES for r in RATES}

# --- Y4 Q1 DCF (v8 base: reinvested)
# Remaining cash flows: Y4 (0.75y), Y5 (1.75y), Y6 (2.75y), grant (3.25y)
y4q1_base = {}
for c in CASES:
    for r in RATES:
        pv = (
            DATA[3][CASE_IDX[c]] / (1 + r) ** 0.75 +
            DATA[4][CASE_IDX[c]] / (1 + r) ** 1.75 +
            DATA[5][CASE_IDX[c]] / (1 + r) ** 2.75 +
            GRANT / (1 + r) ** 3.25
        )
        y4q1_base[(c, r)] = pv + INTANGIBLE_Y3

# --- Y4 Q1 retained (Y1-Y3 profits as cash)
cum_y1_y3 = {c: sum(DATA[t][CASE_IDX[c]] for t in range(3)) for c in CASES}
y4q1_ret = {(c, r): y4q1_base[(c, r)] + cum_y1_y3[c] for c in CASES for r in RATES}

# --- Investor ask
ask_pre = SEED_CHEQUE / INVESTOR_ASK - SEED_CHEQUE

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
    ("Re:", "The Crossover and the Cheque — Two Chapters, One Document"),
    ("Status:", "INTERNAL — negotiation preparation."),
    ("Version:", "v12 — integrated v8 (operational) + v11 (strategic) + bridge."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    p.add_run(value)

doc.add_paragraph()

# ---- 0. Purpose
doc.add_heading("0. Purpose — Two Chapters", level=1)
doc.add_paragraph(
    "This document supersedes both Seed_Valuation_Y4Q1 (v8) and Crossover_Valuation "
    "(v11). They answer different questions and should be read as two chapters of one "
    "argument:"
)
bullet(doc, [("Chapter 1 — the strategic range. ", True),
             ("Where the company sits between the price of a dream (Y0, no proof) and "
              "the value of a proven business (Y7 Q2, no need for VC). The crossover "
              "is the midpoint.", False)])
bullet(doc, [("Chapter 2 — the operational price. ", True),
             ("What the HK$1.0M cheque is worth at Y4 Q1, computed from the remaining "
              "cash flows and the intangible already on the balance sheet.", False)])
bullet(doc, [("Chapter 3 — the bridge. ", True),
             ("Why the two chapters produce different numbers, which one to use when, "
              "and how to explain the difference in the room.", False)])

# ===========================================================================
# PART I — STRATEGIC RANGE
# ===========================================================================

doc.add_heading("1. The Strategic Range — Dream to Proven", level=1)

doc.add_heading("1.1 The two extremes", level=2)
rows = [
    ["Proof", "None — a projection", "6 years audited", "3 years audited"],
    ["Sales record", "None", "Full history", "Traceable"],
    ["Who needs whom", "Founder needs VC", "VC approaches founder", "Mutual"],
    ["Rate applied", "50%+", "~12% or lower", "~12%"],
    ["Valuation method", "Narrative DCF", "Comparables / book", "Crossover"],
]
table(doc, ["Dimension", "Y0 — Dream", "Y7 Q2 — Proven", "Midpoint — Crossover"],
      rows, [3.5, 4.0, 4.2, 4.2], highlight_col=3)
note(doc, "The Y0 dream has a price, not no price — it is the Y0 NPV at a 50% "
          "discount. The Y7 Q2 value is book value if every year lands and nothing "
          "is distributed. Neither is zero and neither is infinite.")

doc.add_heading("1.2 The Y0 dream price", level=2)
rows = []
for c in CASES:
    rows.append([c] + [money(y0[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)
note(doc, "Adjusted basis: grant discounted from Y7 Q2; HK$0.58M intangible at Y6 "
          "discounted 6 years. This is what an investor pays at Y0 for a projection "
          "with no proof.")

doc.add_heading("1.3 The Y7 Q2 proven value", level=2)
rows = []
for c in CASES:
    rows.append([c, money(y7q2[c]), ps_fmt(y7q2[c])])
table(doc, ["Case", "Accumulated value", "Per share"], rows,
      [3.5, 4.5, 4.5], highlight_col=1)
note(doc, "Cumulative Y1–Y6 undiscounted + grant + intangible. No multiple, no "
          "discount. What the company has if every year lands.")

doc.add_heading("1.4 The midpoint — crossover", level=2)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        v = mid[(c, r)]
        row.append([money(v), f"{stake_at(v)*100:.2f}% at HK$1M"])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)
note(doc, "The midpoint is not an average of two arbitrary numbers. It implies a "
          "discount rate of 8–12% — the rate appropriate for a business with an "
          "audited track record. That is the rate the company will deserve at Y4 Q1.")

# ===========================================================================
# PART II — OPERATIONAL PRICE
# ===========================================================================

doc.add_heading("2. The Operational Price — The Cheque at Y4 Q1", level=1)

doc.add_heading("2.1 The Y4 Q1 DCF (v8 base — Y1–Y3 reinvested)", level=2)
doc.add_paragraph(
    "What remains at Y4 Q1 is Y4–Y6 operating profit, the Y7 Q2 grant, and the "
    "HK$0.29M intangible already on the balance sheet. The Y1–Y3 profits are assumed "
    "to have been reinvested into operations, so they are already embedded in the "
    "Y4–Y6 cash flows."
)
rows = []
for c in CASES:
    rows.append([c] + [money(y4q1_base[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("2.2 The Y4 Q1 DCF (retained variant — Y1–Y3 as cash)", level=2)
doc.add_paragraph(
    "If Y1–Y3 profits are instead retained as cash on the balance sheet, they add "
    "directly to the Y4 Q1 valuation."
)
rows = []
for c in CASES:
    rows.append([c, money(cum_y1_y3[c])] +
                [money(y4q1_ret[(c, r)]) for r in RATES])
table(doc, ["Case", "Y1–Y3 cumulative", "25%", "50%", "75%"], rows,
      [2.5, 3.0, 3.3, 3.3, 3.3], highlight_col=3)

doc.add_heading("2.3 The HK$1.0M cheque", level=2)
doc.add_paragraph(
    f"At each case and rate, the ownership transferred for a {money(SEED_CHEQUE)} "
    f"cheque on the retained basis. This is the position the founder should defend: "
    f"the company has the accumulated cash, so the cheque is for acceleration, not "
    f"survival."
)
rows = []
for c in CASES:
    rows.append([c] + [f"{stake_at(y4q1_ret[(c, r)])*100:.2f}%" for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ===========================================================================
# PART III — BRIDGE
# ===========================================================================

doc.add_heading("3. The Bridge — Why the Numbers Differ", level=1)

doc.add_heading("3.1 Reconciliation table (Centre × 50%)", level=2)
rows = [
    ["Y0 adjusted (dream price)", money(y0[("Centre", 0.50)]), "—",
     "Entire stream discounted to today"],
    ["Y4 Q1 base (reinvested)", money(y4q1_base[("Centre", 0.50)]),
     money(y4q1_base[("Centre", 0.50)] - mid[("Centre", 0.50)]),
     "Y1–Y3 profits consumed by operations"],
    ["Y4 Q1 retained", money(y4q1_ret[("Centre", 0.50)]),
     money(y4q1_ret[("Centre", 0.50)] - mid[("Centre", 0.50)]),
     "Y1–Y3 profits sit as cash"],
    ["Midpoint (crossover)", money(mid[("Centre", 0.50)]), "0.00",
     "Average of dream price and proven value"],
    ["Y7 Q2 nominal (proven)", money(y7q2["Centre"]),
     money(y7q2["Centre"] - mid[("Centre", 0.50)]),
     "Everything realised"],
]
table(doc, ["Basis", "Centre × 50%", "Gap to midpoint", "Assumption"],
      rows, [4.0, 3.2, 3.2, 6.0], highlight_col=1, shade_rows=[3])
note(doc, "The retained variant and the midpoint are within 3.3% of each other. The "
          "base case is the outlier — it assumes Y1–Y3 profits are gone, which "
          "understates the Y4 Q1 value by roughly 32%.")

doc.add_heading("3.2 The reconciliation across all cases", level=2)
rows = []
for c in CASES:
    rows.append([
        c,
        money(y4q1_base[(c, 0.50)]),
        money(y4q1_ret[(c, 0.50)]),
        money(mid[(c, 0.50)]),
        f"{(y4q1_ret[(c, 0.50)] - mid[(c, 0.50)]) / mid[(c, 0.50)] * 100:+.1f}%",
        f"{(y4q1_base[(c, 0.50)] - mid[(c, 0.50)]) / mid[(c, 0.50)] * 100:+.1f}%",
    ])
table(doc, ["Case", "Base (reinvested)", "Retained", "Midpoint",
            "Retained vs midpoint", "Base vs midpoint"],
      rows, [2.5, 3.0, 2.8, 2.8, 3.2, 3.2], highlight_col=2)
note(doc, "The retained variant tracks the midpoint within 3.5% in every case. The "
          "base case is consistently ~32% below. This is not noise — it is the "
          "HK$38M of Y1–Y3 retained profits in the Centre case.")

doc.add_heading("3.3 Which number to use when", level=2)
rows = [
    ["Opening the negotiation",
     "Midpoint Ceiling × 50%", money(mid[("Ceiling", 0.50)]),
     "Upside case, supported by the strategic range"],
    ["Target",
     "Midpoint Centre × 50%", money(mid[("Centre", 0.50)]),
     "Confirmed by retained Y4 Q1 DCF (within 3.3%)"],
    ["Settle",
     "Midpoint Lower × 50%", money(mid[("Lower", 0.50)]),
     "Conservative but coherent"],
    ["Walk-away",
     "Midpoint Lower × 75%", money(mid[("Lower", 0.75)]),
     "Floor; below this, decline"],
    ["Never use",
     "v8 base (reinvested)", money(y4q1_base[("Centre", 0.50)]),
     "Understates value by ~32%; invites the investor to believe the company has no "
     "accumulated cash"],
]
table(doc, ["Use", "Basis", "Figure", "Reason"],
      rows, [3.5, 4.0, 3.2, 5.8], highlight_col=2, shade_rows=[4])

# ===========================================================================
# 4. NEGOTIATION
# ===========================================================================

doc.add_heading("4. Negotiation", level=1)

doc.add_heading("4.1 Positions", level=2)
rows = [
    ["Opening ask", money(mid[("Ceiling", 0.50)]),
     f"{stake_at(mid[('Ceiling', 0.50)])*100:.2f}%",
     "Crossover Ceiling × 50%"],
    ["Target", money(mid[("Centre", 0.50)]),
     f"{stake_at(mid[('Centre', 0.50)])*100:.2f}%",
     "Crossover Centre × 50% (≈ retained DCF)"],
    ["Settle", money(mid[("Lower", 0.50)]),
     f"{stake_at(mid[('Lower', 0.50)])*100:.2f}%",
     "Crossover Lower × 50%"],
    ["Walk-away", money(mid[("Lower", 0.75)]),
     f"{stake_at(mid[('Lower', 0.75)])*100:.2f}%",
     "Crossover Lower × 75%"],
]
table(doc, ["Position", "Pre-money", "HK$1M buys", "Basis"],
      rows, [3.0, 3.5, 3.5, 8.0], highlight_col=1, shade_rows=[3])
note(doc, f"The investor's {INVESTOR_ASK*100:.2f}% ask implies {money(ask_pre)} — "
          f"below even the walk-away. That is the central finding: the ask does not "
          f"correspond to any forecast at the crossover rate.")

doc.add_heading("4.2 Script — opening with the strategic range", level=2)
quote_para(doc, "\"There are two prices for a company like this. At Y0, with no "
                "proof, the price is what someone pays for a dream — about HK$37M "
                "at a 50% rate. At Y7 Q2, with six years of audits and a profitable "
                "business, the company doesn't need VC money — it's a target. The "
                "honest price for a Y4 Q1 cheque is between them. That midpoint is "
                "HK$134M at the Centre case, and it's confirmed by a straightforward "
                "Y4 Q1 DCF once you count the accumulated cash.\"")

doc.add_heading("4.3 Script — when the investor anchors at 2.16%", level=2)
quote_para(doc, f"\"Your {INVESTOR_ASK*100:.2f}% implies HK$45.3M — which is below "
                f"every cell in the crossover matrix, including the most "
                f"conservative. It's also below the Y4 Q1 DCF even on the base case "
                f"that assumes we've spent every dollar of Y1–Y3 profit. So the "
                f"question isn't whether my number is too high. It's whether you "
                f"believe the company will have accumulated anything by Y4 Q1. I "
                f"think the audit will show it has.\"")

doc.add_heading("4.4 If the ask holds — the walk", level=2)
quote_para(doc, "\"If we can't bridge, I'd rather not raise. The company is "
                "profitable by Y3 and the accumulated cash means we don't need the "
                "HK$1.0M to survive. I'd rather retain earnings for twelve months "
                "and come back at the proven end of the range than accept a "
                "dream-frame price after the dream has already been executed.\"")

doc.add_heading("4.5 What to extract if you concede", level=2)
bullet(doc, [("Non-participating preference only. ", True),
             ("1× non-participating. Worth more than the stake concession.", False)])
bullet(doc, [("No board seat. ", True),
             ("Information rights and observer seat at most for a 2.16% holder.",
              False)])
bullet(doc, [("Follow-on price band agreed now. ", True),
             ("Lock in the next-round price band before the pilot data lands.",
              False)])

# ===========================================================================
# 5. CAVEATS
# ===========================================================================

doc.add_heading("5. Caveats", level=1)
bullet(doc, [("The midpoint is an approximation, not a method. ", True),
             ("What gives it force is the implied rate (~12%) and the independent "
              "confirmation from the retained Y4 Q1 DCF. State both when challenged.",
              False)])
bullet(doc, [("The Lower case is still Centre ÷ 2.57. ", True),
             ("Must be rebuilt with a distinct cost structure before external "
              "circulation. The crossover argument depends on Lower being a real "
              "forecast.", False)])
bullet(doc, [("Three years of audits must actually land. ", True),
             ("The crossover assumes Y1–Y3 will be audited and profitable. If they "
              "underperform, the rate rises and the midpoint falls.", False)])
bullet(doc, [("The retained-profits assumption is factual, not analytical. ", True),
             ("Whether Y1–Y3 profits sit as cash at Y4 Q1 depends on what the company "
              "actually does with them. Confirm before presenting the retained variant "
              "as the basis.", False)])
bullet(doc, [("Per-share values are on 100 shares. ", True),
             ("Confirm the authorised count with counsel before quoting externally.",
              False)])

# ===========================================================================
# 6. RECOMMENDATION
# ===========================================================================

doc.add_heading("6. Recommendation", level=1)
bold_para(doc, [
    ("Lead with the strategic range, price with the operational DCF. ", True),
    ("Open with the dream-to-proven frame — it establishes that the company sits at "
     "the crossover, not at either extreme. Then land on the retained Y4 Q1 DCF as "
     "the confirmation: the midpoint and the DCF give the same answer within 3.3%, "
     "which is why the price is defensible.", False),
])
bold_para(doc, [
    ("Use the midpoint for positions. ", True),
    (f"Open {money(mid[('Ceiling', 0.50)])}, target {money(mid[('Centre', 0.50)])}, "
     f"settle {money(mid[('Lower', 0.50)])}, walk below "
     f"{money(mid[('Lower', 0.75)])}. All four are supported by both the strategic "
     f"range and the retained DCF.", False),
])
bold_para(doc, [
    ("Never quote the v8 base case. ", True),
    (f"HK${y4q1_base[('Centre', 0.50)]:.2f}M assumes the company has no accumulated "
     f"cash at Y4 Q1. It understates value by ~32% and hands the investor the "
     f"argument that the first three years produced nothing. Use the retained "
     f"variant (HK${y4q1_ret[('Centre', 0.50)]:.2f}M) or the midpoint "
     f"(HK${mid[('Centre', 0.50)]:.2f}M) instead.", False),
])
bold_para(doc, [
    ("Before the meeting. ", True),
    ("Confirm the audit engagement type; rebuild the Lower case with a distinct cost "
     "structure; and decide whether the company can decline HK$1.0M if the terms do "
     "not land at or above the walk-away. The third is what makes the walk credible.",
     False),
])

out = "Crossover_and_Cheque_v12.docx"
doc.save(out)

# ---- Console
print(f"Saved: {out}\n")
print(f"{'Basis':<28} {'Ceiling':>10} {'Centre':>10} {'Lower':>10}")
print("-" * 62)
print(f"{'Y0 adjusted (dream)':<28} "
      f"{y0[('Ceiling', 0.50)]:>10.2f} {y0[('Centre', 0.50)]:>10.2f} "
      f"{y0[('Lower', 0.50)]:>10.2f}")
print(f"{'Y4 Q1 base (reinvested)':<28} "
      f"{y4q1_base[('Ceiling', 0.50)]:>10.2f} {y4q1_base[('Centre', 0.50)]:>10.2f} "
      f"{y4q1_base[('Lower', 0.50)]:>10.2f}")
print(f"{'Y4 Q1 retained':<28} "
      f"{y4q1_ret[('Ceiling', 0.50)]:>10.2f} {y4q1_ret[('Centre', 0.50)]:>10.2f} "
      f"{y4q1_ret[('Lower', 0.50)]:>10.2f}")
print(f"{'Midpoint (crossover)':<28} "
      f"{mid[('Ceiling', 0.50)]:>10.2f} {mid[('Centre', 0.50)]:>10.2f} "
      f"{mid[('Lower', 0.50)]:>10.2f}")
print(f"{'Y7 Q2 nominal (proven)':<28} "
      f"{y7q2['Ceiling']:>10.2f} {y7q2['Centre']:>10.2f} {y7q2['Lower']:>10.2f}")
print()
print(f"Investor ask {INVESTOR_ASK*100:.2f}% -> pre {money(ask_pre)}")
print(f"Ask below walk-away ({money(mid[('Lower', 0.75)])}): "
      f"{ask_pre < mid[('Lower', 0.75)]}")