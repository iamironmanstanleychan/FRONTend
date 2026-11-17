"""
Kung Fu AI Football — The Crossover Memo (.docx)
Requires: pip install python-docx

Frames the midpoint not as an arithmetic average but as the crossover point
at which the company transitions from "dream" (Y0, founder is supplicant)
to "proven" (Y7 Q2, founder is target). Prices the Y4 Q1 seed at that crossover.
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
    return SEED_CHEQUE / (pre_m + SEED_CHEQUE)


def implied_rate(case, target):
    lo, hi = 0.01, 3.0
    for _ in range(80):
        mid_r = (lo + hi) / 2
        pv = sum(DATA[t][CASE_IDX[case]] / (1 + mid_r) ** (t + 1) for t in range(6))
        pv += GRANT / (1 + mid_r) ** GRANT_YEARS
        pv += INTANGIBLE_Y6 / (1 + mid_r) ** INTANGIBLE_YEARS
        if pv > target:
            lo = mid_r
        else:
            hi = mid_r
    return (lo + hi) / 2

# ===========================================================================
# COMPUTATIONS
# ===========================================================================

disc_grant = {r: 1 / (1 + r) ** GRANT_YEARS for r in RATES}
disc_int = {r: 1 / (1 + r) ** INTANGIBLE_YEARS for r in RATES}

y0 = {}
for c in CASES:
    for r in RATES:
        y0[(c, r)] = REPORTED_NPV[c][r] - GRANT + GRANT * disc_grant[r] + INTANGIBLE_Y6 * disc_int[r]

y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE_Y6 for c in CASES}
mid = {(c, r): (y0[(c, r)] + y7q2[c]) / 2 for c in CASES for r in RATES}
implied = {c: implied_rate(c, mid[(c, 0.50)]) for c in CASES}

ask_pre = SEED_CHEQUE / INVESTOR_ASK - SEED_CHEQUE
lowest_cell = min(mid.values())
lowest_key = min(mid, key=mid.get)

y0_stake = stake_at(y0[("Centre", 0.50)])
mid_stake = stake_at(mid[("Centre", 0.50)])
dream_cost = y0_stake / mid_stake

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
    ("Re:", "The Crossover — From Dream to Target"),
    ("Status:", "INTERNAL — negotiation preparation."),
    ("Version:", "v11 — crossover framing; midpoint as power-dynamic transition."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    p.add_run(value)

doc.add_paragraph()

# ---- 1. Two Regimes
doc.add_heading("1. Two Regimes, Not Two Valuations", level=1)
doc.add_paragraph(
    "The Y0 and Y7 Q2 valuations are not two points on a single curve. They belong "
    "to different markets, different investment logic, and different power dynamics."
)
rows = [
    ["Audited accounts", "None", "6 years", "3 years"],
    ["Sales record", "None — a projection", "Full history", "Traceable"],
    ["Discount rate applied", "50%+", "~12% or lower", "~12%"],
    ["Who needs whom", "Founder needs VC", "VC approaches founder", "Mutual"],
    ["Valuation method", "Narrative DCF", "Comparables / book", "Crossover"],
    ["Founder posture", "Supplicant", "Target", "Equal"],
]
table(doc, ["Dimension", "Y0 — Dream", "Y7 Q2 — Proven", "Midpoint — Crossover"],
      rows, [3.5, 4.0, 4.2, 4.2], highlight_col=3)
note(doc, "The midpoint is not an average of the two endpoints. It is the price at "
          "which the posture changes from supplicant to equal — the first moment the "
          "founder can decline a cheque without consequence.")

bold_para(doc, [
    ("Why this matters for the negotiation. ", True),
    ("A dream-frame valuation is not a low valuation — it is a different kind of "
     "valuation. It prices the company as though none of the evidence exists. If the "
     "company will have three years of audits by Y4 Q1, that frame is factually "
     "wrong, not merely conservative.", False),
])

# ---- 2. The Crossover
doc.add_heading("2. The Crossover — What the Midpoint Actually Prices", level=1)
doc.add_paragraph(
    "The midpoint between the adjusted Y0 NPV and the Y7 Q2 accumulated value is the "
    "valuation that corresponds to a business with a verified operating history. The "
    "implied discount rates below are the proof."
)
rows = []
for c in CASES:
    rows.append([c, money(mid[(c, 0.50)]), f"{implied[c]*100:.1f}%",
                 money(y0[("Centre", 0.50)]) if c == "Centre" else "—"])
table(doc, ["Case", "Midpoint value (50% column)", "Implied discount rate",
            "For comparison: Y0 Centre × 50%"],
      rows, [3.0, 4.5, 4.5, 4.5], highlight_col=1)
note(doc, "The implied rates of 8–12% are mature-business rates. They are not seed "
          "rates. The midpoint therefore prices the company as though it already has "
          "the track record — which is exactly what it will have by Y4 Q1.")

bold_para(doc, [
    ("The cleanest way to say it. ", True),
    (f"The midpoint isn't a compromise between your number and mine. It's the price "
     f"your own method produces once the company has three years of audits. You are "
     f"pricing the dream. I am asking you to price the trajectory.", False),
])

# ---- 3. Midpoint Matrix
doc.add_heading("3. The Midpoint Matrix", level=1)
doc.add_paragraph(
    f"Pre-money at the crossover and the stake a {money(SEED_CHEQUE)} cheque buys at "
    f"each cell. The 50% column is the benchmark; the 25% and 75% columns bracket it."
)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        v = mid[(c, r)]
        row.append([money(v), f"{stake_at(v)*100:.2f}% at HK$1M"])
    rows.append(row)
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [2.5, 4.3, 4.3, 4.3], highlight_col=2)

doc.add_heading("Per share on 100 shares", level=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(mid[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ---- 4. Cost of the Dream Frame
doc.add_heading("4. The Cost of Accepting the Dream Frame", level=1)
doc.add_paragraph(
    "The difference between the two regimes is not philosophical. It is measured in "
    "equity given away for the same cheque."
)
rows = [
    ["Y0 — Dream (Centre × 50%)", money(y0[("Centre", 0.50)]),
     f"{y0_stake*100:.2f}%", "None — no track record"],
    ["Midpoint — Crossover (Centre × 50%)", money(mid[("Centre", 0.50)]),
     f"{mid_stake*100:.2f}%", "3 years audited, profitable"],
]
table(doc, ["Frame", "Pre-money", "HK$1M buys", "Evidence priced in"],
      rows, [5.0, 3.0, 3.0, 5.5], highlight_col=2)
note(doc, f"Accepting the dream frame costs {dream_cost:.2f}× the equity for the same "
          f"cheque. The investor's {INVESTOR_ASK*100:.2f}% ask sits almost exactly at "
          f"the Y0 price — it is a dream-frame valuation dressed as a Y4 Q1 offer.")

bold_para(doc, [
    ("The ask, stated plainly. ", True),
    (f"The {INVESTOR_ASK*100:.2f}% ask implies {money(ask_pre)} pre-money. That is "
     f"below every one of the nine midpoint cells — the lowest being "
     f"{lowest_key[0]} × {int(lowest_key[1]*100)}% at {money(lowest_cell)}, which "
     f"still buys the investor only {stake_at(lowest_cell)*100:.2f}%. The ask does "
     f"not correspond to any point on the crossover curve. It corresponds to the "
     f"dream.", False),
])

# ---- 5. The Leverage Shift
doc.add_heading("5. The Leverage Shift", level=1)
doc.add_paragraph(
    "The valuation is downstream of the leverage. What changes between the two "
    "regimes is not the model but who needs whom."
)
bold_para(doc, [
    ("At Y0. ", True),
    ("No audits, no sales, no proof. The investor sets the price and the founder "
     "accepts or walks. Declining a cheque is not credible, because there is no "
     "alternative. This is why the dream frame is coherent at Y0 — it reflects real "
     "weakness in the founder's position.", False),
])
bold_para(doc, [
    ("At the crossover. ", True),
    ("Three years of audits, traceable sales, profitable in every case. The founder "
     "can decline a HK$1.0M cheque without impairing operations. That is the first "
     "moment in the company's life when the walk is credible — and the only moment "
     "when the investor must price the evidence rather than the founder's need.", False),
])
bold_para(doc, [
    ("At Y7 Q2. ", True),
    ("Six years of audits, a proven R&D business, an intangible on the balance sheet. "
     "The company is a target. Investors approach. The founder chooses among them. "
     "This is not a round the founder runs; it is a round the market runs.", False),
])

# ---- 6. Milestones
doc.add_heading("6. What Moves You Across", level=1)
doc.add_paragraph(
    "Each milestone shifts the applicable discount rate down and the valuation up. "
    "They are the milestones that move the posture, not just the price."
)
table(doc, ["Milestone", "What it proves", "Effect on posture"],
      [["First pilot academy signed", "Third-party validation",
        "From supplicant to negotiator"],
       ["Coach agreement benchmark", "The AI kernel works",
        "Supports Centre over Lower forecast"],
       ["First subscription revenue", "Customers pay",
        "Introduces a revenue multiple"],
       ["10k labelled kick events", "Dataset at scale",
        "Supports the licensing thesis"],
       ["First brand LOI", "B2B channel works",
        "Justifies the 12–15× exit multiple"],
       ["3 years audited, profitable", "Operating history",
        "The crossover — founder can decline"],
       ["6 years audited, scaled", "Proven business",
        "The company is a target"]],
      [4.5, 5.5, 6.0])

# ---- 7. Negotiation
doc.add_heading("7. Negotiation", level=1)

doc.add_heading("7.1 The frame to open with", level=2)
quote_para(doc, "\"I want to value this on the crossover basis, not the dream. At Y0 "
                "the company has nothing traceable — the discount rate has to be 50% "
                "because there is no evidence to price. By Y4 Q1 it will have three "
                "years of audits and be profitable in every case. The right rate for "
                "that business is around 12%, not 50%. The midpoint is what your own "
                "method produces at that rate.\"")

doc.add_heading("7.2 When the investor anchors at 2.16%", level=2)
quote_para(doc, "\"Your stake implies HK$45.3M pre-money. That's below every cell in "
                "the midpoint matrix — including the most conservative cell, which is "
                "the Lower case at a 75% rate, at HK$49.2M. So the question isn't "
                "whether my number is too high. It's which forecast your 2.16% "
                "encodes. Because it doesn't encode any forecast that survives the "
                "crossover rate.\"")

doc.add_heading("7.3 The landing zone", level=2)
quote_para(doc, f"\"What I can do is accept the crossover at the Lower case at 50% — "
                f"HK${mid[('Lower', 0.50)]:.1f}M pre-money, which gives you "
                f"{stake_at(mid[('Lower', 0.50)])*100:.2f}% for your HK$1.0M. That "
                f"prices the business at the conservative forecast with the track "
                f"record priced in. It's below the Centre crossover, but it's above "
                f"the floor. It's a fair landing.\"")

doc.add_heading("7.4 The walk, if the ask holds", level=2)
quote_para(doc, "\"If we can't bridge, I'd rather not raise at this price. The "
                "company is profitable by Y3. Retaining earnings for another 12 "
                "months costs less than accepting a cap table that can't be repaired "
                "by later performance. That's not a threat — it's the math of what "
                "the audits are worth.\"")

doc.add_heading("7.5 What to extract if you concede", level=2)
bullet(doc, [("Non-participating preference only. ", True),
             ("1× non-participating. The gap between participating and "
              "non-participating is worth more than the 1.42% stake concession.", False)])
bullet(doc, [("No board seat. ", True),
             ("A 2.16% holder should have information rights and a board observer "
              "seat at most.", False)])
bullet(doc, [("Pro-rata, not super pro-rata. ", True),
             ("Standard pre-emptive rights only. Keeps the cap table clean for the "
              "next round.", False)])
bullet(doc, [("Follow-on price band agreed now. ", True),
             ("If the investor holds at 2.16%, lock in the price band for their "
              "follow-on at the next round. Converts a price concession into a "
              "relationship.", False)])

# ---- 8. Positions
doc.add_heading("8. Positions", level=1)
rows = [
    ["Opening ask", money(mid[("Ceiling", 0.50)]),
     f"{stake_at(mid[('Ceiling', 0.50)])*100:.2f}%",
     "Crossover Ceiling × 50% — upside case"],
    ["Target", money(mid[("Centre", 0.50)]),
     f"{stake_at(mid[('Centre', 0.50)])*100:.2f}%",
     "Crossover Centre × 50% — defensible base"],
    ["Settle", money(mid[("Lower", 0.50)]),
     f"{stake_at(mid[('Lower', 0.50)])*100:.2f}%",
     "Crossover Lower × 50% — conservative but coherent"],
    ["Walk-away", money(mid[("Lower", 0.75)]),
     f"{stake_at(mid[('Lower', 0.75)])*100:.2f}%",
     "Crossover Lower × 75% — the floor"],
]
table(doc, ["Position", "Pre-money", "HK$1M buys", "Basis"],
      rows, [3.0, 3.5, 3.5, 8.0], highlight_col=1, shade_rows=[3])
note(doc, f"Even the walk-away ({stake_at(mid[('Lower', 0.75)])*100:.2f}%) is below "
          f"the investor's ask ({INVESTOR_ASK*100:.2f}%). At every position above the "
          f"floor, the founder is offering the investor less equity than they asked "
          f"for — and every position is supported by the crossover method.")

# ---- 9. Caveats
doc.add_heading("9. Caveats", level=1)
bullet(doc, [("The midpoint is a framing device, not a valuation method. ", True),
             ("It averages a discounted present value with an undiscounted accumulated "
              "value. What gives it force is the implied rate: ~12% is the rate "
              "appropriate for a business with an audited track record. The midpoint "
              "is not a compromise — it is what that rate produces.", False)])
bullet(doc, [("The Lower case is still Centre ÷ 2.57. ", True),
             ("This flaw must be fixed before external circulation. The crossover "
              "argument depends on the Lower case being a real forecast, not the base "
              "case scaled down.", False)])
bullet(doc, [("Three years of audits must actually land. ", True),
             ("The crossover assumes Y1–Y3 will be audited and profitable. If they "
              "underperform, the rate rises and the midpoint falls. The matrix is "
              "conditional on execution.", False)])
bullet(doc, [("The investor's ask may reflect a different frame. ", True),
             ("2.16% could be a deliberate anchor, or a different method entirely. "
              "Ask which. The crossover argument lands only if the investor is "
              "willing to engage on method.", False)])
bullet(doc, [("Per-share values are on 100 shares. ", True),
             ("Confirm the authorised count with counsel before quoting externally.",
              False)])

# ---- 10. Recommendation
doc.add_heading("10. Recommendation", level=1)
bold_para(doc, [
    ("Open with the crossover, not the endpoints. ", True),
    (f"Open at {money(mid[('Ceiling', 0.50)])} (Ceiling × 50%), target "
     f"{money(mid[('Centre', 0.50)])} (Centre × 50%), settle at "
     f"{money(mid[('Lower', 0.50)])} (Lower × 50%), and do not close below "
     f"{money(mid[('Lower', 0.75)])} (Lower × 75%).", False),
])
bold_para(doc, [
    ("Lead with the implied rate. ", True),
    (f"\"The crossover implies a 12% rate\" is a stronger opening than \"the midpoint "
     f"is HK$134M.\" The first engages the investor on method; the second invites a "
     f"counter-number. Method arguments are winnable. Number arguments are not.", False),
])
bold_para(doc, [
    ("The 2.16% ask is a dream-frame valuation. ", True),
    ("That is the finding to carry into the meeting. The investor is pricing the "
     "company as though none of the evidence exists. The response is not a "
     "counter-number — it is the observation that the ask encodes no forecast that "
     "survives the crossover rate.", False),
])
bold_para(doc, [
    ("Before the meeting. ", True),
    ("Confirm the audit engagement type; rebuild the Lower case with a distinct cost "
     "structure; and decide whether the company can decline HK$1.0M if the terms do "
     "not land at or above the walk-away. The third is what makes the walk credible.",
     False),
])

out = "Crossover_Valuation_v11.docx"
doc.save(out)

# ---- Console
print(f"Saved: {out}\n")
print("Y0 (dream):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {y0[(c, r)]:>7.2f}"
    print(line)
print("\nY7 Q2 (proven):")
for c in CASES:
    print(f"  {c:<8} {y7q2[c]:>8.2f}")
print("\nMidpoint (crossover) — HK$M / stake at HK$1M:")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        v = mid[(c, r)]
        line += f"  {int(r*100):>3}%: {v:>7.2f} ({stake_at(v)*100:>4.2f}%)"
    print(line)
print("\nImplied discount rate at midpoint (50% column):")
for c in CASES:
    print(f"  {c:<8} {implied[c]*100:>5.1f}%")
print(f"\nInvestor {INVESTOR_ASK*100:.2f}% ask -> pre {money(ask_pre)}")
print(f"Lowest midpoint cell: {lowest_key[0]} x {int(lowest_key[1]*100)}% = "
      f"{money(lowest_cell)} ({stake_at(lowest_cell)*100:.2f}%)")
print(f"Ask below every midpoint cell: {ask_pre < lowest_cell}")
print(f"\nDream frame cost: {y0_stake*100:.2f}% vs {mid_stake*100:.2f}% "
      f"= {dream_cost:.2f}x equity")