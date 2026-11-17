"""
Kung Fu AI Football — Pre-money Valuation: Y0 and Y7 Q2 (.docx)
Requires: pip install python-docx

Presents per-share pre-money valuation at two dates:
  Y0      seed round pricing basis (discounted, adjusted for grant timing + intangible)
  Y7 Q2   undiscounted accumulated value (reference only, not a pre-money)

Source figures are the founder's reported NPVs, reproduced exactly, then reconciled
to a consistent Y0 basis. Per-share values on 100 authorised shares.
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
GRANT_YEARS = 6.25          # Y7 Q2 from Y0
INTANGIBLE = 0.58           # Y6 HKAS 38 capitalised
INTANGIBLE_YEARS = 6
RATES = [0.25, 0.50, 0.75]
CASES = ["Ceiling", "Centre", "Lower"]

SHARE_COUNT = 100
RAISE_HKD_M = 8.0

# Reported NPVs as supplied by the founder (grant included at face)
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


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

# Discount factors
disc_grant = {r: 1 / (1 + r) ** GRANT_YEARS for r in RATES}
disc_intangible = {r: 1 / (1 + r) ** INTANGIBLE_YEARS for r in RATES}
pv_grant = {r: GRANT * disc_grant[r] for r in RATES}
pv_intangible = {r: INTANGIBLE * disc_intangible[r] for r in RATES}

# Adjusted Y0 basis: reported NPV, remove face grant, add discounted items
adjusted_y0 = {}
for c in CASES:
    for r in RATES:
        adjusted_y0[(c, r)] = (
            REPORTED_NPV[c][r] - GRANT + pv_grant[r] + pv_intangible[r]
        )

# Y7 Q2: undiscounted cumulative (grant already in it) + intangible
y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE for c in CASES}

# Negotiation positions on adjusted Y0
POSITIONS = {
    "Opening ask":  ("Ceiling", 0.50),
    "Target":       ("Centre",  0.50),
    "Reservation":  ("Centre",  0.75),
}
pos = {k: adjusted_y0[v] for k, v in POSITIONS.items()}

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
    ("Re:", "Pre-money Valuation — Kung Fu AI Football (Y0 and Y7 Q2)"),
    ("Status:", "INTERNAL — negotiation preparation."),
    ("Version:", "v6 — dual-date valuation, reconciled basis."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    p.add_run(value)

doc.add_paragraph()

# ---- 1. Purpose
doc.add_heading("1. Purpose and Summary", level=1)
doc.add_paragraph(
    "This memo presents the per-share pre-money valuation on two dates, from the "
    "same P&L projections. The two dates answer different questions, and conflating "
    "them is the most common error in seed-stage valuation documents."
)
bullet(doc, [("Y0 — today, the seed round pricing basis. ", True),
             ("Discounted cash flow at 25%, 50% and 75%, adjusted for the timing of "
              "the Y7 Q2 grant and the Y6 capitalised intangible. This is the number "
              "that goes into a term sheet.", False)])
bullet(doc, [("Y7 Q2 — the end of the forecast horizon. ", True),
             ("Everything realised at face value: six years of profit, the grant "
              "received, the intangible on the balance sheet. This is a reference "
              "point, not a valuation. It applies no multiple, no discount and no "
              "probability.", False)])
bold_para(doc, [
    ("Headline. ", True),
    (f"At the Centre case and the 50% standard seed rate, the Y0 pre-money is "
     f"{money(adjusted_y0[('Centre', 0.50)])} — {ps_fmt(adjusted_y0[('Centre', 0.50)])} "
     f"per share on {SHARE_COUNT} shares. The Y7 Q2 accumulated value in the same "
     f"case is {money(y7q2['Centre'])} — {ps_fmt(y7q2['Centre'])} per share. The "
     f"seed round is priced on the first number.", False),
])

# ---- 2. Source P&L
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
table(doc,
      ["Year", "Lower (40% mkt / P10)", "Centre (35% mkt / P50)",
       "Ceiling (30% mkt / P90)"],
      rows, [4.0, 4.3, 4.3, 4.3], shade_rows=[7])
note(doc,
     "Y6 capitalised intangible of HK$0.58M sits on the balance sheet separately and "
     "is added at the Y7 Q2 reference point; it is not a profit line.")

# ---- 3. The two dates explained
doc.add_heading("3. The Two Valuation Dates", level=1)
bold_para(doc, [
    ("Y0. ", True),
    ("The seed round closes today. Shares are issued and priced at the close. "
     "Every future cash flow is discounted back to this date at a rate that prices "
     "the risk the investor is taking. The 50% column is the standard seed target "
     "return; 75% prices execution risk at the maximum; 25% is an upside case.", False),
])
bold_para(doc, [
    ("Y7 Q2. ", True),
    ("The end of the forecast. Every projection year has landed, the grant has been "
     "received and the intangible is on the balance sheet. No discounting applies, "
     "because nothing is being valued forward — it has all happened. The figure "
     "shows what the company accumulates if the plan is executed exactly as modelled.", False),
])
bold_para(doc, [
    ("The difference between the two is the value of time and risk. ", True),
    (f"Centre case: Y7 Q2 value of {money(y7q2['Centre'])} discounts to "
     f"{money(adjusted_y0[('Centre', 0.50)])} at Y0 on a 50% rate. That gap — "
     f"{money(y7q2['Centre'] - adjusted_y0[('Centre', 0.50)])} — is what the investor "
     f"is being compensated for.", False),
])

# ---- 4. As-calculated NPV
doc.add_heading("4. Y0 — As Calculated (founder's figures)", level=1)
doc.add_paragraph(
    "The reported NPVs below, with the grant included at face value. These are the "
    "figures on record and are reproduced exactly for reconciliation."
)
rows = []
for c in CASES:
    rows.append([c] + [money(REPORTED_NPV[c][r]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("Per Share on 100 Shares (as calculated)", level=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(REPORTED_NPV[c][r]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ---- 5. Adjusted Y0
doc.add_heading("5. Y0 — Adjusted to a Consistent Basis", level=1)
doc.add_paragraph(
    "The reported NPVs treat the HK$0.24M grant as if it were received today. It is "
    "received at Y7 Q2, so it must be discounted. They also omit the HK$0.58M "
    "capitalised intangible, which is an asset on the balance sheet and should be "
    "counted. The adjustment below makes both consistent with the operating cash "
    "flows, which are already discounted year by year."
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
table(doc,
      ["Rate", "DF grant (6.25y)", "PV grant", "DF intangible (6y)", "PV intangible"],
      rows, [2.0, 3.2, 2.8, 3.5, 3.0])

doc.add_heading("Reconciliation — reported to adjusted", level=2)
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
table(doc,
      ["Case", "Rate", "Reported", "Remove grant at face",
       "Add PV grant", "Add PV intangible", "Adjusted"],
      rows, [2.2, 1.5, 2.4, 3.2, 2.4, 2.8, 2.4], highlight_col=6)

doc.add_heading("Adjusted Y0 pre-money", level=2)
rows = []
for c in CASES:
    rows.append([c] + [money(adjusted_y0[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("Per share on 100 shares (adjusted)", level=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(adjusted_y0[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)
note(doc,
     "The adjustment is small: −0.03 at 25%, −0.17 at 50%, −0.21 at 75% relative to "
     "the reported figures. At 25% the discounted grant nearly offsets the "
     "discounted intangible; at higher rates the grant loses more value than the "
     "intangible adds. No conclusion changes.")

# ---- 6. Y7 Q2
doc.add_heading("6. Y7 Q2 — Undiscounted Accumulated Value", level=1)
doc.add_paragraph(
    "At Y7 Q2 everything in the forecast has been realised. No discounting is "
    "applied, because nothing is being valued forward. The figures are the "
    "cumulative undiscounted profit (which already includes the grant at face) plus "
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
table(doc,
      ["Case", "Cumulative Y1–Y6 + grant", "Intangible", "Total at Y7 Q2", "Per share"],
      rows, [2.5, 4.3, 2.3, 3.2, 3.5], highlight_col=3)
note(doc,
     "This is the book value of accumulated assets at Y7 Q2, not enterprise value. "
     "It assumes every forecast year lands exactly as modelled and all profit is "
     "retained as cash. No multiple is applied. Quoting it as a pre-money figure "
     "will invite the response that the company does not need external capital.")

# ---- 7. Per-share summary
doc.add_heading("7. Per-Share Summary — All Three Bases", level=1)
doc.add_paragraph(
    f"Price per share on {SHARE_COUNT} authorised shares. A 100-share structure is a "
    f"seed-stage convention that keeps the cap table administrable; the absolute "
    f"figure carries no meaning, only the proportion each share represents."
)
rows = [
    ["Y0 — as calculated (25%)", ps_fmt(REPORTED_NPV['Ceiling'][0.25]),
     ps_fmt(REPORTED_NPV['Centre'][0.25]), ps_fmt(REPORTED_NPV['Lower'][0.25])],
    ["Y0 — as calculated (50%)", ps_fmt(REPORTED_NPV['Ceiling'][0.50]),
     ps_fmt(REPORTED_NPV['Centre'][0.50]), ps_fmt(REPORTED_NPV['Lower'][0.50])],
    ["Y0 — as calculated (75%)", ps_fmt(REPORTED_NPV['Ceiling'][0.75]),
     ps_fmt(REPORTED_NPV['Centre'][0.75]), ps_fmt(REPORTED_NPV['Lower'][0.75])],
    ["Y0 — adjusted (25%)", ps_fmt(adjusted_y0[('Ceiling', 0.25)]),
     ps_fmt(adjusted_y0[('Centre', 0.25)]), ps_fmt(adjusted_y0[('Lower', 0.25)])],
    ["Y0 — adjusted (50%)", ps_fmt(adjusted_y0[('Ceiling', 0.50)]),
     ps_fmt(adjusted_y0[('Centre', 0.50)]), ps_fmt(adjusted_y0[('Lower', 0.50)])],
    ["Y0 — adjusted (75%)", ps_fmt(adjusted_y0[('Ceiling', 0.75)]),
     ps_fmt(adjusted_y0[('Centre', 0.75)]), ps_fmt(adjusted_y0[('Lower', 0.75)])],
    ["Y7 Q2 — undiscounted", ps_fmt(y7q2['Ceiling']),
     ps_fmt(y7q2['Centre']), ps_fmt(y7q2['Lower'])],
]
table(doc, ["Basis", "Ceiling", "Centre", "Lower"], rows,
      [5.0, 3.8, 3.8, 3.8], shade_rows=[6])

# ---- 8. Explanation
doc.add_heading("8. Explanation — Reading the Numbers", level=1)

add_bold_para = lambda segs: bold_para(doc, segs)

add_bold_para([
    ("The \"0\" column is not a valuation. ", True),
    ("It is the sum of nominal cash flows — what you would have if every year landed "
     "and the cash sat in a drawer. It has no time value and applies no risk. Its "
     "only use is as the Y7 Q2 base, before adding the intangible.", False),
])
add_bold_para([
    ("The 25/50/75% columns are present values. ", True),
    ("They answer: what is the entire future stream worth today, if I require this "
     "return? At 50%, a HK$1 received six years out is worth HK$0.088 today. The "
     "discount rate is a statement of required return, not a probability of failure.", False),
])
add_bold_para([
    ("Per-share values are just pre-money ÷ 100. ", True),
    ("If the authorised share count is later set at 1,000 or 10,000, every per-share "
     "figure divides by 10 or 100 and nothing else changes. The share price is a "
     "convention; the pre-money is the economic figure.", False),
])
add_bold_para([
    ("Why the Y7 Q2 figure is much larger. ", True),
    (f"It is undiscounted and it includes six years of profit. Centre case: "
     f"{money(y7q2['Centre'])} at Y7 Q2 versus {money(adjusted_y0[('Centre', 0.50)])} "
     f"at Y0 on a 50% rate. The difference is what an investor earns for taking the "
     f"risk of waiting.", False),
])

# ---- 9. Intangible
doc.add_heading("9. The Intangible Asset — Treatment and Defence", level=1)
bold_para(doc, [
    ("What it is. ", True),
    ("HK$0.58M of HKAS 38 capitalised R&D carried on the balance sheet at Y6. It is "
     "an accounting recognition of spending already incurred, not new cash.", False),
])
bold_para(doc, [
    ("How it is treated here. ", True),
    ("Added to the operating NPV at its present value (HK$0.152M / HK$0.051M / "
     "HK$0.020M at 25/50/75%) and added at face value to the Y7 Q2 reference "
     "(HK$0.58M).", False),
])
bold_para(doc, [
    ("Why that is defensible. ", True),
    ("A DCF values the cash flows the IP generates. Adding the capitalised cost of "
     "creating that IP risks double-counting. The adjustment is defensible here "
     "because the DCF covers Y1–Y6 operating profit only and assigns no terminal "
     "value — no part of the forecast captures the balance-sheet IP separately.", False),
])
bold_para(doc, [
    ("The defence if challenged. ", True),
    ("It is a balance-sheet floor argument, not an earnings argument. HK$0.58M is "
     "capitalised R&D already spent and carried as an asset. At Y0 it floors "
     "enterprise value at accumulated IP, independent of whether the Y1–Y6 forecast "
     "is believed. It supports the reservation price, not the opening ask.", False),
])

# ---- 10. Negotiation
doc.add_heading("10. Negotiation Strategy", level=1)

doc.add_heading("10.1 Positions", level=2)
rows = []
for k in ["Opening ask", "Target", "Reservation"]:
    v = pos[k]
    basis = {
        "Opening ask": "Ceiling × 50% — upside case at standard seed return",
        "Target": "Centre × 50% — base case at standard seed return",
        "Reservation": "Centre × 75% — base case priced for execution risk",
    }[k]
    rows.append([k, money(v), ps_fmt(v), basis])
table(doc, ["Position", "Pre-money at Y0", "Per share", "Basis"], rows,
      [3.0, 3.2, 3.6, 6.7], highlight_col=1)

doc.add_heading("10.2 Principles", level=2)
bullet(doc, [("Anchor on the discount rate, not the profit forecast. ", True),
             ("50% is the standard seed target return. If the investor argues for a "
              "higher rate, they are arguing about risk, not about the model — a "
              "narrower and more winnable debate than one about profit forecasts.", False)])
bullet(doc, [("Disclose the adjustment. ", True),
             ("Show both the reported NPVs and the adjusted figures. Volunteering the "
              "small downward reconciliation builds credibility for everything else.", False)])
bullet(doc, [("Use Y7 Q2 as narrative, never as an ask. ", True),
             (f"Centre-case {money(y7q2['Centre'])} accumulated by Y7 Q2 justifies the "
              f"Ceiling column and the exit multiple. Introduced as a pre-money "
              f"number, it invites the reply that the company does not need capital.", False)])
bullet(doc, [("Trade terms, not price. ", True),
             ("If pushed toward the 75% column, concede valuation only for structure: "
              "milestone tranching, a smaller option pool, or a lower liquidation "
              "preference.", False)])
bullet(doc, [("Let the pilot set the price. ", True),
             ("Each month of pilot data moves the case from Centre toward Ceiling. If "
              "runway allows the wait, the 75% column becomes indefensible for the "
              "investor.", False)])

doc.add_heading("10.3 Round size and ownership", level=2)
doc.add_paragraph(
    f"At an {money(RAISE_HKD_M)} raise, implied investor ownership is:"
)
rows = []
for k in ["Opening ask", "Target", "Reservation"]:
    v = pos[k]
    own = RAISE_HKD_M / (v + RAISE_HKD_M)
    rows.append([k, money(v), money(RAISE_HKD_M), money(v + RAISE_HKD_M),
                 f"{own*100:.1f}%"])
table(doc, ["Position", "Pre-money", "Raise", "Post-money", "Investor ownership"],
      rows, [3.6, 3.4, 3.0, 3.6, 3.9], highlight_col=4)
raise_for_20 = 0.25 * pos["Target"]
bold_para(doc, [
    (f"At the Target, an {money(RAISE_HKD_M)} raise buys "
     f"{RAISE_HKD_M / (pos['Target'] + RAISE_HKD_M) * 100:.1f}% — within the normal "
     f"seed range. For exactly 20%, raise {money(raise_for_20)}.", False),
])

doc.add_heading("10.4 Instruments", level=2)
table(doc, ["Instrument", "When to use", "Watch-out"],
      [["SAFE with cap + discount", "Speed matters; valuation unresolved",
        "Caps compound across multiple SAFEs — model the full stack"],
       ["Convertible note", "Bridge to pilot milestones",
        "Interest and maturity can force a down round"],
       ["Priced equity round", "Pilot data exists; valuation defensible",
        "Sets a public benchmark that is hard to reprice downward"],
       ["Milestone tranching", "Investor discounts the ceiling case",
        "Define milestones in the SHA with objective, auditable tests"]],
      [4.2, 5.3, 7.0])

# ---- 11. Caveats
doc.add_heading("11. Caveats", level=1)
bullet(doc, [("The projections are unvalidated. ", True),
             ("No pilot has run. The discount rates price that uncertainty; the cash "
              "flows themselves are model outputs.", False)])
bullet(doc, [("Paper profitability is not a business. ", True),
             ("Even the Lower case is profitable in every year. An investor will test "
              "it against realistic CAC and a longer sales cycle. Prepare that "
              "sensitivity before the first meeting.", False)])
bullet(doc, [("The three cases are scalar multiples. ", True),
             ("From Y2 onward, Centre is approximately 2.57× Lower and Ceiling "
              "approximately 6.25× Lower in every year. The cases vary size, not "
              "shape. A conservative case that fails in no year is not a downside "
              "case. Rebuild Lower with a different structure before diligence.", False)])
bullet(doc, [("Share count is a convention. ", True),
             (f"Per-share prices are on {SHARE_COUNT} shares and scale inversely with "
              f"the authorised count. Confirm with counsel before quoting a price per "
              f"share externally.", False)])
bullet(doc, [("Tax incentives are not earnings. ", True),
             ("DIPN 55 and HKAS 38 improve cash timing and the balance sheet. They do "
              "not belong in a profit-based valuation as income.", False)])

# ---- 12. Recommendation
doc.add_heading("12. Recommendation", level=1)
bold_para(doc, [
    ("Price the round on the adjusted Y0 basis. ", True),
    (f"Target {money(pos['Target'])} pre-money ({ps_fmt(pos['Target'])}/share) at "
     f"Centre × 50%. Open at {money(pos['Opening ask'])} "
     f"({ps_fmt(pos['Opening ask'])}/share) at Ceiling × 50%. Do not close below "
     f"{money(pos['Reservation'])} ({ps_fmt(pos['Reservation'])}/share) at "
     f"Centre × 75%.", False),
])
bold_para(doc, [
    ("Quote the Y7 Q2 figure only as narrative. ", True),
    (f"Centre-case {money(y7q2['Centre'])} accumulated by Y7 Q2 justifies the Ceiling "
     f"column and the exit multiple. It is not a pre-money number and should never "
     f"appear in a term sheet.", False),
])
bold_para(doc, [
    ("Before the first meeting. ", True),
    ("Rebuild the Lower case with a different cost structure so the columns differ "
     "in kind and not only in scale. The current Lower case is Centre divided by "
     "2.57 — that will be identified and challenged.", False),
])

out = "PreMoney_Valuation_Y0_Y7Q2.docx"
doc.save(out)

# ---- Console verification
print(f"Saved: {out}\n")
print("Discount factors:")
for r in RATES:
    print(f"  {int(r*100):>3}%  grant(6.25y) {disc_grant[r]:.4f}  "
          f"intangible(6y) {disc_intangible[r]:.4f}")
print("\nAdjusted Y0 (HK$M / per share):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        v = adjusted_y0[(c, r)]
        line += f"  {int(r*100):>3}%: {v:>7.2f} ({ps_fmt(v):>12})"
    print(line)
print("\nY7 Q2 (undiscounted, HK$M / per share):")
for c in CASES:
    print(f"  {c:<8} {y7q2[c]:>8.2f} ({ps_fmt(y7q2[c])})")
print("\nPositions at Y0 (HK$8.00M raise):")
for k in ["Opening ask", "Target", "Reservation"]:
    v = pos[k]
    print(f"  {k:<12} {money(v):>12}  {ps_fmt(v):>16}  "
          f"ownership {RAISE_HKD_M/(v+RAISE_HKD_M)*100:>5.1f}%")
print(f"\nRaise for 20% at Target: {money(raise_for_20)}")