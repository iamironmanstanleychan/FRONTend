"""
Kung Fu AI Football — Y4 Q1 Seed Valuation (.docx)
Requires: pip install python-docx

Funding plan: HK$1.0M seed cheque at Y4 Q1; HK$0.29M capitalised intangible
on the balance sheet from Y3 year-end. Prices the round at Y4 Q1, not Y0,
and shows the step-up the pilot must earn.
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
GRANT_Y7Q2_FROM_Y0 = 6.25
INTANGIBLE_Y3 = 0.29          # balance sheet at Y3 year-end
INTANGIBLE_Y6 = 0.58          # balance sheet at Y6 (for Y0 reference)
INTANGIBLE_Y6_YEARS = 6
RATES = [0.25, 0.50, 0.75]
CASES = ["Ceiling", "Centre", "Lower"]
CASE_IDX = {"Lower": 1, "Centre": 2, "Ceiling": 3}

SEED_Y4Q1 = 1.0               # HK$1M seed cheque at Y4 Q1
Y4Q1_YEARS = 3.25             # from Y0
SHARE_COUNT = 100

REPORTED_NPV = {
    "Ceiling": {0.25: 199.02, 0.50: 90.85, 0.75: 49.15},
    "Centre":  {0.25:  81.96, 0.50: 37.51, 0.75: 20.37},
    "Lower":   {0.25:  32.04, 0.50: 14.70, 0.75:  8.01},
}

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


def own(pre, raise_m):
    return raise_m / (pre + raise_m)


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

cash_flows = {c: [r[CASE_IDX[c]] for r in DATA] for c in CASES}

# ---- Y0 reference (v7 basis: adjusted Y0)
disc_grant_y0 = {r: 1 / (1 + r) ** GRANT_Y7Q2_FROM_Y0 for r in RATES}
disc_intangible_y6 = {r: 1 / (1 + r) ** INTANGIBLE_Y6_YEARS for r in RATES}
pv_grant_y0 = {r: GRANT * disc_grant_y0[r] for r in RATES}
pv_intangible_y6 = {r: INTANGIBLE_Y6 * disc_intangible_y6[r] for r in RATES}
y0 = {}
for c in CASES:
    for r in RATES:
        y0[(c, r)] = REPORTED_NPV[c][r] - GRANT + pv_grant_y0[r] + pv_intangible_y6[r]

# ---- Y4 Q1 valuation
# Cash flows remaining: Y4 (0.75y), Y5 (1.75y), Y6 (2.75y), grant (3.25y)
# Intangible 0.29M is already on the balance sheet at Y4 Q1
disc_y4 = {r: 1 / (1 + r) ** 0.75 for r in RATES}
disc_y5 = {r: 1 / (1 + r) ** 1.75 for r in RATES}
disc_y6 = {r: 1 / (1 + r) ** 2.75 for r in RATES}
disc_grant_y4 = {r: 1 / (1 + r) ** Y4Q1_YEARS for r in RATES}

y4q1_base = {}       # intangible only, no retained profits
y4q1_retained = {}   # plus cumulative Y1-Y3 retained as cash
cumulative_y1_y3 = {c: sum(cash_flows[c][:3]) for c in CASES}

for c in CASES:
    for r in RATES:
        pv_fwd = (
            cash_flows[c][3] * disc_y4[r] +
            cash_flows[c][4] * disc_y5[r] +
            cash_flows[c][5] * disc_y6[r] +
            GRANT * disc_grant_y4[r]
        )
        y4q1_base[(c, r)] = pv_fwd + INTANGIBLE_Y3
        y4q1_retained[(c, r)] = pv_fwd + INTANGIBLE_Y3 + cumulative_y1_y3[c]

# ---- Step-up Y4 Q1 / Y0
step_up = {(c, r): y4q1_base[(c, r)] / y0[(c, r)] for c in CASES for r in RATES}
annual_return = {k: v ** (1 / Y4Q1_YEARS) - 1 for k, v in step_up.items()}

# ---- Ownership at HK$1M seed
ownership = {(c, r): own(y4q1_base[(c, r)], SEED_Y4Q1) for c in CASES for r in RATES}
ownership_ret = {(c, r): own(y4q1_retained[(c, r)], SEED_Y4Q1) for c in CASES for r in RATES}

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
    ("Re:", "Y4 Q1 Seed Valuation — Kung Fu AI Football"),
    ("Status:", "INTERNAL — negotiation preparation."),
    ("Version:", "v8 — HK$1M seed at Y4 Q1; intangible HK$0.29M at Y3 year-end."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    p.add_run(value)

doc.add_paragraph()

# ---- 1. Purpose and Summary
doc.add_heading("1. Purpose and Summary", level=1)
doc.add_paragraph(
    "The funding plan has moved: the seed cheque of HK$1.0M is received at Y4 Q1, "
    "not at Y0, by which point HK$0.29M of capitalised intangible sits on the balance "
    "sheet from Y3 year-end. The pricing date for the round is therefore Y4 Q1, not "
    "Y0, and every figure below is computed at that date."
)
bold_para(doc, [
    ("Primary pricing basis. ", True),
    (f"At the Centre case and the 50% standard seed rate, Y4 Q1 pre-money is "
     f"{money(y4q1_base[('Centre', 0.50)])} — {ps_fmt(y4q1_base[('Centre', 0.50)])} "
     f"per share on {SHARE_COUNT} shares. The HK$1.0M cheque buys "
     f"{ownership[('Centre', 0.50)]*100:.2f}% of the company.", False),
])
bold_para(doc, [
    ("Step-up from Y0. ", True),
    (f"The Y4 Q1 price is {step_up[('Centre', 0.50)]:.2f}× the Y0 price of "
     f"{money(y0[('Centre', 0.50)])}, equivalent to "
     f"{annual_return[('Centre', 0.50)]*100:.1f}% annualised over "
     f"{Y4Q1_YEARS} years. That is the return the Y0 backer earns if the plan lands, "
     f"and the benchmark the Y4 Q1 investor will test.", False),
])
bold_para(doc, [
    ("Why this is the right funding date. ", True),
    ("At Y4 Q1 the company has three years of operating history, is profitable in all "
     "three cases, carries HK$0.29M of capitalised IP, and should have pilot data. "
     "The HK$1.0M cheque is a growth instrument, not a survival instrument — and "
     "pricing it against unvalidated Y0 forecasts would be neither credible nor "
     "necessary.", False),
])

# ---- 2. P&L
doc.add_heading("2. P&L Projections", level=1)
rows = [[r[0], f"{r[1]:.2f}", f"{r[2]:.2f}", f"{r[3]:.2f}"] for r in DATA]
rows.append(["Y7 Q2 grant", f"{GRANT:.2f}", f"{GRANT:.2f}", f"{GRANT:.2f}"])
table(doc, ["Year", "Lower (40% mkt / P10)", "Centre (35% mkt / P50)",
            "Ceiling (30% mkt / P90)"],
      rows, [4.0, 4.3, 4.3, 4.3], shade_rows=[6])
note(doc, f"Intangible on balance sheet: HK${INTANGIBLE_Y3:.2f}M at Y3 year-end; "
          f"HK${INTANGIBLE_Y6:.2f}M at Y6. Accrues approximately "
          f"HK${(INTANGIBLE_Y6 - INTANGIBLE_Y3)/3:.3f}M per year.")

# ---- 3. Y0 reference
doc.add_heading("3. Y0 Valuation — Reference", level=1)
doc.add_paragraph(
    "The adjusted Y0 basis from the previous memo, reproduced for comparison. It is "
    "the price a Y0 investor would have paid; it is not the pricing basis for this "
    "round."
)
rows = []
for c in CASES:
    rows.append([c] + [money(y0[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(y0[(c, r)]) for r in RATES])
table(doc, ["Per share", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ---- 4. Y4 Q1 valuation
doc.add_heading("4. Y4 Q1 Valuation — Primary Pricing Basis", level=1)
doc.add_paragraph(
    "At Y4 Q1 the Y1–Y3 profits are in the past. What remains is Y4–Y6 operating "
    "profit, the Y7 Q2 grant, and the HK$0.29M intangible already on the balance "
    "sheet. Discounting is from Y4 Q1, not Y0."
)
doc.add_heading("Discount factors from Y4 Q1", level=2)
rows = []
for r in RATES:
    rows.append([
        f"{int(r*100)}%",
        f"{disc_y4[r]:.4f}", f"{disc_y5[r]:.4f}",
        f"{disc_y6[r]:.4f}", f"{disc_grant_y4[r]:.4f}",
    ])
table(doc, ["Rate", "Y4 (0.75y)", "Y5 (1.75y)", "Y6 (2.75y)", "Grant (3.25y)"],
      rows, [2.0, 3.0, 3.0, 3.0, 3.3])

doc.add_heading("Pre-money at Y4 Q1", level=2)
rows = []
for c in CASES:
    rows.append([c] + [money(y4q1_base[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("Per share at Y4 Q1 on 100 shares", level=2)
rows = []
for c in CASES:
    rows.append([c] + [ps_fmt(y4q1_base[(c, r)]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

# ---- 4.5 Audit impact on valuation
doc.add_heading("4.5 What the Y1–Y3 Audit Does to the Valuation", level=1)
doc.add_paragraph(
    "The Y4 Q1 valuation depends almost entirely on the discount rate, because the "
    "remaining cash flows are fixed by the forecast. An audit of Y1–Y3 does not "
    "change those cash flows — it changes the rate a reasonable investor can apply "
    "to them, by converting three years of founder projections into verified history."
)

doc.add_heading("Discount rate sensitivity at Y4 Q1 (Centre case)", level=2)
rate_rows = []
for rr in [0.25, 0.30, 0.35, 0.40, 0.50]:
    pv_fwd_r = (
        cash_flows["Centre"][3] / (1 + rr) ** 0.75 +
        cash_flows["Centre"][4] / (1 + rr) ** 1.75 +
        cash_flows["Centre"][5] / (1 + rr) ** 2.75 +
        GRANT / (1 + rr) ** Y4Q1_YEARS
    )
    pre_r = pv_fwd_r + INTANGIBLE_Y3
    rate_rows.append([
        f"{int(rr*100)}%",
        money(pre_r),
        ps_fmt(pre_r),
        f"{SEED_Y4Q1 / (pre_r + SEED_Y4Q1) * 100:.2f}%",
    ])
table(doc, ["Discount rate", "Centre pre-money", "Per share",
            "HK$1M buys"],
      rate_rows, [3.0, 4.0, 4.0, 4.0], highlight_col=2)
note(doc, "50% is the seed target return for a pre-revenue company. 30–35% is "
          "consistent with a profitable private company with three years of audited "
          "operation. The gap between those rows — HK$18M to HK$34M — is the "
          "quantified value of the audit.")

doc.add_heading("What the audit validates — and what it does not", level=2)
bullet(doc, [("Validates: ", True),
             ("Y1–Y3 revenue, cost and net profit; the HK$0.29M capitalisation and "
              "the IAS 38 recognition criteria; the opening balance sheet.", False)])
bullet(doc, [("Does not validate: ", True),
             ("the Y4–Y6 forecast; the market opportunity; whether the intangible "
              "will generate the cash flows it represents; or the valuation itself. "
              "An auditor gives an opinion on historical statements, not on future "
              "prospects or fair value.", False)])

doc.add_heading("The engagement type matters", level=2)
table(doc, ["Engagement", "Standard", "Assurance level", "Weight in negotiation"],
      [["Audit", "HKSA", "Reasonable",
        "Full — supports the 30–35% rate argument"],
       ["Review", "HKRE 2400", "Limited",
        "Partial — expect the rate argument to be discounted"],
       ["Compilation", "HKSA 4410 / none", "None",
        "Minimal — treat as founder-prepared statements"]],
      [3.0, 3.5, 3.5, 7.5], highlight_col=3)
note(doc, "State the engagement type explicitly in the information memorandum. If "
          "the statements are reviewed or compiled rather than audited, the discount-"
          "rate argument weakens considerably and the memo's conclusions should be "
          "adjusted accordingly.")

doc.add_heading("The narrative shift", level=2)
bold_para(doc, [
    ("At Y0 the company was raising to survive. ", True),
    ("At Y4 Q1 it is raising to accelerate. Those are different conversations, and "
     "the second is stronger: the investor is buying a small position in a proven "
     "business, not a lottery ticket, and the return comes from a shorter, "
     "lower-risk step-up rather than a large multiple on day one.", False),
])
bold_para(doc, [
    ("A 0.9–1.1% stake for HK$1.0M is a strategic or bridge position. ", True),
    ("Some investors will find that attractive; others will pass because the cheque "
     "is too small to move their fund. Both are fine — the round does not need to "
     "close with any single investor, and the audit is what makes that indifference "
     "credible.", False),
])

# ---- 5. Retained profits variant
doc.add_heading("5. Variant — If Y1–Y3 Profits Are Retained", level=1)
doc.add_paragraph(
    "The base case assumes Y1–Y3 profits are reinvested, so they are already captured "
    "in the Y4–Y6 cash flows. If they are instead retained as cash on the balance "
    "sheet, they add to the Y4 Q1 valuation directly."
)
rows = []
for c in CASES:
    rows.append([
        c,
        f"{cumulative_y1_y3[c]:.2f}",
        money(y4q1_base[(c, 0.50)]),
        money(y4q1_retained[(c, 0.50)]),
        f"{ownership_ret[(c, 0.50)]*100:.2f}%",
    ])
table(doc, ["Case", "Cumulative Y1–Y3",
            "Y4 Q1 (reinvested)", "Y4 Q1 (retained)",
            "Ownership at HK$1M (retained)"],
      rows, [2.5, 3.3, 3.6, 3.6, 4.0], highlight_col=3)
note(doc, "If profits are retained, an HK$1M cheque buys less. The base case is the "
          "conservative reading and should be the position in negotiation.")

# ---- 6. Step-up
doc.add_heading("6. Step-up from Y0 to Y4 Q1", level=1)
doc.add_paragraph(
    "The step-up is the multiple of Y4 Q1 value over Y0 value, expressed both as a "
    "total multiple and as an annualised return over the 3.25-year interval."
)
rows = []
for c in CASES:
    for r in RATES:
        rows.append([
            c, f"{int(r*100)}%",
            money(y0[(c, r)]),
            money(y4q1_base[(c, r)]),
            f"{step_up[(c, r)]:.2f}×",
            f"{annual_return[(c, r)]*100:.1f}%",
        ])
table(doc, ["Case", "Rate", "Y0", "Y4 Q1", "Step-up", "Annualised"],
      rows, [2.5, 1.8, 3.2, 3.4, 2.5, 2.5], highlight_col=4)
note(doc, "The step-up is consistent across cases at each rate because the three "
          "cases share the same cash-flow shape. At 50%, the annualised step-up is "
          "31.9% — the return a Y0 backer earns if the plan lands and Y4 Q1 prices at "
          "the same rate.")

# ---- 7. HK$1M seed
doc.add_heading("7. The HK$1.0M Seed at Y4 Q1", level=1)
doc.add_paragraph(
    f"At each case and rate, the ownership transferred for an {money(SEED_Y4Q1)} "
    f"cheque. This is a small dilution by design — the company is profitable by Y4 Q1 "
    f"and the cheque is for acceleration, not survival."
)
rows = []
for c in CASES:
    rows.append([c] + [f"{ownership[(c, r)]*100:.2f}%" for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("Post-money and investor position (Centre × 50%)", level=2)
pre = y4q1_base[("Centre", 0.50)]
post = pre + SEED_Y4Q1
rows = [
    ["Pre-money", money(pre), ps_fmt(pre)],
    ["Seed cheque", money(SEED_Y4Q1), "—"],
    ["Post-money", money(post), ps_fmt(post)],
    ["Investor ownership", f"{SEED_Y4Q1/post*100:.2f}%", "—"],
    ["Founder ownership", f"{(1-SEED_Y4Q1/post)*100:.2f}%", "—"],
]
table(doc, ["Item", "Value", "Per share"], rows,
      [5.0, 5.0, 5.0])

# ---- 8. Negotiation
doc.add_heading("8. Negotiation at Y4 Q1", level=1)

doc.add_heading("8.1 Positions", level=2)
rows = [
    ["Opening ask", money(y4q1_base[("Ceiling", 0.50)]),
     ps_fmt(y4q1_base[("Ceiling", 0.50)]),
     f"{own(y4q1_base[('Ceiling', 0.50)], SEED_Y4Q1)*100:.2f}%",
     "Ceiling × 50% — upside case at standard seed rate"],
    ["Target", money(y4q1_base[("Centre", 0.50)]),
     ps_fmt(y4q1_base[("Centre", 0.50)]),
     f"{own(y4q1_base[('Centre', 0.50)], SEED_Y4Q1)*100:.2f}%",
     "Centre × 50% — base case at standard seed rate"],
    ["Reservation", money(y4q1_base[("Centre", 0.75)]),
     ps_fmt(y4q1_base[("Centre", 0.75)]),
     f"{own(y4q1_base[('Centre', 0.75)], SEED_Y4Q1)*100:.2f}%",
     "Centre × 75% — base case at max execution risk"],
    ["Walk-away", money(y4q1_base[("Lower", 0.75)]),
     ps_fmt(y4q1_base[("Lower", 0.75)]),
     f"{own(y4q1_base[('Lower', 0.75)], SEED_Y4Q1)*100:.2f}%",
     "Lower × 75% — below this, do not take the cheque"],
]
table(doc, ["Position", "Pre-money", "Per share",
            "Ownership at HK$1M", "Basis"],
      rows, [2.8, 3.0, 3.2, 3.0, 6.3], highlight_col=1, shade_rows=[3])

doc.add_heading("8.2 The leverage is different from Y0", level=2)
bold_para(doc, [
    ("At Y4 Q1 the company does not need the money. ", True),
    ("Every case is profitable by Y3. The HK$1.0M accelerates growth; it does not "
     "fund survival. This is the single most important fact in the negotiation — a "
     "founder who can decline a small cheque has a stronger position than one who "
     "cannot.", False),
])
bold_para(doc, [
    ("The intangible is a floor. ", True),
    (f"HK${INTANGIBLE_Y3:.2f}M of capitalised R&D on the balance sheet is an "
     f"independent, auditable asset that does not depend on the forecast. Even if the "
     f"investor discounts every projection year, the floor is not zero.", False),
])
bold_para(doc, [
    ("A small cheque for a small stake is not a problem to solve. ", True),
    (f"At the Target, HK$1.0M buys {ownership[('Centre', 0.50)]*100:.2f}%. If the "
     f"investor needs a larger position, the correct response is to take a larger "
     f"cheque at the same price — not to accept a lower price for a larger stake.", False),
])

doc.add_heading("8.3 If the investor anchors low", level=2)
doc.add_paragraph(
    "The likely anchor at Y4 Q1 is not HK$7.80M — that cell is irrelevant now because "
    "the funding date has moved. The likely anchor is a discount to the Y0 price on "
    "the theory that the company is still pre-revenue in reality. Three responses:"
)
bullet(doc, [("Name the asymmetry. ", True),
             ("You are asking me to accept a Y4 Q1 price below the Y0 price, after "
              "three years of profits and HK$0.29M of capitalised IP. That is not a "
              "discount for risk — it is a discount for my having executed.", False)])
bullet(doc, [("Offer a structure instead. ", True),
             ("A SAFE with a cap at the Target, converting at the next round, keeps "
              "the cap table clean and defers the argument to when the pilot data "
              "exists.", False)])
bullet(doc, [("Decline and compound. ", True),
             ("The company is profitable. Declining HK$1.0M and retaining earnings "
              "for another 12 months costs less than accepting a permanently depressed "
              "reference price.", False)])

# ---- 9. Risks
doc.add_heading("9. Risks and Caveats", level=1)
bullet(doc, [("Y1–Y3 profits are model outputs. ", True),
             ("The entire Y4 Q1 case rests on three years of profitability that have "
              "not happened yet. If Y1–Y3 land below plan, the Y4 Q1 valuation falls "
              "and the HK$1.0M may not be available on these terms.", False)])
bullet(doc, [("The Y4 Q1 investor prices the residual, not the history. ", True),
             ("Past profits are sunk. The valuation is set by Y4–Y6 plus the "
              "intangible. A Y4 Q1 investor who believes only the Lower case will "
              "price at the walk-away row regardless of how well Y1–Y3 went.", False)])
bullet(doc, [("Discount rates at Y4 Q1 should arguably be lower. ", True),
             ("The company will have three years of profitable operation and an "
              "audited intangible. A 25–35% rate may be more appropriate than the "
              "50% used here. The memo uses 50% because it is the standard seed target "
              "return and preserves comparability with Y0, but expect the investor to "
              "argue for a lower rate — which favours the founder.", False)])
bullet(doc, [("The three cases remain scalar multiples. ", True),
             ("Centre is approximately 2.57× Lower and Ceiling 6.25× Lower in every "
              "year. This flaw is unchanged from earlier memos and should be fixed "
              "before any Y4 Q1 negotiation.", False)])
bullet(doc, [("Share count is a convention. ", True),
             (f"Per-share prices are on {SHARE_COUNT} shares. Confirm the authorised "
              f"count with counsel before quoting a price externally.", False)])

# ---- 10. Recommendation
doc.add_heading("10. Recommendation", level=1)
bold_para(doc, [
    ("Price the HK$1.0M at Y4 Q1 on the base case. ", True),
    (f"Target {money(y4q1_base[('Centre', 0.50)])} pre-money "
     f"({ps_fmt(y4q1_base[('Centre', 0.50)])}/share), giving the investor "
     f"{ownership[('Centre', 0.50)]*100:.2f}%. Open at "
     f"{money(y4q1_base[('Ceiling', 0.50)])}. Do not close below "
     f"{money(y4q1_base[('Centre', 0.75)])}.", False),
])
bold_para(doc, [
    ("Quote the step-up, not the absolute. ", True),
    (f"The Y4 Q1 price is {step_up[('Centre', 0.50)]:.2f}× the Y0 price. That is the "
     f"number a Y4 Q1 investor will test against their own return requirement, and it "
     f"is the number that justifies the higher entry point.", False),
])
bold_para(doc, [
    ("Do not negotiate against the Y0 anchor. ", True),
    ("The Y0 price is a historical reference. An investor who offers less than the Y0 "
     "price for a de-risked company is not pricing risk — they are pricing the "
     "founder's need for the cheque. If the need is not real, decline.", False),
])
bold_para(doc, [
    ("Before the meeting. ", True),
    ("Confirm that Y1–Y3 are on track, that the intangible is audited at HK$0.29M, "
     "and that the company can decline the HK$1.0M without impairing operations. The "
     "third condition is the one that determines the outcome.", False),
])

out = "Seed_Valuation_Y4Q1_v8.docx"
doc.save(out)

# ---- Console verification
print(f"Saved: {out}\n")
print("Y0 reference (HK$M / per share):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        v = y0[(c, r)]
        line += f"  {int(r*100):>3}%: {v:>7.2f} ({ps_fmt(v):>12})"
    print(line)
print("\nY4 Q1 (HK$M / per share):")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        v = y4q1_base[(c, r)]
        line += f"  {int(r*100):>3}%: {v:>7.2f} ({ps_fmt(v):>12})"
    print(line)
print("\nStep-up Y4 Q1 / Y0:")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {step_up[(c, r)]:>5.2f}x"
    print(line)
print("\nOwnership at HK$1.0M seed:")
for c in CASES:
    line = f"  {c:<8}"
    for r in RATES:
        line += f"  {int(r*100):>3}%: {ownership[(c, r)]*100:>5.2f}%"
    print(line)
print(f"\nTarget (Centre x 50%): pre {money(y4q1_base[('Centre', 0.50)])}, "
      f"post {money(y4q1_base[('Centre', 0.50)] + SEED_Y4Q1)}, "
      f"investor {ownership[('Centre', 0.50)]*100:.2f}%")