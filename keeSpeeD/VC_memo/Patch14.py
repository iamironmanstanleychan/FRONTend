"""
Kung Fu AI Football — Global Scenario Stress-Test (v14, INTERNAL)
Requires: pip install python-docx
Standalone. Run directly. Outputs Global_Scenarios_v14_INTERNAL.docx

Companion to Valuation_Route_v13_INTERNAL.docx. Designed to be appended as
Section 9 of v13, or read as a standalone addendum.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ===========================================================================
# PARAMETERS — edit the scenario assumptions here
# ===========================================================================

CHEQUE = 1.0                 # HK$M, base seed cheque
SHARE_COUNT = 100
MIDPOINT_CENTRE = 134.23     # HK$M, v13 Step D Centre x 50%
MIDPOINT_LOWER = 52.54
MIDPOINT_CEILING = 326.11

# Scenario: (label, SA_low, SA_base, SA_high, survival, required_mult, note)
SCENARIOS = [
    ("S0  Base (v13 assumptions)", 134.0, 200.0, 400.0, 0.05, 20.0,
     "Series A clears at the midpoint or better; 5% survival; 20x fund floor."),
    ("S1  Extended funding winter", 70.0, 100.0, 150.0, 0.04, 25.0,
     "Later, smaller Series A; funds concentrate, so required multiple rises."),
    ("S2  China youth-sports tailwind", 200.0, 300.0, 450.0, 0.06, 15.0,
     "Policy-driven demand; more competitors, but exit prices rise."),
    ("S3  Consumer / education contraction", 50.0, 70.0, 100.0, 0.03, 30.0,
     "Renewals fall, CAC rises; survival drops and multiples harden."),
    ("S4  AI compute / export constraint", 90.0, 130.0, 180.0, 0.05, 22.0,
     "Training and edge-inference costs rise; margin compression, capex up."),
    ("S5  HK financial tailwind", 150.0, 250.0, 400.0, 0.06, 15.0,
     "18C listing route, GBA schemes, HK co-investment lower cost of capital."),
]

COUNTER_CYCLICAL = [
    ("DIPN 55 — 300% R&D deduction",
     "Up to HK$2M qualifying spend",
     "HK$0.40–0.60M cash saving per year",
     "Worth more when cash is scarce; unchanged in all scenarios"),
    ("HKAS 38 — capitalised development",
     "HK$0.29M at Y3, HK$0.58M at Y6",
     "Balance-sheet floor of HK$0.29–0.58M",
     "Provides a valuation floor independent of forecast"),
    ("PolyU Entrepreneurship Investment Fund",
     "Up to HK$4M co-investment",
     "Up to HK$4M matched capital",
     "Reduces reliance on the seed cheque entirely"),
    ("Research Talent Hub / ITF",
     "Researcher salary support",
     "HK$0.3–0.8M per researcher-year",
     "Directly lowers the burn that drives the Lowest case"),
    ("BUD / GBA subsidy schemes",
     "Project-based, up to HK$7M cumulative",
     "HK$1–3M over 3 years",
     "Funds channel development without diluting equity"),
]

# ===========================================================================
# HELPERS
# ===========================================================================

def set_cell_background(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
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
    t.style = "Table Grid"; t.alignment = WD_TABLE_ALIGNMENT.CENTER
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


def stake(C, M, SA):
    """Governing equation: stake = C*M/SA."""
    return C * M / SA


def pre(C, M, SA):
    """Pre-money implied by the fund's required multiple."""
    return SA / M - C


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

base_mult = 20.0
rows_founder = []
for label, lo, mid_sa, hi, surv, M, _ in SCENARIOS:
    rows_founder.append({
        "label": label, "lo": lo, "mid": mid_sa, "hi": hi,
        "surv": surv, "M": M,
        "pre_20_base": pre(CHEQUE, M, mid_sa),
        "stake_base": stake(CHEQUE, M, mid_sa),
        "pre_lo": pre(CHEQUE, M, lo),
        "stake_lo": stake(CHEQUE, M, lo),
        "pre_hi": pre(CHEQUE, M, hi),
        "stake_hi": stake(CHEQUE, M, hi),
    })

# ===========================================================================
# DOCUMENT
# ===========================================================================

doc = Document()

sec = doc.sections[0]
sec.different_first_page_header_footer = True
hdr = sec.header.paragraphs[0]; hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = hdr.add_run("INTERNAL  ·  CONFIDENTIAL  ·  NOT FOR DISTRIBUTION")
r.bold = True; r.font.size = Pt(8); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

ftr = sec.footer.paragraphs[0]; ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = ftr.add_run("Scenario assumptions are internal planning parameters, not "
                "forecasts of macro events. Verify current data before external use.  "
                "v14 addendum to v13.")
r.italic = True; r.font.size = Pt(7); r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

cb = doc.add_paragraph(); cb.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = cb.add_run("!!   INTERNAL — CONFIDENTIAL   !!")
r.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

cs = doc.add_paragraph(); cs.alignment = WD_ALIGN_PARAGRAPH.CENTER
cs.paragraph_format.space_after = Pt(18)
r = cs.add_run("Global scenario stress-test. Contains planning assumptions and "
               "negotiation strategy.")
r.italic = True; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("GLOBAL SCENARIO STRESS-TEST"); r.bold = True; r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()
for label, value in [
    ("To:", "Founders / Board"),
    ("From:", "Finance & Strategy"),
    ("Re:", "Seed Pricing Under Six Macro Scenarios"),
    ("Classification:", "INTERNAL — CONFIDENTIAL — NOT FOR DISTRIBUTION"),
    ("Status:", "Addendum to v13. All macro parameters are planning assumptions."),
    ("Version:", "v14 — scenario framework, governing equation, decision rules."),
]:
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(2)
    rr = p.add_run(label + " "); rr.bold = True
    if label == "Classification":
        rr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        v = p.add_run(value); v.bold = True
        v.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    else:
        p.add_run(value)

doc.add_paragraph()

# ---- 1. The governing equation
doc.add_heading("1. The Governing Equation", level=1)
bold_para(doc, [
    ("Founder dilution at seed = (cheque × required multiple) ÷ Series A price.", True),
])
doc.add_paragraph(
    "If a fund requires M× on its position, it must own M·C ÷ SA of the company, "
    "where C is the cheque and SA is the Series A pre-money. The pre-money follows "
    "as SA ÷ M − C. Everything else in this memo is a consequence of that identity."
)

rows = []
for chk, M, SA in [(1.0, 20, 134.23), (1.0, 20, 200.0), (0.5, 20, 134.23),
                   (1.0, 25, 100.0), (1.0, 15, 300.0)]:
    rows.append([
        money(chk), f"{M:.0f}×", money(SA),
        f"{stake(chk, M, SA)*100:.2f}%",
        money(pre(chk, M, SA)),
    ])
table(doc, ["Cheque", "Required multiple", "Series A pre-money",
            "Founder dilution", "Implied seed pre-money"],
      rows, [2.2, 3.0, 3.6, 3.2, 3.6], highlight_col=3)

note(doc, "Verify against v13: HK$1.0M × 20 ÷ HK$134.23M = 14.90%, matching the "
          "Scenario A stake in v13 G.9. The identity holds.")

bold_para(doc, [
    ("Three consequences. ", True),
    ("(i) Dilution is independent of the pre-money you negotiate — it is fixed by "
     "the Series A price and the fund's required multiple. (ii) A downturn raises "
     "dilution for the same cheque, because SA falls while M holds. (iii) The only "
     "levers are C, M and SA — smaller cheque, lower required multiple, or a higher "
     "Series A clearing price.", False),
])

# ---- 2. The six scenarios
doc.add_heading("2. The Six Scenarios", level=1)
doc.add_paragraph(
    "Each scenario is defined by four parameters: the Series A clearing range, the "
    "survival rate, and the multiple the fund must return on a surviving position. "
    "They are planning assumptions, not predictions. The purpose is to test whether "
    "the seed price holds up across worlds, not to forecast which world arrives."
)

rows = []
for label, lo, mid_sa, hi, surv, M, desc in SCENARIOS:
    rows.append([label, f"{money(lo)} – {money(hi)}", f"{surv*100:.0f}%",
                 f"{M:.0f}×", desc])
table(doc, ["Scenario", "Series A range", "Survival", "Required mult", "Rationale"],
      rows, [4.0, 3.2, 1.8, 2.0, 7.5])

note(doc, "S1 and S3 are the downside worlds. S2 and S5 are the upside worlds. S4 is "
          "a cost shock that compresses margin without collapsing demand. S0 is the "
          "v13 base case and is retained for comparison.")

# ---- 3. Scenario outcomes
doc.add_heading("3. Scenario Outcomes — Pre-money and Dilution", level=1)
doc.add_paragraph(
    "For each scenario, the seed pre-money and founder dilution at the low, base and "
    "high ends of the Series A range, holding the cheque at HK$1.0M and applying each "
    "scenario's required multiple."
)

rows = []
for s in rows_founder:
    rows.append([
        s["label"],
        f"{s['M']:.0f}×",
        f"{money(s['pre_lo'])} / {money(s['pre_20_base'])} / {money(s['pre_hi'])}",
        f"{s['stake_lo']*100:.1f}% / {s['stake_base']*100:.1f}% / {s['stake_hi']*100:.1f}%",
    ])
table(doc, ["Scenario", "Required mult",
            "Seed pre-money (low / base / high)",
            "Dilution (low / base / high)"],
      rows, [4.2, 2.0, 6.0, 5.0], highlight_col=1)

bold_para(doc, [
    ("The downside world costs more than it looks. ", True),
    ("In S3 (contraction), a HK$1.0M cheque requires a 30× return at a HK$70M Series "
     "A — dilution of 42.9%. In S0 (base), the same cheque at a HK$200M Series A "
     "costs 10.0%. The difference is 33 percentage points of the company for the "
     "same money, driven entirely by the market, not by the negotiation.", False),
])

# ---- 4. Cheque-size lever
doc.add_heading("4. The Cheque-Size Lever", level=1)
doc.add_paragraph(
    "Because dilution = C·M/SA, halving the cheque halves the dilution. This is the "
    "only lever fully within the founder's control, and it becomes most valuable "
    "exactly when the market is worst."
)

cheques = [1.0, 0.75, 0.5, 0.25]
rows = []
for s in rows_founder:
    if s["label"].startswith("S3") or s["label"].startswith("S0"):
        row = [s["label"], f"{s['M']:.0f}×"]
        for C in cheques:
            row.append(f"{stake(C, s['M'], s['mid'])*100:.1f}%")
        rows.append(row)
table(doc, ["Scenario", "Required mult"] + [money(c) for c in cheques],
      rows, [4.2, 2.0, 2.8, 2.8, 2.8, 2.8], highlight_col=3)

note(doc, "In S3, halving the cheque from HK$1.0M to HK$0.5M cuts dilution from "
          "42.9% to 21.4%. The pre-money rises, because it is SA/M − C. The fund's "
          "absolute return halves, which is why the fund may refuse — a portfolio-"
          "construction constraint, not a valuation disagreement. Ask before the "
          "meeting whether a HK$0.5M position is above their minimum.")

# ---- 5. The promotion problem
doc.add_heading("5. The Promotion Problem — Downturn Promotes the Lower Case", level=1)
bold_para(doc, [
    ("A downturn does not change your model. It promotes your Lower case to base "
     "case. ", True),
    ("In S1 and S3, the Series A clears at HK$70–150M — the range that corresponds "
     "to the Lower column of the v13 midpoint matrix, not the Centre. The case the "
     "founder treats as the downside becomes the planning case. That has a direct "
     "consequence for this document's credibility.", False),
])
bold_para(doc, [
    ("The Lower case is still Centre ÷ 2.57. ", True),
    ("Flagged in every version since v9 and still outstanding. In a benign "
     "environment it is a methodological weakness. In S1/S3 it is fatal: if the "
     "Lower case becomes the base case and it is arithmetically the Centre case "
     "scaled to 39%, then in a downturn the company has no credible downside at all "
     "— it has a smaller version of its upside. A fund doing diligence in a downturn "
     "will test exactly this case and will find the proportionality in one pass, "
     "because Step E of v13 displays it openly.", False),
])
bold_para(doc, [
    ("This is now the highest-priority fix in the entire pack. ", True),
    ("Higher than the audit, higher than the negotiation script. The global scenarios "
     "make it urgent because they identify the specific world in which the flaw "
     "becomes decisive. Rebuild the Lower case with a distinct cost structure — "
     "channel failure, subscription-attach failure, back-loaded ramp — before the "
     "next investor conversation, not before the next draft.", False),
])

# ---- 6. Counter-cyclical levers
doc.add_heading("6. Counter-Cyclical Levers — HK Incentives in a Downturn", level=1)
doc.add_paragraph(
    "The HK R&D and co-investment incentives are worth more in a downturn, not less. "
    "They reduce the cash burn that drives the Lower case, and they substitute for "
    "the equity cheque the market is unwilling to price."
)
rows = [[a, b, c, d] for a, b, c, d in COUNTER_CYCLICAL]
table(doc, ["Instrument", "Basis", "Estimated contribution", "Why it matters in a downturn"],
      rows, [4.5, 3.5, 3.5, 4.5], highlight_col=2)

bold_para(doc, [
    ("The PolyU fund is the most important line. ", True),
    ("Up to HK$4M in co-investment capital, against a HK$1.0M seed cheque, means the "
     "incentive pool is four times the round. In S1/S3, pursuing the PolyU route "
     "instead of a priced seed may be strictly better: less dilution, no price "
     "anchor set at a depressed valuation, and the pilot still gets funded. That "
     "option should be costed before the next VC meeting, not after.", False),
])

# ---- 7. Decision rules
doc.add_heading("7. Decision Rules by Scenario", level=1)
rows = [
    ["S0  Base", "Raise HK$1.0M at HK$5.7–13.4M",
     "Proceed. Target Scenario B or C from v13 G.9."],
    ["S1  Extended winter", "Raise HK$0.5M, or defer and use PolyU/BUD",
     "Do not set a price anchor at HK$70–100M. Half-cheque or non-dilutive first."],
    ["S2  Youth-sports tailwind", "Raise HK$1.0M at the higher end",
     "Demand is there; expect competitors. Use the tailwind to justify a higher SA."],
    ["S3  Contraction", "Do not raise a priced round",
     "Retain earnings; pursue non-dilutive incentives. A priced round here sets a "
     "floor that is hard to repair."],
    ["S4  Compute constraint", "Raise HK$1.0M but ring-fence capex",
     "Model the margin compression explicitly; the Centre case may not survive."],
    ["S5  HK financial tailwind", "Raise HK$1.0M at the higher end",
     "Best environment for a priced round: higher SA, lower required multiple."],
]
table(doc, ["Scenario", "Action", "Rationale"],
      rows, [3.5, 5.5, 7.0], highlight_col=1)

note(doc, "The rule that generalises: raise a priced round when SA is high and the "
          "required multiple is low (S2, S5). Use non-dilutive capital and retained "
          "earnings when SA is low and the multiple is high (S1, S3). Do not raise "
          "into a depressed price, because the anchor persists into every subsequent "
          "round.")

# ---- 8. Positions by scenario
doc.add_heading("8. Revised Positions by Scenario", level=1)
doc.add_paragraph(
    "For each scenario, the seed pre-money that delivers the fund its required "
    "multiple at the scenario's base Series A price, and the founder stake that "
    "results. Compare against the v13 positions."
)
rows = []
for s in rows_founder:
    rows.append([
        s["label"],
        money(s["mid"]),
        f"{s['M']:.0f}×",
        money(s["pre_20_base"]),
        f"{s['stake_base']*100:.1f}%",
    ])
table(doc, ["Scenario", "Series A (base)", "Required mult",
            "Seed pre-money", "Founder dilution"],
      rows, [4.2, 2.8, 2.0, 3.2, 3.2], highlight_col=3)

bold_para(doc, [
    ("Read this table before the meeting, not the v13 matrix. ", True),
    ("The v13 matrix tells you what the company is worth when proven. This table "
     "tells you what the market will pay in each world, and what it costs you. If "
     "the founder walks in with HK$134M and the world is S1, the conversation is "
     "over before it starts. The table is the preparation.", False),
])

# ---- 9. Caveats
doc.add_heading("9. Caveats and Verification", level=1)
bullet(doc, [("The macro parameters are assumptions. ", True),
             ("The Series A ranges, survival rates and required multiples in Section 2 "
              "are internal planning figures. They are not forecasts and not data. "
              "Verify against current market conditions before any external use.",
              False)])
bullet(doc, [("Specific policy references need checking. ", True),
             ("DIPN 55, HKAS 38, the PolyU fund, Research Talent Hub and the BUD "
              "scheme are real instruments, but the qualifying amounts and conditions "
              "change. Confirm current terms with the relevant body before relying on "
              "any figure in Section 6.", False)])
bullet(doc, [("FX exposure is not captured. ", True),
             ("HK$ is pegged to USD, so the cheque has no HK$/USD risk. But if revenue "
              "is RMB-denominated and the investor's return is USD-denominated, RMB "
              "depreciation reduces the USD value of the exit. That channel is "
              "material in S1/S3 and should be modelled separately.", False)])
bullet(doc, [("Survival rates are portfolio-level, not company-level. ", True),
             ("5% is a fund-construction assumption. It says nothing about this "
              "company's probability of success. Do not let an investor use a "
              "portfolio survival rate to argue a company-specific discount.", False)])
bullet(doc, [("The Lower case is still not rebuilt. ", True),
             ("Section 5 explains why this is now urgent. Until it is fixed, the "
              "downturn scenarios cannot be modelled honestly.", False)])

# ---- 10. Recommendation
doc.add_heading("10. Recommendation", level=1)
bold_para(doc, [
    ("Prepare three scenarios, not six. ", True),
    ("S0 (base), S1 (winter) and S2 (tailwind) cover the realistic range. S3–S5 are "
     "sensitivities to discuss if asked, not positions to present. Six scenarios in "
     "the room reads as hedging.", False),
])
bold_para(doc, [
    ("Lead with the governing equation. ", True),
    ("Dilution = C·M/SA is the single most useful sentence in the document. It "
     "converts the conversation from \"what is my company worth?\" to \"which of the "
     "three levers can we move?\" — and every one of those is winnable.", False),
])
bold_para(doc, [
    ("In S1/S3, do not raise a priced round. ", True),
    ("Use the PolyU co-investment, BUD and retained earnings. A depressed price "
     "anchor persists into every subsequent round; the cost of raising in a downturn "
     "is not the dilution today, it is the reference price for the next five years.",
     False),
])
bold_para(doc, [
    ("Fix the Lower case before the next conversation. ", True),
    ("This is the outstanding item and the scenarios have made it decisive. In a "
     "downturn the Lower case becomes the base case; if it is Centre ÷ 2.57, the "
     "company has no credible downside in exactly the world where it needs one.",
     False),
])

out = "Global_Scenarios_v14_INTERNAL.docx"
doc.save(out)

# ---- Console verification
print(f"Saved: {out}\n")
print(f"{'Scenario':<34}{'SA base':>9}{'M':>5}{'Pre':>10}{'Dilution':>10}")
print("-" * 68)
for s in rows_founder:
    print(f"{s['label']:<34}{money(s['mid']):>9}{s['M']:>4.0f}x"
          f"{money(s['pre_20_base']):>10}{s['stake_base']*100:>9.1f}%")
print("\nCheque-size lever (base SA per scenario):")
print(f"{'Scenario':<34}" + "".join(f"{money(c):>10}" for c in cheques))
for s in rows_founder:
    line = f"{s['label']:<34}"
    for C in cheques:
        line += f"{stake(C, s['M'], s['mid'])*100:>9.1f}%"
    print(line)
print("\nGoverning equation check (v13 G.9 Scenario A):")
print(f"  1.0 x 20 / 134.23 = {stake(1.0, 20, 134.23)*100:.2f}%  "
      f"(expected 14.90%)")