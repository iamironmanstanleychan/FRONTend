"""
Kung Fu AI Football — The Valuation Route (v13, INTERNAL)
Requires: pip install python-docx
Standalone script — run directly. Outputs Valuation_Route_v13_INTERNAL.docx
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


def ps_fmt(val_m):
    return f"HK${val_m * 1_000_000 / SHARE_COUNT:,.0f}"


def stake_post(pre_m):
    return SEED_CHEQUE / (pre_m + SEED_CHEQUE)


def shares_at_pre(pre_m):
    return SHARE_COUNT * SEED_CHEQUE / pre_m


def shares_at_post(pre_m):
    return SHARE_COUNT * SEED_CHEQUE / (pre_m + SEED_CHEQUE)


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

y7q2 = {c: REPORTED_UNDISCOUNTED[c] + INTANGIBLE_Y6 for c in CASES}

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

mid = {(c, r): (y0[(c, r)] + y7q2[c]) / 2 for c in CASES for r in RATES}


def y0_value(case, r):
    pv = sum(DATA[t][CASE_IDX[case]] / (1 + r) ** (t + 1) for t in range(6))
    pv += GRANT / (1 + r) ** GRANT_YEARS
    pv += INTANGIBLE_Y6 / (1 + r) ** INTANGIBLE_YEARS
    return pv


def implied_rate(case, target):
    lo, hi = 0.0005, 3.0
    if y0_value(case, lo) < target:
        return None
    if y0_value(case, hi) > target:
        return None
    for _ in range(120):
        m = (lo + hi) / 2
        if y0_value(case, m) > target:
            lo = m
        else:
            hi = m
    return (lo + hi) / 2


implied = {(c, r): implied_rate(c, mid[(c, r)]) for c in CASES for r in RATES}
shares_pre = {(c, r): shares_at_pre(mid[(c, r)]) for c in CASES for r in RATES}
shares_post = {(c, r): shares_at_post(mid[(c, r)]) for c in CASES for r in RATES}

POSITIONS = {
    "Opening ask": ("Ceiling", 0.50),
    "Target":      ("Centre",  0.50),
    "Settle":      ("Lower",   0.50),
    "Walk-away":   ("Lower",   0.75),
}
pos = {k: mid[v] for k, v in POSITIONS.items()}
ask_pre = SEED_CHEQUE / INVESTOR_ASK - SEED_CHEQUE

# ===========================================================================
# DOCUMENT — Patch 1: classification header / footer / cover block
# ===========================================================================

doc = Document()

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

st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

# Cover banner
cover_box = doc.add_paragraph()
cover_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
cover_box.paragraph_format.space_before = Pt(0)
r = cover_box.add_run("!!   INTERNAL — CONFIDENTIAL   !!")
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

# Title
t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("MEMORANDUM"); r.bold = True; r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()

# Patch 2: meta table
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

# ===========================================================================
# SECTION 0 — The route
# ===========================================================================

doc.add_heading("0. The Route", level=1)
doc.add_paragraph(
    "This document derives the midpoint valuation in seven steps, each consuming the "
    "output of the one before. Nothing is asserted that is not built from the prior "
    "step."
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
carry(doc, "the three Y7 Q2 values form the upper bound in Step D.")

# ===========================================================================
# STEP C
# ===========================================================================

doc.add_heading("C. Y0 Adjusted — The Dream Price", level=1)

doc.add_heading("C.1 The reported NPVs", level=2)
rows = []
for c in CASES:
    rows.append([c] + [money(REPORTED_NPV[c][r]) for r in RATES])
table(doc, ["Case", "25%", "50%", "75%"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

doc.add_heading("C.2 The two adjustments", level=2)
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
carry(doc, "the nine adjusted cells form the lower bound in Step D.")

# ===========================================================================
# STEP D
# ===========================================================================

doc.add_heading("D. Midpoint of C and B — The Nine-Cell Matrix", level=1)
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
carry(doc, "the nine midpoint values are converted to implied rates in Step E and to "
           "share allocations in Step F.")

# ===========================================================================
# STEP E
# ===========================================================================

doc.add_heading("E. Implied Discount Rate for Each Cell", level=1)
rows = []
for c in CASES:
    row = [c]
    for r in RATES:
        ir = implied[(c, r)]
        row.append(f"{ir*100:.1f}%" if ir is not None else "—")
    rows.append(row)
table(doc, ["Case", "25% column", "50% column", "75% column"], rows,
      [3.5, 4.0, 4.0, 4.0], highlight_col=2)

bold_para(doc, [
    ("The 50% column is the centre of gravity. ", True),
    (f"Implied rates there are {implied[('Lower', 0.50)]*100:.0f}%, "
     f"{implied[('Centre', 0.50)]*100:.0f}% and {implied[('Ceiling', 0.50)]*100:.0f}% "
     f"for Lower, Centre and Ceiling. A rate of roughly 12% is what you apply to a "
     f"profitable private business with three years of audited accounts.", False),
])
carry(doc, "the rate matrix establishes that the midpoint is defensible. Step F "
           "converts the midpoint values into equity.")

# ===========================================================================
# STEP F
# ===========================================================================

doc.add_heading("F. Shares Allocated to the VC for HK$1.0M", level=1)

doc.add_heading("F.1 Priced off pre-money (the trap)", level=2)
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

doc.add_heading("F.4 Founder and investor position at each midpoint cell", level=2)
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
# Patch 3a — Section G opening divider
# ===========================================================================

doc.add_page_break()

stop_bar = doc.add_paragraph()
stop_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_bar.add_run("═" * 58)
r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_title = doc.add_paragraph()
stop_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_title.add_run("STOP  ·  SECTION G BEGINS HERE")
r.bold = True; r.font.size = Pt(14); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_warn = doc.add_paragraph()
stop_warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_warn.add_run(
    "DO NOT DISTRIBUTE THIS SECTION.\n"
    "Contains founder positions, counterpart responses, and negotiation scripts.\n"
    "Strip this section in full before any investor communication."
)
r.italic = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_bar2 = doc.add_paragraph()
stop_bar2.alignment = WD_ALIGN_PARAGRAPH.CENTER
stop_bar2.paragraph_format.space_after = Pt(14)
r = stop_bar2.add_run("═" * 58)
r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_heading("G. Positioning and Counterpart", level=1)

# ===========================================================================
# SECTION G
# ===========================================================================

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
        k, money(v), ps_fmt(v),
        f"{shares_post[(POSITIONS[k][0], POSITIONS[k][1])]:.2f}",
        f"{stake_post(v)*100:.2f}%",
        basis_text[k],
    ])
table(doc, ["Position", "Pre-money", "Per share",
            "VC shares of 100", "VC %", "Basis"],
      rows, [2.6, 2.8, 2.8, 2.8, 2.2, 6.0], highlight_col=4, shade_rows=[3])

doc.add_heading("G.2 Counterpart positions and how to answer", level=2)
rows = [
    [f"{INVESTOR_ASK*100:.2f}% stake at HK$1.0M", money(ask_pre),
     f"{INVESTOR_ASK*100:.2f}%",
     "Below every cell in the Step D matrix, including the most conservative."],
    ["Y0 dream frame", money(y0[("Centre", 0.50)]),
     f"{stake_post(y0[('Centre', 0.50)])*100:.2f}%",
     "Prices the company as though the Y1–Y3 audits will not exist."],
    ["Lower × 75% (the old anchor)", money(mid[("Lower", 0.75)]),
     f"{stake_post(mid[('Lower', 0.75)])*100:.2f}%",
     "Below the walk-away. Decline — the company is profitable by Y3."],
]
table(doc, ["Counterpart position", "Implied pre-money", "VC stake", "Response"],
      rows, [4.0, 3.0, 2.2, 7.0], highlight_col=2)

doc.add_heading("G.3 The three arguments, in order", level=2)

doc.add_paragraph("Argument 1 — the rate.").runs[0].bold = True
quote_para(doc, f"\"The midpoint at the Centre case implies a "
                f"{implied[('Centre', 0.50)]*100:.0f}% discount rate. That is the rate "
                f"you apply to a profitable business with three years of audited "
                f"accounts. It is not a seed rate.\"")

doc.add_paragraph("Argument 2 — the forecast.").runs[0].bold = True
quote_para(doc, f"\"Your {INVESTOR_ASK*100:.2f}% implies {money(ask_pre)}. That is "
                f"below every cell in the matrix. The question is which forecast "
                f"your stake encodes — because it encodes none of them.\"")

doc.add_paragraph("Argument 3 — the walk.").runs[0].bold = True
quote_para(doc, "\"If we cannot bridge, I would rather not raise. The company is "
                "profitable by Y3 and the accumulated cash means the HK$1.0M is "
                "for acceleration, not survival.\"")

doc.add_heading("G.4 What to extract if the stake holds", level=2)
bullet(doc, [("Non-participating preference only. ", True),
             ("1× non-participating. Worth more than the stake concession.", False)])
bullet(doc, [("No board seat. ", True),
             ("Information rights and observer seat at most.", False)])
bullet(doc, [("Pro-rata, not super pro-rata. ", True),
             ("Standard pre-emptive rights only.", False)])
bullet(doc, [("Follow-on price band agreed now. ", True),
             ("Lock in the next-round band before pilot data lands.", False)])

# ===========================================================================
# G.5 — Market Reality Check
# ===========================================================================

doc.add_heading("G.5 Market Reality Check — The 10% Ask", level=2)
doc.add_paragraph(
    "The matrix in Step D is a planning tool. The market will not price against it. "
    "A realistic seed offer for HK$1.0M is 10% of the company — HK$9.0M pre-money, "
    "HK$90,000 per share. That is below every cell in the Step D matrix."
)

rows = [
    ["Y0 Lower × 75% (absolute floor)", money(7.80), "12.8%",
     "The only benchmark below the market offer"],
    ["Market offer (10%)", money(9.00), "10.0%",
     "Fund-arithmetic price, not a DCF output"],
    ["Y0 Lower × 50%", money(14.53), "6.9%",
     "What the Lower case implies at a standard seed rate"],
    ["Midpoint Lower × 75%", money(mid[("Lower", 0.75)]),
     f"{stake_post(mid[('Lower', 0.75)])*100:.2f}%",
     "Lowest cell in the matrix — still 82% above the market offer"],
    ["Midpoint Centre × 50% (target)", money(mid[("Centre", 0.50)]),
     f"{stake_post(mid[('Centre', 0.50)])*100:.2f}%",
     "The crossover target — 93% above the market offer"],
]
table(doc, ["Basis", "Pre-money", "HK$1M buys", "Comment"], rows,
      [5.0, 2.8, 2.5, 6.5], highlight_col=1, shade_rows=[1])

note(doc, "The market offer corresponds to roughly a 70% discount rate applied to "
          "the Lower case — higher than any rate in the matrix. It is fund "
          "arithmetic: a HK$1.0M cheque in a HK$100M fund needs 10% to be worth "
          "managing.")

# ===========================================================================
# G.6 — What this means for negotiation
# ===========================================================================

doc.add_heading("G.6 What this means for the negotiation", level=2)
bullet(doc, [("Do not argue the DCF against the 10% ask. ", True),
             ("Argue fund structure instead: smaller cheque, tranched cheques, "
              "or a follow-on price band agreed now.", False)])
bullet(doc, [("Accept the 10% if the terms are clean. ", True),
             ("HK$9.0M pre-money with non-participating preference and no board "
              "seat is manageable. It sets a reference price, but a clean one.",
              False)])
bullet(doc, [("Reduce the cheque if HK$1.0M is more than the plan needs. ", True),
             ("HK$500K at HK$9.0M pre-money is 5.3% dilution — half the equity "
              "for half the runway.", False)])
bullet(doc, [("Reserve the matrix for the next round. ", True),
             ("The midpoint cells are what the company should be worth at Y5 or "
              "Y6, after the pilot and the audit.", False)])

# ===========================================================================
# G.7 — VC Return Mechanics
# ===========================================================================

doc.add_heading("G.7 VC Return Mechanics — Why the 10× Ask Is Rational", level=2)
bold_para(doc, [
    ("The fund's return equation. ", True),
    ("Return multiple = Series A pre-money ÷ Seed post-money. The Series A "
     "clearing price is set by the market, not by the seed valuation.", False),
])

rows = []
series_a = [(100, "Conservative"), (200, "Base"), (500, "Strong")]
entries = [
    ("Midpoint Centre × 50%", mid[("Centre", 0.50)]),
    ("Midpoint Lower × 50%", mid[("Lower", 0.50)]),
    ("Y0 Lower × 50%", y0[("Lower", 0.50)]),
    ("VC 10% ask", 9.00),
]
for label, pre in entries:
    row = [label, money(pre), f"{stake_post(pre)*100:.2f}%"]
    for pa, _ in series_a:
        mult = pa / (pre + SEED_CHEQUE)
        row.append(f"{mult:.2f}×")
    rows.append(row)
table(doc, ["Seed pre-money", "Pre-money", "HK$1M buys",
            "Return @ HK$100M", "@ HK$200M", "@ HK$500M"],
      rows, [3.2, 2.6, 2.2, 2.8, 2.6, 2.6], highlight_col=2, shade_rows=[0])

note(doc, "At the midpoint, even a strong Series A at HK$500M returns only 3.7×, "
          "and a conservative Series A at HK$100M is a loss. No seed fund can "
          "hold a position that loses money when the company succeeds.")

# ===========================================================================
# G.8 — Three counters
# ===========================================================================

doc.add_heading("G.8 Three counters — none of them is the DCF", level=2)
bullet(doc, [("Argue the Series A clearing price. ", True),
             ("If the pilot supports a HK$200M Series A, post-money ≤ HK$20M gives "
              "the fund 10×. That is HK$19M pre-money — 2.1× their offer.", False)])
bullet(doc, [("Argue the target return. ", True),
             ("10× is the right hurdle for a binary pre-revenue bet. For a "
              "profitable company with three years of audits, 5× is defensible.",
              False)])
bullet(doc, [("Restructure so the fund's downside is covered without 10% at seed.",
              True),
             ("Milestone tranche or a ratchet — both address Series A price risk "
              "without giving away 10% today.", False)])

note(doc, "The practical conclusion: expect to raise at HK$9–19M pre-money for "
          "5–10%, and use the Step D matrix to plan the Series A step-up, not to "
          "fight the seed price.")

# ===========================================================================
# G.9 — 20× ladder (NEW)
# ===========================================================================

doc.add_heading("G.9 Seed Return Scenarios — The 20× Ladder", level=2)

doc.add_paragraph(
    "The seed fund models a 5% survival rate. With $100 committed across 100 "
    "companies at $1 each, five survivors must return $20 each to cover the $100 "
    "cost — a 20× gross multiple per survivor just to break even."
)

bold_para(doc, [
    ("The 20× breakeven math. ", True),
    ("$100 fund ÷ 5% survival = 5 winners required to return $100, each at 20×. "
     "Zero fees and zero carry assumed; with 2% fees and 20% carry, the gross "
     "multiple required per survivor rises toward 25–30×. 20× is the floor.", False),
])

doc.add_heading("The four seed valuation scenarios", level=3)

scenarios = [
    ("A — 20× at midpoint", 5.71),
    ("B — 20× at HK$176M Series A", 7.80),
    ("C — 20× at HK$200M Series A", 9.00),
    ("D — 10× at midpoint", 13.42),
]

rows = []
for label, pre in scenarios:
    stake = SEED_CHEQUE / (pre + SEED_CHEQUE)
    shares = SHARE_COUNT * stake
    mult_mid = mid[("Centre", 0.50)] / (pre + SEED_CHEQUE)
    sa_20 = 20 * (pre + SEED_CHEQUE)
    sa_60 = 60 * (pre + SEED_CHEQUE)
    rows.append([
        label, money(pre), f"{stake*100:.2f}%", f"{shares:.2f}",
        f"{mult_mid:.2f}×", money(sa_20), money(sa_60),
    ])
table(doc, ["Scenario", "Seed pre-money", "HK$1M buys",
            "Shares of 100", "Multiple at SA = HK$134.23M",
            "Series A for 20×", "Series A for 60× (3× fund)"],
      rows, [4.2, 2.6, 2.2, 2.2, 2.6, 2.8, 3.0], highlight_col=4)

note(doc, "Only Scenario A delivers 20× at the midpoint. Scenarios B, C and D "
          "deliver 15.3×, 13.4× and 9.3× respectively — each requires a "
          "progressively higher Series A clearing price to reach 20×.")

doc.add_heading("The multiple matrix — four scenarios across Series A outcomes",
                level=3)

sa_prices = [134.23, 176.00, 200.00, 288.40, 400.00]
header = ["Seed pre-money"] + [money(p) for p in sa_prices]
rows = []
for label, pre in scenarios:
    row = [money(pre)]
    for sa in sa_prices:
        mult = sa / (pre + SEED_CHEQUE)
        marker = " ◄" if abs(mult - 20.0) < 0.15 else ""
        row.append(f"{mult:.1f}×{marker}")
    rows.append(row)
table(doc, header, rows, [3.0, 3.0, 3.0, 3.0, 3.0, 3.0], highlight_col=2)

doc.add_heading("What each scenario means for the negotiation", level=3)

bold_para(doc, [
    ("A — HK$5.71M (14.90% for HK$1M). ", True),
    ("The only seed price at which 20× is delivered at the midpoint. Fund's "
     "return is secured at the base case, but costs the founder 14.90% — roughly "
     "three times the dilution of Scenario D.", False),
])
bold_para(doc, [
    ("B — HK$7.80M (11.36% for HK$1M). ", True),
    ("The anchor the fund will likely propose. Delivers 15.25× at the midpoint; "
     "20× if the Series A clears at HK$176M.", False),
])
bold_para(doc, [
    ("C — HK$9.00M (10.00% for HK$1M). ", True),
    ("The round-number ask. Delivers 13.42× at the midpoint; requires HK$200M at "
     "Series A for 20×.", False),
])
bold_para(doc, [
    ("D — HK$13.42M (6.93% for HK$1M). ", True),
    ("The price the midpoint matrix implies if the fund accepts 10× at the "
     "midpoint. Requires HK$288M at Series A for 20×.", False),
])

doc.add_heading("The pragmatic landing zone", level=3)
bullet(doc, [
    ("Milestone tranche (most likely). ", True),
    ("HK$250K at HK$5.71M (4.2%) + HK$750K at HK$25M post-pilot (2.9%). Blended "
     "entry ≈ HK$14.4M, blended stake ≈ 6.7%.", False),
])
bullet(doc, [
    ("Ratchet. ", True),
    ("HK$1M at HK$13.42M (6.93%), with anti-dilution resetting to HK$5.71M if "
     "Series A clears below HK$176M.", False),
])
bullet(doc, [
    ("Scenario B (acceptable). ", True),
    ("HK$1M at HK$7.80M (11.36%). 15.25× at midpoint is close enough to "
     "breakeven that most funds will treat it as a strong outcome.", False),
])
bullet(doc, [
    ("Scenario C (acceptable only with clean terms). ", True),
    ("HK$1M at HK$9.00M (10.00%). Insist on non-participating preference, no "
     "board seat, standard pro-rata.", False),
])
bullet(doc, [
    ("Scenario A alone (unlikely). ", True),
    ("HK$1M at HK$5.71M (14.90%). Do not accept without a milestone step-up on "
     "the next round.", False),
])

doc.add_heading("The number to argue for", level=3)
bold_para(doc, [
    ("Open at Scenario D (HK$13.42M), land at C or B, refuse A. ", True),
    ("The midpoint matrix says the company is worth HK$134.23M when proven. At "
     "HK$13.42M the fund gets 10× at that valuation. At HK$5.71M the fund gets "
     "20× on the base case — three times the fair return, extracted as dilution.",
     False),
])

# ===========================================================================
# Patch 3b — End of Section G
# ===========================================================================

end_bar = doc.add_paragraph()
end_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
end_bar.paragraph_format.space_before = Pt(18)
r = end_bar.add_run("═" * 58)
r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

end_title = doc.add_paragraph()
end_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = end_title.add_run("END OF SECTION G  ·  INTERNAL CONTENT")
r.bold = True; r.font.size = Pt(11); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

end_bar2 = doc.add_paragraph()
end_bar2.alignment = WD_ALIGN_PARAGRAPH.CENTER
end_bar2.paragraph_format.space_after = Pt(18)
r = end_bar2.add_run("═" * 58)
r.bold = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_page_break()

# ===========================================================================
# SECTION 7 — Caveats
# ===========================================================================

doc.add_heading("7. Caveats", level=1)
bullet(doc, [("The three cases are scalar multiples. ", True),
             ("From Y2 onward Centre is approximately 2.57× Lower and Ceiling "
              "approximately 6.25× Lower in every year. Rebuild the Lower case "
              "with a distinct cost structure before external circulation.", False)])
bullet(doc, [("The midpoint is an average, not a method. ", True),
             ("What gives it force is Step E. If challenged, present the rate, "
              "not the arithmetic.", False)])
bullet(doc, [("The grant and intangible treatment must be consistent. ", True),
             ("Confirm with the auditor that the capitalised amount is the "
              "amortised carrying value, not gross R&D spend.", False)])
bullet(doc, [("Per-share values assume 100 shares. ", True),
             ("Confirm the authorised count with counsel before quoting "
              "externally.", False)])
bullet(doc, [("Three years of audits must actually land. ", True),
             ("The entire matrix is conditional on Y1–Y3 being audited and "
              "profitable.", False)])

# ===========================================================================
# SECTION 8 — Recommendation
# ===========================================================================

doc.add_heading("8. Recommendation", level=1)
bold_para(doc, [
    ("Walk the route in the room. ", True),
    ("Lead with Step A — the P&L — then Step B, then Step C. By the time Step D "
     "appears, the investor has watched it being built. A number that is derived "
     "is harder to dispute than a number that is announced.", False),
])
bold_para(doc, [
    ("Lead with the implied rate. ", True),
    (f"\"The midpoint implies {implied[('Centre', 0.50)]*100:.0f}%\" is a stronger "
     f"opening than any valuation figure. Method arguments are winnable; number "
     f"arguments are not.", False),
])
bold_para(doc, [
    ("Use the midpoint cells for positions. ", True),
    (f"Open {money(pos['Opening ask'])}, target {money(pos['Target'])}, settle "
     f"{money(pos['Settle'])}, walk below {money(pos['Walk-away'])}.", False),
])
bold_para(doc, [
    ("Expect the market to price at HK$9–19M, not at the midpoint. ", True),
    ("The midpoint matrix is a planning tool for the Series A, not a negotiating "
     "position for the seed. The seed price is set by the fund's required return, "
     "which is a different calculation from intrinsic value.", False),
])

# ===========================================================================
# SAVE — Patch 4
# ===========================================================================

out = "Valuation_Route_v13_INTERNAL.docx"
doc.save(out)

# ---- Console verification
print(f"Saved: {out}\n")
print("C — Y0 adjusted (9 cells):")
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
        line += f"  {int(r*100):>3}%: " + (f"{ir*100:>5.1f}%" if ir else "  n/a ")
    print(line)
print("\nG.9 — Seed → 20× ladder:")
for label, pre in scenarios:
    stake = SEED_CHEQUE / (pre + SEED_CHEQUE)
    print(f"  {label:<28} pre {money(pre):>10}  stake {stake*100:>5.2f}%  "
          f"SA@20× {money(20*(pre+SEED_CHEQUE)):>12}")