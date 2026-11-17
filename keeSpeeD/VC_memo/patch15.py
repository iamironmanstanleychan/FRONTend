"""
Seed Structure Comparison (.docx) — one cheque vs two
Requires: pip install python-docx
Standalone. Outputs Seed_Structure_Comparison_INTERNAL.docx
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

SHARE_COUNT = 100
SA_MIDPOINT = 134.23       # HK$M, v13 Step D Centre x 50%
REQUIRED_MULT = 20.0       # fund floor
POST_TARGET = SA_MIDPOINT / REQUIRED_MULT   # HK$6.71M — fixed
PPS = POST_TARGET * 1_000_000 / SHARE_COUNT  # HK$67,100 per share

OPTIONS = [
    ("A — One seed", [0.5]),
    ("B — Two seeds (equal)", [0.5, 0.5]),
    ("C — Two seeds (lead + follower)", [0.75, 0.25]),
    ("D — Two tranches", [0.5, 0.5]),   # second tranche at higher pre
]
TRANCHE_2_PRE = 12.0       # HK$M, second tranche price after milestones

GOVERNANCE_THRESHOLDS = [
    ("Information rights only", "below 5%"),
    ("Observer seat customary", "5–9.9%"),
    ("Board seat customary", "10–19.9%"),
    ("Blocking rights on reserved matters", "20%+"),
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


# ===========================================================================
# COMPUTATIONS
# ===========================================================================

def structure(cheques, second_tranche_pre=None):
    """Return per-investor stake, combined stake, founder retained."""
    stakes = []
    for i, c in enumerate(cheques):
        pre = POST_TARGET - sum(cheques)   # simultaneous
        if second_tranche_pre is not None and i == 1:
            pre = second_tranche_pre
        stakes.append(c / (pre + c))
    combined = 1 - (1 - sum(stakes)) if len(stakes) == 1 else sum(stakes)
    return stakes, combined

structures = []
for label, cheques in OPTIONS:
    if label.startswith("D"):
        # tranche 1 at 6.21 pre, tranche 2 at 12.0 pre
        s1 = 0.5 / (6.21 + 0.5)
        s2 = 0.5 / (TRANCHE_2_PRE + 0.5)
        founder = (1 - s1) * (1 - s2)
        structures.append({
            "label": label, "cheques": cheques,
            "total": sum(cheques),
            "stakes": [s1, s2],
            "combined": 1 - founder,
            "founder": founder,
            "pre_blended": (sum(cheques) / (1 - founder)) - sum(cheques),
        })
    else:
        total = sum(cheques)
        pre = POST_TARGET - total
        stakes = [c / POST_TARGET for c in cheques]
        founder = 1 - sum(stakes)
        structures.append({
            "label": label, "cheques": cheques, "total": total,
            "stakes": stakes, "combined": sum(stakes),
            "founder": founder, "pre_blended": pre,
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
r = ftr.add_run("Structure comparison — addendum to v13 and v14.")
r.italic = True; r.font.size = Pt(7); r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

st = doc.styles["Normal"]
st.font.name = "Calibri"; st.font.size = Pt(10.5)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")

t = doc.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = t.add_run("SEED STRUCTURE COMPARISON"); r.bold = True; r.font.size = Pt(18)
r.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

doc.add_paragraph()
for label, value in [
    ("To:", "Founders / Board"),
    ("Re:", "One cheque vs two — same price, different governance"),
    ("Classification:", "INTERNAL — CONFIDENTIAL — NOT FOR DISTRIBUTION"),
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

# ---- 1. The identity
doc.add_heading("1. Both Options Are the Same Price per Share", level=1)
bold_para(doc, [
    ("Post-money is fixed by the fund's required multiple, not by the cheque size. ", True),
    (f"At a HK$134.23M Series A and a {REQUIRED_MULT:.0f}× requirement, post-money "
     f"must be {money(POST_TARGET)}. One cheque of HK$0.5M and two cheques of "
     f"HK$0.5M both land on the same post-money, because pre-money adjusts to absorb "
     f"the difference in cash raised.", False),
])
rows = [
    ["A — One seed", money(6.21), money(0.5), money(6.71),
     f"HK${PPS:,.0f}", "7.45%"],
    ["B — Two seeds", money(5.71), money(1.0), money(6.71),
     f"HK${PPS:,.0f}", "14.90%"],
]
table(doc, ["Structure", "Pre-money", "Cash raised", "Post-money",
            "Price per share", "Combined stake"],
      rows, [4.0, 2.8, 2.8, 2.8, 3.2, 3.0], highlight_col=4)
note(doc, "The HK$6.21M vs HK$5.71M pre-money difference is an artifact of cheque "
          "size. The founder is not getting a better price in Option A — they are "
          "raising less at the same price. Post-money and price per share are "
          "identical.")

# ---- 2. Ownership and governance
doc.add_heading("2. Ownership and Governance Outcomes", level=1)
rows = []
for s in structures:
    if len(s["stakes"]) == 1:
        stakes_str = f"1 × {s['stakes'][0]*100:.2f}%"
    else:
        stakes_str = " + ".join(f"{x*100:.2f}%" for x in s["stakes"])
    rows.append([
        s["label"],
        money(s["total"]),
        stakes_str,
        f"{s['combined']*100:.2f}%",
        f"{s['founder']*100:.2f}%",
        f"{max(s['stakes'])*100:.2f}%",
    ])
table(doc, ["Structure", "Raised", "Per-investor stake",
            "Combined dilution", "Founder retained", "Largest holder"],
      rows, [4.5, 2.2, 3.2, 3.0, 3.0, 3.0], highlight_col=5)

bold_para(doc, [
    ("The governance difference is the point. ", True),
    ("Two seeds of 7.45% each keep both investors below the 10% threshold at which a "
     "board seat is customary. A single HK$1.0M investor at HK$5.71M pre holds "
     "14.90% and will reasonably expect a board observer or seat. That is the "
     "'play safe' logic — it is a governance argument, not a valuation argument.",
     False),
])

doc.add_heading("Customary governance thresholds", level=2)
table(doc, ["Threshold", "Stake range"],
      [[a, b] for a, b in GOVERNANCE_THRESHOLDS],
      [8.0, 6.0])
note(doc, "Thresholds vary by investor and jurisdiction, but 10% is the common line "
          "for a board seat and 20% for blocking rights on reserved matters. Two "
          "7.45% holders sit comfortably below both.")

# ---- 3. Fund-side constraints
doc.add_heading("3. Fund-Side Constraints — Why Two Seeds May Be Harder to Close", level=1)
bullet(doc, [("Minimum position size. ", True),
             ("Many seed funds have a minimum cheque of HK$0.75–1.5M. A HK$0.5M "
              "position may be below the minimum at which the fund can justify "
              "diligence and reserve. Ask before proposing a split.", False)])
bullet(doc, [("Two independent decisions. ", True),
             ("Two seeds means two diligence processes, two IC approvals and two "
              "timelines. The risk is that one closes and the other does not, "
              "leaving the company at HK$0.5M raised against a cap table built for "
              "HK$1.0M — and a pre-money that the remaining investor may no longer "
              "accept.", False)])
bullet(doc, [("Doubled legal and administrative cost. ", True),
             ("Each investor needs SHA, SPA and side letters. Expect HK$50–100K of "
              "additional legal cost for the second investor, against HK$0.5M of "
              "incremental capital — 10–20% of the extra cheque consumed by "
              "friction.", False)])
bullet(doc, [("Cap table clutter. ", True),
             ("Some Series A investors dislike many small holders. Two seeds at "
              "7.45% is manageable; four or five at 3–4% starts to look "
              "fragmented.", False)])

# ---- 4. Three practical structures
doc.add_heading("4. Three Practical Structures for Two Seeds", level=1)

doc.add_heading("4.1 Equal seeds, simultaneous close", level=2)
doc.add_paragraph(
    "Two investors of HK$0.5M each, both signing on the same date at the same price, "
    "cross-conditioned so neither closes without the other. Each holds 7.45%. "
    "Cleanest outcome if both are genuinely committed."
)
bullet(doc, [("Advantage: ", True), ("Neither investor crosses 10%; cap table simple.", False)])
bullet(doc, [("Risk: ", True), ("One drops out and the other renegotiates the pre-money.", False)])

doc.add_heading("4.2 Lead and follower", level=2)
doc.add_paragraph(
    "One investor takes HK$0.75M (10.0%) and the other HK$0.25M (3.6%). The lead sets "
    "terms and conducts diligence; the follower joins on the same documents. Faster "
    "to close than two equals, but the lead crosses 10% and will expect an observer "
    "seat."
)

doc.add_heading("4.3 Two tranches (the dilution-saving structure)", level=2)
doc.add_paragraph(
    "HK$0.5M now at HK$6.21M pre (7.45%), and HK$0.5M after the pilot at a higher "
    "pre-money — modelled here at HK$12.0M. The second tranche prices against "
    "evidence rather than against the plan."
)
rows = []
for s in structures:
    if s["label"].startswith("D"):
        rows.append([
            s["label"], money(s["total"]),
            f"{s['stakes'][0]*100:.2f}% + {s['stakes'][1]*100:.2f}%",
            f"{s['combined']*100:.2f}%",
            f"{s['founder']*100:.2f}%",
            money(s["pre_blended"]),
        ])
table(doc, ["Structure", "Raised", "Per-tranche stake",
            "Combined dilution", "Founder retained", "Blended pre-money"],
      rows, [4.0, 2.2, 4.0, 3.0, 3.0, 3.2], highlight_col=3)
note(doc, "The tranche structure raises the same HK$1.0M as two simultaneous seeds "
          "but dilutes the founder by roughly 11% instead of 15% — because the second "
          "cheque prices against post-pilot evidence, not against the plan. The cost "
          "is timing: the company runs at HK$0.5M until the pilot lands.")

# ---- 5. Decision rule
doc.add_heading("5. Decision Rule", level=1)
rows = [
    ["Company needs only HK$0.5M to reach milestones",
     "One seed (Option A)", "Least dilution for sufficient capital."],
    ["Company needs HK$1.0M and one investor will take it",
     "Two seeds if a second can be found (Option B)",
     "Governance advantage: no 10%+ holder."],
    ["Company needs HK$1.0M but a second seed is uncertain",
     "Lead + follower (Option C)",
     "Faster close; accept the 10% lead with a clean observer arrangement."],
    ["Pilot data is 6–9 months away and runway allows the wait",
     "Two tranches (Option D)",
     "Saves 3–4 points of dilution; prices the second cheque on evidence."],
    ["Downside global scenario (v14 S1/S3)",
     "One seed or non-dilutive first",
     "Do not set a HK$5.71M anchor in a depressed market."],
]
table(doc, ["Situation", "Recommended structure", "Rationale"],
      rows, [5.5, 4.5, 6.0], highlight_col=1)

# ---- 6. Recommendation
doc.add_heading("6. Recommendation", level=1)
bold_para(doc, [
    ("If the company needs HK$1.0M, pursue two seeds — but structure them as two "
     "tranches, not two simultaneous cheques. ", True),
    (f"Two simultaneous seeds give you the governance benefit (neither above 10%) "
     f"but cost 14.90% dilution. Two tranches give you the same HK$1.0M at roughly "
     f"11% dilution, because the second cheque prices after the pilot. The tranche "
     f"is strictly better if the runway allows the gap.", False),
])
bold_para(doc, [
    ("If a second seed cannot be found, take one cheque — but reduce it. ", True),
    (f"HK$0.5M at HK$6.21M pre (7.45%) is better for the founder than HK$1.0M at "
     f"HK$5.71M pre (14.90%), on every axis except capital deployed. The pre-money "
     f"is higher, the dilution is half, and the investor still gets their 20× at the "
     f"midpoint. The question is whether HK$0.5M is enough runway — if it is, this "
     f"is the cleanest outcome in the document.", False),
])
bold_para(doc, [
    ("Ask the fund two questions before proposing any structure. ", True),
    ("(i) What is your minimum position size? (ii) Would you accept a HK$0.5M "
     "tranche now with a pre-agreed price band for a second tranche post-pilot? The "
     "answers determine whether Options B, C or D are available at all — and they "
     "take five minutes to ask, versus weeks to discover during diligence.", False),
])

# ===========================================================================
# SECTION 7 — Installment structures
# ===========================================================================

doc.add_heading("7. Installment Structures — Whose Option?", level=1)
doc.add_paragraph(
    "Installments are standard practice. The structure determines whether they help "
    "or hurt the founder, and the determinant is whose option the second payment is."
)

rows = [
    ["Deferred funding (one price)",
     "HK$1.0M committed at signing; HK$0.5M drawn at close, HK$0.5M at a defined "
     "date. Full 14.90% issued at close.",
     "Neither — binding on both",
     "14.90% / HK$1.0M",
     "Acceptable if the commitment is legally binding with default remedies"],
    ["Separate closings (two prices)",
     "HK$0.5M at HK$6.21M pre now; HK$0.5M at HK$12.0M pre post-pilot",
     "Founder can reprice upward; investor can decline second",
     "~11.15% / HK$1.0M",
     "Best structure — second cheque prices against evidence"],
    ["Milestone-linked, investor option",
     "HK$1.0M committed at one price; tranche 2 at investor's discretion",
     "Investor only",
     "14.90% if fully drawn / HK$0.5M if declined",
     "Refuse — dilution committed, capital optional"],
]
table(doc, ["Structure", "Mechanics", "Whose option",
            "Founder dilution / cash", "Assessment"],
      rows, [3.8, 5.5, 3.5, 3.5, 4.5], highlight_col=3, shade_rows=[2])

bold_para(doc, [
    ("The test. ", True),
    ("\"If tranche 2 does not fund, what is my cap table and what did I receive?\" "
     "If the answer is 14.90% dilution for HK$0.5M of cash, the structure transfers "
     "value to the investor and should be refused. If the answer is 7.45% for "
     "HK$0.5M, nothing is lost and the structure is acceptable.", False),
])

doc.add_heading("7.1 The favourable variants", level=2)
bold_para(doc, [
    ("Structure 1 — deferred funding with a binding commitment. ", True),
    ("One subscription for HK$1.0M at HK$5.71M pre; cash drawn in two installments "
     "at the same price. Full 14.90% issued at close. Acceptable only if the second "
     "drawdown is contractually certain with default remedies — otherwise it is the "
     "investor-option structure in disguise.", False),
])
bold_para(doc, [
    ("Structure 2 — separate closings at different prices. ", True),
    (f"HK$0.5M at HK$6.21M pre (7.45%), HK$0.5M at HK${TRANCHE_2_PRE:.1f}M pre "
     f"post-pilot (4.0%). Blended ~11.15% for HK$1.0M. This is Option D and the "
     f"strongest structure in the document, because the second cheque prices "
     f"against evidence rather than against the plan.", False),
])
bold_para(doc, [
    ("Structure 3 — milestone-linked with a price ratchet. ", True),
    ("HK$0.5M now at HK$6.21M pre; HK$0.5M later at a price determined by pilot "
     "outcome, with a floor and ceiling agreed now. Fallback if the investor insists "
     "on downside protection but will not accept separate closings.", False),
])

doc.add_heading("7.2 The line to use", level=2)
p = doc.add_paragraph()
p.paragraph_format.left_indent = Cm(0.8)
p.paragraph_format.space_after = Pt(6)
r = p.add_run(
    "\"Installments work for me, but the pricing has to match the risk. If the "
    "second cheque is priced today, before the pilot, then it is the same risk as "
    "the first — and it should be a binding commitment, not an option. If it is at "
    "your option after the pilot, then it should be priced after the pilot, at "
    "whatever the evidence supports. I can accept either. I cannot accept today's "
    "price with tomorrow's option.\""
)
r.italic = True; r.font.size = Pt(10)

doc.add_heading("7.3 HK-specific note", level=2)
doc.add_paragraph(
    "Installments linked to performance targets are common in HK and mainland "
    "practice. Where the second tranche is conditional on financial targets, the "
    "arrangement may be characterised as a valuation adjustment mechanism. VAM "
    "structures involving listed-company targets have been restricted in mainland "
    "China since 2020, and the tax and accounting treatment differs from a simple "
    "deferred drawdown. Take advice before signing a milestone-linked tranche that "
    "has a financial-target condition — the structure may be functionally a VAM even "
    "if the documents do not use that term."
)

out = "Seed_Structure_Comparison_INTERNAL.docx"
doc.save(out)

print(f"Saved: {out}\n")
print(f"Post-money target: {money(POST_TARGET)}  |  Price per share: HK${PPS:,.0f}\n")
print(f"{'Structure':<34}{'Raised':>9}{'Combined':>10}{'Founder':>9}{'Largest':>9}")
print("-" * 72)
for s in structures:
    print(f"{s['label']:<34}{money(s['total']):>9}"
          f"{s['combined']*100:>9.2f}%{s['founder']*100:>8.2f}%"
          f"{max(s['stakes'])*100:>8.2f}%")