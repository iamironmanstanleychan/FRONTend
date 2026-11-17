# ===========================================================================
# SECTION 7 — Installment structures (append before save)
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
     f"14.90% / HK$1.0M",
     "Acceptable if the commitment is legally binding with default remedies"],
    ["Separate closings (two prices)",
     "HK$0.5M at HK$6.21M pre now; HK$0.5M at HK$12.0M pre post-pilot",
     "Founder can reprice upward; investor can decline second",
     f"~11.15% / HK$1.0M",
     "Best structure — second cheque prices against evidence"],
    ["Milestone-linked, investor option",
     "HK$1.0M committed at one price; tranche 2 at investor's discretion",
     "Investor only",
     f"14.90% if fully drawn / HK$0.5M if declined",
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