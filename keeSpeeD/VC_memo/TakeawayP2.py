# --- REPLACE the "Applied to the Current Proposal" section on Page 4 ---

section_heading(doc, "The Causal Chain — What Tranche 1 Actually Buys")

data_table(
    doc,
    ["Step", "What happens", "What it costs", "What it produces"],
    [
        ["1", "5 academies' coaches capture and label kick events",
         "HK$63K (capture + labelling)",
         "10,000 labelled kick events"],
        ["2", "Independent coach assesses the AI kernel against ground truth",
         "HK$32K (benchmark assessment)",
         "Coach-agreement benchmark (≥80%)"],
        ["3", "Benchmark shown to each academy using that academy's own players' data",
         "HK$30K (coordination and travel)",
         "5 signed LOIs — confirmation, not a sale"],
        ["4", "Audit of Y1–Y3 statements",
         "HK$80–150K",
         "Audited track record — the SA price argument"],
        ["5", "Legal documentation for Tranche 2",
         "HK$40–60K",
         "Tranche 2 at HK$12.0M pre-money"],
    ],
    widths=[1.0, 6.5, 4.0, 6.0],
    font_size=8,
    highlight_col=3
)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(3)
r = p.add_run(
    "Total deployed across tranche 1: HK$349–459K against the HK$0.5M cheque. "
    "Headroom of HK$40–150K. The milestone is sized to the cheque, not the cheque "
    "to the milestone."
)
r.italic = True
r.font.size = Pt(8)
r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

section_heading(doc, "What the 5 LOIs Are — and Are Not")

data_table(
    doc,
    ["The LOIs are NOT", "The LOIs ARE"],
    [
        ["A sales motion — no cold outreach, no pipeline",
         "The academy's confirmation that the benchmark on its own players is valuable"],
        ["Uncertain — dependent on finding 5 willing buyers",
         "Low-risk — the academies already participated in the capture"],
        ["A separate spend from the 10k kicks",
         "The direct consequence of the 10k kicks — one spend, two outputs"],
        ["A guarantee — an academy can decline to sign",
         "Contingent on one thing only: the benchmark is ≥80%"],
    ],
    widths=[8.5, 8.5],
    font_size=8
)

callout_box(
    doc,
    "The uncomfortable lesson, applied:  HK$0.5M buys 10,000 labelled kicks, an "
    "independent benchmark, and five signed LOIs. It does not buy 'general progress.' "
    "If the benchmark falls below 80%, the LOIs do not follow, and Tranche 2 does not "
    "price at HK$12.0M. That is the risk the tranche exists to resolve — and it is a "
    "single, measurable risk, not a portfolio of hopes.",
    fill_color="E2EFDA",
    text_color=GREEN,
    border_color=GREEN,
    size=9,
    bold=False,
    italic=True
)