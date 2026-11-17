# ===========================================================================
# PAGE 4 — LESSONS FROM PONGBOT & STEPOUT  (v2: positioning, chain, verification)
# ===========================================================================

page_break(doc)

page_title(
    doc,
    "PAGE 4  ·  LESSONS FROM PONGBOT & STEPOUT",
    "Both companies funded their step-up with the previous round. Neither "
    "argued for it at the next one."
)

callout_box(
    doc,
    "A valuation step-up is not negotiated — it is purchased with the first cheque.",
    fill_color="E2EFDA",
    text_color=GREEN,
    border_color=GREEN,
    size=11,
    bold=True
)

p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(2)
r = p.add_run(
    "PongBot's later round and StepOut's Pre-A were priced higher because a specific, "
    "externally verifiable event had occurred between rounds. Each company bought the "
    "evidence with the previous round's capital. Applied to us: Tranche 1 is not "
    "runway — it is the budget for producing the milestone that prices Tranche 2."
)
r.italic = True
r.font.size = Pt(8.5)
r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

section_heading(doc, "The Milestone-Gated Structure")

chart5 = chart_timeline()
doc.add_picture(chart5, width=Cm(17.5))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

# --- Item 2: the causal chain replaces the two-parallel-milestones framing ---

section_heading(doc, "The Causal Chain — What Tranche 1 Buys")

data_table(
    doc,
    ["Step", "What happens", "Cost (HK$)", "Output"],
    [
        ["1",
         "Five academies' coaches capture and label kick events",
         "63K (capture + labelling)",
         "10,000 labelled kick events"],
        ["2",
         "Independent coach scores the AI kernel against ground truth",
         "32K (benchmark assessment)",
         "Coach-agreement benchmark (≥80%)"],
        ["3",
         "Each academy reviews the benchmark on its own players' data",
         "30K (coordination, travel, review sessions)",
         "5 signed LOIs"],
        ["4",
         "Y1–Y3 statements audited",
         "80–150K",
         "Audited track record — the SA price argument"],
        ["5",
         "Tranche 2 documentation",
         "40–60K",
         "Tranche 2 at HK$12.0M pre-money"],
    ],
    widths=[1.0, 6.6, 4.2, 5.7],
    font_size=8,
    highlight_col=3
)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(3)
r = p.add_run(
    "Total deployed: HK$349–459K against the HK$0.5M cheque, leaving HK$40–150K of "
    "headroom. The milestone is sized to the cheque, not the cheque to the milestone."
)
r.italic = True
r.font.size = Pt(8)
r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

section_heading(doc, "Why the 5 LOIs Are a Consequence, Not a Sale")

data_table(
    doc,
    ["Not", "Yes"],
    [
        ["A sales motion — no cold outreach, no pipeline",
         "The academy reviews the benchmark on its own players' data"],
        ["Dependent on finding five willing buyers",
         "Low execution risk — the five academies already participated in the capture"],
        ["A separate spend from the 10k kicks",
         "The direct output of the 10k kicks — one spend, two deliverables"],
        ["A guarantee",
         "Contingent on one variable only: the benchmark reaching 80%"],
    ],
    widths=[8.5, 8.5],
    font_size=8
)

# --- Item 1: corrected PongBot positioning ---

section_heading(doc, "PongBot — What Transfers")

data_table(
    doc,
    ["#", "What happened", "What to copy"],
    [
        ["1", "Two crowdfunding campaigns (PACE, then Aura), both in the millions",
         "One validation event is an anecdote; two is a pattern"],
        ["2", "Shipping began roughly two months after campaign close",
         "Delivery converts a commitment into a record"],
        ["3", "Each raise funded the next campaign",
         "Round sized to the next de-risking event, not to abstract runway"],
        ["4", "Product video drove customer acquisition",
         "In sports hardware, the demonstration is the marketing"],
    ],
    widths=[0.8, 8.0, 8.8],
    font_size=8
)

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(2)
p.paragraph_format.space_after = Pt(2)
r = p.add_run(
    "What does not transfer:  PongBot's primary market is consumer hardware, with a "
    "growing B2B training segment. Kung Fu AI Football is a pure B2B technology "
    "enabler. Crowdfunding is not a natural analogue — but a named academy pilot "
    "serves the same validation function."
)
r.italic = True
r.font.size = Pt(7.5)
r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

section_heading(doc, "StepOut — What Transfers")

data_table(
    doc,
    ["#", "What happened", "What to copy"],
    [
        ["1", "Seed round, then Pre-A priced on operating metrics (growth, renewal rate)",
         "Software can step up on metrics alone — no hardware required"],
        ["2", "Payty clubs and federations across multiple countries",
         "Institutional logos are the strongest reference asset"],
        ["3", "Volume metrics on matches analysed and players tracked",
         "Volume metrics are the dataset story — direct analogue to the 10k kicks"],
        ["4", "Modest absolute round sizes with meaningful percentage step-ups",
         "A step-up is a ratio, not an absolute figure"],
    ],
    widths=[0.8, 8.0, 8.8],
    font_size=8
)

section_heading(doc, "Applied to the Current Proposal")

data_table(
    doc,
    ["Item", "Current", "Position"],
    [
        ["Tranche 1",
         "HK$0.5M at HK$6.21M pre (7.45%)",
         "Unchanged in size and price — allocated to the five-step chain above"],
        ["Tranche 2 trigger",
         "Post-pilot",
         "Ten thousand labelled kicks + ≥80% benchmark + five LOIs, with a long-stop date"],
        ["Tranche 2 price",
         "HK$12.0M pre (4.00%)",
         "Binding clause, agreed before Tranche 1 clears"],
        ["Fallback if milestone missed",
         "Undefined",
         "Investor may decline Tranche 2; founder retains 92.55%; no shares issued "
         "against capital not received"],
    ],
    widths=[3.0, 5.8, 8.7],
    font_size=8,
    highlight_col=2
)

callout_box(
    doc,
    "Both cases show the same pattern: the step-up was funded by the previous round, "
    "not argued at the next one. The equivalent for us is the 10,000-kick dataset, "
    "the coach benchmark, and the five LOIs that follow. That is what Tranche 1 is for.",
    fill_color="E2EFDA",
    text_color=GREEN,
    border_color=GREEN,
    size=9,
    bold=False,
    italic=True
)

# --- Item 3: verification footnote ---

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(4)
p.paragraph_format.space_after = Pt(0)
r = p.add_run(
    "Figure status:  PongBot figures to be confirmed against pongbotsports.com "
    "(campaign totals, shipping date, round size). StepOut figures to be confirmed "
    "against stepout.ai (round sizes, growth rate, renewal rate, club count). "
    "Claims that cannot be confirmed against the primary source are to be dropped "
    "rather than footnoted — the lessons do not depend on the precise numbers."
)
r.italic = True
r.font.size = Pt(7)
r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)