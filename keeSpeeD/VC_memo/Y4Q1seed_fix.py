# --- add after Section 4 (Y4 Q1 Valuation), before Section 5 ---

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