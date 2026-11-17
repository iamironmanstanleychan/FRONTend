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
          "the Lower case — higher than any rate in the matrix. The matrix was "
          "never designed to contain this cell, because it is not a valuation. It "
          "is fund arithmetic: a HK$1.0M cheque in a HK$100M fund needs 10% to be "
          "worth managing.")

doc.add_heading("G.6 What this means for the negotiation", level=2)
bullet(doc, [("Do not argue the DCF against the 10% ask. ", True),
             ("The ask is not a DCF output. Argue fund structure instead: smaller "
              "cheque, tranched cheques, or a follow-on price band agreed now.",
              False)])
bullet(doc, [("Accept the 10% if the terms are clean. ", True),
             (f"A HK$9.0M pre-money with non-participating preference and no board "
              f"seat is a manageable outcome. It sets a reference price, but a "
              f"clean one.", False)])
bullet(doc, [("Reduce the cheque if HK$1.0M is more than the plan needs. ", True),
             ("HK$500K at HK$9.0M pre-money is 5.3% dilution — half the equity for "
              "half the runway. This is the single strongest lever available.",
              False)])
bullet(doc, [("Reserve the matrix for the next round. ", True),
             ("The midpoint cells are what the company should be worth at Y5 or Y6, "
              "after the pilot and the audit. Use them to plan the step-up, not to "
              "fight the current round.", False)])