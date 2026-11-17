# --- replace the RAISE_HKD_M constant block ---

RAISE_Y0_HKD_M = 1.5        # realistic Y0 instrument: angel / SAFE
RAISE_Y1_HKD_M = 8.0        # priced seed target, post-pilot
Y0_ANGEL_PRE = 30.0         # Y0 angel pre-money
Y1_SEED_PRE = 37.34         # Y1 priced seed target (adjusted Centre x 50%)

def own_v2(pre, raise_m):
    return raise_m / (pre + raise_m)

# --- replace Section 11.2 (round size) ---

doc.add_heading("11.2 Round size — two instruments, two dates", level=2)
bold_para(doc, [
    ("Y0. ", True),
    (f"Raise {money(RAISE_Y0_HKD_M)} at {money(Y0_ANGEL_PRE)} pre-money "
     f"(≈{own_v2(Y0_ANGEL_PRE, RAISE_Y0_HKD_M)*100:.1f}% dilution). This funds the "
     f"5–10 academy pilot and the paired dataset. It is deliberately small: an "
     f"HK$8M cheque into a pre-pilot company is not a realistic Y0 event, and "
     f"pricing one now would set a low reference that the pilot cannot repair.", False),
])
bold_para(doc, [
    ("Y1. ", True),
    (f"Raise {money(RAISE_Y1_HKD_M)} at the Target of {money(Y1_SEED_PRE)} pre-money "
     f"once pilot data exists (≈{own_v2(Y1_SEED_PRE, RAISE_Y1_HKD_M)*100:.1f}% "
     f"dilution). This is the round the valuation matrix in Sections 4–7 is written "
     f"for. It should not be attempted before the coach-agreement benchmarks land.", False),
])

# --- insert as new Section 12.0, before the anchor defence ---

doc.add_heading("12. The Y0 Risk Is the Pass, Not the Anchor", level=1)
doc.add_paragraph(
    "The negotiation framework in this section addresses the least likely Y0 outcome. "
    "The dominant risk at Y0 is that no priced round is offered at all, because an "
    "investor with HK$8M has alternatives that already carry proof. Planning should "
    "reflect the probability distribution, not the most colourful scenario."
)
rows = [
    ["Pass, no term sheet", "Most likely",
     "No negotiation. Pilot-first path becomes the plan."],
    ["Smaller instrument (HK$1–2M SAFE)", "Likely",
     "Sections 11.2 (Y0) applies. Cap table stays clean."],
    ["Priced seed at HK$20–30M", "Possible",
     "Between Reservation and Target. Acceptable."],
    ["Anchor at Lower × 75% (HK$7.80M)", "Least likely",
     "Section 12.4 applies. Do not concede."],
]
table(doc, ["Y0 outcome", "Likelihood", "Consequence"], rows,
      [5.0, 3.0, 9.5])

doc.add_heading("12.0 Readiness diagnostic — before seeking a priced seed", level=2)
doc.add_paragraph(
    "A priced seed at Y0 is only attemptable if all four are true. If any is false, "
    "run the angel instrument instead and revisit after the pilot."
)
num_bullet(doc, [("A named academy has agreed to pilot terms in writing. ", True),
                 ("An LOI with a timetable, not a verbal expression of interest.", False)])
num_bullet(doc, [("The pilot budget is fully specified in months of runway. ", True),
                 ("HK$1.5M buys how many months? If the answer is under 9, the raise "
                  "is too small for the milestones it must buy.", False)])
num_bullet(doc, [("At least one brand partner has engaged substantively. ", True),
                 ("A meeting is not engagement. A pilot scope discussion is.", False)])
num_bullet(doc, [("The Lower case has been rebuilt with a distinct cost structure. ", True),
                 ("If Lower is still Centre ÷ 2.57, an analyst will say so and the "
                  "priced round will not close on this schedule.", False)])