def chart_timeline():
    """Page 4: milestone-gated structure."""
    fig, ax = plt.subplots(figsize=(6.7, 2.2), dpi=140)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 3)
    ax.axis("off")

    nodes = [
        (1.0, 1.5, "TRANCHE 1", "HK$0.5M\n7.45%", NAVY,
         "PongBot: two campaigns\nStepOut: SaaS launch"),
        (3.5, 1.5, "MILESTONE", "10K kicks → 5 LOIs\n80% coach agreement", GOLD,
         "PongBot: shipping\nStepOut: operating metrics"),
        (6.0, 1.5, "TRANCHE 2", "HK$0.5M\n4.00%", GREEN,
         "PongBot: second campaign\nStepOut: Pre-A step-up"),
        (8.5, 1.5, "SERIES A", "HK$134M+\n10–20× return", RED,
         "Priced by evidence"),
    ]
    # ... rest of function unchanged ...