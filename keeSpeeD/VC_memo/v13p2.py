# ===========================================================================
# COMBINED PATCH — Patches 2, 3, 4
# Apply AFTER Patch 1 (classification block, page header/footer, cover banner).
# Each sub-patch gives the exact anchor line(s) to find and the block to paste.
# ===========================================================================

# ---------------------------------------------------------------------------
# PATCH 2 — Meta table: add Classification, Next version, red status lines
# ---------------------------------------------------------------------------
# FIND:
#     for label, value in [
#         ("To:", "Founders / Board"),
#         ("From:", "Finance & Strategy"),
#         ("Re:", "The Valuation Route — From P&L to the Midpoint Matrix"),
#         ("Status:", "INTERNAL — negotiation preparation."),
#         ("Version:", "v13 — seven-step derivation, A through G."),
#     ]:
#
# REPLACE WITH:

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


# ---------------------------------------------------------------------------
# PATCH 3a — Section G opening divider (page break + STOP banner + heading)
# ---------------------------------------------------------------------------
# FIND:
#     doc.add_heading("G. Positioning and Counterpart", level=1)
#
# REPLACE WITH:

# ---- Section G divider — DO NOT DISTRIBUTE ----
doc.add_page_break()

stop_bar = doc.add_paragraph()
stop_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
stop_bar.paragraph_format.space_before = Pt(0)
r = stop_bar.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_title = doc.add_paragraph()
stop_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_title.add_run("STOP  ·  SECTION G BEGINS HERE")
r.bold = True
r.font.size = Pt(14)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_warn = doc.add_paragraph()
stop_warn.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = stop_warn.add_run(
    "DO NOT DISTRIBUTE THIS SECTION.\n"
    "Contains founder positions, counterpart responses, and negotiation scripts.\n"
    "Strip this section in full before any investor communication."
)
r.italic = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

stop_bar2 = doc.add_paragraph()
stop_bar2.alignment = WD_ALIGN_PARAGRAPH.CENTER
stop_bar2.paragraph_format.space_after = Pt(14)
r = stop_bar2.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_heading("G. Positioning and Counterpart", level=1)


# ---------------------------------------------------------------------------
# PATCH 3b — Section G closing marker + page break before Caveats
# ---------------------------------------------------------------------------
# FIND:
#     doc.add_heading("7. Caveats", level=1)
#
# REPLACE WITH:

# ---- End of Section G ----
end_bar = doc.add_paragraph()
end_bar.alignment = WD_ALIGN_PARAGRAPH.CENTER
end_bar.paragraph_format.space_before = Pt(18)
r = end_bar.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

end_title = doc.add_paragraph()
end_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = end_title.add_run("END OF SECTION G  ·  INTERNAL CONTENT")
r.bold = True
r.font.size = Pt(11)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

end_bar2 = doc.add_paragraph()
end_bar2.alignment = WD_ALIGN_PARAGRAPH.CENTER
end_bar2.paragraph_format.space_after = Pt(18)
r = end_bar2.add_run("═" * 58)
r.bold = True
r.font.size = Pt(10)
r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

doc.add_page_break()
doc.add_heading("7. Caveats", level=1)


# ---------------------------------------------------------------------------
# PATCH 4 — Output filename gets _INTERNAL suffix
# ---------------------------------------------------------------------------
# FIND:
#     out = "Valuation_Route_v13.docx"
#
# REPLACE WITH:

out = "Valuation_Route_v13_INTERNAL.docx"