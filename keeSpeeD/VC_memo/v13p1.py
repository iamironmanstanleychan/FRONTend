# ===========================================================================
# CLASSIFICATION — header, footer, cover block
# ===========================================================================

# Section-level header/footer (suppressed on cover page, shown on pages 2+)
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

# Cover-page classification block (above the title)
cover_box = doc.add_paragraph()
cover_box.alignment = WD_ALIGN_PARAGRAPH.CENTER
cover_box.paragraph_format.space_before = Pt(0)
r = cover_box.add_run("⚠   INTERNAL — CONFIDENTIAL   ⚠")
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