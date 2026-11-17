# --- REPLACE the added sentence under the three-sentences block ---
# NEW:

p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(3)
p.paragraph_format.space_after = Pt(2)
r = p.add_run(
    "Tranche 1 is allocated to the 10,000-kick dataset and the benchmark it produces. "
    "The 5 LOIs follow from showing each academy its own players' data. That is the "
    "milestone that prices Tranche 2."
)
r.italic = True
r.font.size = Pt(8.5)
r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)