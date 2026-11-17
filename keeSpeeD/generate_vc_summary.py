"""
KungFu AI Football — VC Executive Summary Generator
Creates a concise, narrative-driven Word document.
"""

from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT

FONT = 'Times New Roman'
ACCENT_COLOR = RGBColor(0x1F, 0x4E, 0x78)

def set_run(run, size, bold=False, italic=False, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', FONT)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    size = 16 if level == 1 else 13
    set_run(p.add_run(text), size, bold=True, color=ACCENT_COLOR)
    return p

def add_body(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        set_run(p.add_run(bold_prefix), 11, bold=True)
    set_run(p.add_run(text), 11)
    return p

def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_after = Pt(2)
    if bold_prefix:
        set_run(p.add_run(bold_prefix), 11, bold=True)
    set_run(p.add_run(text), 11)
    return p

def build_doc():
    doc = Document()
    
    # A4 Page Setup
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)
    section.left_margin = Mm(22)
    section.right_margin = Mm(22)
    
    # Base Style
    style = doc.styles['Normal']
    style.font.name = FONT
    style.font.size = Pt(11)

    # ─── TITLE PAGE ───
    for _ in range(6):
        doc.add_paragraph()
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("Kung Fu AI Football"), 24, bold=True, color=ACCENT_COLOR)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("Executive Summary & Investment Thesis"), 14, bold=True)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("A Deep-Tech Sports AI Platform enabling the next generation of football training"), 11, italic=True)
    
    for _ in range(8):
        doc.add_paragraph()
        
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run("Prepared for Venture Capital Review"), 10)
    
    doc.add_page_break()

    # ─── SECTION 1: MARKET POTENTIAL ───
    add_heading(doc, "1. Market Potential & Tailwinds", 1)
    
    add_body(doc, "The convergence of massive market capital and a post-pandemic shift in consumer priorities creates a unique window for sports technology.")
    
    add_bullet(doc, " The China football training market alone is valued at RMB 48.76 billion (2025), projected to reach RMB 82.84 billion by 2030. Football remains the single largest sports market globally.", "Capital:")
    add_bullet(doc, " High-net-worth consumers are shifting spend from material luxury (watches, jewelry) to health, wellness, and sports experiences. The Hurun 2026 report notes health spending is now the #2 consumption category, with service experiences rising 12%.", "Willingness to Shift:")
    add_bullet(doc, " We have a clear funnel from TAM to SOM:", "Market Sizing:")
    add_bullet(doc, " 380 million football fans in China.", "Total Addressable Market (TAM):")
    add_bullet(doc, " 13.89 million active youth and amateur participants (campus + elite + amateur).", "Serviceable Available Market (SAM):")
    add_bullet(doc, " By partnering with top-tier brands (Anta, Li Ning, Amer Sports) and targeting the 8.7% of parents willing to support a long-term football path, we convert this SAM into a highly engaged paying base (SOM). Even at our most conservative penetration (P10), we capture a highly profitable slice of this market.", "Serviceable Obtainable Market (SOM):")
    
    # ─── SECTION 2: COMPETITIVENESS ───
    add_heading(doc, "2. Competitiveness & Uniqueness", 1)
    
    add_body(doc, "Kung Fu AI Football is not another basic hardware tracker. We provide a comprehensive digital coaching platform powered by proprietary datasets.")
    
    add_bullet(doc, " Competitors like INSAIT Joy (US$139) offer basic hardware with no subscription or feedback loop. We offer a hybrid system: a Smart Ball with IoT sensors (HK$980) plus a recurring subscription (HK$1,592/year) that delivers actionable physics and biomechanics feedback.", "Price Analysis:")
    add_bullet(doc, " Our AI Kernel fuses raw data from the Smart Football (ball kinetics, spin, speed) with Computer Vision (player biomechanics, foot angle, hip rotation, kick style).", "Proprietary Data:")
    add_bullet(doc, " We don't just tell players 'how hard' they kicked. We tell them 'why' it missed and 'how' to fix it. This is a data-driven digital coaching moat that hardware-only competitors cannot replicate.", "The Moos (Moat):")
    
    # ─── SECTION 3: LEVERAGING HKSAR INCENTIVES ───
    add_heading(doc, "3. Leveraging HKSAR Incentives", 1)
    
    add_body(doc, "We heavily leverage Hong Kong’s R&D ecosystem to subsidize our development and improve our capital efficiency.")
    
    add_bullet(doc, " Hong Kong offers an enhanced tax deduction for R&D expenditures. By paying local universities (DLRIs) for our deep-tech R&D (HK$120k/year), we qualify for a 300% tax deduction on the first HK$2 million. This significantly reduces our annual tax burden.", "DIPN 55 (Enhanced Tax Reduction):")
    add_bullet(doc, " We capitalize our accumulated R&D expenses (HK$360k every 3 years) into an Intangible Asset on our balance sheet. This not only strengthens our balance sheet but reflects the true economic value of the IP we are building, while the tax deduction still applies to the cash spent.", "HKAS 38 (Intangible Assets):")
    
    # ─── SECTION 4: PROFITABILITY ───
    add_heading(doc, "4. Potential 6-Year Cumulative Net Profits", 1)
    
    add_body(doc, "Our financial model is built to be highly profitable even in the most adverse scenarios. By year 6, we project the following B2B cumulative net profits (after all operating costs, HKAS 38 amortization, and HKSAR taxes):")
    
    add_bullet(doc, " 40% Marketing Cost (Highest) + P10 Adoption (Lowest). Net Profit: HK$89.97 million.", "Conservative Case:")
    add_bullet(doc, " 35% Marketing Cost + P50 Adoption. Net Profit: HK$230.54 million.", "Middle Case:")
    add_bullet(doc, " 30% Marketing Cost (Lowest) + P90 Adoption (Highest). Net Profit: HK$560.96 million.", "Ceiling Case:")
    
    add_body(doc, "The business is designed to be cash-generative and highly scalable, turning from a hardware-plus-subscription model into a high-margin data and IP licensing business.")
    
    # ─── CONCLUSION ───
    add_heading(doc, "Summary for VC", 1)
    add_body(doc, "Kung Fu AI Football is not just a consumer gadget. It is a comprehensive, data-driven sports AI platform with a defensible technical moat, strong market tailwinds, and highly favourable government R&D incentives. We are seeking [Pilot Partners / Seed Capital] to execute our 5–10 academy pilot, validate our AI Edge, and build the ultimate digital coaching engine for the world's most popular sport.")
    
    doc.save("KungFu_AI_Football_VC_Executive_Summary.docx")
    print("✅ Saved: KungFu_AI_Football_VC_Executive_Summary.docx")

if __name__ == '__main__':
    build_doc()