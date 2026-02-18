"""
NAPE 2026 Final Presentation — 15 slides
Strictly follows user framework. PNG charts from QMD. Minimal text. Big fonts.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# ═══════════════════════════ COLORS ═══════════════════════════
DARK_GREEN   = RGBColor(0x3A, 0x4E, 0x42)
MED_GREEN    = RGBColor(0x5E, 0x9A, 0x66)
LIGHT_GREEN  = RGBColor(0xAF, 0xD4, 0x80)
BRIGHT_GREEN = RGBColor(0x6A, 0xC7, 0x58)
CREAM        = RGBColor(0xFF, 0xF4, 0xDA)
WARM_GRAY    = RGBColor(0xE9, 0xE3, 0xDE)
CHARCOAL     = RGBColor(0x32, 0x32, 0x32)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
RED          = RGBColor(0xC0, 0x39, 0x2B)
WARM_TAN     = RGBColor(0xB0, 0xA5, 0x99)

# ═══════════════════════════ SETUP ═══════════════════════════
prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height

IMG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Output")

def bg(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def rect(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def pill(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def txt(slide, l, t, w, h, text, sz=18, color=CHARCOAL, bold=False, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(sz); p.font.color.rgb = color; p.font.bold = bold
    p.font.name = "Calibri"; p.alignment = align
    return tb

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

def header_bar(slide, title, subtitle=None):
    """Standard dark green top bar with title."""
    rect(slide, Inches(0), Inches(0), SW, Inches(1.0), DARK_GREEN)
    txt(slide, Inches(0.5), Inches(0.15), Inches(10), Inches(0.7),
        title, sz=28, color=WHITE, bold=True)
    if subtitle:
        p = pill(slide, Inches(9.5), Inches(0.2), Inches(3.5), Inches(0.5), MED_GREEN)
        p.text_frame.paragraphs[0].text = subtitle
        p.text_frame.paragraphs[0].font.size = Pt(12)
        p.text_frame.paragraphs[0].font.color.rgb = WHITE
        p.text_frame.paragraphs[0].font.bold = True
        p.text_frame.paragraphs[0].font.name = "Calibri"
        p.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

def add_png(slide, filename, left, top, width=None, height=None):
    """Insert a PNG from Output/. Returns shape or None if not found."""
    path = os.path.join(IMG_DIR, filename)
    if os.path.exists(path):
        return slide.shapes.add_picture(path, left, top, width, height)
    return None

def placeholder_box(slide, left, top, width, height, label):
    """Empty box with label for teammate to insert image later."""
    box = pill(slide, left, top, width, height, WHITE)
    box.line.color.rgb = MED_GREEN; box.line.width = Pt(2)
    box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]
    p.text = f"[ {label} ]"
    p.font.size = Pt(16); p.font.color.rgb = MED_GREEN
    p.font.bold = True; p.font.name = "Calibri"
    p.alignment = PP_ALIGN.CENTER
    return box

def build_table(slide, rows, cols, left, top, width, height):
    return slide.shapes.add_table(rows, cols, left, top, width, height).table

def cell(table, r, c, text, sz=12, color=CHARCOAL, bold=False,
         align=PP_ALIGN.LEFT, bg_color=None):
    cl = table.cell(r, c)
    cl.text = ""
    p = cl.text_frame.paragraphs[0]
    p.text = text; p.font.size = Pt(sz); p.font.color.rgb = color
    p.font.bold = bold; p.font.name = "Calibri"; p.alignment = align
    if bg_color:
        cl.fill.solid(); cl.fill.fore_color.rgb = bg_color

# ═══════════════════════════════════════════════════════════════
# SLIDE 1 — TITLE + TEAM
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, DARK_GREEN)
rect(s, Inches(0), Inches(0), Inches(0.35), SH, MED_GREEN)

# Decorative circles
for x, y, r, c in [(Inches(11.5), Inches(0.3), Inches(1.2), LIGHT_GREEN),
                    (Inches(12.0), Inches(1.0), Inches(0.6), MED_GREEN)]:
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, x, y, r, r)
    sh.fill.solid(); sh.fill.fore_color.rgb = c; sh.line.fill.background()

txt(s, Inches(1.0), Inches(0.6), Inches(10), Inches(0.6),
    "2026 SMU Cox NAPE Energy Innovation Case Competition",
    sz=18, color=LIGHT_GREEN)
txt(s, Inches(1.0), Inches(1.5), Inches(10.5), Inches(1.8),
    "Strategic Response to\nDissident Investor Activism",
    sz=44, color=WHITE, bold=True)
txt(s, Inches(1.0), Inches(3.5), Inches(10), Inches(0.7),
    "JV/PPA with AWS  |  Nuclear-Powered Data Center Strategy",
    sz=18, color=CREAM)

rect(s, Inches(1.0), Inches(4.5), Inches(8), Pt(3), LIGHT_GREEN)

txt(s, Inches(1.0), Inches(4.8), Inches(4), Inches(0.4),
    "TEAM MEMBERS", sz=14, color=LIGHT_GREEN, bold=True)

for i, name in enumerate(["Anh Bui", "Cuong Nguyen", "Minh Nguyen"]):
    p = pill(s, Inches(1.0 + i * 3.5), Inches(5.3), Inches(3.0), Inches(0.55), MED_GREEN)
    p.text_frame.paragraphs[0].text = name
    p.text_frame.paragraphs[0].font.size = Pt(16)
    p.text_frame.paragraphs[0].font.color.rgb = WHITE
    p.text_frame.paragraphs[0].font.name = "Calibri"
    p.text_frame.paragraphs[0].font.bold = True
    p.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

txt(s, Inches(1.0), Inches(6.5), Inches(5), Inches(0.4),
    "February 18, 2026  |  Finals Presentation", sz=13, color=WARM_TAN)
rect(s, Inches(0), Inches(7.2), SW, Inches(0.3), BRIGHT_GREEN)

notes(s, "Welcome. We're presenting our strategic recommendation for an IPP facing a 9% dissident investor. Our recommendation: Execute a JV/PPA with AWS, modeled on Talen Energy's transformative partnership.")

# ═══════════════════════════════════════════════════════════════
# SLIDE 2 — COMPANY BACKGROUND + TALEN COMPARISON
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Company Background & Real-World Precedent")

# Left: Our Company
txt(s, Inches(0.5), Inches(1.2), Inches(5.5), Inches(0.5),
    "OUR COMPANY", sz=16, color=DARK_GREEN, bold=True)

company_items = [
    "13,000 MW total capacity (Nuclear + Gas + Coal)",
    "2,200 MW nuclear plant — PJM RTO",
    "$20B market cap  |  45M shares  |  $444/share",
    "30x EV/EBITDA  |  BB credit rating",
    "9% dissident investor → 10% = special meeting",
    "Target: 30% Adj FCF/share growth",
]
for i, item in enumerate(company_items):
    txt(s, Inches(0.7), Inches(1.8 + i * 0.45), Inches(5.3), Inches(0.4),
        f"•  {item}", sz=13, color=CHARCOAL)

# Right: Talen Energy
txt(s, Inches(6.8), Inches(1.2), Inches(6), Inches(0.5),
    "REAL CASE: TALEN ENERGY (TLN)", sz=16, color=DARK_GREEN, bold=True)

talen_items = [
    "13,100 MW capacity — nearly identical fleet",
    "2,500 MW Susquehanna nuclear — PJM RTO",
    "Sold Cumulus DC to AWS for $650M (Mar 2024)",
    "20-year nuclear PPA with AWS through 2042+",
    "Stock: $60 → $389 — 6.5x appreciation",
    "Market cap: $3B → $17.6B post-deal",
]
for i, item in enumerate(talen_items):
    txt(s, Inches(7.0), Inches(1.8 + i * 0.45), Inches(5.8), Inches(0.4),
        f"•  {item}", sz=13, color=CHARCOAL)

# Talen stock chart
add_png(s, "fig_talen_stock.png", Inches(0.5), Inches(4.7), width=Inches(12.0))

notes(s, """Our company is a $20B IPP with 13,000 MW in PJM — virtually identical to Talen Energy. Talen sold its Cumulus data center campus to AWS for $650M, signed a 20-year nuclear PPA, and stock went from $60 to $389. The parallels are striking: same market, same fuel mix, same scale.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 3 — OPTIONS 1 & 2 PROS/CONS (MARKET & FINANCIAL)
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Market & Financial — Options 1 & 2", "WHY WE PASSED")

# Option 1 table
txt(s, Inches(0.5), Inches(1.2), Inches(6), Inches(0.5),
    "Option 1: Acquire Data Center ($4–6B)", sz=18, color=DARK_GREEN, bold=True)

tbl1 = build_table(s, 6, 2, Inches(0.5), Inches(1.8), Inches(5.8), Inches(2.8))
cell(tbl1, 0, 0, "✓  PROS", sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=MED_GREEN)
cell(tbl1, 0, 1, "✗  CONS", sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=RED)
for i, (pro, con) in enumerate([
    ("Full value chain control", "$4–6B cost → massive debt/dilution"),
    ("Multiple expansion potential", "Dilutive to FCF/share Years 1–2"),
    ("Direct AI/DC exposure", "No IPP has done this successfully"),
    ("Revenue diversification", "Integration risk (power ≠ DCs)"),
    ("Strategic optionality", "BB rating at risk of downgrade"),
]):
    bg_c = WARM_GRAY if i % 2 == 0 else WHITE
    cell(tbl1, i+1, 0, pro, sz=11, bg_color=bg_c)
    cell(tbl1, i+1, 1, con, sz=11, color=RED, bg_color=bg_c)

# Option 2 table
txt(s, Inches(6.8), Inches(1.2), Inches(6), Inches(0.5),
    "Option 2: Sell to Shell/ExxonMobil", sz=18, color=DARK_GREEN, bold=True)

tbl2 = build_table(s, 6, 2, Inches(6.8), Inches(1.8), Inches(5.8), Inches(2.8))
cell(tbl2, 0, 0, "✓  PROS", sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=MED_GREEN)
cell(tbl2, 0, 1, "✗  CONS", sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=RED)
for i, (pro, con) in enumerate([
    ("Immediate 20–30% premium", "Company ceases to exist"),
    ("Satisfies dissident with exit", "0% FCF growth — no upside"),
    ("Certain value realization", "ESG funds divest (oil major)"),
    ("De-risks shareholders", "Hyperscalers may cancel PPAs"),
    ("No execution risk", "NRC transfer: 12–18 months"),
]):
    bg_c = WARM_GRAY if i % 2 == 0 else WHITE
    cell(tbl2, i+1, 0, pro, sz=11, bg_color=bg_c)
    cell(tbl2, i+1, 1, con, sz=11, color=RED, bg_color=bg_c)

# Verdict bar
v = rect(s, Inches(0.5), Inches(4.9), Inches(12.0), Inches(0.6), DARK_GREEN)
v.text_frame.paragraphs[0].text = "VERDICT: Option 1 too expensive ($4–6B) and dilutive.  Option 2 destroys long-term value.  Neither unlocks the AI premium."
v.text_frame.paragraphs[0].font.size = Pt(14)
v.text_frame.paragraphs[0].font.color.rgb = WHITE
v.text_frame.paragraphs[0].font.bold = True
v.text_frame.paragraphs[0].font.name = "Calibri"
v.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Charts from QMD
add_png(s, "fig_option1_waterfall.png", Inches(0.3), Inches(5.7), width=Inches(6.2))
add_png(s, "fig_option2_premium.png", Inches(6.7), Inches(5.7), width=Inches(6.2))

notes(s, """Option 1: $4–6B acquisition is prohibitively expensive. Dilutive for Years 1–2, threatens BB credit rating.
Option 2: Shareholders get 20–30% premium but company dies. Zero future upside, ESG backlash, hyperscalers may cancel PPAs.
Neither option captures the AI premium our nuclear assets command.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 4 — OPTION 3 (WHY PROS OUTWEIGH CONS)
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Option 3: JV/PPA with AWS — Our Recommendation", "RECOMMENDED")

# Key deal terms (left side, concise)
txt(s, Inches(0.5), Inches(1.2), Inches(5.5), Inches(0.5),
    "Deal Structure (Talen Playbook)", sz=18, color=DARK_GREEN, bold=True)

terms = [
    ("PPA Partner:", "Amazon Web Services (AWS)"),
    ("Capacity:", "1,500–2,000 MW nuclear"),
    ("Price:", "$90–100/MWh (vs $51 merchant) → +86%"),
    ("Duration:", "20 years (2026–2046)"),
    ("Investment:", "$200–400M (from cash — ZERO dilution)"),
    ("DC Campus:", "100–200 MW IT → sale to AWS $750–900M"),
]
for i, (label, value) in enumerate(terms):
    y = Inches(1.8 + i * 0.42)
    tb = txt(s, Inches(0.7), y, Inches(5.3), Inches(0.4), "", sz=13)
    tf = tb.text_frame
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = label + " "; r1.font.size = Pt(13)
    r1.font.bold = True; r1.font.color.rgb = MED_GREEN; r1.font.name = "Calibri"
    r2 = p.add_run(); r2.text = value; r2.font.size = Pt(13)
    r2.font.color.rgb = CHARCOAL; r2.font.name = "Calibri"

# Right: Why pros outweigh cons
txt(s, Inches(6.5), Inches(1.2), Inches(6), Inches(0.5),
    "Why Pros Outweigh Cons", sz=18, color=DARK_GREEN, bold=True)

tbl3 = build_table(s, 7, 2, Inches(6.5), Inches(1.8), Inches(6.2), Inches(3.5))
cell(tbl3, 0, 0, "✓  PROS (11)", sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=MED_GREEN)
cell(tbl3, 0, 1, "✗  CONS (6)", sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=RED)
for i, (pro, con) in enumerate([
    ("+65% FCF growth (vs 30% target)", "Concentrated on single partner"),
    ("$200M cost vs $4–6B (Option 1)", "PPA locks in price (upside cap)"),
    ("Proven by Talen: 6.5x stock", "FERC scrutiny on BTM arrangement"),
    ("Zero equity dilution", "Nuclear ops risk during transition"),
    ("Preserves independence", "DC campus construction timeline"),
    ("Satisfies dissident immediately", "DC acquisition integration risk"),
]):
    bg_c = WARM_GRAY if i % 2 == 0 else WHITE
    cell(tbl3, i+1, 0, pro, sz=11, bg_color=bg_c)
    cell(tbl3, i+1, 1, con, sz=11, color=RED, bg_color=bg_c)

# Bottom: financial impact summary
for i, (metric, value, desc) in enumerate([
    ("Year 1 FCF", "+65%", "$10.20 → $16.87/sh"),
    ("Revenue", "+$513M/yr", "PPA incremental"),
    ("Market Cap", "$32B", "from $20B"),
    ("Share Price", "$711/sh", "from $444"),
]):
    x = Inches(0.5 + i * 3.15)
    card = pill(s, x, Inches(5.7), Inches(2.85), Inches(1.5), DARK_GREEN)
    tf = card.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = metric; p.font.size = Pt(11)
    p.font.color.rgb = LIGHT_GREEN; p.font.bold = True; p.font.name = "Calibri"; p.alignment = PP_ALIGN.CENTER
    p2 = tf.add_paragraph(); p2.text = value; p2.font.size = Pt(24)
    p2.font.color.rgb = WHITE; p2.font.bold = True; p2.font.name = "Calibri"; p2.alignment = PP_ALIGN.CENTER
    p3 = tf.add_paragraph(); p3.text = desc; p3.font.size = Pt(10)
    p3.font.color.rgb = CREAM; p3.font.name = "Calibri"; p3.alignment = PP_ALIGN.CENTER

notes(s, """Option 3 — JV/PPA with AWS. Modeled on Talen's March 2024 deal.
PPA at $90-100/MWh (vs $51 merchant). +$513M/yr incremental revenue. FCF goes from $10.20 to $16.87/share (+65%).
Only $200-400M investment, funded from cash. Zero dilution. Preserves independence.
Pros clearly outweigh cons: the downside risks are manageable (FERC has Talen precedent, single-partner risk mitigated by competitive bidding). The upside is transformative.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 5 — CHART (OPTION 3 FINANCIAL DETAIL)
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Option 3 — Revenue & FCF Impact")

add_png(s, "fig_option3_revenue.png", Inches(0.3), Inches(1.1), width=Inches(6.3))
add_png(s, "fig_option3_fcf.png", Inches(6.7), Inches(1.1), width=Inches(6.3))

# Shareholder value trajectory at bottom
add_png(s, "fig_shareholder_value.png", Inches(0.3), Inches(4.2), width=Inches(12.5))

notes(s, """Left: Nuclear revenue jumps from $637M (merchant at $51/MWh) to $1,150M (PPA at $95/MWh).
Right: FCF/share goes from $10.20 to $16.87 in Year 1 — 65% growth, far exceeding the dissident's 30% target.
Bottom: 5-year shareholder value trajectory shows Option 3 reaching $52B market cap by Year 5 vs $32B for Option 1 and flat $24B for Option 2.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 6 — PLACEHOLDER: TEAMMATE'S TALEN FINANCIAL TIME SERIES
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Talen Energy — Financial Transformation", "REAL-WORLD DATA")

placeholder_box(s, Inches(0.5), Inches(1.2), Inches(12.0), Inches(6.0),
    "TEAMMATE'S CHART — Talen Financial Time Series\n\n"
    "Insert: Stock price trajectory, Revenue/EBITDA growth,\n"
    "Market cap expansion ($3B → $17.6B)")

notes(s, """Teammate inserts Talen Energy time series analysis: stock trajectory ($60→$389), revenue/EBITDA growth post-AWS deal, market cap expansion ($3B→$17.6B), and our company's projected overlay.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 7 — OPTIONS 1 & 2 EXECUTION PROS/CONS
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Operational & Execution — Options 1 & 2", "RESOURCES & SUPPLY CHAIN")

# Option 1 execution challenges
txt(s, Inches(0.5), Inches(1.2), Inches(6), Inches(0.5),
    "Option 1: Acquire DC — Execution", sz=18, color=DARK_GREEN, bold=True)

tbl_e1 = build_table(s, 5, 3, Inches(0.5), Inches(1.8), Inches(5.8), Inches(2.5))
for i, h in enumerate(["Challenge", "Severity", "Detail"]):
    cell(tbl_e1, 0, i, h, sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=DARK_GREEN)
for r, (ch, sev, det) in enumerate([
    ("Integration complexity", "HIGH", "Power gen ≠ DC operations"),
    ("Capital deployment", "HIGH", "$4–6B — 18–24 month fundraise"),
    ("Talent gap", "MEDIUM", "Need DC engineers, not plant ops"),
    ("Regulatory", "MEDIUM", "FERC, PUC, antitrust: 6–12 months"),
]):
    bg_c = WARM_GRAY if r % 2 == 0 else WHITE
    cell(tbl_e1, r+1, 0, ch, sz=11, bold=True, bg_color=bg_c)
    sc = RED if sev == "HIGH" else MED_GREEN
    cell(tbl_e1, r+1, 1, sev, sz=11, bold=True, color=sc, align=PP_ALIGN.CENTER, bg_color=bg_c)
    cell(tbl_e1, r+1, 2, det, sz=11, bg_color=bg_c)

# Option 2 execution challenges
txt(s, Inches(6.8), Inches(1.2), Inches(6), Inches(0.5),
    "Option 2: Sale — Execution Risks", sz=18, color=DARK_GREEN, bold=True)

tbl_e2 = build_table(s, 5, 3, Inches(6.8), Inches(1.8), Inches(5.8), Inches(2.5))
for i, h in enumerate(["Challenge", "Severity", "Detail"]):
    cell(tbl_e2, 0, i, h, sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=DARK_GREEN)
for r, (ch, sev, det) in enumerate([
    ("NRC license transfer", "HIGH", "12–18 months, $10–20M legal"),
    ("ESG backlash", "HIGH", "Oil major → ESG funds divest"),
    ("PPA termination", "HIGH", "AWS/MSFT may exit nuclear PPAs"),
    ("Culture clash", "MEDIUM", "Oil major parent disruption"),
]):
    bg_c = WARM_GRAY if r % 2 == 0 else WHITE
    cell(tbl_e2, r+1, 0, ch, sz=11, bold=True, bg_color=bg_c)
    sc = RED if sev == "HIGH" else MED_GREEN
    cell(tbl_e2, r+1, 1, sev, sz=11, bold=True, color=sc, align=PP_ALIGN.CENTER, bg_color=bg_c)
    cell(tbl_e2, r+1, 2, det, sz=11, bg_color=bg_c)

# Timeline comparison
v = rect(s, Inches(0.5), Inches(4.5), Inches(12.0), Inches(0.6), DARK_GREEN)
v.text_frame.paragraphs[0].text = "Option 1: 18–24 months  |  Option 2: 12–18 months  |  Option 3: 1–3 months to close"
v.text_frame.paragraphs[0].font.size = Pt(16)
v.text_frame.paragraphs[0].font.color.rgb = WHITE
v.text_frame.paragraphs[0].font.bold = True
v.text_frame.paragraphs[0].font.name = "Calibri"
v.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

# Timeline bars
for i, (opt, time, w_bar, color) in enumerate([
    ("Option 1", "18–24 months — Fundraise → Acquire → Integrate", Inches(9.5), RED),
    ("Option 2", "12–18 months — NRC Review → FERC → Antitrust → Close", Inches(8.0), RED),
    ("Option 3", "1–3 months — Negotiate PPA → Sign → Revenue Starts", Inches(3.0), BRIGHT_GREEN),
]):
    y = Inches(5.4 + i * 0.65)
    lbl = pill(s, Inches(0.5), y, Inches(1.5), Inches(0.55), DARK_GREEN)
    lbl.text_frame.paragraphs[0].text = opt
    lbl.text_frame.paragraphs[0].font.size = Pt(12); lbl.text_frame.paragraphs[0].font.color.rgb = WHITE
    lbl.text_frame.paragraphs[0].font.bold = True; lbl.text_frame.paragraphs[0].font.name = "Calibri"
    lbl.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    bar = pill(s, Inches(2.2), y, w_bar, Inches(0.55), color)
    bar.text_frame.paragraphs[0].text = time
    bar.text_frame.paragraphs[0].font.size = Pt(11); bar.text_frame.paragraphs[0].font.color.rgb = WHITE
    bar.text_frame.paragraphs[0].font.bold = True; bar.text_frame.paragraphs[0].font.name = "Calibri"

notes(s, """Option 1: Integration complexity is extreme — power gen and data centers are different businesses. $4–6B fundraising takes 18–24 months.
Option 2: NRC license transfer alone takes 12–18 months and costs $10–20M. ESG funds would divest.
Option 3 closes in 1–3 months — it's a bilateral PPA contract.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 8 — OPTION 3 EXECUTION PLAN
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Option 3 — Execution Roadmap")

# Regulatory dual-track chart from QMD
add_png(s, "fig_regulatory_dual_track.png", Inches(0.3), Inches(1.1), width=Inches(12.5))

# Phase summary table
tbl_ph = build_table(s, 6, 5, Inches(0.3), Inches(4.3), Inches(12.5), Inches(3.0))
for i, h in enumerate(["Phase", "Timeline", "Action", "Cost", "Value Created"]):
    cell(tbl_ph, 0, i, h, sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=DARK_GREEN)
phases = [
    ["Phase 1", "Month 0–3", "Negotiate & Sign AWS PPA", "$0", "+$513M/yr revenue"],
    ["Phase 2", "Month 3–12", "Develop DC Campus", "$200–400M", "100–200 MW IT capacity"],
    ["Phase 3", "Month 12–18", "Campus Sale to AWS", "+$750–900M", "Net cash +$350–500M"],
    ["Phase 4", "Year 2–5", "Expansion Rights", "+$500M–1B", "300–500 MW additional"],
    ["Phase 5", "Year 3–7", "DC Acquisitions (Talen)", "$250–530M", "$750M–1.6B value"],
]
for r, row in enumerate(phases):
    bg_c = WARM_GRAY if r % 2 == 0 else WHITE
    for c, val in enumerate(row):
        bold = c == 0
        cell(tbl_ph, r+1, c, val, sz=11, bold=bold, align=PP_ALIGN.CENTER, bg_color=bg_c)

notes(s, """5-phase execution roadmap:
Phase 1 (Months 0–3): Bilateral PPA with AWS. No regulatory approval needed for front-of-meter. Revenue starts.
Phase 2 (Months 3–12): DC campus development. $200–400M, AWS funds majority.
Phase 3 (Months 12–18): Campus sale to AWS for $750–900M. Talen got $650M for smaller campus.
Phase 4 (Years 2–5): Exercise expansion rights. Add 300–500 MW.
Phase 5 (Years 3–7): Acquire DCs near power plants, following Talen's Cumulus Growth playbook.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 9 — CHART: SENSITIVITY/TORNADO
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Sensitivity Analysis — What Drives Value")

add_png(s, "fig_tornado_opt3.png", Inches(0.3), Inches(1.1), width=Inches(7.5), height=Inches(5.5))

# Key takeaways
txt(s, Inches(8.3), Inches(1.2), Inches(4.5), Inches(0.4),
    "Key Takeaways", sz=18, color=DARK_GREEN, bold=True)

insights = [
    ("PPA Price = #1 Driver", "$11B NPV swing. Run competitive bidding among AWS, MSFT, Google."),
    ("Multiple Re-Rating", "$7.7B swing. Talen 28x is our floor; 40x achievable."),
    ("Nuclear CF Stable", "92%+ CF industry-proven. Lowest-risk input."),
    ("Even Low Case Wins", "$75/MWh → +$290M/yr vs merchant. $5.8B NPV."),
]
for i, (title, desc) in enumerate(insights):
    y = Inches(1.8 + i * 1.4)
    card = pill(s, Inches(8.3), y, Inches(4.5), Inches(1.2), WHITE)
    card.line.color.rgb = MED_GREEN; card.line.width = Pt(1.5)
    tf = card.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = title; p.font.size = Pt(13)
    p.font.bold = True; p.font.color.rgb = DARK_GREEN; p.font.name = "Calibri"
    p2 = tf.add_paragraph(); p2.text = desc; p2.font.size = Pt(11)
    p2.font.color.rgb = CHARCOAL; p2.font.name = "Calibri"

notes(s, """PPA Price is the #1 value driver — $11B NPV swing between $75 and $110/MWh. This is why we run competitive bidding.
Multiple re-rating is #2 — Talen trades at 28x, which is our floor.
Even in the worst case ($75/MWh), Option 3 still generates $5.8B NPV.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 10 — PLACEHOLDER: TEAMMATE'S TALEN EXECUTION TIME SERIES
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Talen Energy — Execution Timeline", "REAL-WORLD DATA")

placeholder_box(s, Inches(0.5), Inches(1.2), Inches(12.0), Inches(6.0),
    "TEAMMATE'S CHART — Talen Execution Timeline\n\n"
    "Insert: Deal timeline (negotiation → close → expansion),\n"
    "Cumulus campus milestones, post-deal DC acquisitions")

notes(s, """Teammate inserts Talen execution timeline: deal negotiation through close, Cumulus campus development, post-deal DC acquisitions through Cumulus Growth subsidiary.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 11 — COMPARATIVE ANALYSIS: ALL THREE OPTIONS
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Comparative Analysis — All Three Options")

# Decision matrix chart
add_png(s, "fig_three_options.png", Inches(0.3), Inches(1.1), width=Inches(6.5))

# Radar chart
add_png(s, "fig_radar.png", Inches(7.0), Inches(1.0), width=Inches(5.8), height=Inches(5.5))

# Winner bar at bottom
w_bar = rect(s, Inches(0.3), Inches(6.7), Inches(12.5), Inches(0.6), DARK_GREEN)
w_bar.text_frame.paragraphs[0].text = "WINNER: Option 3 — JV/PPA with AWS  |  Score: 4.80/5  |  Only option delivering +65% FCF growth + independence + proven precedent"
w_bar.text_frame.paragraphs[0].font.size = Pt(14)
w_bar.text_frame.paragraphs[0].font.color.rgb = WHITE
w_bar.text_frame.paragraphs[0].font.bold = True
w_bar.text_frame.paragraphs[0].font.name = "Calibri"
w_bar.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

notes(s, """Weighted decision matrix: Option 1 scores 2.45/5, Option 2 scores 2.15/5, Option 3 scores 4.80/5.
Option 3 dominates every criterion: FCF growth (5/5), multiple expansion (5/5), execution feasibility (4/5), shareholder value (5/5), independence (5/5), satisfies dissident (5/5).
The radar chart visually shows Option 3's complete dominance across all dimensions.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 12 — RISK ASSESSMENT
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Risk Assessment & Mitigation")

tbl_r = build_table(s, 7, 4, Inches(0.3), Inches(1.2), Inches(12.5), Inches(5.5))
for i, h in enumerate(["Risk", "Probability", "Impact", "Mitigation"]):
    cell(tbl_r, 0, i, h, sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=DARK_GREEN)
risks = [
    ["PPA negotiation fails", "Low (20%)", "High", "Competitive bid: AWS, MSFT, Google, Meta all want nuclear"],
    ["FERC blocks BTM", "Med (30%)", "Medium", "Front-of-meter PPA primary (no FERC needed); BTM is backup only"],
    ["Nuclear plant outage", "Low (10%)", "High", "92%+ CF; dual-unit redundancy; force majeure clauses"],
    ["PPA price below target", "Med (25%)", "Low", "Even $75/MWh = +$290M/yr vs merchant"],
    ["Dissident escalates first", "Med (35%)", "Medium", "Board seat offer; announce framework early"],
    ["DC construction delay", "Med (30%)", "Low", "PPA revenue starts Day 1 regardless of campus"],
]
for r, (risk, prob, impact, mit) in enumerate(risks):
    bg_c = WARM_GRAY if r % 2 == 0 else WHITE
    cell(tbl_r, r+1, 0, risk, sz=12, bold=True, bg_color=bg_c)
    pc = DARK_GREEN if "Low" in prob else MED_GREEN
    cell(tbl_r, r+1, 1, prob, sz=12, bold=True, color=pc, align=PP_ALIGN.CENTER, bg_color=bg_c)
    ic = RED if impact == "High" else MED_GREEN
    cell(tbl_r, r+1, 2, impact, sz=12, bold=True, color=ic, align=PP_ALIGN.CENTER, bg_color=bg_c)
    cell(tbl_r, r+1, 3, mit, sz=11, bg_color=bg_c)

# Bottom bar
rb = rect(s, Inches(0.3), Inches(6.9), Inches(12.5), Inches(0.5), DARK_GREEN)
rb.text_frame.paragraphs[0].text = "Even worst-case ($75/MWh): +$290M/yr revenue, $5.8B NPV  |  Talen faced greater risks (bankruptcy) and succeeded"
rb.text_frame.paragraphs[0].font.size = Pt(13)
rb.text_frame.paragraphs[0].font.color.rgb = WHITE
rb.text_frame.paragraphs[0].font.bold = True
rb.text_frame.paragraphs[0].font.name = "Calibri"
rb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

notes(s, """Key risks and mitigations. Most important: PPA negotiation failure (20% prob) — mitigated by competitive bidding among 4 hyperscalers.
FERC risk mitigated by using front-of-meter PPA as primary track (no approval needed).
Even in worst case, Option 3 generates +$290M/yr and $5.8B NPV. Talen faced far greater risks (bankruptcy, OTC listing) and still succeeded.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 13 — ESG CONSIDERATION
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "ESG & Sustainability")

add_png(s, "fig_esg_comparison.png", Inches(0.3), Inches(1.1), width=Inches(12.5))

# Key insight cards
txt(s, Inches(0.3), Inches(4.3), Inches(12), Inches(0.5),
    "Hyperscaler Sustainability Mandates — Why Nuclear Wins", sz=16, color=DARK_GREEN, bold=True)

for i, (company, target) in enumerate([
    ("AWS", "Net-zero by 2040  |  35+ GW needed"),
    ("Microsoft", "Carbon-negative by 2030  |  30+ GW needed"),
    ("Google", "24/7 carbon-free by 2030  |  25+ GW needed"),
]):
    x = Inches(0.3 + i * 4.2)
    card = pill(s, x, Inches(4.9), Inches(3.8), Inches(0.7), DARK_GREEN)
    tf = card.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = f"{company}: {target}"
    p.font.size = Pt(12); p.font.color.rgb = WHITE; p.font.bold = True
    p.font.name = "Calibri"; p.alignment = PP_ALIGN.CENTER

# ESG value bar
eb = rect(s, Inches(0.3), Inches(5.9), Inches(12.5), Inches(1.3), DARK_GREEN)
tf = eb.text_frame; tf.word_wrap = True
p = tf.paragraphs[0]; p.text = "ESG Value: $2–4B Over 5 Years"
p.font.size = Pt(18); p.font.color.rgb = WHITE; p.font.bold = True; p.font.name = "Calibri"; p.alignment = PP_ALIGN.CENTER
p2 = tf.add_paragraph()
p2.text = "+1–2x EV/EBITDA from ESG premium  |  Green bond savings $15–25M/yr  |  45Q credits $85–350M/yr"
p2.font.size = Pt(12); p2.font.color.rgb = CREAM; p2.font.name = "Calibri"; p2.alignment = PP_ALIGN.CENTER
p3 = tf.add_paragraph()
p3.text = "Option 2 DESTROYS ESG value (oil major ownership → ESG funds divest)"
p3.font.size = Pt(12); p3.font.color.rgb = BRIGHT_GREEN; p3.font.bold = True; p3.font.name = "Calibri"; p3.alignment = PP_ALIGN.CENTER

notes(s, """Nuclear = 0g CO2/MWh. Option 3 aligns perfectly with hyperscaler sustainability mandates.
AWS, Microsoft, Google all have aggressive carbon-free targets and need 100+ GW of clean power.
ESG value: $2–4B over 5 years from multiple premium, green bonds, 45Q credits.
Option 2 destroys ESG value — oil major ownership causes ESG fund divestment.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 14 — CCUS
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "CCUS Integration — Gas Fleet Value Unlock")

add_png(s, "fig_ccus_economics.png", Inches(0.3), Inches(1.1), width=Inches(12.5))

# Phased deployment table
tbl_c = build_table(s, 5, 5, Inches(0.3), Inches(4.3), Inches(12.5), Inches(2.5))
for i, h in enumerate(["Phase", "Capacity", "45Q Credits/yr", "Capex", "Timeline"]):
    cell(tbl_c, 0, i, h, sz=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=DARK_GREEN)
for r, row in enumerate([
    ["1: Pilot", "800 MW", "$85M", "$400M", "2027–29"],
    ["2: Scale", "2,000 MW", "$210M", "$900M", "2029–31"],
    ["3: Full Fleet", "3,500 MW", "$350M", "$1.5B", "2031–34"],
    ["Total", "6,300 MW", "$350M/yr", "$2.8B", "2027–34"],
]):
    bg_c = DARK_GREEN if r == 3 else (WARM_GRAY if r % 2 == 0 else WHITE)
    fc = WHITE if r == 3 else CHARCOAL
    for c, val in enumerate(row):
        cell(tbl_c, r+1, c, val, sz=12, bold=(r==3), color=fc, align=PP_ALIGN.CENTER, bg_color=bg_c)

# Competitive moat
mb = rect(s, Inches(0.3), Inches(7.0), Inches(12.5), Inches(0.4), DARK_GREEN)
mb.text_frame.paragraphs[0].text = "Nuclear (0g CO₂) + CCUS Gas (~72 lb/MWh) = Only IPP offering 24/7 near-zero carbon in PJM"
mb.text_frame.paragraphs[0].font.size = Pt(13)
mb.text_frame.paragraphs[0].font.color.rgb = WHITE
mb.text_frame.paragraphs[0].font.bold = True
mb.text_frame.paragraphs[0].font.name = "Calibri"
mb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

notes(s, """CCUS adds $28/MWh to gas generation through 45Q credits, carbon cost avoidance, and premium PPA pricing.
3-phase deployment: pilot (800 MW), scale (2,000 MW), full fleet (3,500 MW). Total: $2.8B capex, $350M/yr in credits.
Combined nuclear + CCUS gas = only IPP in PJM offering true 24/7 near-zero carbon energy.""")

# ═══════════════════════════════════════════════════════════════
# SLIDE 15 — NUCLEAR REGULATORY & RISK SOLUTION
# ═══════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6])
bg(s, CREAM)
header_bar(s, "Nuclear Regulatory & Risk Management")

# Regulatory table (concise)
tbl_rg = build_table(s, 5, 4, Inches(0.3), Inches(1.2), Inches(12.5), Inches(2.8))
for i, h in enumerate(["Body", "Concern", "Risk", "Our Solution"]):
    cell(tbl_rg, 0, i, h, sz=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER, bg_color=DARK_GREEN)
for r, (body, concern, risk, sol) in enumerate([
    ["NRC", "License amendment for DC", "Low–Med", "No change for FoM PPA; BTM has Talen precedent"],
    ["FERC", "BTM reduces grid capacity", "Medium", "FoM PPA primary (no approval); Talen survived FERC"],
    ["State PUC", "DC siting permits", "Low", "Nuclear sites have industrial zoning already"],
    ["PJM RTO", "Capacity obligation", "Medium", "Maintain FoM to preserve capacity revenue"],
]):
    bg_c = WARM_GRAY if r % 2 == 0 else WHITE
    cell(tbl_rg, r+1, 0, body, sz=13, bold=True, align=PP_ALIGN.CENTER, bg_color=bg_c)
    cell(tbl_rg, r+1, 1, concern, sz=12, bg_color=bg_c)
    rc = DARK_GREEN if "Low" in risk else MED_GREEN
    cell(tbl_rg, r+1, 2, risk, sz=13, bold=True, color=rc, align=PP_ALIGN.CENTER, bg_color=bg_c)
    cell(tbl_rg, r+1, 3, sol, sz=12, bg_color=bg_c)

# Dual-track visual (from QMD PNG)
add_png(s, "fig_regulatory_dual_track.png", Inches(0.3), Inches(4.2), width=Inches(7.0))

# Key advantages (right side)
txt(s, Inches(7.8), Inches(4.2), Inches(5), Inches(0.4),
    "Key Advantages vs Options 1 & 2", sz=16, color=DARK_GREEN, bold=True)

for i, (title, desc) in enumerate([
    ("No NRC Transfer", "Opt 2 needs 12–18 months. We retain ownership."),
    ("No Antitrust Review", "Opt 1's $4–6B triggers HSR Act. PPA doesn't."),
    ("FERC De-Risked", "FoM PPA is standard bilateral trade."),
    ("1–3 Month Close", "Revenue starts immediately."),
]):
    y = Inches(4.7 + i * 0.7)
    card = pill(s, Inches(7.8), y, Inches(5.0), Inches(0.6), WHITE)
    card.line.color.rgb = MED_GREEN; card.line.width = Pt(1.5)
    tf = card.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = f"✓ {title}: "; r1.font.size = Pt(12)
    r1.font.bold = True; r1.font.color.rgb = DARK_GREEN; r1.font.name = "Calibri"
    r2 = p.add_run(); r2.text = desc; r2.font.size = Pt(11)
    r2.font.color.rgb = CHARCOAL; r2.font.name = "Calibri"

notes(s, """Dual-track regulatory strategy:
Track A (Primary): Front-of-meter PPA — standard bilateral trade, no regulatory approval needed.
Track B (Backup): Behind-the-meter — higher value but requires NRC amendment. Talen precedent exists.
Key advantage: Option 3 avoids the 12–18 month NRC license transfer (Option 2) and the antitrust review (Option 1). Closes in 1–3 months.""")

# ═══════════════════ SAVE ═══════════════════
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "NAPE_Final_Presentation.pptx")
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
