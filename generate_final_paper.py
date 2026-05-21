from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import docx

def set_font(run, size=12, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    r = run._element
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:cs"),    "Times New Roman")

def add_paragraph(doc, text, bold=False, italic=False, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                  space_before=0, space_after=0, first_indent=True, size=12, line_spacing=2.0):
    p = doc.add_paragraph()
    p.alignment = alignment
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after  = Pt(space_after)
    pf.line_spacing = Pt(size * line_spacing)
    if first_indent:
        pf.first_line_indent = Inches(0.5)
    else:
        pf.first_line_indent = Inches(0)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold, italic=italic)
    return p

def add_heading(doc, text, size=12, center=False, bold=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after  = Pt(0)
    pf.first_line_indent = Inches(0)
    pf.line_spacing = Pt(size * 2.0)
    run = p.add_run(text)
    set_font(run, size=size, bold=bold)
    return p

def add_ref(doc, text, italic_range=None):
    """Add a hanging-indent reference entry."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after  = Pt(0)
    pf.first_line_indent = Inches(-0.5)
    pf.left_indent = Inches(0.5)
    pf.line_spacing = Pt(24)

    if italic_range is None:
        run = p.add_run(text)
        set_font(run, size=12)
    else:
        # italic_range = (start_char, end_char) within text
        start, end = italic_range
        r1 = p.add_run(text[:start])
        set_font(r1, size=12)
        r2 = p.add_run(text[start:end])
        set_font(r2, size=12, italic=True)
        r3 = p.add_run(text[end:])
        set_font(r3, size=12)
    return p

# ── Build document ──────────────────────────────────────────────────────────
doc = Document()

# Page margins (1 inch all sides)
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1)
    section.right_margin  = Inches(1)

# ── Header block (centered, double-spaced, no indent) ──
for line in [
    "William Robert Schreiner Jr",
    "Economics 104 — Principles of Macroeconomics",
    "Youssuf Abdelatif",
    "May 22, 2026",
]:
    add_heading(doc, line, size=12, center=True, bold=False)

# Title
add_heading(doc, "Japan Economic Analysis: An IMF Advisor Report", size=12, center=True, bold=True)

# ── Introduction ────────────────────────────────────────────────────────────
add_heading(doc, "Introduction", bold=True)

add_paragraph(doc,
    "Japan presents one of the most compelling macroeconomic case studies in the modern world. It is "
    "simultaneously an advanced, high-income economy with a GDP per capita of approximately $35,700 "
    "(IMF, 2026), one of the lowest unemployment rates globally at 2.5 percent, and a high Human "
    "Development Index score of 0.920 — and yet it faces a set of structural challenges that threaten "
    "its long-run trajectory. A shrinking population, the highest government debt-to-GDP ratio of any "
    "major advanced economy at roughly 255 percent, and slow GDP growth of just 0.6 percent in 2025 "
    "all demand careful policy attention. This report analyzes Japan's current macroeconomic position "
    "and provides both short-term and long-term policy recommendations from the perspective of IMF advisors."
)

# ── Section 1 ────────────────────────────────────────────────────────────────
add_heading(doc, "Economic Output and Growth", bold=True)

add_paragraph(doc,
    "Japan's real GDP per capita of $35,700 places it approximately 32nd globally, well above the "
    "world average of $15,680 but significantly below the United States at $94,430 (IMF, 2026). The "
    "gap reflects higher U.S. productivity, a larger and more flexible workforce, and a dominant "
    "services sector. Real GDP growth in Japan has decelerated sharply over recent decades; after "
    "the extraordinary Post-WWII Economic Miracle — during which GDP expanded at nearly 10 percent "
    "annually from 1945 to 1973 — growth has settled near zero to one percent in recent years. "
    "Japan's population is declining at approximately 0.5 percent per year due to low birth rates "
    "and minimal immigration, which directly constrains labor supply and long-run output potential."
)

add_paragraph(doc,
    "Japan's main export sector is automotive manufacturing, contributing over $100 billion "
    "annually and reflecting a strong comparative advantage in capital-intensive precision production. "
    "Companies such as Toyota, Honda, and Nissan dominate global markets. This export structure is "
    "sustained by Japan's deep stock of industrial capital and its culture of kaizen, or continuous "
    "process improvement, which produces quality and cost-control advantages difficult for other "
    "nations to replicate."
)

# ── Section 2 ────────────────────────────────────────────────────────────────
add_heading(doc, "Long-Run Growth: Capital, Institutions, and Human Capital", bold=True)

add_paragraph(doc,
    "From a production function perspective, Japan's capital stock per worker is already large, "
    "placing it on the flat portion of the diminishing-returns curve. Additional physical capital "
    "investment would yield only marginal output gains; meaningful growth must come from upward "
    "shifts of the production function — that is, improvements in technology and total factor "
    "productivity. Japan's institutional quality is favorable for this: the country scores "
    "approximately 0.73 out of 1.0 on the Liberal Political Institutions Index (Our World in Data, "
    "2024), reflecting a stable democracy, independent judiciary, and secure property rights that "
    "support innovation and investment."
)

add_paragraph(doc,
    "Japan's human capital base is also strong. Its Human Development Index score of 0.920 ranks "
    "24th globally (UNDP, 2023), driven by a life expectancy of approximately 84 years and an "
    "average of 13.4 years of schooling. This highly skilled workforce underpins Japan's strength "
    "in advanced manufacturing and technology. The challenge is demographic: as the workforce ages "
    "and shrinks, sustaining the human capital stock requires deliberate policy intervention."
)

# ── Section 3 ────────────────────────────────────────────────────────────────
add_heading(doc, "Monetary Policy and Labor Markets", bold=True)

add_paragraph(doc,
    "Japan's unemployment rate of 2.5 percent is well below the United States' estimated natural "
    "rate of 4 percent. This reflects Japan's lifetime employment culture, which sharply reduces "
    "both frictional unemployment — workers cycling between jobs — and structural unemployment. "
    "With the labor market running tight and inflation at 2.7 percent, slightly above the Bank of "
    "Japan's 2 percent target, monetary tightening is appropriate. The Bank of Japan moved "
    "decisively: it ended its negative interest rate policy in March 2024, raised rates to 0.25 "
    "percent in July 2024, and again to 0.5 percent in January 2025 — its first sustained "
    "rate-hiking cycle in decades (Bank of Japan, 2025). This response is consistent with the "
    "Federal Reserve's dual mandate logic: raise rates when inflation is above target and "
    "unemployment is at or below the natural rate."
)

# ── Section 4 ────────────────────────────────────────────────────────────────
add_heading(doc, "Fiscal Policy and Government Debt", bold=True)

add_paragraph(doc,
    "Japan's government expenditure represents approximately 38 to 40 percent of GDP, and the "
    "country runs budget deficits of roughly 5 to 6 percent annually (IMF, 2024). With the output "
    "gap near zero — the economy is operating close to its potential — counter-cyclical best "
    "practice calls for fiscal consolidation. Japan is not following this practice, partly because "
    "demographic pressures create unavoidable structural spending on pensions and healthcare. The "
    "resulting debt-to-GDP ratio of approximately 255 percent is the highest among major advanced "
    "economies. While roughly 90 percent of this debt is held domestically, limiting the risk of a "
    "foreign-driven crisis, the long-term trajectory is unsustainable without reform. A shrinking "
    "tax base and rising social security costs create a structural imbalance that will worsen "
    "without intervention."
)

# ── Section 5: Policy Recommendations ───────────────────────────────────────
add_heading(doc, "Policy Recommendations", bold=True)

add_paragraph(doc,
    "In the short term, the Bank of Japan should continue its gradual rate-normalization path to "
    "bring inflation back to the 2 percent target while monitoring employment closely. Simultaneously, "
    "the government should pursue modest fiscal consolidation — reducing discretionary spending "
    "and allowing automatic stabilizers to function — to begin stabilizing the debt-to-GDP ratio "
    "before it deteriorates further."
)

add_paragraph(doc,
    "In the long term, Japan's most urgent priority is addressing demographic decline. Expanding "
    "immigration pathways for skilled workers would directly grow the labor supply and broaden the "
    "tax base. Structural labor market reforms — increasing flexibility, reducing the cultural "
    "dependence on lifetime employment, and expanding female workforce participation — would also "
    "raise potential output. To sustain growth on the production function's flat portion, Japan "
    "must invest heavily in R&D, artificial intelligence, and automation to shift the production "
    "function upward through technology. Finally, pension and healthcare reform is essential to "
    "reduce the structural deficit and place the debt-to-GDP ratio on a downward path. Together, "
    "these measures would position Japan to maintain its status as a high-income, stable economy "
    "over the next generation."
)

# ── References ───────────────────────────────────────────────────────────────
add_heading(doc, "References", bold=True)

refs = [
    ("Bank of Japan. (2025). ",
     "Monetary policy decisions — Interest rate announcements 2024–2025.",
     " Bank of Japan. https://www.boj.or.jp/en/mopo/index.htm"),
    ("Congressional Budget Office. (2024). ",
     "The budget and economic outlook: 2024 to 2034.",
     " U.S. Congressional Budget Office. https://www.cbo.gov/publication/59946"),
    ("International Monetary Fund. (2026). ",
     "World Economic Outlook DataMapper — GDP per capita, current prices.",
     " IMF. https://www.imf.org/external/datamapper/NGDPDPC@WEO/OEMDC/ADVEC/WEOWORLD"),
    ("International Monetary Fund. (2024). ",
     "World Economic Outlook Database — October 2024.",
     " IMF. https://www.imf.org/en/Publications/WEO"),
    ("Our World in Data. (2024). ",
     "Liberal political institutions index.",
     " Global Change Data Lab. https://ourworldindata.org/grapher/liberal-political-institutions-index"),
    ("United Nations Development Programme. (2023). ",
     "Human Development Report 2023/2024 — Japan.",
     " UNDP. https://hdr.undp.org/data-center/specific-country-data#/countries/JPN"),
    ("The World Bank. (2024). ",
     "World Bank Open Data — Japan.",
     " The World Bank Group. https://data.worldbank.org/country/japan"),
]

for before, italic_part, after in refs:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after  = Pt(0)
    pf.first_line_indent = Inches(-0.5)
    pf.left_indent = Inches(0.5)
    pf.line_spacing = Pt(24)
    r1 = p.add_run(before)
    set_font(r1, size=12)
    r2 = p.add_run(italic_part)
    set_font(r2, size=12, italic=True)
    r3 = p.add_run(after)
    set_font(r3, size=12)

# ── Save ─────────────────────────────────────────────────────────────────────
out = r"c:\Users\willf\OneDrive\Documents\Python Class\word\FinalProject_Paper_Japan.docx"
doc.save(out)
print(f"Saved: {out}")
