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
    "For this final project, I chose Japan as my country because it is one of the most interesting "
    "economies to look at right now. Japan is a rich country with a strong history of economic "
    "success, but it is also dealing with some serious problems that do not have easy answers. Its "
    "real GDP per capita is around $35,700 according to the IMF, and unemployment is only 2.5 "
    "percent, which sounds great on the surface. But at the same time, its population is shrinking, "
    "its government debt is the highest of any major economy in the world, and growth has basically "
    "stalled out. As IMF advisors, the goal of this report is to explain where Japan stands "
    "economically and give some recommendations for what it should do going forward."
)

# ── Section 1 ────────────────────────────────────────────────────────────────
add_heading(doc, "Economic Output and Growth", bold=True)

add_paragraph(doc,
    "Japan's real GDP per capita is around $35,700, which sounds like a lot, but when you compare "
    "it to the U.S. at $94,430, there is a pretty big gap (IMF, 2026). Part of the reason the U.S. "
    "is so much higher is because American workers tend to be more productive and the economy is "
    "much larger overall. Japan used to grow incredibly fast — in the decades after World War II, "
    "the country's economy grew at almost 10 percent per year, which economists call the 'Economic "
    "Miracle.' But since then, growth has slowed way down to basically 0 to 1 percent per year. "
    "A big part of the reason for this is that Japan's population is actually getting smaller, "
    "shrinking by about 0.5 percent per year because the birth rate is low and the country does "
    "not let many immigrants in. Fewer workers means less output, which makes it really hard to "
    "grow the economy."
)

add_paragraph(doc,
    "Japan's biggest export is cars. Companies like Toyota, Honda, and Nissan ship over $100 "
    "billion worth of vehicles every year to countries all over the world. Japan has a real "
    "advantage here because its workers are incredibly skilled at making high-quality products "
    "efficiently. This comes partly from a Japanese work philosophy called kaizen, which means "
    "always trying to improve. It is hard for other countries to compete with that level of "
    "precision and quality."
)

# ── Section 2 ────────────────────────────────────────────────────────────────
add_heading(doc, "Long-Run Growth: Capital, Institutions, and Human Capital", bold=True)

add_paragraph(doc,
    "One of the concepts we covered in class is the production function, and Japan is a good "
    "example of what happens when a country already has a lot of capital. Because Japan has been "
    "building up factories, machines, and technology for decades, adding more capital does not "
    "help as much anymore because of diminishing returns. To actually grow, Japan needs to shift "
    "the whole production function upward by getting better technology and improving how "
    "productively workers use capital. The good news is that Japan has strong institutions that "
    "make this possible — it scores about 0.73 out of 1.0 on the Liberal Political Institutions "
    "Index (Our World in Data, 2024), which means it has a stable government, courts that work, "
    "and property rights that businesses can count on. That kind of environment encourages "
    "investment and innovation."
)

add_paragraph(doc,
    "Japan also has a really educated and healthy population. Its Human Development Index score "
    "is 0.920, which puts it at 24th in the world (UNDP, 2023). People in Japan live to about "
    "84 years old on average, and the average person gets over 13 years of schooling. This is "
    "great for human capital. The problem is that this workforce is getting older and smaller "
    "every year because of the low birth rate. If Japan does not find a way to bring in more "
    "workers or make existing workers more productive, the economy will keep slowing down."
)

# ── Section 3 ────────────────────────────────────────────────────────────────
add_heading(doc, "Monetary Policy and Labor Markets", bold=True)

add_paragraph(doc,
    "Japan has an unemployment rate of only 2.5 percent, which is lower than the U.S. natural "
    "rate of about 4 percent. One reason for this is that Japanese companies traditionally offer "
    "lifetime employment, so workers do not change jobs as often. This cuts down on frictional "
    "unemployment, which is the kind that happens when people are between jobs. But even though "
    "unemployment is very low, Japan has been dealing with inflation of about 2.7 percent, which "
    "is slightly above the Bank of Japan's 2 percent target. Because of this, the central bank "
    "started raising interest rates — from basically zero in early 2024, to 0.25 percent in July "
    "2024, and then 0.5 percent in January 2025 (Bank of Japan, 2025). This is the right move "
    "because when inflation is above target and unemployment is already low, raising rates is the "
    "standard tool to cool the economy down a little. It is basically the same logic behind what "
    "the Federal Reserve does in the U.S."
)

# ── Section 4 ────────────────────────────────────────────────────────────────
add_heading(doc, "Fiscal Policy and Government Debt", bold=True)

add_paragraph(doc,
    "Japan's government spends a lot — about 38 to 40 percent of GDP — and still runs a deficit "
    "of around 5 to 6 percent every year (IMF, 2024). Because the economy is already running "
    "close to its potential output, this kind of deficit spending is not really needed and just "
    "adds to the debt. Japan's debt-to-GDP ratio is now around 255 percent, which is the highest "
    "of any major country in the world. That number sounds alarming, but the situation is a "
    "little less scary than it looks because about 90 percent of the debt is owned by Japanese "
    "citizens and institutions, not foreign lenders. Still, the problem is getting worse every "
    "year. As the population gets older, the government has to spend more on pensions and "
    "healthcare, but fewer people are working and paying taxes. Without some kind of reform, "
    "this is going to become a really serious problem."
)

# ── Section 5: Policy Recommendations ───────────────────────────────────────
add_heading(doc, "Policy Recommendations", bold=True)

add_paragraph(doc,
    "In the short term, Japan should keep raising interest rates slowly to get inflation back down "
    "to the 2 percent target. The Bank of Japan needs to be careful not to raise rates too fast "
    "though, because that could hurt growth. The government should also try to spend a little less "
    "money where it can so the deficit does not keep growing. Even small steps toward balancing "
    "the budget would help stop the debt from getting even worse."
)

add_paragraph(doc,
    "In the long run, the biggest thing Japan needs to fix is its population problem. The easiest "
    "way to do that would be to let more skilled immigrants into the country, which would add "
    "workers and taxpayers at the same time. Japan should also make its labor market more flexible "
    "— the lifetime employment system works in some ways, but it also makes it hard for companies "
    "to adapt and for workers to move into higher-productivity jobs. Getting more women into the "
    "workforce full-time would also help a lot. On top of that, Japan needs to keep investing in "
    "technology like AI and automation so that each worker can produce more output. This is the "
    "only real way to keep growing when you have fewer people. Finally, Japan has to reform its "
    "pension and healthcare systems to reduce long-term spending. None of these things are easy, "
    "but without them, Japan is going to keep falling further behind other advanced economies."
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
