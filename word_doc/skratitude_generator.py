import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

doc = docx.Document()

# Page setup - Standard Letter with clean margins
for section in doc.sections:
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

# Color Palette: Calming Sage, Warm Sand, Deep Charcoal, Soft Muted Teal
SAGE_DARK = RGBColor(58, 90, 64)  # #3A5A40 - Deep forest/sage
SAGE_MED = RGBColor(88, 129, 87)  # #588157 - Muted leaf green
CHARCOAL = RGBColor(45, 49, 46)  # #2D312E - Soft dark text
MUTED_GRAY = RGBColor(115, 122, 117)  # #737A75 - Subtitles / prompts
HEX_CREAM = "F6F5F0"  # Warm paper tint for callout boxes
HEX_SAGE_LIGHT = "E9EFEA"  # Soft sage tint for headers
HEX_BORDER = "C7D1C9"  # Subtle sage-gray border
HEX_GOLD_ACCENT = "C5A880"  # Warm sand/gold accent

# Set default Normal font
style_normal = doc.styles["Normal"]
style_normal.font.name = "Georgia"
style_normal.font.size = Pt(10.5)
style_normal.font.color.rgb = CHARCOAL


# Helper functions for XML styling
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_table_borders(table, color=HEX_BORDER, sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f"""
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    """
    )
    tblPr.append(borders)


def set_cell_bottom_border_only(cell, color=HEX_BORDER, sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f"""
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
            <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tcBorders>
    """
    )
    tcPr.append(borders)


def add_section_heading(doc, text, subtitle=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(2 if subtitle else 6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text.upper())
    run.font.name = "Calibri"
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = SAGE_DARK

    # Add subtle bottom border to heading paragraph
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f"""
        <w:pBdr {nsdecls("w")}>
            <w:bottom w:val="single" w:sz="6" w:space="3" w:color="{HEX_GOLD_ACCENT}"/>
        </w:pBdr>
    """
    )
    pPr.append(pBdr)

    if subtitle:
        sub = doc.add_paragraph()
        sub.paragraph_format.space_before = Pt(2)
        sub.paragraph_format.space_after = Pt(6)
        sub.paragraph_format.keep_with_next = True
        r_sub = sub.add_run(subtitle)
        r_sub.font.name = "Georgia"
        r_sub.font.size = Pt(9.5)
        r_sub.font.italic = True
        r_sub.font.color.rgb = MUTED_GRAY


def add_writing_lines(container, num_lines=3, line_spacing_pt=24):
    tbl = container.add_table(rows=num_lines, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for row in tbl.rows:
        row.height = Pt(line_spacing_pt)
        cell = row.cells[0]
        cell.width = Inches(6.8)
        set_cell_bottom_border_only(cell, color=HEX_BORDER, sz="4")
        set_cell_margins(cell, top=40, bottom=20, left=40, right=40)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("")
        r.font.size = Pt(10)


# ==============================================================================
# PAGE 1: COVER PAGE
# ==============================================================================
p_spacer = doc.add_paragraph()
p_spacer.paragraph_format.space_before = Pt(90)

# Decorative Top Accent Line
p_top_line = doc.add_paragraph()
p_top_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_line = p_top_line.add_run("—  ✦  —")
r_line.font.name = "Georgia"
r_line.font.size = Pt(14)
r_line.font.color.rgb = RGBColor(197, 168, 128)

# Main Title
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.paragraph_format.space_before = Pt(18)
p_title.paragraph_format.space_after = Pt(8)
r_title = p_title.add_run("MINDFULNESS & INTENTION")
r_title.font.name = "Georgia"
r_title.font.size = Pt(28)
r_title.font.bold = True
r_title.font.color.rgb = SAGE_DARK

# Subtitle
p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_before = Pt(0)
p_sub.paragraph_format.space_after = Pt(36)
r_sub = p_sub.add_run("A Guided Journal for Daily Presence, Clarity, and Growth")
r_sub.font.name = "Georgia"
r_sub.font.size = Pt(13)
r_sub.font.italic = True
r_sub.font.color.rgb = MUTED_GRAY

# Quote Box on Cover
quote_tbl = doc.add_table(rows=1, cols=1)
quote_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
q_cell = quote_tbl.cell(0, 0)
q_cell.width = Inches(5.4)
set_cell_background(q_cell, HEX_CREAM)
set_cell_margins(q_cell, top=220, bottom=220, left=280, right=280)
set_table_borders(quote_tbl, color=HEX_GOLD_ACCENT, sz="6")

qp = q_cell.paragraphs[0]
qp.alignment = WD_ALIGN_PARAGRAPH.CENTER
qp.paragraph_format.space_after = Pt(6)
qr = qp.add_run("“The present moment is the only time over which we have dominion.”")
qr.font.name = "Georgia"
qr.font.size = Pt(11.5)
qr.font.italic = True
qr.font.color.rgb = CHARCOAL

qp2 = q_cell.add_paragraph()
qp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
qp2.paragraph_format.space_before = Pt(0)
qp2.paragraph_format.space_after = Pt(0)
qr2 = qp2.add_run("— Thích Nhất Hạnh")
qr2.font.name = "Calibri"
qr2.font.size = Pt(10)
qr2.font.color.rgb = SAGE_MED

# Owner / Date Block at Bottom of Cover
p_bot_spacer = doc.add_paragraph()
p_bot_spacer.paragraph_format.space_before = Pt(130)

meta_tbl = doc.add_table(rows=2, cols=2)
meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
meta_fields = [
    ("THIS JOURNAL BELONGS TO:", "[Enter Your Name]"),
    ("DATE STARTED:", "[Month / Year]"),
]
for idx, (label, placeholder) in enumerate(meta_fields):
    row = meta_tbl.rows[idx]
    row.height = Pt(28)
    c0, c1 = row.cells[0], row.cells[1]
    c0.width = Inches(2.4)
    c1.width = Inches(3.2)
    set_cell_bottom_border_only(c1, color=HEX_BORDER, sz="4")

    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r0 = p0.add_run(label + "  ")
    r0.font.name = "Calibri"
    r0.font.size = Pt(9.5)
    r0.font.bold = True
    r0.font.color.rgb = SAGE_DARK

    p1 = c1.paragraphs[0]
    r1 = p1.add_run(placeholder)
    r1.font.name = "Georgia"
    r1.font.size = Pt(10)
    r1.font.italic = True
    r1.font.color.rgb = MUTED_GRAY

doc.add_page_break()

# ==============================================================================
# PAGE 2: HOW TO USE & INTENTIONAL GOAL-SETTING
# ==============================================================================
# Header for Page 2
p_h2 = doc.add_paragraph()
p_h2.paragraph_format.space_after = Pt(4)
r_h2 = p_h2.add_run("INTENTIONAL GOAL-SETTING & VISION")
r_h2.font.name = "Georgia"
r_h2.font.size = Pt(18)
r_h2.font.bold = True
r_h2.font.color.rgb = SAGE_DARK

p_intro = doc.add_paragraph()
p_intro.paragraph_format.space_after = Pt(12)
r_intro = p_intro.add_run(
    "Mindful goal-setting focuses less on rigid outcomes and more on how you wish to feel, show up, and grow. "
    "Use this customizable template to anchor your intentions before diving into your daily pages. "
    "Tip: Duplicate the Daily Practice pages as needed in Word."
)
r_intro.font.size = Pt(10)
r_intro.font.color.rgb = MUTED_GRAY

# Core Values & Word of the Season Box
val_tbl = doc.add_table(rows=1, cols=2)
val_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(val_tbl, color=HEX_BORDER, sz="4")

c_left, c_right = val_tbl.rows[0].cells[0], val_tbl.rows[0].cells[1]
c_left.width = Inches(3.4)
c_right.width = Inches(3.4)
set_cell_background(c_left, HEX_CREAM)
set_cell_background(c_right, HEX_CREAM)
set_cell_margins(c_left, top=140, bottom=140, left=160, right=160)
set_cell_margins(c_right, top=140, bottom=140, left=160, right=160)

pl = c_left.paragraphs[0]
pl.paragraph_format.space_after = Pt(6)
rl = pl.add_run("MY WORD OR THEME FOR THIS SEASON")
rl.font.name = "Calibri"
rl.font.size = Pt(10)
rl.font.bold = True
rl.font.color.rgb = SAGE_DARK

pl2 = c_left.add_paragraph()
rl2 = pl2.add_run(
    "[e.g., Stillness, Balance, Courage, Presence]\n\nWhy this resonates with me right now:\n"
)
rl2.font.size = Pt(9.5)
rl2.font.italic = True
rl2.font.color.rgb = MUTED_GRAY

pr = c_right.paragraphs[0]
pr.paragraph_format.space_after = Pt(6)
rr = pr.add_run("CORE VALUES TO HONOR")
rr.font.name = "Calibri"
rr.font.size = Pt(10)
rr.font.bold = True
rr.font.color.rgb = SAGE_DARK

pr2 = c_right.add_paragraph()
rr2 = pr2.add_run(
    "1. [Value 1 — e.g., Authenticity]\n2. [Value 2 — e.g., Compassion]\n3. [Value 3 — e.g., Curiosity]\n4. [Value 4 — e.g., Rest]"
)
rr2.font.size = Pt(9.5)
rr2.font.color.rgb = CHARCOAL

# Mindful Goal Matrix Table
add_section_heading(
    doc, "Mindful Goal Pillars", "Define 1–2 meaningful intentions across key areas of your life."
)

goal_tbl = doc.add_table(rows=5, cols=3)
goal_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(goal_tbl, color=HEX_BORDER, sz="4")

headers = ["LIFE PILLAR", "MY INTENTION / GOAL", "SMALL MINDFUL STEP THIS WEEK"]
col_widths = [Inches(1.6), Inches(2.7), Inches(2.5)]

hdr_row = goal_tbl.rows[0]
hdr_row.height = Pt(26)
for i, h_text in enumerate(headers):
    cell = hdr_row.cells[i]
    cell.width = col_widths[i]
    set_cell_background(cell, HEX_SAGE_LIGHT)
    set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(h_text)
    r.font.name = "Calibri"
    r.font.size = Pt(9.5)
    r.font.bold = True
    r.font.color.rgb = SAGE_DARK

pillars = [
    (
        "Mind & Emotional Well-being",
        "[e.g., Cultivate calm mornings without screens]",
        "[e.g., 10 mins of quiet breathing before checking phone]",
    ),
    (
        "Body & Physical Vitality",
        "[e.g., Move with joy and listen to energy cues]",
        "[e.g., Take a 20-minute outdoor walk 4x this week]",
    ),
    (
        "Relationships & Connection",
        "[e.g., Practice deep, distraction-free listening]",
        "[e.g., Put devices away during dinner conversations]",
    ),
    (
        "Work, Purpose & Creativity",
        "[e.g., Focus on one meaningful task at a time]",
        "[e.g., Block 60 mins of uninterrupted deep work daily]",
    ),
]

for r_idx, (pillar, goal_ex, step_ex) in enumerate(pillars, start=1):
    row = goal_tbl.rows[r_idx]
    row.height = Pt(54)
    vals = [pillar, goal_ex, step_ex]
    for c_idx, val in enumerate(vals):
        cell = row.cells[c_idx]
        cell.width = col_widths[c_idx]
        set_cell_margins(cell, top=90, bottom=90, left=100, right=100)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(val)
        if c_idx == 0:
            r.font.name = "Calibri"
            r.font.size = Pt(10)
            r.font.bold = True
            r.font.color.rgb = CHARCOAL
        else:
            r.font.name = "Georgia"
            r.font.size = Pt(9.5)
            r.font.italic = True
            r.font.color.rgb = MUTED_GRAY

# Daily Affirmation Bank
add_section_heading(
    doc,
    "Personal Affirmation Bank",
    "Keep a rotating list of grounded affirmations you can copy onto your daily pages.",
)

aff_tbl = doc.add_table(rows=1, cols=1)
aff_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(aff_tbl, color=HEX_BORDER, sz="4")
ac = aff_tbl.cell(0, 0)
ac.width = Inches(6.8)
set_cell_background(ac, HEX_CREAM)
set_cell_margins(ac, top=120, bottom=120, left=160, right=160)

affirmations_list = [
    "•  I release the need to rush; everything meant for me unfolds in its own time.",
    "•  I am allowed to take up space, rest without guilt, and honor my boundaries.",
    "•  I meet this moment with curiosity rather than judgment.",
    "•  [Add your custom affirmation here...]",
    "•  [Add your custom affirmation here...]",
]
for idx, aff in enumerate(affirmations_list):
    p = ac.paragraphs[0] if idx == 0 else ac.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(aff)
    r.font.size = Pt(10)
    r.font.italic = True if "[" in aff else False
    r.font.color.rgb = MUTED_GRAY if "[" in aff else CHARCOAL

doc.add_page_break()

# ==============================================================================
# PAGE 3: DAILY MINDFULNESS PRACTICE (MORNING & EVENING TEMPLATE)
# ==============================================================================
# Top Date & Mood Bar
top_bar = doc.add_table(rows=1, cols=2)
top_bar.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(top_bar, color=HEX_BORDER, sz="4")

tb_left, tb_right = top_bar.rows[0].cells[0], top_bar.rows[0].cells[1]
tb_left.width = Inches(3.4)
tb_right.width = Inches(3.4)
set_cell_background(tb_left, HEX_SAGE_LIGHT)
set_cell_background(tb_right, HEX_SAGE_LIGHT)
set_cell_margins(tb_left, top=100, bottom=100, left=140, right=140)
set_cell_margins(tb_right, top=100, bottom=100, left=140, right=140)

p_date = tb_left.paragraphs[0]
p_date.paragraph_format.space_after = Pt(0)
r_d1 = p_date.add_run("DATE: ")
r_d1.font.name = "Calibri"
r_d1.font.bold = True
r_d1.font.size = Pt(10.5)
r_d1.font.color.rgb = SAGE_DARK
r_d2 = p_date.add_run("[Day, Month DD, YYYY]")
r_d2.font.size = Pt(10)
r_d2.font.color.rgb = CHARCOAL

p_mood = tb_right.paragraphs[0]
p_mood.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p_mood.paragraph_format.space_after = Pt(0)
r_m1 = p_mood.add_run("ENERGY / MOOD: ")
r_m1.font.name = "Calibri"
r_m1.font.bold = True
r_m1.font.size = Pt(10.5)
r_m1.font.color.rgb = SAGE_DARK
r_m2 = p_mood.add_run("☐ Calm   ☐ Energized   ☐ Tired   ☐ Anxious   ☐ Grateful")
r_m2.font.name = "Calibri"
r_m2.font.size = Pt(9.5)
r_m2.font.color.rgb = CHARCOAL

# MORNING SECTION
p_morn = doc.add_paragraph()
p_morn.paragraph_format.space_before = Pt(14)
p_morn.paragraph_format.space_after = Pt(6)
r_morn = p_morn.add_run("☀  MORNING GROUNDING & INTENTION")
r_morn.font.name = "Georgia"
r_morn.font.size = Pt(13)
r_morn.font.bold = True
r_morn.font.color.rgb = SAGE_DARK

# Daily Affirmation Callout
aff_box = doc.add_table(rows=1, cols=1)
aff_box.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(aff_box, color=HEX_GOLD_ACCENT, sz="6")
abc = aff_box.cell(0, 0)
abc.width = Inches(6.8)
set_cell_background(abc, HEX_CREAM)
set_cell_margins(abc, top=100, bottom=100, left=140, right=140)

p_ab1 = abc.paragraphs[0]
p_ab1.paragraph_format.space_after = Pt(3)
r_ab1 = p_ab1.add_run("TODAY'S AFFIRMATION")
r_ab1.font.name = "Calibri"
r_ab1.font.size = Pt(9.5)
r_ab1.font.bold = True
r_ab1.font.color.rgb = SAGE_DARK

p_ab2 = abc.add_paragraph()
p_ab2.paragraph_format.space_after = Pt(0)
r_ab2 = p_ab2.add_run(
    "“I am grounded in the present moment, and I trust myself to navigate today with clarity and ease.”"
)
r_ab2.font.size = Pt(10.5)
r_ab2.font.italic = True
r_ab2.font.color.rgb = CHARCOAL

# Morning Prompts
add_section_heading(doc, "Three Things I Am Grateful For Today")
add_writing_lines(doc, num_lines=3, line_spacing_pt=22)

add_section_heading(
    doc,
    "Today's Mindful Focus & Top Priorities",
    "What would make today feel purposeful and aligned?",
)
prio_tbl = doc.add_table(rows=3, cols=2)
prio_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(prio_tbl, color=HEX_BORDER, sz="4")
for i in range(3):
    row = prio_tbl.rows[i]
    row.height = Pt(24)
    c0, c1 = row.cells[0], row.cells[1]
    c0.width = Inches(0.45)
    c1.width = Inches(6.35)
    set_cell_margins(c0, top=40, bottom=40, left=60, right=40)
    set_cell_margins(c1, top=40, bottom=40, left=80, right=80)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_after = Pt(0)
    r0 = p0.add_run("☐")
    r0.font.size = Pt(12)
    r0.font.color.rgb = SAGE_MED

    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run(f"Priority {i+1}: " if i == 0 else "")
    r1.font.size = Pt(9.5)
    r1.font.italic = True
    r1.font.color.rgb = MUTED_GRAY

# EVENING SECTION
p_eve = doc.add_paragraph()
p_eve.paragraph_format.space_before = Pt(18)
p_eve.paragraph_format.space_after = Pt(6)
r_eve = p_eve.add_run("☾  EVENING REFLECTION & RELEASE")
r_eve.font.name = "Georgia"
r_eve.font.size = Pt(13)
r_eve.font.bold = True
r_eve.font.color.rgb = SAGE_DARK

add_section_heading(
    doc, "Moments of Presence & Joy", "When did I feel most connected, peaceful, or engaged today?"
)
add_writing_lines(doc, num_lines=3, line_spacing_pt=22)

add_section_heading(
    doc, "Gentle Self-Inquiry", "What challenged me today, and what can I release before rest?"
)
add_writing_lines(doc, num_lines=3, line_spacing_pt=22)

# Mindful Habit Check-in Footer Table
p_hab_lbl = doc.add_paragraph()
p_hab_lbl.paragraph_format.space_before = Pt(12)
p_hab_lbl.paragraph_format.space_after = Pt(4)
r_hl = p_hab_lbl.add_run("DAILY MINDFUL RITUALS CHECK-IN")
r_hl.font.name = "Calibri"
r_hl.font.size = Pt(9.5)
r_hl.font.bold = True
r_hl.font.color.rgb = SAGE_DARK

hab_tbl = doc.add_table(rows=1, cols=4)
hab_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(hab_tbl, color=HEX_BORDER, sz="4")
habits = [
    "☐  Stillness / Breathwork",
    "☐  Mindful Movement",
    "☐  Screen-Free Break",
    "☐  Restful Wind-Down",
]
for idx, h in enumerate(habits):
    c = hab_tbl.rows[0].cells[idx]
    c.width = Inches(1.7)
    set_cell_background(c, HEX_CREAM)
    set_cell_margins(c, top=70, bottom=70, left=80, right=80)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(h)
    r.font.name = "Calibri"
    r.font.size = Pt(9)
    r.font.color.rgb = CHARCOAL

doc.add_page_break()

# ==============================================================================
# PAGE 4: WEEKLY MINDFULNESS CHECK-IN & REFLECTION
# ==============================================================================
p_w_title = doc.add_paragraph()
p_w_title.paragraph_format.space_after = Pt(2)
r_wt = p_w_title.add_run("WEEKLY REFLECTION & ALIGNMENT")
r_wt.font.name = "Georgia"
r_wt.font.size = Pt(18)
r_wt.font.bold = True
r_wt.font.color.rgb = SAGE_DARK

p_w_sub = doc.add_paragraph()
p_w_sub.paragraph_format.space_after = Pt(10)
r_ws = p_w_sub.add_run(
    "WEEK OF: [Month DD – Month DD, YYYY]     |     OVERALL PEACE RATING:  1   2   3   4   5   6   7   8   9   10"
)
r_ws.font.name = "Calibri"
r_ws.font.size = Pt(10)
r_ws.font.bold = True
r_ws.font.color.rgb = MUTED_GRAY

# Two-column reflection grid
wk_grid = doc.add_table(rows=2, cols=2)
wk_grid.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(wk_grid, color=HEX_BORDER, sz="4")

wk_prompts = [
    (
        "WINS & MOMENTS OF GRATITUDE",
        "What went well this week? Celebrate small acts of presence and progress.",
    ),
    (
        "ENERGY DRAINS & BOUNDARIES",
        "What depleted my energy, and how can I protect my peace next week?",
    ),
    (
        "LESSONS & PATTERNS NOTICED",
        "What did my thoughts, habits, or emotions teach me over the past seven days?",
    ),
    (
        "INTENTION FOR THE WEEK AHEAD",
        "How do I want to feel next week, and what is one supportive habit I will prioritize?",
    ),
]

idx = 0
for r in range(2):
    row = wk_grid.rows[r]
    row.height = Inches(2.5)
    for c in range(2):
        cell = row.cells[c]
        cell.width = Inches(3.4)
        set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
        if (r + c) % 2 == 0:
            set_cell_background(cell, HEX_CREAM)
        title_txt, desc_txt = wk_prompts[idx]
        idx += 1

        p1 = cell.paragraphs[0]
        p1.paragraph_format.space_after = Pt(3)
        r1 = p1.add_run(title_txt)
        r1.font.name = "Calibri"
        r1.font.size = Pt(10)
        r1.font.bold = True
        r1.font.color.rgb = SAGE_DARK

        p2 = cell.add_paragraph()
        p2.paragraph_format.space_after = Pt(8)
        r2 = p2.add_run(desc_txt)
        r2.font.size = Pt(9)
        r2.font.italic = True
        r2.font.color.rgb = MUTED_GRAY

# Open Free-Writing Space
add_section_heading(
    doc,
    "Open Mindful Space",
    "Unstructured room for stream-of-consciousness thoughts, sketches, or notes.",
)
add_writing_lines(doc, num_lines=7, line_spacing_pt=24)

output_filename = "Mindfulness_Journal_Template.docx"
doc.save(output_filename)
print(f"Successfully generated {output_filename}")
