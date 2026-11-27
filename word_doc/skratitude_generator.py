from pydoc import doc

import docx
from docx import section
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Mm, Pt, RGBColor

import word_doc.colors as COLORS
import word_doc.helpers as helpers


class JournalGenerator:
    def __init__(self):
        self.journal = docx.Document()
        self.setup_pages()
        self.setup_style()

    def setup_pages(self):
        # Page setup - A5 with balanced margins
        for section in self.journal.sections:
            section.page_width = Mm(148)
            section.page_height = Mm(210)
            section.top_margin = Mm(18)
            section.bottom_margin = Mm(18)
            section.left_margin = Mm(20)
            section.right_margin = Mm(20)

    def setup_style(self):
        # Set default Normal font
        style_normal = self.journal.styles["Normal"]
        style_normal.font.name = "Georgia"
        style_normal.font.size = Pt(10.5)
        style_normal.font.color.rgb = COLORS.CHARCOAL

    def add_cover_page(self):
        p_spacer = self.journal.add_paragraph()
        p_spacer.paragraph_format.space_before = Pt(50)

        # Decorative Top Accent Line
        p_top_line = self.journal.add_paragraph()
        p_top_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_line = p_top_line.add_run("—  ✦  —")
        r_line.font.name = "Georgia"
        r_line.font.size = Pt(14)
        r_line.font.color.rgb = RGBColor(197, 168, 128)

        # Main Title
        p_title = self.journal.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(18)
        p_title.paragraph_format.space_after = Pt(8)
        r_title = p_title.add_run("MINDFULNESS & INTENTION")
        r_title.font.name = "Georgia"
        r_title.font.size = Pt(28)
        r_title.font.bold = True
        r_title.font.color.rgb = COLORS.SAGE_DARK

        # Subtitle
        p_sub = self.journal.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(36)
        r_sub = p_sub.add_run("A Guided Journal for Daily Presence, Clarity, and Growth")
        r_sub.font.name = "Georgia"
        r_sub.font.size = Pt(13)
        r_sub.font.italic = True
        r_sub.font.color.rgb = COLORS.MUTED_GRAY

        # Quote Box on Cover
        quote_tbl = self.journal.add_table(rows=1, cols=1)
        quote_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        q_cell = quote_tbl.cell(0, 0)
        q_cell.width = Inches(5.4)
        helpers.set_cell_background(q_cell, COLORS.HEX_CREAM)
        helpers.set_cell_margins(q_cell, top=220, bottom=220, left=280, right=280)
        helpers.set_table_borders(quote_tbl, color=COLORS.HEX_GOLD_ACCENT, sz="6")

        qp = q_cell.paragraphs[0]
        qp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        qp.paragraph_format.space_after = Pt(6)
        qr = qp.add_run("“The present moment is the only time over which we have dominion.”")
        qr.font.name = "Georgia"
        qr.font.size = Pt(11.5)
        qr.font.italic = True
        qr.font.color.rgb = COLORS.CHARCOAL

        qp2 = q_cell.add_paragraph()
        qp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        qp2.paragraph_format.space_before = Pt(0)
        qp2.paragraph_format.space_after = Pt(0)
        qr2 = qp2.add_run("— Thích Nhất Hạnh")
        qr2.font.name = "Calibri"
        qr2.font.size = Pt(10)
        qr2.font.color.rgb = COLORS.SAGE_MED

        # Owner / Date Block at Bottom of Cover
        p_bot_spacer = self.journal.add_paragraph()
        p_bot_spacer.paragraph_format.space_before = Pt(130)

        meta_tbl = self.journal.add_table(rows=2, cols=2)
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
            helpers.set_cell_bottom_border_only(c1, color=COLORS.HEX_BORDER, sz="4")

            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            r0 = p0.add_run(label + "  ")
            r0.font.name = "Calibri"
            r0.font.size = Pt(9.5)
            r0.font.bold = True
            r0.font.color.rgb = COLORS.SAGE_DARK

            p1 = c1.paragraphs[0]
            r1 = p1.add_run(placeholder)
            r1.font.name = "Georgia"
            r1.font.size = Pt(10)
            r1.font.italic = True
            r1.font.color.rgb = COLORS.MUTED_GRAY

        self.journal.add_page_break()

    def add_goal_setting_page(self):
        # Header for Page 2
        p_h2 = self.journal.add_paragraph()
        p_h2.paragraph_format.space_after = Pt(4)
        r_h2 = p_h2.add_run("INTENTIONAL GOAL-SETTING & VISION")
        r_h2.font.name = "Georgia"
        r_h2.font.size = Pt(18)
        r_h2.font.bold = True
        r_h2.font.color.rgb = COLORS.SAGE_DARK

        p_intro = self.journal.add_paragraph()
        p_intro.paragraph_format.space_after = Pt(12)
        r_intro = p_intro.add_run(
            "Mindful goal-setting focuses less on rigid outcomes and more on how you wish to feel, show up, and grow. "
            "Use this customizable template to anchor your intentions before diving into your daily pages. "
            "Tip: Duplicate the Daily Practice pages as needed in Word."
        )
        r_intro.font.size = Pt(10)
        r_intro.font.color.rgb = COLORS.MUTED_GRAY

        # Core Values & Word of the Season Box
        val_tbl = self.journal.add_table(rows=1, cols=2)
        val_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(val_tbl, color=COLORS.HEX_BORDER, sz="4")

        c_left, c_right = val_tbl.rows[0].cells[0], val_tbl.rows[0].cells[1]
        c_left.width = Inches(3.4)
        c_right.width = Inches(3.4)
        helpers.set_cell_background(c_left, COLORS.HEX_CREAM)
        helpers.set_cell_background(c_right, COLORS.HEX_CREAM)
        helpers.set_cell_margins(c_left, top=140, bottom=140, left=160, right=160)
        helpers.set_cell_margins(c_right, top=140, bottom=140, left=160, right=160)

        pl = c_left.paragraphs[0]
        pl.paragraph_format.space_after = Pt(6)
        rl = pl.add_run("MY WORD OR THEME FOR THIS SEASON")
        rl.font.name = "Calibri"
        rl.font.size = Pt(10)
        rl.font.bold = True
        rl.font.color.rgb = COLORS.SAGE_DARK

        pl2 = c_left.add_paragraph()
        rl2 = pl2.add_run(
            "[e.g., Stillness, Balance, Courage, Presence]\n\nWhy this resonates with me right now:\n"
        )
        rl2.font.size = Pt(9.5)
        rl2.font.italic = True
        rl2.font.color.rgb = COLORS.MUTED_GRAY

        pr = c_right.paragraphs[0]
        pr.paragraph_format.space_after = Pt(6)
        rr = pr.add_run("CORE VALUES TO HONOR")
        rr.font.name = "Calibri"
        rr.font.size = Pt(10)
        rr.font.bold = True
        rr.font.color.rgb = COLORS.SAGE_DARK

        pr2 = c_right.add_paragraph()
        rr2 = pr2.add_run(
            "1. [Value 1 — e.g., Authenticity]\n2. [Value 2 — e.g., Compassion]\n3. [Value 3 — e.g., Curiosity]\n4. [Value 4 — e.g., Rest]"
        )
        rr2.font.size = Pt(9.5)
        rr2.font.color.rgb = COLORS.CHARCOAL

        # Mindful Goal Matrix Table
        helpers.add_section_heading(
            self.journal,
            "Mindful Goal Pillars",
            "Define 1–2 meaningful intentions across key areas of your life.",
        )

        goal_tbl = self.journal.add_table(rows=5, cols=3)
        goal_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(goal_tbl, color=COLORS.HEX_BORDER, sz="4")

        headers = ["LIFE PILLAR", "MY INTENTION / GOAL", "SMALL MINDFUL STEP THIS WEEK"]
        col_widths = [Inches(1.6), Inches(2.7), Inches(2.5)]

        hdr_row = goal_tbl.rows[0]
        hdr_row.height = Pt(26)
        for i, h_text in enumerate(headers):
            cell = hdr_row.cells[i]
            cell.width = col_widths[i]
            helpers.set_cell_background(cell, COLORS.HEX_SAGE_LIGHT)
            helpers.set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(h_text)
            r.font.name = "Calibri"
            r.font.size = Pt(9.5)
            r.font.bold = True
            r.font.color.rgb = COLORS.SAGE_DARK

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
                helpers.set_cell_margins(cell, top=90, bottom=90, left=100, right=100)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(val)
                if c_idx == 0:
                    r.font.name = "Calibri"
                    r.font.size = Pt(10)
                    r.font.bold = True
                    r.font.color.rgb = COLORS.CHARCOAL
                else:
                    r.font.name = "Georgia"
                    r.font.size = Pt(9.5)
                    r.font.italic = True
                    r.font.color.rgb = COLORS.MUTED_GRAY

        # Daily Affirmation Bank
        helpers.add_section_heading(
            self.journal,
            "Personal Affirmation Bank",
            "Keep a rotating list of grounded affirmations you can copy onto your daily pages.",
        )

        aff_tbl = self.journal.add_table(rows=1, cols=1)
        aff_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(aff_tbl, color=COLORS.HEX_BORDER, sz="4")
        ac = aff_tbl.cell(0, 0)
        ac.width = Inches(6.8)
        helpers.set_cell_background(ac, COLORS.HEX_CREAM)
        helpers.set_cell_margins(ac, top=120, bottom=120, left=160, right=160)

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
            r.font.color.rgb = COLORS.MUTED_GRAY if "[" in aff else COLORS.CHARCOAL

        self.journal.add_page_break()

    def add_mindfulness_page(self):
        # Top Date & Mood Bar
        top_bar = self.journal.add_table(rows=1, cols=2)
        top_bar.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(top_bar, color=COLORS.HEX_BORDER, sz="4")

        tb_left, tb_right = top_bar.rows[0].cells[0], top_bar.rows[0].cells[1]
        tb_left.width = Inches(3.4)
        tb_right.width = Inches(3.4)
        helpers.set_cell_background(tb_left, COLORS.HEX_SAGE_LIGHT)
        helpers.set_cell_background(tb_right, COLORS.HEX_SAGE_LIGHT)
        helpers.set_cell_margins(tb_left, top=100, bottom=100, left=140, right=140)
        helpers.set_cell_margins(tb_right, top=100, bottom=100, left=140, right=140)

        p_date = tb_left.paragraphs[0]
        p_date.paragraph_format.space_after = Pt(0)
        r_d1 = p_date.add_run("DATE: ")
        r_d1.font.name = "Calibri"
        r_d1.font.bold = True
        r_d1.font.size = Pt(10.5)
        r_d1.font.color.rgb = COLORS.SAGE_DARK
        r_d2 = p_date.add_run("[Day, Month DD, YYYY]")
        r_d2.font.size = Pt(10)
        r_d2.font.color.rgb = COLORS.CHARCOAL

        p_mood = tb_right.paragraphs[0]
        p_mood.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_mood.paragraph_format.space_after = Pt(0)
        r_m1 = p_mood.add_run("ENERGY / MOOD: ")
        r_m1.font.name = "Calibri"
        r_m1.font.bold = True
        r_m1.font.size = Pt(10.5)
        r_m1.font.color.rgb = COLORS.SAGE_DARK
        r_m2 = p_mood.add_run("☐ Calm   ☐ Energized   ☐ Tired   ☐ Anxious   ☐ Grateful")
        r_m2.font.name = "Calibri"
        r_m2.font.size = Pt(9.5)
        r_m2.font.color.rgb = COLORS.CHARCOAL

        # MORNING SECTION
        p_morn = self.journal.add_paragraph()
        p_morn.paragraph_format.space_before = Pt(14)
        p_morn.paragraph_format.space_after = Pt(6)
        r_morn = p_morn.add_run("☀  MORNING GROUNDING & INTENTION")
        r_morn.font.name = "Georgia"
        r_morn.font.size = Pt(13)
        r_morn.font.bold = True
        r_morn.font.color.rgb = COLORS.SAGE_DARK

        # Daily Affirmation Callout
        aff_box = self.journal.add_table(rows=1, cols=1)
        aff_box.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(aff_box, color=COLORS.HEX_GOLD_ACCENT, sz="6")
        abc = aff_box.cell(0, 0)
        abc.width = Inches(6.8)
        helpers.set_cell_background(abc, COLORS.HEX_CREAM)
        helpers.set_cell_margins(abc, top=100, bottom=100, left=140, right=140)

        p_ab1 = abc.paragraphs[0]
        p_ab1.paragraph_format.space_after = Pt(3)
        r_ab1 = p_ab1.add_run("TODAY'S AFFIRMATION")
        r_ab1.font.name = "Calibri"
        r_ab1.font.size = Pt(9.5)
        r_ab1.font.bold = True
        r_ab1.font.color.rgb = COLORS.SAGE_DARK

        p_ab2 = abc.add_paragraph()
        p_ab2.paragraph_format.space_after = Pt(0)
        r_ab2 = p_ab2.add_run(
            "“I am grounded in the present moment, and I trust myself to navigate today with clarity and ease.”"
        )
        r_ab2.font.size = Pt(10.5)
        r_ab2.font.italic = True
        r_ab2.font.color.rgb = COLORS.CHARCOAL

        # Morning Prompts
        helpers.add_section_heading(self.journal, "Three Things I Am Grateful For Today")
        helpers.add_writing_lines(self.journal, num_lines=3, line_spacing_pt=22)

        helpers.add_section_heading(
            self.journal,
            "Today's Mindful Focus & Top Priorities",
            "What would make today feel purposeful and aligned?",
        )
        prio_tbl = self.journal.add_table(rows=3, cols=2)
        prio_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(prio_tbl, color=COLORS.HEX_BORDER, sz="4")
        for i in range(3):
            row = prio_tbl.rows[i]
            row.height = Pt(24)
            c0, c1 = row.cells[0], row.cells[1]
            c0.width = Inches(0.45)
            c1.width = Inches(6.35)
            helpers.set_cell_margins(c0, top=40, bottom=40, left=60, right=40)
            helpers.set_cell_margins(c1, top=40, bottom=40, left=80, right=80)
            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p0.paragraph_format.space_after = Pt(0)
            r0 = p0.add_run("☐")
            r0.font.size = Pt(12)
            r0.font.color.rgb = COLORS.SAGE_MED

            p1 = c1.paragraphs[0]
            p1.paragraph_format.space_after = Pt(0)
            r1 = p1.add_run(f"Priority {i+1}: " if i == 0 else "")
            r1.font.size = Pt(9.5)
            r1.font.italic = True
            r1.font.color.rgb = COLORS.MUTED_GRAY

        # EVENING SECTION
        p_eve = self.journal.add_paragraph()
        p_eve.paragraph_format.space_before = Pt(18)
        p_eve.paragraph_format.space_after = Pt(6)
        r_eve = p_eve.add_run("☾  EVENING REFLECTION & RELEASE")
        r_eve.font.name = "Georgia"
        r_eve.font.size = Pt(13)
        r_eve.font.bold = True
        r_eve.font.color.rgb = COLORS.SAGE_DARK

        helpers.add_section_heading(
            self.journal,
            "Moments of Presence & Joy",
            "When did I feel most connected, peaceful, or engaged today?",
        )
        helpers.add_writing_lines(self.journal, num_lines=3, line_spacing_pt=22)

        helpers.add_section_heading(
            self.journal,
            "Gentle Self-Inquiry",
            "What challenged me today, and what can I release before rest?",
        )
        helpers.add_writing_lines(self.journal, num_lines=3, line_spacing_pt=22)

        # Mindful Habit Check-in Footer Table
        p_hab_lbl = self.journal.add_paragraph()
        p_hab_lbl.paragraph_format.space_before = Pt(12)
        p_hab_lbl.paragraph_format.space_after = Pt(4)
        r_hl = p_hab_lbl.add_run("DAILY MINDFUL RITUALS CHECK-IN")
        r_hl.font.name = "Calibri"
        r_hl.font.size = Pt(9.5)
        r_hl.font.bold = True
        r_hl.font.color.rgb = COLORS.SAGE_DARK

        hab_tbl = self.journal.add_table(rows=1, cols=4)
        hab_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(hab_tbl, color=COLORS.HEX_BORDER, sz="4")
        habits = [
            "☐  Stillness / Breathwork",
            "☐  Mindful Movement",
            "☐  Screen-Free Break",
            "☐  Restful Wind-Down",
        ]
        for idx, h in enumerate(habits):
            c = hab_tbl.rows[0].cells[idx]
            c.width = Inches(1.7)
            helpers.set_cell_background(c, COLORS.HEX_CREAM)
            helpers.set_cell_margins(c, top=70, bottom=70, left=80, right=80)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(h)
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            r.font.color.rgb = COLORS.CHARCOAL

        self.journal.add_page_break()

    def add_weekly_reflection_page(self):
        p_w_title = self.journal.add_paragraph()
        p_w_title.paragraph_format.space_after = Pt(2)
        r_wt = p_w_title.add_run("WEEKLY REFLECTION & ALIGNMENT")
        r_wt.font.name = "Georgia"
        r_wt.font.size = Pt(18)
        r_wt.font.bold = True
        r_wt.font.color.rgb = COLORS.SAGE_DARK

        p_w_sub = self.journal.add_paragraph()
        p_w_sub.paragraph_format.space_after = Pt(10)
        r_ws = p_w_sub.add_run(
            "WEEK OF: [Month DD – Month DD, YYYY]     |     OVERALL PEACE RATING:  1   2   3   4   5   6   7   8   9   10"
        )
        r_ws.font.name = "Calibri"
        r_ws.font.size = Pt(10)
        r_ws.font.bold = True
        r_ws.font.color.rgb = COLORS.MUTED_GRAY

        # Two-column reflection grid
        wk_grid = self.journal.add_table(rows=2, cols=2)
        wk_grid.alignment = WD_TABLE_ALIGNMENT.CENTER
        helpers.set_table_borders(wk_grid, color=COLORS.HEX_BORDER, sz="4")

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
                helpers.set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
                if (r + c) % 2 == 0:
                    helpers.set_cell_background(cell, COLORS.HEX_CREAM)
                title_txt, desc_txt = wk_prompts[idx]
                idx += 1

                p1 = cell.paragraphs[0]
                p1.paragraph_format.space_after = Pt(3)
                r1 = p1.add_run(title_txt)
                r1.font.name = "Calibri"
                r1.font.size = Pt(10)
                r1.font.bold = True
                r1.font.color.rgb = COLORS.SAGE_DARK

                p2 = cell.add_paragraph()
                p2.paragraph_format.space_after = Pt(8)
                r2 = p2.add_run(desc_txt)
                r2.font.size = Pt(9)
                r2.font.italic = True
                r2.font.color.rgb = COLORS.MUTED_GRAY


if __name__ == "__main__":

    journal = JournalGenerator()

    # ==============================================================================
    # PAGE 1: COVER PAGE
    # ==============================================================================
    journal.add_cover_page()

    # ==============================================================================
    # PAGE 2: HOW TO USE & INTENTIONAL GOAL-SETTING
    # ==============================================================================
    journal.add_goal_setting_page()

    # ==============================================================================
    # PAGE 3: DAILY MINDFULNESS PRACTICE (MORNING & EVENING TEMPLATE)
    # ==============================================================================
    journal.add_mindfulness_page()

    # ==============================================================================
    # PAGE 4: WEEKLY MINDFULNESS CHECK-IN & REFLECTION
    # ==============================================================================
    journal.add_weekly_reflection_page()

    # Open Free-Writing Space
    helpers.add_section_heading(
        journal.journal,
        "Open Mindful Space",
        "Unstructured room for stream-of-consciousness thoughts, sketches, or notes.",
    )
    helpers.add_writing_lines(journal.journal, num_lines=7, line_spacing_pt=24)

    output_filename = "Mindfulness_Journal_Template.docx"
    journal.journal.save(output_filename)
    print(f"Successfully generated {output_filename}")
