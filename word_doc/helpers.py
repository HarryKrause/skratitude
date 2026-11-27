"""
Helper functions for XML styling
"""

from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt

import word_doc.colors as COLORS


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


def set_table_borders(table, color=COLORS.HEX_BORDER, sz="4", val="single"):
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


def set_cell_bottom_border_only(cell, color=COLORS.HEX_BORDER, sz="4"):
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
    run.font.color.rgb = COLORS.SAGE_DARK

    # Add subtle bottom border to heading paragraph
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f"""
        <w:pBdr {nsdecls("w")}>
            <w:bottom w:val="single" w:sz="6" w:space="3" w:color="{COLORS.HEX_GOLD_ACCENT}"/>
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
        r_sub.font.color.rgb = COLORS.MUTED_GRAY


def add_writing_lines(container, num_lines=3, line_spacing_pt=24):
    tbl = container.add_table(rows=num_lines, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for row in tbl.rows:
        row.height = Pt(line_spacing_pt)
        cell = row.cells[0]
        cell.width = Inches(6.8)
        set_cell_bottom_border_only(cell, color=COLORS.HEX_BORDER, sz="4")
        set_cell_margins(cell, top=40, bottom=20, left=40, right=40)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run("")
        r.font.size = Pt(10)
