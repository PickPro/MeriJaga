#!/usr/bin/env python3
"""Safely update ONLY the existing Slide 2 native credit-risk table.

Supports:
1. In-place update of original Slide 2 (default)
2. Dedicated Slide 3 (--new-table-slide)
3. Same-Slide layout fix (--same-slide): Slide 2 upper-left group-count summary,
   upper-right compact mini-table, and lower region detailed 9x7 table with adaptive fonts.

Requires verified audit JSON from build_credit_risk_table.py.
"""
import argparse
import copy
import json
import math
import os
import re
import sys
import uuid
from pathlib import Path

from pptx import Presentation
from pptx.oxml import parse_xml
from pptx.oxml.xmlchemy import OxmlElement
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

GROUPS = [
    ("Total KeyCorp", 1, "Asset Quality"),
    ("Total KeyCorp", 1, "Concentration"),
    ("Line of Business", 2, "Asset Quality"),
    ("Line of Business", 3, "Asset Quality"),
    ("Commercial Industry", 2, "Concentration"),
    ("Commercial Industry", 3, "Concentration"),
    ("Commercial Industry", 2, "Asset Quality"),
    ("Commercial Industry", 3, "Asset Quality"),
]


def norm(value):
    return re.sub(r"\s+", " ", str(value or "")).strip().lower()



def setup_empty_slide_for_credit_risk(slide, report, prs=None):
    """Automatically construct the full native Credit Risk Slide 2 structure from scratch
    when an empty slide or arbitrary non-template presentation is provided.
    First cleans all existing stale shapes, charts, or footers from the slide to avoid ghosting.
    Constructs:
      1. Title & Commentary Text Frame (Upper Left)
      2. Top-Right Summary Counts Table (6 rows x 8 cols)
      3. Main Detailed Table (9 rows x 7 cols) with executive colors and clean fills
      4. Footer Notes Text Box
      5. KeyBank Red Key Logo (Bottom Left)
    """
    # 0. Clean all existing shapes from the slide to avoid ghost overlays and double footers
    for shape in list(slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    quarter = report.get('quarter', '3Q26')
    prior_quarter = get_prior_quarter(quarter)

    # 1. Title & Commentary Text Frame (Left=0.40", Top=0.35", Width=8.2", Height=1.85")
    tb_shape = slide.shapes.add_textbox(Inches(0.40), Inches(0.11), Inches(8.2), Inches(1.75))
    tf = tb_shape.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p0 = tf.paragraphs[0]
    p0.text = f"{quarter} Credit Risk Metrics"
    p0.font.bold = True
    p0.font.size = Pt(20)
    p0.font.name = "Arial"
    p0.space_after = Pt(3)

    row0 = report['rows'][0] if report.get('rows') else None
    if row0:
        total_breaches = len(row0.get('yellow', [])) + len(row0.get('red', []))
        l1_status = f"{total_breaches} Exceeds Warning" if total_breaches > 0 else "All Below Risk Appetite Limits"
    else:
        l1_status = "1 Exceeds Warning"

    p1 = tf.add_paragraph()
    p1.space_after = Pt(2)
    r1 = p1.add_run()
    r1.text = "• Total KeyCorp Level 1 Asset Quality: "
    r1.font.bold = True
    r1.font.size = Pt(11.0)
    r1.font.name = "Arial"
    r2 = p1.add_run()
    r2.text = l1_status
    r2.font.bold = False
    r2.font.size = Pt(11.0)
    r2.font.name = "Arial"

    bullets = [
        ("• LOB Level 2 and 3 Asset Quality: ", "Majority Below Risk Appetite Limits", True),
        ("    ✓ Stable asset quality with six Limit breaches", "", False),
        ("• Commercial Industry Level 2 and 3 Concentration: ", "Majority Within Risk Appetite Limits", True),
        ("    ✓ Portfolio growth and declining CET1 resulted in a stable number of metrics exceeding Warning level", "", False),
        ("• Commercial Industry Level 2 and 3 Asset Quality: ", "Majority Within Risk Appetite Limits", True),
        ("    ✓ Stable asset quality with two Limit breach", "", False),
    ]
    for b_pre, b_val, is_hdr in bullets:
        p = tf.add_paragraph()
        p.space_after = Pt(2)
        r = p.add_run()
        r.text = b_pre + b_val
        r.font.bold = is_hdr
        r.font.size = Pt(11.0)
        r.font.name = "Arial"

    dark_navy = RGBColor(38, 48, 64)
    light_green = RGBColor(204, 255, 204)  # #CCFFCC from template inspection
    light_yellow = RGBColor(255, 255, 153) # #FFFF99 from template inspection
    light_red = RGBColor(255, 124, 128)    # #FF7C80 from template inspection
    white = RGBColor(255, 255, 255)

    # 2. Top-Right Summary Counts Table (6 rows x 8 cols)
    # Exact geometry from template inspection: Left=8.6167", Top=0.345", Width=4.6", Height=1.45"
    summary_shape = slide.shapes.add_table(6, 8, Inches(8.6167), Inches(0.10), Inches(4.6), Inches(1.45))
    s_tbl = summary_shape.table
    try:
        s_tbl._tbl.tblPr.set('bandRow', '0')
        s_tbl._tbl.tblPr.set('firstRow', '0')
    except Exception:
        pass

    s_col_widths = [0.65, 0.48, 0.48, 0.49, 0.49, 0.49, 0.49, 1.03]
    for i, w in enumerate(s_col_widths):
        s_tbl.columns[i].width = Inches(w)
    s_row_heights = [0.46, 0.22, 0.19, 0.19, 0.19, 0.20]
    for i, h in enumerate(s_row_heights):
        s_tbl.rows[i].height = Inches(h)

    # Merges on Summary Table (5 merges from template inspection)
    s_tbl.cell(0, 0).merge(s_tbl.cell(1, 0))
    s_tbl.cell(0, 1).merge(s_tbl.cell(0, 2))
    s_tbl.cell(0, 3).merge(s_tbl.cell(0, 4))
    s_tbl.cell(0, 5).merge(s_tbl.cell(0, 6))
    s_tbl.cell(0, 7).merge(s_tbl.cell(1, 7))

    # Header Text on Summary Table
    s_tbl.cell(0, 1).text = "Within\nRisk Appetite"
    s_tbl.cell(0, 3).text = "Exceeds\nRisk Appetite\nWarning"
    s_tbl.cell(0, 5).text = "Breaches\nRisk Appetite\nLimit"
    s_tbl.cell(0, 7).text = "Total"

    s_tbl.cell(1, 1).text = prior_quarter
    s_tbl.cell(1, 2).text = quarter
    s_tbl.cell(1, 3).text = prior_quarter
    s_tbl.cell(1, 4).text = quarter
    s_tbl.cell(1, 5).text = prior_quarter
    s_tbl.cell(1, 6).text = quarter

    # Row labels on Summary Table
    s_tbl.cell(2, 0).text = "Level 1"
    s_tbl.cell(3, 0).text = "Level 2"
    s_tbl.cell(4, 0).text = "Level 3"
    s_tbl.cell(5, 0).text = "Total"

    # Format all cells of Summary Table with clean fills & font styling
    for r_i, r in enumerate(s_tbl.rows):
        for c_i, cell in enumerate(r.cells):
            if cell.is_spanned:
                continue
            cell.margin_top = Inches(0.01)
            cell.margin_bottom = Inches(0.01)
            cell.margin_left = Inches(0.02)
            cell.margin_right = Inches(0.02)
            cell.fill.solid()

            # Background colors
            if r_i in (0, 1):
                if c_i in (1, 2):
                    cell.fill.fore_color.rgb = light_green
                elif c_i in (3, 4):
                    cell.fill.fore_color.rgb = light_yellow
                elif c_i in (5, 6):
                    cell.fill.fore_color.rgb = light_red
                elif c_i == 7:
                    cell.fill.fore_color.rgb = dark_navy
                else:
                    cell.fill.fore_color.rgb = white
            else:
                if c_i > 0 and (r_i == 5 or c_i == 7):
                    cell.fill.fore_color.rgb = RGBColor(246, 247, 249)
                else:
                    cell.fill.fore_color.rgb = white

            # Font styling
            is_hdr_cell = (r_i in (0, 1)) or (r_i == 5) or (c_i in (0, 7))
            txt_clr = white if (r_i in (0, 1) and c_i == 7) else dark_navy
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.RIGHT if c_i == 0 else PP_ALIGN.CENTER
                p.font.name = "Arial"
                p.font.size = Pt(8.5)
                p.font.bold = is_hdr_cell
                p.font.color.rgb = txt_clr
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(8.5)
                    run.font.bold = is_hdr_cell
                    run.font.color.rgb = txt_clr

    # 3. Main Detailed Table (9x7)
    # Exact geometry from template inspection: Left=0.2406", Top=2.24", Width=12.8521", Height=4.535"
    t_shape = slide.shapes.add_table(9, 7, Inches(0.2406), Inches(1.85), Inches(12.8521), Inches(4.535))
    table = t_shape.table
    try:
        table._tbl.tblPr.set('bandRow', '0')
        table._tbl.tblPr.set('firstRow', '0')
        
    except Exception:
        pass

    cw = [1.584, 0.7093, 1.3256, 1.6987, 1.8023, 2.2985, 3.4338]
    for i, w in enumerate(cw):
        table.columns[i].width = Inches(w)
    rh = [0.40, 0.29, 0.29, 0.29, 1.0578, 0.6731, 0.29, 0.8038, 0.44]
    for i, h in enumerate(rh):
        table.rows[i].height = Inches(h)

    # Headers
    headers = [
        "Credit Risk Metrics", "Level", "Measure",
        "Within\nRisk Appetite", "Exceeds\nRisk Appetite Warning", "",
        "Breaches\nRisk Appetite Limit"
    ]
    for i, h_txt in enumerate(headers):
        if h_txt:
            table.cell(0, i).text = h_txt

    for c in range(3):
        table.cell(0, c).fill.solid()
        table.cell(0, c).fill.fore_color.rgb = dark_navy
        for p in table.cell(0, c).text_frame.paragraphs:
            p.font.name = "Arial"
            p.font.size = Pt(10.0)
            p.font.bold = True
            p.font.color.rgb = white

    table.cell(0, 3).fill.solid()
    table.cell(0, 3).fill.fore_color.rgb = light_green
    for p in table.cell(0, 3).text_frame.paragraphs:
        p.font.name = "Arial"
        p.font.size = Pt(10.0)
        p.font.bold = True
        p.font.color.rgb = dark_navy

    table.cell(0, 4).fill.solid()
    table.cell(0, 4).fill.fore_color.rgb = light_yellow
    for p in table.cell(0, 4).text_frame.paragraphs:
        p.font.name = "Arial"
        p.font.size = Pt(10.0)
        p.font.bold = True
        p.font.color.rgb = dark_navy

    table.cell(0, 5).fill.solid()
    table.cell(0, 5).fill.fore_color.rgb = light_yellow

    table.cell(0, 6).fill.solid()
    table.cell(0, 6).fill.fore_color.rgb = light_red
    for p in table.cell(0, 6).text_frame.paragraphs:
        p.font.name = "Arial"
        p.font.size = Pt(10.0)
        p.font.bold = True
        p.font.color.rgb = dark_navy

    # Merge Col 4 & Col 5 on Header
    table.cell(0, 4).merge(table.cell(0, 5))

    # Row 1-8 Merges and Labels from template inspection:
    # R1C0 -> R2C0 ("Total KeyCorp")
    # R3C0 -> R4C0 ("Line Of Business")
    # R5C0 -> R8C0 ("Commercial Industry")
    row_data = [
        ("Total KeyCorp", "1", "Asset Quality"),
        ("", "1", "Concentration"),
        ("Line Of Business", "2", "Asset Quality"),
        ("", "3", "Asset Quality"),
        ("Commercial Industry", "2", "Concentration"),
        ("", "3", "Concentration"),
        ("", "2", "Asset Quality"),
        ("", "3", "Asset Quality")
    ]
    for r_i, (cat, lvl, meas) in enumerate(row_data, start=1):
        if cat:
            table.cell(r_i, 0).text = cat
        table.cell(r_i, 1).text = lvl
        table.cell(r_i, 2).text = meas
        if r_i != 5:
            table.cell(r_i, 4).merge(table.cell(r_i, 5))

    # Category merges
    table.cell(1, 0).merge(table.cell(2, 0))
    table.cell(3, 0).merge(table.cell(4, 0))
    table.cell(5, 0).merge(table.cell(8, 0))

    # Ensure all data cells have explicit clean white fill to eliminate theme shading
    for r_i in range(1, 9):
        for c_i in range(7):
            cell = table.cell(r_i, c_i)
            if not cell.is_spanned:
                cell.fill.solid()
                cell.fill.fore_color.rgb = white

    # 4. Footer Elements (if not present)
    has_footer = any(s.has_text_frame and 'classification: confidential' in s.text_frame.text.lower() for s in slide.shapes)
    if not has_footer:
        tb_f = slide.shapes.add_textbox(Inches(1.5), Inches(7.08), Inches(10.3), Inches(0.38))
        tff = tb_f.text_frame
        tff.word_wrap = True
        tff.margin_left = tff.margin_right = tff.margin_top = tff.margin_bottom = 0
        pf0 = tff.paragraphs[0]
        pf0.text = "Supporting information on metrics exceeding Risk Appetite Warning and in breach Risk Appetite Limit at the end of the deck"
        pf0.font.name = "Arial"; pf0.font.size = Pt(8.0); pf0.font.bold = False; pf0.font.italic = True; pf0.font.color.rgb = RGBColor(0, 0, 0)
        pf0.alignment = PP_ALIGN.CENTER; pf0.space_after = Pt(0)
        pf1 = tff.add_paragraph()
        pf1.text = "Note: Comparisons above reflect current ERM and Credit risk metric framework. For prior period breaches refer to historic Credit Risk Metric reports"
        pf1.font.name = "Arial"; pf1.font.size = Pt(8.0); pf1.font.bold = False; pf1.font.italic = True; pf1.font.color.rgb = RGBColor(0, 0, 0)
        pf1.alignment = PP_ALIGN.CENTER; pf1.space_after = Pt(0)
        pf2 = tff.add_paragraph()
        pf2.text = "CLASSIFICATION: CONFIDENTIAL"
        pf2.font.name = "Arial"; pf2.font.size = Pt(7.5); pf2.font.bold = False; pf2.font.italic = False; pf2.font.color.rgb = RGBColor(0, 0, 0)
        pf2.alignment = PP_ALIGN.CENTER

    has_num = any(s.has_text_frame and s.text_frame.text.strip() == "2" for s in slide.shapes)
    if not has_num:
        tb_num = slide.shapes.add_textbox(Inches(12.75), Inches(7.02), Inches(0.4), Inches(0.3))
        p_num = tb_num.text_frame.paragraphs[0]
        p_num.text = "2"
        p_num.font.name = "Arial"; p_num.font.size = Pt(9.0); p_num.font.color.rgb = RGBColor(0, 0, 0)

    # 5. KeyBank Red Key Logo
    copy_or_insert_logo(prs, slide)

    return t_shape


def find_main_table_and_slide(prs):
    """Dynamically locate the credit-risk detailed table across all slides.
    Checks Slide 2 first (corporate standard), then checks all other slides."""
    # 1. Standard: Check Slide 2 for 9x7 table
    if len(prs.slides) > 1:
        tables = [s for s in prs.slides[1].shapes if s.has_table and len(s.table.rows) == 9 and len(s.table.columns) == 7]
        if tables:
            return prs.slides[1], tables[0], 2

    # 2. Check any slide for 9x7 table
    for idx, slide in enumerate(prs.slides, start=1):
        tables = [s for s in slide.shapes if s.has_table and len(s.table.rows) == 9 and len(s.table.columns) == 7]
        if tables:
            return slide, tables[0], idx

    # 3. Check Slide 2 for any table with >= 6 rows
    if len(prs.slides) > 1:
        tables = [s for s in prs.slides[1].shapes if s.has_table and len(s.table.rows) >= 6]
        if tables:
            tables.sort(key=lambda s: len(s.table.rows), reverse=True)
            return prs.slides[1], tables[0], 2

    # 4. Check any slide for any table with >= 6 rows
    candidates = []
    for idx, slide in enumerate(prs.slides, start=1):
        for s in slide.shapes:
            if s.has_table and len(s.table.rows) >= 6:
                candidates.append((len(s.table.rows), slide, s, idx))
    if candidates:
        candidates.sort(key=lambda x: x[0], reverse=True)
        return candidates[0][1], candidates[0][2], candidates[0][3]

    # Detailed diagnostic if no table is found
    slide_summary = []
    for idx, s in enumerate(prs.slides, start=1):
        t_count = len([sh for sh in s.shapes if sh.has_table])
        slide_summary.append(f"Slide {idx} (shapes={len(s.shapes)}, tables={t_count})")
    summary_str = "; ".join(slide_summary)
    raise ValueError(f"No detailed credit risk table found in presentation ({summary_str}). "
                     "Please ensure the template slide has the 9x7 detailed table pasted into it.")


def pick_main_table(slide, shape_id=None, verify_geometry=False):
    """Dynamically find the main 9x7 detailed table on Slide 2 regardless of shape ID.
    Supports original template PPTX as well as user's new custom deck slides seamlessly."""
    # 1. Prefer finding table by exact 9x7 dimensions (works across ANY PPT, old or new)
    candidates = [s for s in slide.shapes if s.has_table and len(s.table.rows) == 9 and len(s.table.columns) == 7]
    if candidates:
        return candidates[0]

    # 2. If shape_id was explicitly requested, try finding by that ID
    if shape_id is not None:
        id_candidates = [s for s in slide.shapes if s.shape_id == shape_id and s.has_table]
        if id_candidates:
            return id_candidates[0]

    # 3. Fallback: Find table by largest row count on the slide
    all_tables = [s for s in slide.shapes if s.has_table]
    if all_tables:
        all_tables.sort(key=lambda s: len(s.table.rows), reverse=True)
        return all_tables[0]

    raise ValueError("Slide 2 does not contain any table. Ensure Slide 2 has the detailed credit risk table.")


def paragraphs_with_rich_text(cell):
    return [copy.deepcopy(p._p) for p in cell.text_frame.paragraphs if p.text.strip()]


def raw_paragraph_text(p_element):
    return ''.join(e.text or '' for e in p_element.xpath('.//a:t'))


def is_bulleted_p(p_element):
    return bool(p_element.xpath('./a:pPr/a:buChar') or p_element.xpath('./a:pPr/a:buAutoNum'))


def styled_lines(cell, lines, fallback=None, bullet=False):
    if cell.is_spanned:
        return
    options = paragraphs_with_rich_text(cell)
    if not options and fallback is not None:
        options = paragraphs_with_rich_text(fallback)
    if not options:
        default_p = OxmlElement('a:p')
        p_pr = OxmlElement('a:pPr')
        default_p.append(p_pr)
        first_run = OxmlElement('a:r')
        rpr = OxmlElement('a:rPr')
        rpr.set('sz', '1000')
        first_run.append(rpr)
        default_p.append(first_run)
        options = [default_p]
    body = cell.text_frame._txBody
    for p in list(body):
        if p.tag.endswith('}p'):
            body.remove(p)
    for k, item in enumerate(lines or ['']):
        template = copy.deepcopy(options[min(k, len(options)-1)])
        first_run = next(iter(template.xpath('./a:r')), None)
        if first_run is None:
            first_run = OxmlElement('a:r')
            sample_rpr = template.xpath('./a:endParaRPr')
            if sample_rpr:
                rpr = copy.deepcopy(sample_rpr[0])
                rpr.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr'
                first_run.append(rpr)
        else:
            first_run = copy.deepcopy(first_run)
        for ch in list(template):
            if ch.tag.endswith(('}r', '}br', '}fld')):
                template.remove(ch)
        for child in list(first_run):
            if child.tag.endswith('}t'):
                first_run.remove(child)

        # When bullet=True, explicitly strip bolding and ensure uniform zero spacing
        if bullet:
            rpr_match = first_run.xpath('./a:rPr')
            if rpr_match:
                rpr_match[0].set('b', '1')
            p_pr_match = template.xpath('./a:pPr')
            if p_pr_match:
                p_pr_match[0].attrib.pop('spcBfr', None)
                p_pr_match[0].attrib.pop('spcAft', None)

        txt = OxmlElement('a:t')
        prefix = ''
        if bullet and item:
            if not is_bulleted_p(template):
                prefix = '\u2022  '
        txt.text = prefix + str(item)
        first_run.append(txt)
        ends = template.xpath('./a:endParaRPr')
        if ends:
            template.insert(list(template).index(ends[0]), first_run)
        else:
            template.append(first_run)
        body.append(template)


def clean_metric_display_name(name, row_idx=None):
    """Transform formal database strings into clean presentation deck standards."""
    if not name:
        return ""
    n = str(name).strip()
    # Row 5 / Concentration under CMT % of CET1 header: strip redundant suffix
    if row_idx == 5 or 'commitments as % of cet1' in n.lower() or '% of cet1' in n.lower():
        if 'construction materials' in n.lower():
            return 'Construction Materials and Contractors'
        n = re.sub(r'\s*-\s*Commitments\s+as\s+%?\s*of\s*CET1.*$', '', n, flags=re.I)
        n = re.sub(r'\s*%?\s*of\s*CET1.*$', '', n, flags=re.I)
        n = re.sub(r'\s*Industry\s*$', '', n, flags=re.I)
        return n.strip()

    # Exact client presentation deck transforms
    transforms = {
        'Institutional Banking - Criticized Commitments': 'Institutional Criticized',
        'Real Estate Capital - Criticized Commitments': 'Real Estate Capital Criticized',
        'Consumer Credit Card - Early Payment Default Rate': 'Consumer Credit Card EPD Rate',
        'Small Business - Early Payment Default Rate': 'Small Business EPD Rate',
        'Student Loans Direct - Early Payment Default Rate': 'Student Loans Direct EPD Rate',
        'Business Banking - Criticized Commitments': 'Business Banking Criticized',
        'Enterprise Leveraged (Total Leveraged) - Interest Coverage': 'Total Leveraged Interest Coverage Ratio',
        'Enterprise Leveraged - Interest Coverage (Core)': 'Enterprise Leveraged - Interest Coverage (Core)',
        'Enterprise Leveraged - Interest Coverage (Non-Core)': 'Non-Core Leveraged Interest Coverage Ratio',
        'KeyCorp - Criticized Commitments %': 'KeyCorp Criticized %',
        'KeyCorp - Criticized': 'KeyCorp Criticized %',
        'Student Loans Direct - 30-89 Days Past Due': 'Student Loans Direct - 30-89 Days Past Due %',
        'Student Loans Direct 30-89 Days Past Due': 'Student Loans Direct - 30-89 Days Past Due %',
        'Student Loans Direct - 90+ Days Past Due': 'Student Loans Direct - 90+ Days Past Due %',
        'Student Loans Direct 90+ Days Past Due': 'Student Loans Direct - 90+ Days Past Due %',
        'Other Consumer - 90+ Days Past Due': 'Other Consumer Secured - 90+ Days Past Due %',
        'Other Consumer Secured - 90+ Days Past Due': 'Other Consumer Secured - 90+ Days Past Due %',
        'Other Consumer Secured 90+ Days Past Due': 'Other Consumer Secured - 90+ Days Past Due %',
        'Construction Materials and Contractors': 'Construction Materials and Contractors',
        'Construction Materials and Contractors - Commitments as % of CET1': 'Construction Materials and Contractors',
        'Construction Materials & Contractors': 'Construction Materials and Contractors',
    }
    if n in transforms:
        return transforms[n]

    # Row 7 / Industry Criticized Commitments % -> Industry Criticized
    if 'Criticized Commitments %' in n:
        return n.replace('Criticized Commitments %', 'Criticized').strip()
    if 'Criticized Commitments' in n:
        return n.replace('Criticized Commitments', 'Criticized').strip()

    return n


def styled_lines_with_header(cell, header_text, lines, fallback=None):
    """Render a header paragraph (or spacer) followed by bulleted items in a cell."""
    if cell.is_spanned:
        return
    options = paragraphs_with_rich_text(cell)
    if not options and fallback is not None:
        options = paragraphs_with_rich_text(fallback)
    if not options:
        default_p = OxmlElement('a:p')
        p_pr = OxmlElement('a:pPr')
        default_p.append(p_pr)
        first_run = OxmlElement('a:r')
        rpr = OxmlElement('a:rPr')
        rpr.set('sz', '1000')
        first_run.append(rpr)
        default_p.append(first_run)
        options = [default_p]
    body = cell.text_frame._txBody
    for p in list(body):
        if p.tag.endswith('}p'):
            body.remove(p)

    # Extract sample run properties
    sample_rpr = None
    first_run_match = options[0].xpath('./a:r')
    if first_run_match:
        rpr_match = first_run_match[0].xpath('./a:rPr')
        if rpr_match:
            sample_rpr = rpr_match[0]
    if sample_rpr is None:
        end_rpr = options[0].xpath('./a:endParaRPr')
        if end_rpr:
            sample_rpr = end_rpr[0]

    # Header paragraph (non-bullet, bold if text, or spacer if empty)
    hdr_tmpl = copy.deepcopy(options[0])
    for ch in list(hdr_tmpl):
        if ch.tag.endswith(('}r', '}br', '}fld')):
            hdr_tmpl.remove(ch)
    r_hdr = OxmlElement('a:r')
    if sample_rpr is not None:
        rpr_hdr = copy.deepcopy(sample_rpr)
        rpr_hdr.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr'
    else:
        rpr_hdr = OxmlElement('a:rPr')
    
    is_spacer = not header_text or not str(header_text).strip()
    if not is_spacer:
        rpr_hdr.set('b', '1')
    else:
        rpr_hdr.attrib.pop('b', None)
    rpr_hdr.set('sz', '1000')
    r_hdr.append(rpr_hdr)
    t_hdr = OxmlElement('a:t')
    t_hdr.text = " " if is_spacer else str(header_text)
    r_hdr.append(t_hdr)
    hdr_tmpl.append(r_hdr)
    
    # Strip space before/after on header
    p_pr_h = hdr_tmpl.xpath('./a:pPr')
    if p_pr_h:
        p_pr_h[0].attrib.pop('spcBfr', None)
        p_pr_h[0].attrib.pop('spcAft', None)
        for b_child in list(p_pr_h[0]):
            if b_child.tag.endswith(('buChar', 'buAutoNum', 'buBlip', 'buClr', 'buSzPct', 'buSzPts', 'buFont')):
                p_pr_h[0].remove(b_child)
    body.append(hdr_tmpl)

    # Bullets underneath: ALL items (Finance, Consumer Services, etc.) use bullet template & sample rpr
    # from fallback (pure yellow/amber bullet) or options[1:], so formatting is 100% identical
    bullet_options = []
    if fallback is not None:
        bullet_options = paragraphs_with_rich_text(fallback)
    if not bullet_options and len(options) > 1:
        bullet_options = options[1:]
    if not bullet_options:
        bullet_options = options

    bullet_sample_rpr = None
    b_first_run = bullet_options[0].xpath('./a:r')
    if b_first_run:
        b_rpr = b_first_run[0].xpath('./a:rPr')
        if b_rpr:
            bullet_sample_rpr = b_rpr[0]
    if bullet_sample_rpr is None:
        b_end_rpr = bullet_options[0].xpath('./a:endParaRPr')
        if b_end_rpr:
            bullet_sample_rpr = b_end_rpr[0]

    for k, item in enumerate(lines or []):
        tmpl = copy.deepcopy(bullet_options[min(k, len(bullet_options)-1)])
        for ch in list(tmpl):
            if ch.tag.endswith(('}r', '}br', '}fld')):
                tmpl.remove(ch)
        r_item = OxmlElement('a:r')
        if bullet_sample_rpr is not None:
            rpr_item = copy.deepcopy(bullet_sample_rpr)
            rpr_item.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr'
        else:
            rpr_item = OxmlElement('a:rPr')
        rpr_item.set('b', '1')
        rpr_item.set('sz', '1000')
        r_item.append(rpr_item)
        t_item = OxmlElement('a:t')
        prefix = '\u2022  ' if not is_bulleted_p(tmpl) else ''
        t_item.text = prefix + str(item)
        r_item.append(t_item)
        tmpl.append(r_item)
        
        p_pr_i = tmpl.xpath('./a:pPr')
        if p_pr_i:
            p_pr_i[0].attrib.pop('spcBfr', None)
            p_pr_i[0].attrib.pop('spcAft', None)
        body.append(tmpl)


def estimate_lines(names, width_in, font_pt=10.0):
    chars = max(12, int((width_in * 72 - 9) / (font_pt * 0.53)))
    return sum(max(1, math.ceil((len(n) + 2) / chars)) for n in names)


def text_font_pt(cell, fallback=None):
    for source in (cell, fallback):
        if source is None:
            continue
        for p in source.text_frame.paragraphs:
            for run in p.runs:
                if run.font.size:
                    return run.font.size.pt
    return 10.0


def distribute_yellow(table, row_i, names):
    left, right = table.cell(row_i, 4), table.cell(row_i, 5)
    if right.is_spanned or left.is_merge_origin:
        return names, []
    if not right.text.strip():
        return names, []
    original_left = [p.text.strip() for p in left.text_frame.paragraphs if p.text.strip()]
    original_right = [p.text.strip() for p in right.text_frame.paragraphs if p.text.strip()]
    nleft = len(original_left)
    nright = len(original_right)
    if nleft + nright == 0:
        return names, []
    fraction = nleft / (nleft + nright)
    k = min(len(names), max(1, round(len(names) * fraction))) if names else 0
    return names[:k], names[k:]


def validate_report(report):
    q = report.get('quarter')
    if not q or not re.match(r'^[1-4]Q\d{2}$', str(q)):
        raise ValueError(f'Invalid quarter format in audit report: {q!r}')
    rows = report.get('rows')
    if not isinstance(rows, list) or len(rows) != 8:
        raise ValueError('Need exactly 8 metric rows in audit JSON')
    draft = report.get('draft', False)
    if not isinstance(draft, bool):
        raise ValueError('Invalid draft flag')
    seen = set(); totals = {'green': 0, 'yellow': 0, 'red': 0, 'total': 0, 'pending': 0, 'verified': 0}
    for row, g in zip(rows, GROUPS):
        if (row.get('category'), row.get('level'), row.get('measure')) != g:
            raise ValueError(f'Unexpected row/order {row}; expected {g}')
        green = row.get('green'); total = row.get('total')
        if not isinstance(green, int) or not isinstance(total, int) or green < 0 or total < 0:
            raise ValueError(f'Invalid green/total count for {g}')
        yellow = row.get('yellow'); red = row.get('red'); pending = row.get('pending', [])
        if not isinstance(yellow, list) or not isinstance(red, list) or not isinstance(pending, list):
            raise ValueError('Yellow/red/pending metrics must be lists')
        if pending and not draft:
            raise ValueError('Pending statuses require a draft audit')
        if green + len(yellow) + len(red) + len(pending) != total:
            raise ValueError(f'Nonreconciling green/yellow/red/pending counts for {g}')
        for item in yellow + red + pending:
            mid = item.get('id'); name = item.get('name')
            if not mid or not isinstance(name, str) or not name.strip() or mid in seen:
                raise ValueError(f'Missing/duplicate metric ID or name: {item!r}')
            seen.add(mid)
        for name, count in [('green', green), ('yellow', len(yellow)), ('red', len(red)),
                            ('pending', len(pending)), ('total', total),
                            ('verified', green + len(yellow) + len(red))]:
            totals[name] += count
    for key, num in totals.items():
        actual = report.get('totals', {}).get(key)
        if actual is None and not draft and key in ('pending', 'verified'):
            actual = 0 if key == 'pending' else (totals['green'] + totals['yellow'] + totals['red'])
        if actual != num:
            raise ValueError(f'Audit overall total {key} does not reconcile ({num})')


def fit_split_yellow(table, row_i, names, font_pt=8.8):
    left, right = table.cell(row_i, 4), table.cell(row_i, 5)
    if right.is_spanned or left.is_merge_origin:
        return list(names), []
    widths = [table.columns[i].width / Inches(1) for i in (4, 5)]
    candidates = []
    for k in range(len(names) + 1):
        a, b = list(names[:k]), list(names[k:])
        la = estimate_lines(a, widths[0], font_pt)
        lb = estimate_lines(b, widths[1], font_pt)
        candidates.append((max(la, lb), abs(la - lb), -len(a), k))
    k = min(candidates)[3]
    return list(names[:k]), list(names[k:])


def dedicated_layout(prs, table, planned, title):
    slide_w = prs.slide_width / Inches(1)
    slide_h = prs.slide_height / Inches(1)
    top = max(0.75, slide_h * 0.11)
    bottom = 0.35
    room = slide_h - top - bottom
    width = min(slide_w - 0.44, 12.85209)
    table_shape = table._graphic_frame
    table_shape.left = Inches((slide_w - width) / 2)
    table_shape.top = Inches(top)
    table_shape.width = Inches(width)
    widths = [c.width / Inches(1) for c in table.columns]
    original_total = sum(widths)
    if abs(original_total - width) > 0.001:
        for i, c in enumerate(table.columns):
            c.width = Inches(widths[i] * width / original_total)
    widths = [c.width / Inches(1) for c in table.columns]

    solutions = None
    for font_pt in (9.0, 8.7, 8.4, 8.0, 7.6, 7.2):
        heights = [0.41]
        for spec in planned:
            r = spec['row']
            g = spec['green']
            left = spec['yellow_left']; right = spec['yellow_right']; red = spec['red']
            merged_y = table.cell(r, 4).is_merge_origin and table.cell(r, 5).is_spanned
            wleft = widths[4] + (widths[5] if merged_y else 0)
            levels = [estimate_lines(left, wleft, font_pt), estimate_lines(red, widths[6], font_pt)]
            if not merged_y: levels.append(estimate_lines(right, widths[5], font_pt))
            levels.append(estimate_lines([g], widths[3], font_pt))
            n = max(levels + [1])
            heights.append(max(0.37, n * font_pt * 1.27 / 72 + 0.095))
        total = sum(heights)
        if total <= room - 0.04:
            solutions = (font_pt, heights)
            break
    if solutions is None:
        raise ValueError('Text still cannot fit on one dedicated slide even at 7.2 pt.')
    font_pt, heights = solutions
    leftover = room - sum(heights) - 0.03
    weights = [max(0.5, h) for h in heights[1:]]
    weight_total = sum(weights)
    for i, h in enumerate(heights):
        if i:
            h += min(0.11, leftover * weights[i-1] / weight_total)
        table.rows[i].height = Inches(h)
    table_shape.height = Inches(sum(r.height for r in table.rows) / Inches(1))
    for r in range(1, 9):
        for c in (3, 4, 5, 6):
            cell = table.cell(r, c)
            if cell.is_spanned: continue
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(font_pt)
    return {'font_pt': font_pt, 'table_height_in': round(table_shape.height / Inches(1), 3),
            'table_top_in': round(top, 3), 'row_heights_in': [round(row.height / Inches(1), 3) for row in table.rows],
            'room_in': round(room, 3)}


def get_prior_quarter(q_str):
    """Compute previous reporting quarter (e.g. 3Q26 -> 2Q26, 1Q27 -> 4Q26, 2Q27 -> 1Q27)."""
    m = re.search(r'([1-4])\s*Q\s*(?:20)?(\d{2})', str(q_str or '').strip(), re.I)
    if not m:
        return '2Q26'
    q, y = int(m.group(1)), int(m.group(2))
    if q == 1:
        return f"4Q{(y - 1) % 100:02d}"
    return f"{q - 1}Q{y:02d}"


def update_slide2_summary_table(slide, report, main_table_shape_id=8):
    """Find and populate ONLY the existing native small summary table on Slide 2."""
    summary_shape = next((s for s in slide.shapes if s.has_table and s.shape_id != main_table_shape_id), None)
    if not summary_shape:
        return None

    table = summary_shape.table
    quarter = report.get('quarter', '3Q26')
    prior_quarter = get_prior_quarter(quarter)

    # Compute current quarter counts from validated audit report
    bylevel = {n: dict(Green=0, Yellow=0, Red=0, Total=0, MasterTotal=0) for n in (1, 2, 3)}
    for row in report['rows']:
        lvl = row['level']
        bylevel[lvl]['Green'] += row['green']
        bylevel[lvl]['Yellow'] += len(row['yellow'])
        bylevel[lvl]['Red'] += len(row['red'])
        bylevel[lvl]['Total'] += row['total']
        bylevel[lvl]['MasterTotal'] += row.get('master_total', row['total'])

    totals = report.get('totals', {})

    # Historical 2Q26 baseline counts (authoritative reference from approved deck)
    baseline_2q26 = {
        'level 1': {'Green': 9, 'Yellow': 0, 'Red': 0, 'Total': 9},
        'level 2': {'Green': 36, 'Yellow': 9, 'Red': 2, 'Total': 48},
        'level 3': {'Green': 53, 'Yellow': 12, 'Red': 6, 'Total': 61},
        'total': {'Green': 98, 'Yellow': 19, 'Red': 8, 'Total': 118},
    }

    sample_cell = next((table.cell(r, c) for r in range(len(table.rows)) for c in range(len(table.columns))
                        if not table.cell(r, c).is_spanned and table.cell(r, c).text.strip()), None)

    # Find the quarter subheader row (the row containing 1Q/2Q/3Q/4Q)
    subhdr_row = 1 if len(table.rows) > 1 else 0
    for r_idx in range(min(2, len(table.rows))):
        row_text = ' '.join(table.cell(r_idx, c).text.lower() for c in range(len(table.columns)))
        if any(q in row_text for q in ['1q', '2q', '3q', '4q']):
            subhdr_row = r_idx
            break

    if len(table.columns) >= 7:
        for c_prior, c_curr in [(1, 2), (3, 4), (5, 6)]:
            if c_prior < len(table.columns) and c_curr < len(table.columns):
                if not table.cell(subhdr_row, c_prior).is_spanned:
                    styled_lines(table.cell(subhdr_row, c_prior), [prior_quarter], fallback=sample_cell)
                if not table.cell(subhdr_row, c_curr).is_spanned:
                    styled_lines(table.cell(subhdr_row, c_curr), [quarter], fallback=sample_cell)
        if len(table.columns) > 7 and not table.cell(subhdr_row, 7).is_spanned and not table.cell(subhdr_row, 7).text.strip():
            styled_lines(table.cell(subhdr_row, 7), ['Total'], fallback=sample_cell)
    elif len(table.columns) == 5:
        for c_idx in (1, 2, 3):
            if not table.cell(subhdr_row, c_idx).is_spanned:
                styled_lines(table.cell(subhdr_row, c_idx), [quarter], fallback=sample_cell)
        if not table.cell(subhdr_row, 4).is_spanned:
            styled_lines(table.cell(subhdr_row, 4), ['Total'], fallback=sample_cell)

    # Level mappings for data rows
    level_map = {
        'level 1': ('level 1', bylevel[1]),
        'level 2': ('level 2', bylevel[2]),
        'level 3': ('level 3', bylevel[3]),
        'total': ('total', totals)
    }

    for r_idx in range(len(table.rows)):
        col0_text = table.cell(r_idx, 0).text.strip().lower()
        matched_key = None
        current_data = None
        for key_prefix, (bkey, counts) in level_map.items():
            if key_prefix in col0_text:
                matched_key = bkey
                current_data = counts
                break
        if not matched_key:
            continue

        b_data = baseline_2q26[matched_key]

        if len(table.columns) >= 7:
            # Prior quarter columns: 1 (Green), 3 (Yellow), 5 (Red)
            # Current quarter columns: 2 (Green), 4 (Yellow), 6 (Red)
            for c_prior, c_curr, col_key in [(1, 2, 'Green'), (3, 4, 'Yellow'), (5, 6, 'Red')]:
                if c_prior < len(table.columns) and not table.cell(r_idx, c_prior).is_spanned:
                    curr_prior_txt = table.cell(r_idx, c_prior).text.strip()
                    if quarter == '3Q26' or not curr_prior_txt or curr_prior_txt == '10':
                        prior_val = b_data[col_key]
                    else:
                        prior_val = curr_prior_txt
                    styled_lines(table.cell(r_idx, c_prior), [str(prior_val)], fallback=sample_cell)
                if c_curr < len(table.columns) and not table.cell(r_idx, c_curr).is_spanned:
                    val_curr = current_data.get(col_key, current_data.get(col_key.lower(), 0))
                    styled_lines(table.cell(r_idx, c_curr), [str(val_curr)], fallback=sample_cell)

            # Total column (Col 7)
            if len(table.columns) > 7 and not table.cell(r_idx, 7).is_spanned:
                tot_val = current_data.get('MasterTotal', current_data.get('master_total', b_data['Total']))
                styled_lines(table.cell(r_idx, 7), [str(tot_val)], fallback=sample_cell)
        elif len(table.columns) == 5:
            if not table.cell(r_idx, 1).is_spanned:
                styled_lines(table.cell(r_idx, 1), [str(current_data.get('Green', current_data.get('green', 0)))], fallback=sample_cell)
            if not table.cell(r_idx, 2).is_spanned:
                styled_lines(table.cell(r_idx, 2), [str(current_data.get('Yellow', current_data.get('yellow', 0)))], fallback=sample_cell)
            if not table.cell(r_idx, 3).is_spanned:
                styled_lines(table.cell(r_idx, 3), [str(current_data.get('Red', current_data.get('red', 0)))], fallback=sample_cell)
            if not table.cell(r_idx, 4).is_spanned:
                styled_lines(table.cell(r_idx, 4), [str(current_data.get('MasterTotal', current_data.get('master_total', b_data['Total'])))], fallback=sample_cell)

    return summary_shape


def copy_or_insert_logo(prs, target_slide):
    """Ensure KeyBank Red Key logo is on Slide 2.
    1. Looks for logo on Slide 3 (next slide), Slide 1 (previous slide), or any slide.
    2. Falls back to local/bundled keybank_red_key_clean.png if no picture shape exists.
    """
    if target_slide is None:
        return
    has_pic = any(s.shape_type == MSO_SHAPE_TYPE.PICTURE for s in target_slide.shapes)
    if has_pic:
        return

    import io
    search_order = []
    if prs is not None:
        try:
            cur_idx = list(prs.slides).index(target_slide)
            if cur_idx + 1 < len(prs.slides):
                search_order.append(prs.slides[cur_idx + 1])
            if cur_idx - 1 >= 0:
                search_order.append(prs.slides[cur_idx - 1])
            for s in prs.slides:
                if s not in search_order and s != target_slide:
                    search_order.append(s)
        except Exception:
            search_order = [s for s in prs.slides if s != target_slide]

    for s in search_order:
        for sh in s.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.top > Inches(5.5):
                try:
                    img_bytes = sh.image.blob
                    target_slide.shapes.add_picture(io.BytesIO(img_bytes), sh.left, sh.top, sh.width, sh.height)
                    return
                except Exception:
                    pass
        for sh in s.shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    img_bytes = sh.image.blob
                    left = sh.left if sh.top > Inches(5.5) else Inches(0.4)
                    top = sh.top if sh.top > Inches(5.5) else Inches(7.0)
                    target_slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, sh.width, sh.height)
                    return
                except Exception:
                    pass

    # Fallback to local image
    possible_paths = [
        'keybank_red_key_clean.png',
        'c:/chatgpt_recovery/keybank_red_key_clean.png',
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'keybank_red_key_clean.png')
    ]
    for p in possible_paths:
        if os.path.exists(p):
            try:
                target_slide.shapes.add_picture(p, Inches(0.4), Inches(7.0), Inches(0.50), Inches(0.22))
                return
            except Exception:
                pass


def set_cell_border_and_fill(cell, fill_hex=None, bottom_border_hex='D0D4DC', bottom_border_pt=0.75):
    """Set cell background fill and bottom border with strictly valid DrawingML element ordering."""
    tcPr = cell._tc.get_or_add_tcPr()
    for el in list(tcPr):
        if any(el.tag.endswith(x) for x in ['lnB', 'lnT', 'lnL', 'lnR', 'solidFill', 'noFill']):
            tcPr.remove(el)

    if bottom_border_hex and bottom_border_pt > 0:
        w_emu = int(bottom_border_pt * 12700)
        lnB = parse_xml(
            f'<a:lnB xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{w_emu}" cmpd="sng">'
            f'  <a:solidFill>'
            f'    <a:srgbClr val="{bottom_border_hex.replace("#", "")}"/>'
            f'  </a:solidFill>'
            f'</a:lnB>'
        )
        tcPr.append(lnB)

    if fill_hex:
        sf = parse_xml(
            f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'  <a:srgbClr val="{fill_hex.replace("#", "")}"/>'
            f'</a:solidFill>'
        )
        tcPr.append(sf)
    else:
        nf = parse_xml('<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>')
        tcPr.append(nf)


def update_slide2_title_and_commentary(slide, report):
    """Update Slide 2 Title and Commentary matching official template exact 11pt typography & single bullets."""
    quarter = report.get('quarter', '3Q26')
    tb = next((s for s in slide.shapes if s.has_text_frame and not s.has_table and s.top < Inches(2.2)), None)
    if not tb: return False
    # Move Title & Commentary text frame UP to give room for the table
    tb.top = Inches(0.10)
    tb.left = Inches(0.40)
    tb.width = Inches(8.20)
    tb.text_frame.margin_top = Inches(0.0)
    tb.text_frame.margin_bottom = Inches(0.0)
    tb.text_frame.margin_left = Inches(0.0)
    tb.text_frame.margin_right = Inches(0.0)
    tf = tb.text_frame
    tf.word_wrap = True

    dark_navy = RGBColor(35, 41, 54)     # #232936
    green_clr = RGBColor(112, 173, 71)   # #70AD47
    amber_clr = RGBColor(242, 174, 0)    # #F2AE00
    red_clr = RGBColor(193, 0, 27)       # #C1001B
    black_clr = RGBColor(0, 0, 0)

    num_words = {0: 'zero', 1: 'one', 2: 'two', 3: 'three', 4: 'four', 5: 'five', 6: 'six', 7: 'seven', 8: 'eight', 9: 'nine', 10: 'ten'}
    rows = report.get('rows', [])

    # Row 0: L1 Asset Quality
    row0 = rows[0] if len(rows) > 0 else {}
    y0 = [clean_metric_display_name(x['name'], 1) for x in row0.get('yellow', [])]
    r0 = [clean_metric_display_name(x['name'], 1) for x in row0.get('red', [])]
    tot0 = len(y0) + len(r0)
    if tot0 == 0:
        p1_runs = [('• Total KeyCorp Level 1 Asset Quality and Concentration: ', True, dark_navy),
                   ('All Below Risk Appetite Limits', True, green_clr)]
    else:
        joined = ', '.join(y0 + r0)
        p1_runs = [('• Total KeyCorp Level 1 Asset Quality: ', True, dark_navy),
                   (f'{tot0} Exceeds Warning ({joined})', False, dark_navy)]

    # Rows 2 & 3: LOB breaches
    row2 = rows[2] if len(rows) > 2 else {}
    row3 = rows[3] if len(rows) > 3 else {}
    lob_red = len(row2.get('red', [])) + len(row3.get('red', []))
    lob_w = num_words.get(lob_red, str(lob_red))
    p3_runs = [(f'    ✓ Stable asset quality with {lob_w} ', False, dark_navy),
               ('Limit', True, red_clr),
               (f" {'breach' if lob_red == 1 else 'breaches'}", False, dark_navy)]

    # Rows 6 & 7: Commercial Industry breaches
    row6 = rows[6] if len(rows) > 6 else {}
    row7 = rows[7] if len(rows) > 7 else {}
    ci_red = len(row6.get('red', [])) + len(row7.get('red', []))
    ci_w = num_words.get(ci_red, str(ci_red))
    p7_runs = [(f'    ✓ Stable asset quality with {ci_w} ', False, dark_navy),
               ('Limit', True, red_clr),
               (f" {'breach' if ci_red == 1 else 'breaches'}", False, dark_navy)]

    specs = [
        [(f"{quarter} Credit Risk Metrics", True, black_clr, 20.0)],
        [(t, b, c, 12.0) for t, b, c in p1_runs],
        [('• LOB Level 2 and 3 Asset Quality: ', True, dark_navy, 12.0),
         ('Majority ', False, dark_navy, 12.0),
         ('Below Risk Appetite Limits', True, green_clr, 12.0)],
        [(t, b, c, 12.0) for t, b, c in p3_runs],
        [('• Commercial Industry Level 2 and 3 Concentration: ', True, dark_navy, 12.0),
         ('Majority ', False, dark_navy, 12.0),
         ('Within Risk Appetite Limits', True, green_clr, 12.0)],
        [('    ✓ Portfolio growth and declining CET1 resulted in a stable number of metrics exceeding ', False, dark_navy, 12.0),
         ('Warning', True, amber_clr, 12.0),
         (' level', False, dark_navy, 12.0)],
        [('• Commercial Industry Level 2 and 3 Asset Quality: ', True, dark_navy, 12.0),
         ('Majority ', False, dark_navy, 12.0),
         ('Within Risk Appetite Limits', True, green_clr, 12.0)],
        [(t, b, c, 12.0) for t, b, c in p7_runs]
    ]

    while len(tf.paragraphs) < len(specs):
        tf.add_paragraph()

    bu_none_xml = '<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
    for p_idx, run_specs in enumerate(specs):
        p = tf.paragraphs[p_idx]
        p.text = ""
        p.space_before = Pt(0)
        p.space_after = Pt(1.0) if p_idx == 0 else Pt(0.5)
        p.line_spacing = Pt(11.5) if p_idx > 0 else Pt(18.0)
        
        # Enforce buNone in pPr so layout does not generate duplicate bullet
        pPr = p._p.get_or_add_pPr()
        for child in list(pPr):
            if any(child.tag.endswith(x) for x in ['buChar', 'buAutoNum', 'buBlip', 'buNone']):
                pPr.remove(child)
        pPr.append(parse_xml(bu_none_xml))

        for text, bold, color, size in run_specs:
            r = p.add_run()
            r.text = text
            r.font.name = "Arial"
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
    return True


def ensure_slide_text_arial(slide):
    """Ensure all text across the entire slide uses Arial font."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                p.font.name = "Arial"
                for r in p.runs:
                    r.font.name = "Arial"
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    if not cell.is_spanned:
                        for p in cell.text_frame.paragraphs:
                            p.font.name = "Arial"
                            for r in p.runs:
                                r.font.name = "Arial"

def apply_same_slide_layout(slide, table, planned, report, count_reference_pptx=None, prs=None):
    """Populate Slide 2 native tables and update title/commentary matching ideal slide.

    1. Updates Slide 2 Title (20pt Bold) and Commentary with exact colored highlights matching ppt.pptx.
    2. Populates existing native top-right summary table with white data cells and subtle bottom divider.
    3. Formats bottom 9x7 detailed table with alternating row fills (#FFFFFF & #F6F7F9),
       executive category dividing lines, exact 8.5pt regular bullet text, and exact row heights.
    4. Copies Red KeyBank logo from presentation (or inserts official logo) on Slide 2.
    """
    copy_or_insert_logo(prs, slide)

    # 1. Update title and commentary text
    update_slide2_title_and_commentary(slide, report)

    # 2. Update existing top-right summary table
    summary_shape = update_slide2_summary_table(slide, report, main_table_shape_id=table._graphic_frame.shape_id)
    if summary_shape:
        summary_shape.top = Inches(0.10)
        summary_shape.left = Inches(8.6167)

    # Clean styling on summary table: white data cells, subtle bottom divider, clean alignment
    if summary_shape:
        s_tbl = summary_shape.table
        try:
            s_tbl._tbl.tblPr.set('bandRow', '0')
            s_tbl._tbl.tblPr.set('firstRow', '0')
        except Exception:
            pass
        s_tbl.cell(0, 0).fill.background()
        for s_r in range(len(s_tbl.rows)):
            for s_c in range(len(s_tbl.columns)):
                s_cell = s_tbl.cell(s_r, s_c)
                if s_cell.is_spanned:
                    continue
                s_cell.margin_top = Inches(0.01)
                s_cell.margin_bottom = Inches(0.01)
                s_cell.margin_left = Inches(0.02)
                s_cell.margin_right = Inches(0.02)

                if s_r < 2:
                    # Headers: vertically and horizontally centered
                    s_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                    for p in s_cell.text_frame.paragraphs:
                        p.alignment = PP_ALIGN.CENTER
                        p.font.name = "Arial"
                        p.font.size = Pt(8.0)
                        p.font.bold = True
                        p.font.color.rgb = RGBColor(255, 255, 255) if s_c == 7 else RGBColor(0, 0, 0)
                        for r in p.runs:
                            r.font.name = "Arial"
                            r.font.size = Pt(8.0)
                            r.font.bold = True
                else:
                    # Data cells: font size 8.0, bold srif latest quarter ke column numbers hai is me.
                    # Col 0: Level labels (Bold), Col 7: Total (Bold)
                    # Col 1, 3, 5: Previous Quarter numbers (NOT Bold)
                    # Col 2, 4, 6: Latest Quarter numbers (Bold)
                    is_prev_q = (s_c in (1, 3, 5))
                    cell_bold = not is_prev_q

                    for p in s_cell.text_frame.paragraphs:
                        p.alignment = PP_ALIGN.RIGHT if s_c == 0 else PP_ALIGN.CENTER
                        p.font.name = "Arial"
                        p.font.size = Pt(8.0)
                        p.font.bold = cell_bold
                        p.font.color.rgb = RGBColor(0, 0, 0)
                        for r in p.runs:
                            r.font.name = "Arial"
                            r.font.size = Pt(8.0)
                            r.font.bold = cell_bold
                            r.font.color.rgb = RGBColor(0, 0, 0)

    # 3. Enforce exact column widths and geometry matching Ideal_slide_ppt.pptx:
    ideal_col_widths = [1.584, 0.7093, 1.3256, 1.6987, 1.8023, 2.2985, 3.4338]
    for i, w in enumerate(ideal_col_widths):
        if i < len(table.columns):
            table.columns[i].width = Inches(w)

    table._graphic_frame.left = Inches(0.2406)
    table._graphic_frame.top = Inches(1.7)
    table._graphic_frame.width = Inches(12.8521)

    # 4. Format ALL detailed table cells matching exact forensic inspection
    amber_color = RGBColor(242, 174, 0)   # #F2AE00 exact presentation warm amber/gold
    red_color = RGBColor(193, 0, 27)        # #C1001B exact KeyBank corporate red
    navy_color = RGBColor(35, 41, 54)      # #232936 exact presentation dark navy

    # Header Row 0: Center horizontally and vertically (10.0pt Bold)
    for c_i in range(len(table.columns)):
        h_cell = table.cell(0, c_i)
        if not h_cell.is_spanned:
            h_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            for p in h_cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                p.font.name = "Arial"
                p.font.size = Pt(10.0)
                p.font.bold = True
                p.font.color.rgb = RGBColor(255, 255, 255) if c_i < 3 else navy_color
                for r in p.runs:
                    r.font.name = "Arial"
                    r.font.size = Pt(10.0)
                    r.font.bold = True

    # Data Rows 1 to 8: Executive category dividing borders and soft alternating fills
    for r_idx in range(1, len(table.rows)):
        is_cat_bottom = r_idx in (2, 4, 8)
        border_hex = '000000' if is_cat_bottom else None
        border_pt = 1.5 if is_cat_bottom else 0
        row_grey = (r_idx % 2 == 0)

        for c_idx in range(len(table.columns)):
            cell = table.cell(r_idx, c_idx)
            if cell.is_spanned:
                continue

            c_fill = None if (c_idx == 0 or not row_grey) else 'F6F7F9'
            # Divider lines are required across all columns including Column 0 ('Credit Risk Metrics')
            set_cell_border_and_fill(cell, fill_hex=c_fill, bottom_border_hex=border_hex, bottom_border_pt=border_pt)

            cell.margin_top = Inches(0.002)
            cell.margin_bottom = Inches(0.002)
            cell.margin_left = Inches(0.03)
            cell.margin_right = Inches(0.03)

            if c_idx in (0, 1, 2, 3):
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                for p in cell.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.CENTER if c_idx == 1 else PP_ALIGN.LEFT
                    p.font.name = "Arial"
                    p.font.size = Pt(10.0)
                    p.font.bold = True
                    p.font.color.rgb = navy_color
                    for run in p.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(10.0)
                        run.font.bold = True
                        run.font.color.rgb = navy_color

            elif c_idx in (4, 5, 6):
                m_font_pt = 10.0
                ln_spc = Pt(11.8)

                for p_idx, para in enumerate(cell.text_frame.paragraphs):
                    para.space_before = Pt(0)
                    para.space_after = Pt(0)
                    para.line_spacing = ln_spc
                    para.alignment = PP_ALIGN.LEFT

                    p_bold = True
                    p_color = amber_color if c_idx in (4, 5) else red_color

                    for run in para.runs:
                        run.font.name = "Arial"
                        run.font.size = Pt(m_font_pt)
                        run.font.bold = p_bold
                        run.font.color.rgb = p_color

    # Enforce Double-Sided Pure Black Dividing Lines at Row 2, Row 4, Row 8
    w_emu = int(1.5 * 12700) # 1.5pt solid pure black
    def _add_line(cell, border_tag):
        tcPr = cell._tc.get_or_add_tcPr()
        for el in list(tcPr):
            if el.tag.endswith(border_tag):
                tcPr.remove(el)
        ln = parse_xml(f'<a:{border_tag} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="{w_emu}" cmpd="sng"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:{border_tag}>')
        first_fill = next((el for el in list(tcPr) if any(el.tag.endswith(x) for x in ['solidFill', 'noFill', 'gradFill', 'blipFill'])), None)
        if first_fill is not None:
            first_fill.addprevious(ln)
        else:
            tcPr.append(ln)

    # 100% Complete, full-width solid pure black dividing lines across ALL columns (0 to 6):
    for c_i in range(len(table.columns)):
        # Category 1: Row 2 bottom & Row 3 top
        _add_line(table.cell(2, c_i), 'lnB')
        if len(table.rows) > 3:
            _add_line(table.cell(3, c_i), 'lnT')
        # Category 2: Row 4 bottom & Row 5 top
        if len(table.rows) > 4:
            _add_line(table.cell(4, c_i), 'lnB')
        if len(table.rows) > 5:
            _add_line(table.cell(5, c_i), 'lnT')
        # Category 3: Row 8 bottom
        if len(table.rows) > 8:
            _add_line(table.cell(8, c_i), 'lnB')

    # Enforce Column 0 vertically merged cells specifically so line is 100% full & complete from left edge:
    # Category 1: Total KeyCorp (R1C0 primary rowSpan=2, R2C0 vMerge=1)
    _add_line(table.cell(1, 0), 'lnB')
    _add_line(table.cell(2, 0), 'lnB')
    _add_line(table.cell(3, 0), 'lnT')
    # Category 2: Line of Business (R3C0 primary rowSpan=2, R4C0 vMerge=1)
    _add_line(table.cell(3, 0), 'lnB')
    _add_line(table.cell(4, 0), 'lnB')
    _add_line(table.cell(5, 0), 'lnT')
    # Category 3: Commercial Industry (R5C0 primary rowSpan=4, R8C0 vMerge=1)
    _add_line(table.cell(5, 0), 'lnB')
    _add_line(table.cell(8, 0), 'lnB')

    # 5. Row heights matching Ideal_slide_ppt.pptx exactly:
    row_heights = [0.4248, 0.3192, 0.3192, 0.3192, 1.0578, 0.6731, 0.3192, 0.8038, 0.4621]
    for i, h in enumerate(row_heights):
        if i < len(table.rows):
            table.rows[i].height = Inches(h)
    table._graphic_frame.height = Inches(sum(r.height for r in table.rows) / Inches(1))

    # 6. Ensure Footer Notes, Logo & Slide Number matching ppt.pptx exactly and positioned cleanly at the bottom
    # Deduplicate footer textboxes if multiple exist on the slide
    conf_shapes = [s for s in slide.shapes if s.has_text_frame and "classification: confidential" in s.text_frame.text.lower()]
    if len(conf_shapes) > 1:
        for extra_s in conf_shapes[1:]:
            sp = extra_s._element
            sp.getparent().remove(sp)

    num_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip() == "2"]
    if len(num_shapes) > 1:
        for extra_num in num_shapes[1:]:
            sp = extra_num._element
            sp.getparent().remove(sp)

    for s in slide.shapes:
        if s.has_text_frame and "classification: confidential" in s.text_frame.text.lower():
            s.top = Inches(7.12)
            s.left = Inches(1.50)
            s.width = Inches(10.30)
            s.height = Inches(0.38)
            s.text_frame.margin_top = s.text_frame.margin_bottom = s.text_frame.margin_left = s.text_frame.margin_right = 0
            for p_idx, p in enumerate(s.text_frame.paragraphs):
                p.space_before = Pt(0)
                p.space_after = Pt(0)
                # Big table ke niche jo text tilted ha wo bold ha
                if p_idx < 2:
                    p.font.name = "Arial"
                    p.font.size = Pt(8.0)
                    p.font.bold = True
                    p.font.italic = True
                    for r in p.runs:
                        r.font.name = "Arial"
                        r.font.size = Pt(8.0)
                        r.font.bold = True
                        r.font.italic = True
        elif s.has_text_frame and s.text_frame.text.strip() == "2":
            s.top = Inches(7.14)
            s.left = Inches(12.75)
        elif s.shape_type == MSO_SHAPE_TYPE.PICTURE and s.top > Inches(6.5):
            s.top = Inches(7.12)
            s.left = Inches(0.40)

    has_confidential = any(s.has_text_frame and "classification: confidential" in s.text_frame.text.lower() for s in slide.shapes)
    if not has_confidential:
        tb_f = slide.shapes.add_textbox(Inches(1.50), Inches(7.08), Inches(10.30), Inches(0.38))
        tff = tb_f.text_frame
        tff.word_wrap = True
        tff.margin_left = tff.margin_right = tff.margin_top = tff.margin_bottom = 0
        tff.clear()
        pf0 = tff.paragraphs[0]
        pf0.text = "Supporting information on metrics exceeding Risk Appetite Warning and in breach Risk Appetite Limit at the end of the deck"
        pf0.font.name = "Arial"; pf0.font.size = Pt(8.0); pf0.font.bold = True; pf0.font.italic = True; pf0.font.color.rgb = RGBColor(0, 0, 0)
        pf0.alignment = PP_ALIGN.CENTER; pf0.space_after = Pt(0); pf0.space_before = Pt(0)

        pf1 = tff.add_paragraph()
        pf1.text = "Note: Comparisons above reflect current ERM and Credit risk metric framework. For prior period breaches refer to historic Credit Risk Metric reports"
        pf1.font.name = "Arial"; pf1.font.size = Pt(8.0); pf1.font.bold = True; pf1.font.italic = True; pf1.font.color.rgb = RGBColor(0, 0, 0)
        pf1.alignment = PP_ALIGN.CENTER; pf1.space_after = Pt(0); pf1.space_before = Pt(0)

        pf2 = tff.add_paragraph()
        pf2.text = "CLASSIFICATION: CONFIDENTIAL"
        pf2.font.name = "Arial"; pf2.font.size = Pt(7.5); pf2.font.bold = False; pf2.font.italic = False; pf2.font.color.rgb = RGBColor(0, 0, 0)
        pf2.alignment = PP_ALIGN.CENTER; pf2.space_after = Pt(0); pf2.space_before = Pt(0)

    has_num = any(s.has_text_frame and s.text_frame.text.strip() == "2" for s in slide.shapes)
    if not has_num:
        tb_num = slide.shapes.add_textbox(Inches(12.75), Inches(7.10), Inches(0.4), Inches(0.3))
        p_num = tb_num.text_frame.paragraphs[0]
        p_num.text = "2"
        p_num.font.name = "Arial"; p_num.font.size = Pt(9.0); p_num.font.color.rgb = RGBColor(0, 0, 0)

    # 7. Ensure ALL text across the entire slide uses Arial font
    ensure_slide_text_arial(slide)

    return {
        'same_slide': True,
        'summary_table_updated': summary_shape is not None,
        'table_height_in': round(table._graphic_frame.height / Inches(1), 3),
        'table_top_in': round(table._graphic_frame.top / Inches(1), 3),
        'table_bottom_in': round((table._graphic_frame.top + table._graphic_frame.height) / Inches(1), 3),
        'extra_tables_created': 0
    }



def copy_slide_from_template_com(template_path, source_path, output_path, slide_index=2):
    """Natively copies slide_index from template_path into source_path at slide_index,
    replacing any placeholder/empty slide at that position, and saves to output_path.
    Uses PowerPoint COM for 100% native fidelity."""
    import win32com.client, shutil, os
    
    t_abs = os.path.abspath(str(template_path))
    s_abs = os.path.abspath(str(source_path))
    o_abs = os.path.abspath(str(output_path))
    
    if not os.path.isfile(t_abs):
        raise FileNotFoundError(f"Template PPTX not found: {t_abs}")
    if not os.path.isfile(s_abs):
        raise FileNotFoundError(f"Source PPTX not found: {s_abs}")
        
    shutil.copy2(s_abs, o_abs)
    
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    pres = None
    try:
        pres = ppt_app.Presentations.Open(o_abs, WithWindow=False)
        total_before = pres.Slides.Count
        insert_after = max(0, slide_index - 1)
        pres.Slides.InsertFromFile(t_abs, insert_after, slide_index, slide_index)
        if pres.Slides.Count > total_before and total_before >= slide_index:
            pres.Slides(slide_index + 1).Delete()
        pres.Save()
    finally:
        if pres is not None:
            try: pres.Close()
            except Exception: pass
        try:
            if ppt_app.Presentations.Count == 0:
                ppt_app.Quit()
        except Exception: pass
        
    return o_abs


def create_standalone_deck(output, audit, single_slide=False):
    """Build a complete, standalone, self-contained Credit Risk Executive PPTX from scratch
    without needing any template PPTX file.
    If single_slide is True: Slide 1 is the Credit Risk Slide.
    If single_slide is False (default): Slide 1 is Executive Cover, Slide 2 is the Credit Risk Slide.
    """
    output = Path(output).resolve()
    report = audit if isinstance(audit, dict) else json.loads(Path(audit).read_text(encoding='utf-8'))
    validate_report(report)
    quarter = report.get('quarter', '3Q26')

    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)

    if single_slide:
        prs.slides.add_slide(prs.slide_layouts[6])
    else:
        # Slide 1: Executive Cover
        s1 = prs.slides.add_slide(prs.slide_layouts[6])
        tb1 = s1.shapes.add_textbox(Inches(1.2), Inches(2.8), Inches(11.0), Inches(2.0))
        tf1 = tb1.text_frame
        p1 = tf1.paragraphs[0]
        p1.text = f"{quarter} Credit Risk Metrics Report"
        p1.font.name = "Arial"; p1.font.size = Pt(36); p1.font.bold = True; p1.font.color.rgb = RGBColor(35, 41, 54)
        p1_sub = tf1.add_paragraph()
        p1_sub.text = "Enterprise Risk Management & Credit Risk Administration"
        p1_sub.font.name = "Arial"; p1_sub.font.size = Pt(18); p1_sub.font.color.rgb = RGBColor(112, 128, 144)

        # Slide 2: Target Slide
        prs.slides.add_slide(prs.slide_layouts[6])

    temp_blank = output.with_name(f"temp_blank_{os.getpid()}_{uuid.uuid4().hex[:6]}.pptx")
    prs.save(str(temp_blank))

    try:
        res = patch_pptx(str(temp_blank), str(output), audit, same_slide=True, dry_run=False)
    finally:
        if temp_blank.exists():
            try:
                temp_blank.unlink()
            except Exception:
                pass
    return res


def patch_pptx(source, output, audit, dry_run=False, allow_possible_overflow=False,
               new_table_slide=False, include_count_slide=False,
               same_slide=False, count_reference_pptx=None):
    if source is None or str(source).lower() in ('none', 'standalone', ''):
        return create_standalone_deck(output, audit)
    source = Path(source).resolve(); output = Path(output).resolve()
    if not source.is_file():
        raise ValueError(f'Original PPTX not found: {source}')
    if source == output:
        raise ValueError('Refusing to overwrite the original PowerPoint. Use a different --output.')
    report = json.loads(Path(audit).read_text(encoding='utf-8'))
    validate_report(report)

    if include_count_slide and not new_table_slide and not same_slide:
        raise ValueError('COUNT summary requires dedicated slide or --same-slide')
    if same_slide and new_table_slide:
        raise ValueError('Cannot combine --same-slide with --new-table-slide')

    prs = Presentation(str(source))
    if len(prs.slides) == 1:
        target_slide = prs.slides[0]
        slide_num = 1
    elif len(prs.slides) >= 2:
        target_slide = prs.slides[1]
        slide_num = 2
    else:
        target_slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide_num = 1
    # Check if target_slide has the 9x7 detailed table
    tables = [s for s in target_slide.shapes if s.has_table and len(s.table.rows) == 9 and len(s.table.columns) == 7]
    if not tables:
        found_anywhere = False
        for idx_s, s in enumerate(prs.slides, start=1):
            t_cand = [sh for sh in s.shapes if sh.has_table and len(sh.table.rows) == 9 and len(sh.table.columns) == 7]
            if t_cand:
                target_slide = s
                tables = t_cand
                slide_num = idx_s
                found_anywhere = True
                break
        if not found_anywhere:
            # Slide is empty or non-template! Construct the full native Credit Risk Slide 2 on target_slide!
            shape = setup_empty_slide_for_credit_risk(target_slide, report, prs=prs)
            table = shape.table
            same_slide = True
        else:
            shape = tables[0]
            table = shape.table
    else:
        shape = tables[0]
        table = shape.table            
    warnings = []
    planned = []

    for idx, row in enumerate(report['rows'], start=1):
        source_metric_label = table.cell(idx, 2).text
        if norm(row['measure']) not in norm(source_metric_label):
            raise ValueError(f'Slide table row {idx} measure {source_metric_label!r} does not match audit {row["measure"]!r}')
        source_level_label = table.cell(idx, 1).text
        if not re.search(r'\b' + str(row['level']) + r'\b', source_level_label):
            raise ValueError(f'Slide table row {idx} Level {source_level_label!r} does not match audit {row["level"]!r}')
        left_names = [clean_metric_display_name(x['name'], idx) for x in row['yellow']]
        red_names = [clean_metric_display_name(x['name'], idx) for x in row['red']]
        yl, yr = fit_split_yellow(table, idx, left_names, font_pt=8.0)
        yellow_cells = [(4, yl)] + ([(5, yr)] if not table.cell(idx, 5).is_spanned and not table.cell(idx, 4).is_merge_origin else [])
        for col, names in yellow_cells + [(6, red_names)]:
            cell = table.cell(idx, col)
            if cell.is_spanned: continue
            font_pt = text_font_pt(cell, table.cell(3, 4) if col in (4, 5) else table.cell(3, 6))
            width = (table.columns[col].width / Inches(1))
            if col == 4 and (table.cell(idx, 4).is_merge_origin and table.cell(idx, 5).is_spanned):
                width += table.columns[5].width / Inches(1)
            lines = estimate_lines(names, width, font_pt)
            max_lines = (table.rows[idx].height / Inches(1)) * 72 / (font_pt * 1.20)
            if lines > max_lines + 0.5:
                warnings.append(f'Row {idx} col {col}: estimated {lines} text lines; room for {max_lines:.1f}. Check overflow.')
        pending_count = len(row.get('pending', []))
        green_label = f"{row['green']} of {row['total']} {'metric' if row['total']==1 else 'metrics'}"
        planned.append({'row': idx, 'group': [row['category'], row['level'], row['measure']],
                        'green': green_label, 'pending': pending_count,
                        'yellow_left': yl, 'yellow_right': yr, 'red': red_names})

    if warnings and not allow_possible_overflow and not new_table_slide and not same_slide:
        raise ValueError('Potential text overflow in original geometry: ' + '; '.join(warnings))

    if dry_run:
        return {'mode': 'dry-run', 'planned': planned, 'warnings': warnings, 'table_shape_id': shape.shape_id,
                'new_table_slide': new_table_slide, 'same_slide': same_slide}


    yellow_sample = next((table.cell(i, j) for i in range(1, 9) for j in (4, 5)
                          if not table.cell(i, j).is_spanned and table.cell(i, j).text.strip()), None)
    red_sample = next((table.cell(i, 6) for i in range(1, 9) if table.cell(i, 6).text.strip()), None)

    for spec in planned:
        r = spec['row']
        if not table.cell(r, 3).is_spanned:
            styled_lines(table.cell(r, 3), [spec['green']])
        
        # Row 5 (Commercial Industry Concentration):
        # Col 4 has 'CMT % of CET1:' header, Col 5 has matching blank spacer so Finance & Technology align at Line 1
        if r == 5 and spec['yellow_left']:
            if not table.cell(r, 4).is_spanned:
                styled_lines_with_header(table.cell(r, 4), 'CMT % of CET1:', spec['yellow_left'], yellow_sample)
            if not table.cell(r, 5).is_spanned and not table.cell(r, 4).is_merge_origin:
                styled_lines_with_header(table.cell(r, 5), ' ', spec['yellow_right'], yellow_sample)
        else:
            if not table.cell(r, 4).is_spanned:
                styled_lines(table.cell(r, 4), spec['yellow_left'], yellow_sample, bullet=True)
            if not table.cell(r, 5).is_spanned and not table.cell(r, 4).is_merge_origin:
                styled_lines(table.cell(r, 5), spec['yellow_right'], yellow_sample, bullet=True)
                
        if not table.cell(r, 6).is_spanned:
            styled_lines(table.cell(r, 6), spec['red'], red_sample, bullet=True)

    if same_slide:
        fit_info = apply_same_slide_layout(target_slide, table, planned, report, count_reference_pptx, prs=prs)
        warnings = []
    elif new_table_slide:
        newslide, moved_table = add_dedicated_slide(prs, shape, report['quarter'], report.get('draft', False),
                                                    report.get('totals', {}).get('pending', 0))
        fit_info = dedicated_layout(prs, moved_table.table, planned, report['quarter'])
        if include_count_slide:
            label_historical_slide(target_slide)
            add_count_slide(prs, report)
        warnings = []
    else:
        fit_info = None

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))

    check = Presentation(str(output))
    if not new_table_slide and len(check.slides) != len(prs.slides):
        raise AssertionError('Slide count changed unexpectedly in same-slide mode')

    return {
        'mode': 'saved',
        'output': str(output),
        'planned': planned,
        'warnings': warnings,
        'table_shape_id': shape.shape_id,
        'same_slide': same_slide,
        'new_table_slide': new_table_slide,
        'table_slide_number': 2 if same_slide else ((4 if include_count_slide else 3) if new_table_slide else 2),
        'layout_info': fit_info
    }


def main():
    p = argparse.ArgumentParser(description='Patch the original native PPT table with single or dedicated slide layout')
    p.add_argument('--pptx', default=None, help='Original CRM Report Slides for BOT.pptx (optional in standalone mode)')
    p.add_argument('--standalone', action='store_true', help='Generate standalone executive deck without template')
    p.add_argument('--audit', required=True, help='Verified 2Q26 or 3Q26 .audit.json from generator')
    p.add_argument('--output', required=True, help='New output PPTX; original will not be overwritten')
    p.add_argument('--dry-run', action='store_true')
    p.add_argument('--allow-possible-overflow', action='store_true')
    p.add_argument('--new-table-slide', action='store_true', help='Move main table to a new dedicated Slide 3')
    p.add_argument('--same-slide', action='store_true', help='Keep count summary and detailed table on existing Slide 2')
    p.add_argument('--count-reference-pptx', help='Optional path to Final_Deck_3Q26_Native_Updated.pptx for compact table design')
    a = p.parse_args()
    if a.standalone or not a.pptx:
        result = create_standalone_deck(a.output, a.audit)
    else:
        result = patch_pptx(a.pptx, a.output, a.audit, a.dry_run, a.allow_possible_overflow,
                            new_table_slide=a.new_table_slide, same_slide=a.same_slide,
                            count_reference_pptx=a.count_reference_pptx)
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr); sys.exit(1)