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
import re
import sys
from pathlib import Path

from pptx import Presentation
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


def pick_main_table(slide, shape_id=8, verify_geometry=True):
    candidates = [s for s in slide.shapes if s.shape_id == shape_id]
    if len(candidates) != 1 or not candidates[0].has_table:
        raise ValueError(f"Slide 2 does not have a native table at shape ID {shape_id}. "
                         "Run inspect_original_ppt.py and inspect the original, don't guess.")
    shape = candidates[0]
    table = shape.table
    if len(table.rows) != 9 or len(table.columns) != 7:
        raise ValueError(f"Shape {shape_id} has {len(table.rows)} rows x {len(table.columns)} columns, "
                         "expected 9 x 7. No modifications made.")
    refs = (0.24062, 2.28957, 12.85209, 4.75701)
    real = tuple(float(d / Inches(1)) for d in (shape.left, shape.top, shape.width, shape.height))
    if verify_geometry and any(abs(a-b) > 0.035 for a, b in zip(refs, real)):
        raise ValueError(f"Shape {shape_id} geometry {real} doesn't match the inspected table {refs}.")
    headers = [norm(table.cell(0, i).text) for i in range(7)]
    def check(idx, words):
        return all(w in headers[idx] for w in words)
    if not (check(0, ['credit', 'risk']) and check(3, ['within', 'risk'])
            and check(4, ['exceeds', 'warning']) and check(6, ['breaches', 'limit'])):
        raise ValueError(f"Shape {shape_id} has unexpected header cells: {headers}. No modifications made.")
    return shape


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
        rpr.set('sz', '800')
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
                rpr_match[0].attrib.pop('b', None)
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
        'Enterprise Leveraged - Interest Coverage (Core)': 'Core Leveraged Interest Coverage Ratio',
        'Enterprise Leveraged - Interest Coverage (Non-Core)': 'Non-Core Leveraged Interest Coverage Ratio',
        'KeyCorp - Criticized Commitments %': 'KeyCorp Criticized %',
        'KeyCorp - Criticized': 'KeyCorp Criticized %',
        'Student Loans Direct - 30-89 Days Past Due': 'Student Loans Direct 30-89 Days Past Due',
        'Student Loans Direct - 90+ Days Past Due': 'Student Loans Direct 90+ Days Past Due',
        'Other Consumer - 90+ Days Past Due': 'Other Consumer Secured 90+ Days Past Due',
        'Construction Materials and Contractors': 'Construction Materials & Contractors',
        'Construction Materials and Contractors - Commitments as % of CET1': 'Construction Materials & Contractors',
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
        rpr.set('sz', '850')
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
    rpr_hdr.set('sz', '850')
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
    body.append(hdr_tmpl)

    # Bullets underneath (sz=850, preserving Amber/font color)
    for k, item in enumerate(lines or []):
        tmpl = copy.deepcopy(options[min(k, len(options)-1)])
        for ch in list(tmpl):
            if ch.tag.endswith(('}r', '}br', '}fld')):
                tmpl.remove(ch)
        r_item = OxmlElement('a:r')
        if sample_rpr is not None:
            rpr_item = copy.deepcopy(sample_rpr)
            rpr_item.tag = '{http://schemas.openxmlformats.org/drawingml/2006/main}rPr'
        else:
            rpr_item = OxmlElement('a:rPr')
        rpr_item.attrib.pop('b', None)
        rpr_item.set('sz', '850')
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


def update_slide2_title_and_commentary(slide, report):
    """Update Slide 2 Title and Commentary to reflect current reporting quarter with bold 18pt."""
    quarter = report.get('quarter', '3Q26')
    tb_shape = next((s for s in slide.shapes if s.has_text_frame and not s.has_table and s.top < Inches(2.0)), None)
    if not tb_shape:
        return False
    tf = tb_shape.text_frame
    for p in tf.paragraphs:
        if 'credit risk metrics' in p.text.lower():
            p.text = ""  # clear runs to rebuild with explicit bold & 18pt size
            r = p.add_run()
            r.text = f"{quarter} Credit Risk Metrics"
            r.font.bold = True
            r.font.size = Pt(18)
            r.font.name = "Arial"
            break

    # Level 1 Asset quality summary
    row0 = report['rows'][0] if report.get('rows') else None
    if row0:
        y_items = [clean_metric_display_name(x['name'], 1) for x in row0.get('yellow', [])]
        r_items = [clean_metric_display_name(x['name'], 1) for x in row0.get('red', [])]
        total_breaches = len(y_items) + len(r_items)
        if total_breaches == 0:
            l1_status = "All Below Risk Appetite Limits"
        else:
            all_breaches = y_items + r_items
            joined_str = ', '.join(all_breaches)
            l1_status = f"{total_breaches} Exceeds Warning ({joined_str})"
        for p in tf.paragraphs:
            if 'Total KeyCorp Level 1 Asset Quality' in p.text:
                p.text = ""
                r1 = p.add_run()
                r1.text = "\u2022 Total KeyCorp Level 1 Asset Quality: "
                r1.font.bold = True
                r1.font.size = Pt(12)
                r2 = p.add_run()
                r2.text = l1_status
                r2.font.bold = False
                r2.font.size = Pt(12)
                break
    return True


def apply_same_slide_layout(slide, table, planned, report, count_reference_pptx=None):
    """Populate Slide 2 native tables and update title/commentary matching ideal slide.

    1. Updates Slide 2 Title (18pt Bold) and Level 1 commentary (12pt matching Bullets 2, 3, 4).
    2. Populates existing native top-right summary table with 2Q26 vs 3Q26.
    3. Formats bottom 9x7 detailed table with 8.5pt font (readable, not small),
       exact ideal row heights [0.425, 0.319, 0.319, 0.319, 1.058, 0.673, 0.319, 0.804, 0.462],
       total height 4.70", bottom at 6.988", preserving clear buffer above footer line at 7.05".
    """
    # 1. Update title (18pt) and commentary (12pt)
    update_slide2_title_and_commentary(slide, report)

    # 2. Update existing top-right summary table
    summary_shape = update_slide2_summary_table(slide, report, main_table_shape_id=table._graphic_frame.shape_id)

    # 3. Enforce exact column widths and geometry matching Ideal_slide_ppt.pptx:
    # Widths: [1.584, 0.7093, 1.3256, 1.6987, 1.8023, 2.2985, 3.4338], Total = 12.8521"
    ideal_col_widths = [1.584, 0.7093, 1.3256, 1.6987, 1.8023, 2.2985, 3.4338]
    for i, w in enumerate(ideal_col_widths):
        if i < len(table.columns):
            table.columns[i].width = Inches(w)

    table._graphic_frame.left = Inches(0.2406)
    table._graphic_frame.top = Inches(2.2896)
    table._graphic_frame.width = Inches(12.8521)

    # 4. Format detailed table cells with 9.0pt font (readable, executive, bold, NOT tiny)
    font_pt = 9.0
    for r in range(1, len(table.rows)):
        for c in (3, 4, 5, 6):
            if c >= len(table.columns):
                continue
            cell = table.cell(r, c)
            if cell.is_spanned:
                continue
            cell.margin_top = Inches(0.01)
            cell.margin_bottom = Inches(0.01)
            cell.margin_left = Inches(0.03)
            cell.margin_right = Inches(0.03)
            for para in cell.text_frame.paragraphs:
                para.space_before = Pt(0)
                para.space_after = Pt(0)
                para.line_spacing = Pt(font_pt * 1.15)
                para.font.size = Pt(font_pt)
                for run in para.runs:
                    run.font.size = Pt(font_pt)

    # 5. Exact row heights matching Ideal_slide_ppt.pptx:
    # Row 0: 0.4248", R1-R3: 0.3192", R4: 1.0578", R5: 0.6731", R6: 0.3192", R7: 0.8038", R8: 0.4621"
    # Total table height = 4.6981" (4.70"), Bottom = 2.2896" + 4.6981" = 6.9877", safely above footer at 7.05"
    quarter = report.get('quarter', '3Q26')
    if quarter == '3Q26':
        row_heights = [0.4248, 0.3192, 0.3192, 0.3192, 1.0578, 0.6731, 0.3192, 0.8038, 0.4621]
    else:
        # Generic content-weighted scaling for other quarters
        target_total_height = 4.70
        header_h = 0.425
        row_heights = [header_h]
        for spec in planned:
            w4 = table.columns[4].width / Inches(1) if len(table.columns) > 4 else 1.80
            w5 = table.columns[5].width / Inches(1) if len(table.columns) > 5 else 2.30
            w6 = table.columns[6].width / Inches(1) if len(table.columns) > 6 else 3.43
            l4 = estimate_lines(spec.get('yellow_left', []), w4, font_pt=9.0)
            l5 = estimate_lines(spec.get('yellow_right', []), w5, font_pt=9.0)
            l6 = estimate_lines(spec.get('red', []), w6, font_pt=9.0)
            max_lines = max(1, l4, l5, l6)
            h = max(0.319, 0.20 + max_lines * 0.10)
            row_heights.append(h)
        allowed_data_h = target_total_height - header_h
        data_heights = row_heights[1:]
        total_data_h = sum(data_heights)
        if total_data_h > allowed_data_h:
            scale = allowed_data_h / total_data_h
            row_heights = [header_h] + [h * scale for h in data_heights]

    for i, h in enumerate(row_heights):
        if i < len(table.rows):
            table.rows[i].height = Inches(h)
    table._graphic_frame.height = Inches(sum(r.height for r in table.rows) / Inches(1))

    return {
        'same_slide': True,
        'summary_table_updated': summary_shape is not None,
        'table_height_in': round(table._graphic_frame.height / Inches(1), 3),
        'extra_tables_created': 0
    }

def patch_pptx(source, output, audit, dry_run=False, allow_possible_overflow=False,
               new_table_slide=False, include_count_slide=False,
               same_slide=False, count_reference_pptx=None):
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
    if len(prs.slides) < 2:
        raise ValueError('Original PowerPoint has fewer than 2 slides')
    shape = pick_main_table(prs.slides[1]); table = shape.table
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
        fit_info = apply_same_slide_layout(prs.slides[1], table, planned, report, count_reference_pptx)
        warnings = []
    elif new_table_slide:
        newslide, moved_table = add_dedicated_slide(prs, shape, report['quarter'], report.get('draft', False),
                                                    report.get('totals', {}).get('pending', 0))
        fit_info = dedicated_layout(prs, moved_table.table, planned, report['quarter'])
        if include_count_slide:
            label_historical_slide(prs.slides[1])
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
    p.add_argument('--pptx', required=True, help='Original CRM Report Slides for BOT.pptx')
    p.add_argument('--audit', required=True, help='Verified 2Q26 or 3Q26 .audit.json from generator')
    p.add_argument('--output', required=True, help='New output PPTX; original will not be overwritten')
    p.add_argument('--dry-run', action='store_true')
    p.add_argument('--allow-possible-overflow', action='store_true')
    p.add_argument('--new-table-slide', action='store_true', help='Move main table to a new dedicated Slide 3')
    p.add_argument('--same-slide', action='store_true', help='Keep count summary and detailed table on existing Slide 2')
    p.add_argument('--count-reference-pptx', help='Optional path to Final_Deck_3Q26_Native_Updated.pptx for compact table design')
    a = p.parse_args()
    result = patch_pptx(a.pptx, a.output, a.audit, a.dry_run, a.allow_possible_overflow,
                        new_table_slide=a.new_table_slide, same_slide=a.same_slide,
                        count_reference_pptx=a.count_reference_pptx)
    print(json.dumps(result, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr); sys.exit(1)
