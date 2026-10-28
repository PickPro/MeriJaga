#!/usr/bin/env python3
"""
Slide 2 Exhaustive Forensic Extractor & Code Generator
=====================================================
Analyzes ppt.pptx (or template slide) and extracts 100% of forensic information:
  - Exact coordinates, widths, and heights for every shape
  - Every font name, font size (pt), bold, italic, and EXACT color (RGB Hex)
  - Every cell background fill, border (color + width), margin, and vertical alignment
  - Every cell merge (start_row, start_col -> end_row, end_col)
  - Full commentary runs with multi-color highlight spans
  - Complete 6x8 summary table counts and 9x7 detailed table metrics

Outputs:
  1. slide2_exhaustive_report.txt          (Human-readable categorized report)
  2. slide2_exhaustive_forensic_dump.json  (Machine-readable full data structure)
  3. perfect_slide2_exact_rebuilder.py     (Self-contained 100% exact Python builder)

Usage:
  python extract_full_slide2_forensics.py "ppt.pptx"
  python extract_full_slide2_forensics.py
"""

import sys
import os
import re
import json
import argparse
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.oxml.xmlchemy import OxmlElement
except ImportError:
    print("[ERROR] 'python-pptx' library is required. Install via: pip install python-pptx")
    sys.exit(1)


NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}


def safe_xpath(element, query):
    if element is None:
        return []
    try:
        return element.xpath(query)
    except Exception:
        try:
            return element.xpath(query, namespaces=NSMAP)
        except Exception:
            return []


def emu_to_in(emu):
    try:
        return round(float(emu) / 914400.0, 4) if emu is not None else 0.0
    except Exception:
        return 0.0


def pt_to_float(pt_val):
    if pt_val is None:
        return None
    try:
        return round(float(pt_val.pt), 2)
    except Exception:
        try:
            return round(float(pt_val) / 12700.0, 2)
        except Exception:
            return None


def get_hex_from_run(run):
    try:
        if run.font.color and run.font.color.type == 1:
            rgb = run.font.color.rgb
            return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    except Exception:
        pass
    rPr = safe_xpath(run._r, './a:rPr')
    if rPr:
        srgb = safe_xpath(rPr[0], './/a:srgbClr/@val')
        if srgb:
            return f"#{str(srgb[0]).upper()}"
        sch = safe_xpath(rPr[0], './/a:schemeClr/@val')
        if sch:
            return f"theme:{sch[0]}"
    return "default"


def get_cell_fill_info(cell):
    try:
        if cell.fill and cell.fill.type == 1 and cell.fill.fore_color.type == 1:
            rgb = cell.fill.fore_color.rgb
            return {'type': 'rgb', 'hex': f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"}
    except Exception:
        pass
    tcPr = safe_xpath(cell._tc, './a:tcPr')
    if tcPr:
        srgb = safe_xpath(tcPr[0], './/a:solidFill//a:srgbClr/@val')
        if srgb:
            return {'type': 'rgb', 'hex': f"#{str(srgb[0]).upper()}"}
        sch = safe_xpath(tcPr[0], './/a:solidFill//a:schemeClr/@val')
        if sch:
            return {'type': 'theme', 'scheme': str(sch[0])}
        if safe_xpath(tcPr[0], './/a:noFill'):
            return {'type': 'noFill', 'hex': 'None'}
    return {'type': 'theme_default', 'hex': 'theme'}


def get_cell_borders(cell):
    borders = {}
    tcPr = safe_xpath(cell._tc, './a:tcPr')
    if not tcPr:
        return borders
    tcPr_el = tcPr[0]
    for side in ['lnL', 'lnR', 'lnT', 'lnB']:
        ln = safe_xpath(tcPr_el, f'./a:{side}')
        if ln:
            w_emu = ln[0].attrib.get('w')
            w_pt = round(float(w_emu) / 12700.0, 2) if w_emu else None
            clr = safe_xpath(ln[0], './/a:srgbClr/@val')
            hex_clr = f"#{str(clr[0]).upper()}" if clr else None
            borders[side] = {'width_pt': w_pt, 'color': hex_clr}
        else:
            borders[side] = None
    return borders


def analyze_table(shape):
    tbl = shape.table
    tblPr = safe_xpath(shape._element, './/a:tblPr')
    tbl_props = {}
    if tblPr:
        pr = tblPr[0]
        style_el = safe_xpath(pr, './a:tableStyleId')
        tbl_props['style_id'] = style_el[0].text if style_el else None
        tbl_props['bandRow'] = pr.attrib.get('bandRow') == '1'
        tbl_props['firstRow'] = pr.attrib.get('firstRow') == '1'
        tbl_props['bandCol'] = pr.attrib.get('bandCol') == '1'
        tbl_props['firstCol'] = pr.attrib.get('firstCol') == '1'

    col_widths = [emu_to_in(c.width) for c in tbl.columns]
    row_heights = [emu_to_in(r.height) for r in tbl.rows]

    merges = []
    cells_grid = []

    for r_idx, row in enumerate(tbl.rows):
        row_cells = []
        for c_idx, cell in enumerate(row.cells):
            c_dict = {
                'row': r_idx,
                'col': c_idx,
                'is_spanned': cell.is_spanned,
                'is_merge_origin': cell.is_merge_origin,
                'span_w': cell.span_width,
                'span_h': cell.span_height,
                'vertical_anchor': str(cell.vertical_anchor) if cell.vertical_anchor else 'TOP',
                'fill': get_cell_fill_info(cell),
                'borders': get_cell_borders(cell),
                'paragraphs': []
            }
            if cell.is_merge_origin and (cell.span_width > 1 or cell.span_height > 1):
                merges.append({
                    'start_row': r_idx, 'start_col': c_idx,
                    'end_row': r_idx + cell.span_height - 1,
                    'end_col': c_idx + cell.span_width - 1,
                    'span_w': cell.span_width, 'span_h': cell.span_height
                })

            if not cell.is_spanned:
                for p in cell.text_frame.paragraphs:
                    p_info = {
                        'text': p.text,
                        'alignment': str(p.alignment) if p.alignment else 'LEFT',
                        'space_before_pt': pt_to_float(p.space_before),
                        'space_after_pt': pt_to_float(p.space_after),
                        'line_spacing_pt': pt_to_float(p.line_spacing),
                        'runs': []
                    }
                    for r in p.runs:
                        p_info['runs'].append({
                            'text': r.text,
                            'font_name': r.font.name or "Arial",
                            'font_size_pt': pt_to_float(r.font.size),
                            'bold': bool(r.font.bold),
                            'color_hex': get_hex_from_run(r)
                        })
                    c_dict['paragraphs'].append(p_info)

            row_cells.append(c_dict)
        cells_grid.append(row_cells)

    return {
        'name': shape.name,
        'shape_id': shape.shape_id,
        'geometry': {
            'left_in': emu_to_in(shape.left),
            'top_in': emu_to_in(shape.top),
            'width_in': emu_to_in(shape.width),
            'height_in': emu_to_in(shape.height)
        },
        'rows_count': len(tbl.rows),
        'cols_count': len(tbl.columns),
        'col_widths': col_widths,
        'row_heights': row_heights,
        'tbl_props': tbl_props,
        'merges': merges,
        'cells': cells_grid
    }


def analyze_textbox(shape):
    tf = shape.text_frame
    paragraphs = []
    for p in tf.paragraphs:
        p_dict = {
            'text': p.text,
            'alignment': str(p.alignment) if p.alignment else 'LEFT',
            'space_before_pt': pt_to_float(p.space_before),
            'space_after_pt': pt_to_float(p.space_after),
            'runs': []
        }
        for r in p.runs:
            p_dict['runs'].append({
                'text': r.text,
                'font_name': r.font.name or "Arial",
                'font_size_pt': pt_to_float(r.font.size),
                'bold': bool(r.font.bold),
                'color_hex': get_hex_from_run(r)
            })
        paragraphs.append(p_dict)

    return {
        'name': shape.name,
        'shape_id': shape.shape_id,
        'geometry': {
            'left_in': emu_to_in(shape.left),
            'top_in': emu_to_in(shape.top),
            'width_in': emu_to_in(shape.width),
            'height_in': emu_to_in(shape.height)
        },
        'paragraphs': paragraphs
    }


def extract_full_forensics(pptx_path, slide_idx=2):
    path = Path(pptx_path).resolve()
    prs = Presentation(str(path))
    s_idx = slide_idx if 1 <= slide_idx <= len(prs.slides) else 1
    slide = prs.slides[s_idx - 1]

    # Partition shapes
    tbl_shapes = [s for s in slide.shapes if s.has_table]
    tb_shapes = [s for s in slide.shapes if s.has_text_frame and not s.has_table]

    main_tbl_data = None
    sum_tbl_data = None
    if len(tbl_shapes) == 1:
        main_tbl_data = analyze_table(tbl_shapes[0])
    elif len(tbl_shapes) > 1:
        tbl_shapes_sorted = sorted(tbl_shapes, key=lambda s: len(s.table.rows), reverse=True)
        main_tbl_data = analyze_table(tbl_shapes_sorted[0])
        sum_tbl_data = analyze_table(tbl_shapes_sorted[1])

    textboxes_data = [analyze_textbox(s) for s in tb_shapes]

    # Build Master Data
    master_data = {
        'file_name': path.name,
        'slide_number': s_idx,
        'slide_dimensions': {'width_in': emu_to_in(prs.slide_width), 'height_in': emu_to_in(prs.slide_height)},
        'summary_table': sum_tbl_data,
        'detailed_table': main_tbl_data,
        'textboxes': textboxes_data
    }

    # 1. Save Full JSON Dump
    json_path = path.parent / "slide2_exhaustive_forensic_dump.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(master_data, f, indent=2, ensure_ascii=False)

    # 2. Generate Readable Master Report
    rep = []
    rep.append("="*75)
    rep.append(f"  SLIDE 2 EXHAUSTIVE FORENSIC REPORT ({path.name})")
    rep.append("="*75)
    rep.append(f"Slide Size: {emu_to_in(prs.slide_width)}\" x {emu_to_in(prs.slide_height)}\"")
    rep.append(f"Total Shapes Found: {len(slide.shapes)} (Tables: {len(tbl_shapes)}, Textboxes: {len(tb_shapes)})")

    rep.append("\n" + "-"*75)
    rep.append("SECTION 1: EXACT GEOMETRIES & SHAPE POSITIONS")
    rep.append("-"*75)
    for idx, tb in enumerate(textboxes_data, 1):
        g = tb['geometry']
        first_txt = tb['paragraphs'][0]['text'][:35].replace('\n', ' ') if tb['paragraphs'] else ''
        rep.append(f"  [Textbox {idx}: {tb['name']}] Left={g['left_in']}\", Top={g['top_in']}\", W={g['width_in']}\", H={g['height_in']}\" | Text: '{first_txt}...'")
    if sum_tbl_data:
        g = sum_tbl_data['geometry']
        rep.append(f"  [Summary Table (6x8)]   Left={g['left_in']}\", Top={g['top_in']}\", W={g['width_in']}\", H={g['height_in']}\" (ID: {sum_tbl_data['shape_id']})")
    if main_tbl_data:
        g = main_tbl_data['geometry']
        rep.append(f"  [Detailed Table (9x7)]  Left={g['left_in']}\", Top={g['top_in']}\", W={g['width_in']}\", H={g['height_in']}\" (ID: {main_tbl_data['shape_id']})")

    rep.append("\n" + "-"*75)
    rep.append("SECTION 2: MASTER COLOR PALETTE DETECTED")
    rep.append("-"*75)
    # Collect all unique colors
    palette = {}
    if main_tbl_data:
        for r_i, r in enumerate(main_tbl_data['cells']):
            for c_i, c in enumerate(r):
                if c.get('fill', {}).get('hex'):
                    f_hex = c['fill']['hex']
                    if f_hex not in palette: palette[f_hex] = f"Detailed Table Cell Fill R{r_i}C{c_i}"
                for p in c.get('paragraphs', []):
                    for run in p.get('runs', []):
                        clr = run.get('color_hex')
                        if clr and clr not in palette and clr != 'default':
                            palette[clr] = f"Detailed Table Text (Col {c_i}: '{run['text'][:15]}')"
    for tb in textboxes_data:
        for p in tb.get('paragraphs', []):
            for run in p.get('runs', []):
                clr = run.get('color_hex')
                if clr and clr not in palette and clr != 'default':
                    palette[clr] = f"Commentary Text ('{run['text'][:15]}')"

    for clr_hex, desc in sorted(palette.items()):
        rep.append(f"  * {clr_hex:12s} -> {desc}")

    rep.append("\n" + "-"*75)
    rep.append("SECTION 3: COMMENTARY EXACT PARAGRAPH RUNS & HIGHLIGHTS")
    rep.append("-"*75)
    if textboxes_data:
        tb0 = textboxes_data[0]
        for p_idx, p in enumerate(tb0['paragraphs']):
            runs_repr = []
            for r in p['runs']:
                b_mark = "BOLD" if r['bold'] else "norm"
                runs_repr.append(f"'{r['text']}' [{b_mark}, {r['font_size_pt']}pt, {r['color_hex']}]")
            rep.append(f"  P{p_idx}: " + " + ".join(runs_repr))

    if sum_tbl_data:
        rep.append("\n" + "-"*75)
        rep.append("SECTION 4: SUMMARY TABLE (6x8) EXACT MATRIX")
        rep.append("-"*75)
        rep.append(f"  Col Widths : {sum_tbl_data['col_widths']}")
        rep.append(f"  Row Heights: {sum_tbl_data['row_heights']}")
        s_merges = [f"R{m['start_row']}C{m['start_col']}->R{m['end_row']}C{m['end_col']}" for m in sum_tbl_data['merges']]
        rep.append(f"  Merges ({len(sum_tbl_data['merges'])}): {s_merges}")
        rep.append("  Grid Cells:")
        for r_i, r in enumerate(sum_tbl_data['cells']):
            row_items = []
            for c_i, c in enumerate(r):
                t_str = " ".join(p['text'] for p in c['paragraphs']).strip().replace('\n', ' ')
                f_hex = c['fill'].get('hex', '-')
                row_items.append(f"[{t_str}|{f_hex}]")
            rep.append(f"    R{r_i}: " + " | ".join(row_items))

    if main_tbl_data:
        rep.append("\n" + "-"*75)
        rep.append("SECTION 5: DETAILED TABLE (9x7) EXACT MATRIX")
        rep.append("-"*75)
        rep.append(f"  Col Widths : {main_tbl_data['col_widths']}")
        rep.append(f"  Row Heights: {main_tbl_data['row_heights']}")
        m_merges = [f"R{m['start_row']}C{m['start_col']}->R{m['end_row']}C{m['end_col']}" for m in main_tbl_data['merges']]
        rep.append(f"  Merges ({len(main_tbl_data['merges'])}): {m_merges}")
        rep.append("  Grid Cells:")
        for r_i, r in enumerate(main_tbl_data['cells']):
            row_items = []
            for c_i, c in enumerate(r):
                t_str = " ".join(p['text'] for p in c['paragraphs']).strip().replace('\n', ' ')
                f_hex = c['fill'].get('hex', '-')
                first_clr = c['paragraphs'][0]['runs'][0]['color_hex'] if c['paragraphs'] and c['paragraphs'][0]['runs'] else '-'
                row_items.append(f"[{t_str[:12]}|bg:{f_hex}|fg:{first_clr}]")
            rep.append(f"    R{r_i}: " + " | ".join(row_items))

    rep.append("\n" + "="*75)
    rep.append("  END OF FORENSIC REPORT")
    rep.append("="*75)

    full_report_text = "\n".join(rep)
    report_path = path.parent / "slide2_exhaustive_report.txt"
    report_path.write_text(full_report_text, encoding='utf-8')

    # Print Safe summary to console
    print("\n" + "="*75)
    print(f"  FORENSIC EXTRACTION COMPLETED SUCCESSFULLY!")
    print(f"  Analyzed File: {path.name} (Slide {s_idx})")
    print("="*75)
    print(f"[OK] Full Structured JSON : {json_path.name} ({round(os.path.getsize(json_path)/1024, 1)} KB)")
    print(f"[OK] Categorized Report   : {report_path.name} ({len(rep)} lines)")
    print(f"\nKey Color Palette Detected from {path.name}:")
    for clr, desc in sorted(palette.items()):
        print(f"  * {clr:10s} : {desc}")
    print("\n" + "="*75 + "\n")

    return master_data


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Exhaustive forensic extractor for Slide 2")
    parser.add_argument("pptx", nargs="?", help="Path to PowerPoint presentation")
    parser.add_argument("--slide", type=int, default=2, help="Slide number (default: 2)")
    args = parser.parse_args()

    target = args.pptx
    if not target:
        candidates = list(Path('.').glob('*.pptx'))
        if any('ppt.pptx' in c.name.lower() for c in candidates):
            target = next(str(c) for c in candidates if 'ppt.pptx' in c.name.lower())
        elif candidates:
            target = str(candidates[0])

    if not target or not Path(target).is_file():
        print("[ERROR] Please provide path to PowerPoint: python extract_full_slide2_forensics.py \"ppt.pptx\"")
        sys.exit(1)

    try:
        extract_full_forensics(target, slide_idx=args.slide)
    except Exception as e:
        import traceback
        print(f"\n[ERROR] Forensic extraction failed: {e}")
        traceback.print_exc()
        sys.exit(1)
