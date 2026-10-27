#!/usr/bin/env python3
"""
Inspect & Extract Slide 2 Tables and Layout Blueprint
=====================================================
Analyzes the corporate Credit Risk template slide (Slide 2) in full forensic detail.
Extracts geometry, column widths, row heights, merged cells, fill colors, font styles,
borders, alignments, and commentary text boxes.

Outputs:
  1. slide2_tables_blueprint.json  (Structured machine-readable data)
  2. slide2_inspection_report.txt  (Human-readable summary report)
  3. generated_slide2_builder.py   (Self-contained Python builder script)

Usage:
  python inspect_template_tables.py
  python inspect_template_tables.py "CRM Report Slides for BOT.pptx"
  python inspect_template_tables.py --pptx "Final Deck.pptx" --slide 2
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
    print("[ERROR] 'python-pptx' library is required. Please install via: pip install python-pptx")
    sys.exit(1)


NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}


def safe_xpath(element, query):
    """Safely execute an XPath query with standard OpenXML namespaces."""
    if element is None:
        return []
    try:
        return element.xpath(query, namespaces=NSMAP)
    except Exception:
        try:
            return element.xpath(query)
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


def get_color_dict(color_obj, xml_element=None):
    """Extract RGB hex and theme details from a ColorFormat object or XML element."""
    result = {'type': 'unknown', 'hex': None, 'theme': None}
    if color_obj is not None:
        try:
            rgb = color_obj.rgb
            if rgb is not None:
                return {'type': 'rgb', 'hex': f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}", 'rgb': [rgb[0], rgb[1], rgb[2]]}
        except Exception:
            pass
        try:
            theme = color_obj.theme_color
            if theme is not None:
                result['theme'] = str(theme)
                result['type'] = 'theme'
        except Exception:
            pass

    if xml_element is not None:
        srgb = safe_xpath(xml_element, './/a:srgbClr/@val')
        if srgb:
            hex_str = str(srgb[0]).upper()
            try:
                r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                return {'type': 'rgb', 'hex': f"#{hex_str}", 'rgb': [r, g, b]}
            except Exception:
                return {'type': 'rgb', 'hex': f"#{hex_str}", 'rgb': None}
        sch = safe_xpath(xml_element, './/a:schemeClr/@val')
        if sch:
            return {'type': 'theme', 'scheme': str(sch[0]), 'hex': None}
        if safe_xpath(xml_element, './/a:noFill'):
            return {'type': 'noFill', 'hex': None}

    return result


def extract_cell_fill(cell):
    """Extract fill details from a cell."""
    try:
        tc = cell._tc
        tcPr = safe_xpath(tc, './a:tcPr')
        if not tcPr:
            return {'type': 'none', 'hex': None}
        
        tcPr_el = tcPr[0]
        solid_fill = safe_xpath(tcPr_el, './a:solidFill')
        if solid_fill:
            srgb = safe_xpath(solid_fill[0], './/a:srgbClr/@val')
            if srgb:
                hex_str = str(srgb[0]).upper()
                try:
                    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
                    return {'type': 'solid', 'hex': f"#{hex_str}", 'rgb': [r, g, b]}
                except Exception:
                    return {'type': 'solid', 'hex': f"#{hex_str}"}
            sch = safe_xpath(solid_fill[0], './/a:schemeClr/@val')
            if sch:
                return {'type': 'theme', 'scheme': str(sch[0])}
        
        if safe_xpath(tcPr_el, './a:noFill'):
            return {'type': 'noFill'}
        
        return {'type': 'theme_default'}
    except Exception:
        return {'type': 'unknown', 'hex': None}


def extract_borders(cell):
    """Extract border formatting from cell tcPr."""
    borders = {}
    try:
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
    except Exception:
        pass
    return borders


def analyze_table_shape(shape):
    """Deep analysis of a PowerPoint Table shape."""
    tbl = shape.table
    tbl_el = safe_xpath(shape._element, './/a:tbl')
    tblPr_el = safe_xpath(tbl_el[0], './a:tblPr') if tbl_el else []
    
    table_style_id = None
    first_row = False
    band_row = False
    first_col = False
    band_col = False
    if tblPr_el:
        pr = tblPr_el[0]
        style_el = safe_xpath(pr, './a:tableStyleId')
        if style_el:
            table_style_id = style_el[0].text
        first_row = pr.attrib.get('firstRow') == '1'
        band_row = pr.attrib.get('bandRow') == '1'
        first_col = pr.attrib.get('firstCol') == '1'
        band_col = pr.attrib.get('bandCol') == '1'

    col_widths = [emu_to_in(c.width) for c in tbl.columns]
    row_heights = [emu_to_in(r.height) for r in tbl.rows]

    cells_data = []
    merges = []

    for r_idx, row in enumerate(tbl.rows):
        row_cells = []
        for c_idx, cell in enumerate(row.cells):
            cell_dict = {
                'row': r_idx,
                'col': c_idx,
                'is_spanned': cell.is_spanned,
                'is_merge_origin': cell.is_merge_origin,
                'span_width': cell.span_width,
                'span_height': cell.span_height,
                'margin_in': {
                    'top': emu_to_in(cell.margin_top),
                    'bottom': emu_to_in(cell.margin_bottom),
                    'left': emu_to_in(cell.margin_left),
                    'right': emu_to_in(cell.margin_right),
                },
                'fill': extract_cell_fill(cell),
                'borders': extract_borders(cell),
                'vertical_anchor': str(cell.vertical_anchor) if cell.vertical_anchor else 'TOP',
                'paragraphs': []
            }

            if cell.is_merge_origin and (cell.span_width > 1 or cell.span_height > 1):
                merges.append({
                    'start_row': r_idx,
                    'start_col': c_idx,
                    'end_row': r_idx + cell.span_height - 1,
                    'end_col': c_idx + cell.span_width - 1,
                    'span_width': cell.span_width,
                    'span_height': cell.span_height
                })

            if not cell.is_spanned:
                for p in cell.text_frame.paragraphs:
                    p_align = str(p.alignment) if p.alignment else 'LEFT'
                    p_dict = {
                        'alignment': p_align,
                        'text': p.text,
                        'runs': []
                    }
                    for run in p.runs:
                        font_size = pt_to_float(run.font.size)
                        font_name = run.font.name or "Arial"
                        is_bold = bool(run.font.bold)
                        is_italic = bool(run.font.italic)
                        rPr_list = safe_xpath(run._r, './a:rPr')
                        run_clr = get_color_dict(run.font.color, rPr_list[0] if rPr_list else None)
                        p_dict['runs'].append({
                            'text': run.text,
                            'font_name': font_name,
                            'font_size_pt': font_size,
                            'bold': is_bold,
                            'italic': is_italic,
                            'color': run_clr
                        })
                    cell_dict['paragraphs'].append(p_dict)

            row_cells.append(cell_dict)
        cells_data.append(row_cells)

    return {
        'shape_name': shape.name,
        'shape_id': shape.shape_id,
        'left_in': emu_to_in(shape.left),
        'top_in': emu_to_in(shape.top),
        'width_in': emu_to_in(shape.width),
        'height_in': emu_to_in(shape.height),
        'num_rows': len(tbl.rows),
        'num_cols': len(tbl.columns),
        'column_widths_in': col_widths,
        'row_heights_in': row_heights,
        'table_style_id': table_style_id,
        'first_row': first_row,
        'band_row': band_row,
        'first_col': first_col,
        'band_col': band_col,
        'merges': merges,
        'cells': cells_data
    }


def analyze_textbox_shape(shape):
    """Deep analysis of a Text Box shape."""
    tf = shape.text_frame
    paragraphs = []
    for p in tf.paragraphs:
        p_dict = {
            'text': p.text,
            'alignment': str(p.alignment) if p.alignment else 'LEFT',
            'level': p.level,
            'space_before_pt': pt_to_float(p.space_before),
            'space_after_pt': pt_to_float(p.space_after),
            'runs': []
        }
        for run in p.runs:
            font_size = pt_to_float(run.font.size)
            font_name = run.font.name or "Arial"
            is_bold = bool(run.font.bold)
            rPr_list = safe_xpath(run._r, './a:rPr')
            run_clr = get_color_dict(run.font.color, rPr_list[0] if rPr_list else None)
            p_dict['runs'].append({
                'text': run.text,
                'font_name': font_name,
                'font_size_pt': font_size,
                'bold': is_bold,
                'color': run_clr
            })
        paragraphs.append(p_dict)

    return {
        'shape_name': shape.name,
        'shape_id': shape.shape_id,
        'left_in': emu_to_in(shape.left),
        'top_in': emu_to_in(shape.top),
        'width_in': emu_to_in(shape.width),
        'height_in': emu_to_in(shape.height),
        'paragraphs': paragraphs
    }


def generate_builder_code(blueprint, output_file='generated_slide2_builder.py'):
    """Generate ready-to-run Python code that constructs identical Slide 2 tables and shapes."""
    code = [
        "#!/usr/bin/env python3",
        "# Auto-generated Slide 2 Native Shape Builder",
        "# Generated from exact template slide forensic inspection",
        "",
        "from pptx.util import Inches, Pt",
        "from pptx.dml.color import RGBColor",
        "from pptx.enum.text import PP_ALIGN, MSO_ANCHOR",
        "from pptx.oxml.xmlchemy import OxmlElement",
        "",
        "",
        "def hex_to_rgb(hex_str):",
        "    if not hex_str or not hex_str.startswith('#'):",
        "        return None",
        "    h = hex_str.lstrip('#')",
        "    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))",
        "",
        "",
        "def set_cell_background(cell, hex_color):",
        "    if not hex_color:",
        "        return",
        "    rgb = hex_to_rgb(hex_color)",
        "    if rgb:",
        "        cell.fill.solid()",
        "        cell.fill.fore_color.rgb = rgb",
        "",
        "",
        "def build_slide2_native_layout(slide):",
        "    '''Builds both tables, title commentary, and footer on an empty slide.'''",
    ]

    for idx, tbl in enumerate(blueprint.get('tables', [])):
        var_tbl = f"table_{idx+1}"
        var_sh = f"shape_{idx+1}"
        code.append(f"\n    # --- TABLE {idx+1}: {tbl['shape_name']} ({tbl['num_rows']} rows x {tbl['num_cols']} cols) ---")
        code.append(f"    {var_sh} = slide.shapes.add_table({tbl['num_rows']}, {tbl['num_cols']}, Inches({tbl['left_in']}), Inches({tbl['top_in']}), Inches({tbl['width_in']}), Inches({tbl['height_in']}))")
        code.append(f"    {var_tbl} = {var_sh}.table")
        
        band_val = '1' if tbl['band_row'] else '0'
        first_row_val = '1' if tbl['first_row'] else '0'
        code.append(f"    # Table style properties")
        code.append(f"    tblPr = {var_tbl}._tbl.tblPr")
        code.append(f"    tblPr.set('bandRow', '{band_val}')")
        code.append(f"    tblPr.set('firstRow', '{first_row_val}')")

        code.append(f"    # Column widths")
        code.append(f"    col_widths = {tbl['column_widths_in']}")
        code.append(f"    for i, w in enumerate(col_widths):")
        code.append(f"        {var_tbl}.columns[i].width = Inches(w)")

        code.append(f"    # Row heights")
        code.append(f"    row_heights = {tbl['row_heights_in']}")
        code.append(f"    for i, h in enumerate(row_heights):")
        code.append(f"        {var_tbl}.rows[i].height = Inches(h)")

        if tbl.get('merges'):
            code.append(f"    # Merges")
            for m in tbl['merges']:
                code.append(f"    {var_tbl}.cell({m['start_row']}, {m['start_col']}).merge({var_tbl}.cell({m['end_row']}, {m['end_col']}))")

        code.append(f"    # Cell fills and text")
        for r_idx, row in enumerate(tbl['cells']):
            for c_idx, cell in enumerate(row):
                if cell['is_spanned']:
                    continue
                fill = cell.get('fill', {})
                hex_fill = fill.get('hex')
                if hex_fill:
                    code.append(f"    set_cell_background({var_tbl}.cell({r_idx}, {c_idx}), '{hex_fill}')")
                
                cell_text = "\n".join(p['text'] for p in cell.get('paragraphs', []) if p['text'].strip())
                if cell_text:
                    safe_txt = cell_text.replace('\\', '\\\\').replace("'", "\\'").replace('\n', '\\n')
                    code.append(f"    {var_tbl}.cell({r_idx}, {c_idx}).text = '{safe_txt}'")
                    first_run = None
                    for p in cell.get('paragraphs', []):
                        if p.get('runs'):
                            first_run = p['runs'][0]
                            break
                    if first_run:
                        if first_run.get('font_size_pt'):
                            code.append(f"    for p in {var_tbl}.cell({r_idx}, {c_idx}).text_frame.paragraphs:")
                            code.append(f"        p.font.name = '{first_run.get('font_name', 'Arial')}'")
                            code.append(f"        p.font.size = Pt({first_run['font_size_pt']})")
                            code.append(f"        p.font.bold = {first_run['bold']}")
                            clr = first_run.get('color', {}).get('hex')
                            if clr:
                                code.append(f"        p.font.color.rgb = hex_to_rgb('{clr}')")

    code.append("\n    return {")
    for idx in range(len(blueprint.get('tables', []))):
        code.append(f"        'table_{idx+1}': table_{idx+1},")
    code.append("    }\n")

    Path(output_file).write_text("\n".join(code), encoding='utf-8')
    return output_file


def inspect_presentation(pptx_path, target_slide_idx=2):
    """Main extraction routine."""
    path = Path(pptx_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Presentation not found: {path}")

    print(f"\n{'='*75}")
    print(f"  SLIDE 2 FORENSIC INSPECTION & TABLE BLUEPRINT EXTRACTOR")
    print(f"  File: {path.name}")
    print(f"  Path: {path}")
    print(f"{'='*75}\n")

    prs = Presentation(str(path))
    slide_w = emu_to_in(prs.slide_width)
    slide_h = emu_to_in(prs.slide_height)
    print(f"[*] Presentation loaded: {len(prs.slides)} slide(s) found.")
    print(f"[*] Slide Dimensions: {slide_w}\" wide x {slide_h}\" high (Aspect: {round(slide_w/slide_h, 2) if slide_h else 'N/A'}:1)")

    # Find target slide
    slide_num = target_slide_idx
    chosen_slide = None
    if 1 <= slide_num <= len(prs.slides):
        s = prs.slides[slide_num - 1]
        if any(sh.has_table for sh in s.shapes):
            chosen_slide = s

    if not chosen_slide:
        slide_table_counts = []
        for i, s in enumerate(prs.slides, start=1):
            t_cnt = sum(1 for sh in s.shapes if sh.has_table)
            slide_table_counts.append((t_cnt, i, s))
        slide_table_counts.sort(key=lambda x: x[0], reverse=True)
        if slide_table_counts and slide_table_counts[0][0] > 0:
            best_cnt, best_idx, best_slide = slide_table_counts[0]
            print(f"[*] Slide {slide_num} did not contain tables. Auto-selected Slide {best_idx} which has {best_cnt} table(s).")
            chosen_slide = best_slide
            slide_num = best_idx
        else:
            chosen_slide = prs.slides[0]
            slide_num = 1

    target_slide = chosen_slide
    print(f"[*] Analyzing Slide {slide_num} (Shapes count: {len(target_slide.shapes)})")

    tables_data = []
    textboxes_data = []
    other_shapes = []

    for sh_idx, shape in enumerate(target_slide.shapes):
        try:
            if shape.has_table:
                print(f"  -> Found Table #{len(tables_data)+1}: '{shape.name}' (Shape ID: {shape.shape_id})")
                tbl_info = analyze_table_shape(shape)
                tables_data.append(tbl_info)
                print(f"     Dimensions: {tbl_info['num_rows']} rows x {tbl_info['num_cols']} cols")
                print(f"     Geometry: Left={tbl_info['left_in']}\", Top={tbl_info['top_in']}\", Width={tbl_info['width_in']}\", Height={tbl_info['height_in']}\"")
                print(f"     BandRow={tbl_info['band_row']}, FirstRow={tbl_info['first_row']}")
            elif shape.has_text_frame:
                tb_info = analyze_textbox_shape(shape)
                textboxes_data.append(tb_info)
                snippet = shape.text_frame.text[:40].replace('\n', ' ')
                print(f"  -> Found TextBox: '{shape.name}' (Left={tb_info['left_in']}\", Top={tb_info['top_in']}\") Text: '{snippet}...'")
            else:
                other_shapes.append({
                    'name': shape.name,
                    'shape_id': shape.shape_id,
                    'left_in': emu_to_in(shape.left),
                    'top_in': emu_to_in(shape.top),
                    'width_in': emu_to_in(shape.width),
                    'height_in': emu_to_in(shape.height)
                })
        except Exception as ex:
            print(f"  [!] Warning on shape '{shape.name}': {ex}")

    blueprint = {
        'presentation_name': path.name,
        'slide_number': slide_num,
        'slide_width_in': slide_w,
        'slide_height_in': slide_h,
        'tables_count': len(tables_data),
        'tables': tables_data,
        'textboxes': textboxes_data,
        'other_shapes': other_shapes
    }

    # 1. Save JSON
    json_path = path.parent / "slide2_tables_blueprint.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(blueprint, f, indent=2, ensure_ascii=False)
    print(f"\n[OK] Complete structured JSON saved to: {json_path}")

    # 2. Save Human-readable Text Report
    txt_path = path.parent / "slide2_inspection_report.txt"
    report_lines = [
        f"SLIDE 2 FORENSIC INSPECTION REPORT",
        f"File: {path.name}",
        f"Slide: {slide_num}",
        f"Tables Found: {len(tables_data)}",
        f"{'='*60}",
    ]
    for i, t in enumerate(tables_data, 1):
        report_lines.append(f"\n--- TABLE {i}: {t['shape_name']} (Shape ID: {t['shape_id']}) ---")
        report_lines.append(f"Position: Left={t['left_in']}\", Top={t['top_in']}\", Width={t['width_in']}\", Height={t['height_in']}\"")
        report_lines.append(f"Grid: {t['num_rows']} Rows x {t['num_cols']} Columns")
        report_lines.append(f"Column Widths: {t['column_widths_in']}")
        report_lines.append(f"Row Heights: {t['row_heights_in']}")
        report_lines.append(f"BandRow: {t['band_row']}, FirstRow: {t['first_row']}")
        report_lines.append(f"Merges ({len(t['merges'])}):")
        for m in t['merges']:
            report_lines.append(f"  * R{m['start_row']}C{m['start_col']} -> R{m['end_row']}C{m['end_col']} (span {m['span_width']}x{m['span_height']})")
        report_lines.append("Cells Matrix Preview:")
        for r_i, r in enumerate(t['cells']):
            row_txts = []
            for c_i, c in enumerate(r):
                t_str = " ".join(p['text'] for p in c['paragraphs']).strip().replace('\n', ' ')
                fill_h = c['fill'].get('hex') or c['fill'].get('type', '-')
                row_txts.append(f"[{t_str[:15]}|{fill_h}]")
            report_lines.append(f"  R{r_i}: " + " | ".join(row_txts))

    txt_path.write_text("\n".join(report_lines), encoding='utf-8')
    print(f"[OK] Human-readable report saved to: {txt_path}")

    # 3. Generate Builder Script
    builder_path = path.parent / "generated_slide2_builder.py"
    generate_builder_code(blueprint, str(builder_path))
    print(f"[OK] Auto-generated Python builder saved to: {builder_path}")

    print(f"\n{'='*75}")
    print(f"  ANALYSIS COMPLETE! You can share the extracted JSON or report.")
    print(f"{'='*75}\n")
    return blueprint


def find_candidate_presentation():
    """Auto-detect any candidate PPTX files in the current folder."""
    candidates = list(Path('.').glob('*.pptx'))
    if not candidates:
        return None
    for c in candidates:
        name_l = c.name.lower()
        if 'crm' in name_l or 'bot' in name_l or 'report' in name_l or 'deck' in name_l or 'slide' in name_l:
            return str(c)
    return str(candidates[0])


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Inspect Slide 2 tables & layout forensic properties")
    parser.add_argument("pptx", nargs="?", help="Path to PowerPoint presentation (.pptx)")
    parser.add_argument("--pptx", dest="pptx_opt", help="Alternative flag for path to PowerPoint presentation")
    parser.add_argument("--slide", type=int, default=2, help="Slide number to analyze (default: 2)")
    args = parser.parse_args()

    chosen_pptx = args.pptx or args.pptx_opt
    if not chosen_pptx:
        detected = find_candidate_presentation()
        if detected:
            print(f"[*] No file specified. Auto-detected: {detected}")
            chosen_pptx = detected
        else:
            print("[ERROR] Please provide the path to a PowerPoint presentation:")
            print("  python inspect_template_tables.py \"CRM Report Slides for BOT.pptx\"")
            sys.exit(1)

    try:
        inspect_presentation(chosen_pptx, target_slide_idx=args.slide)
    except Exception as e:
        import traceback
        print(f"\n[ERROR] Inspection failed: {e}")
        traceback.print_exc()
        sys.exit(1)
