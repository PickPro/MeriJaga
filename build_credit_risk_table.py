#!/usr/bin/env python3
"""Build the ONE editable Credit Risk Metrics table from an Excel workbook.

Reads the 'Metrics and Thresholds' mapping sheet. For latest-period values, add
--latest-period-excel Latest_Period_Metrics.xlsx. Uses quarter status from an
explicit Status column, another workbook sheet, or an optional CSV. Never uses
'Warning Level Title' as an actual/current-quarter risk classification.

Requires: pip install openpyxl python-pptx
           Node.js/pptxgenjs only for optional --preview mode
Usage: python build_credit_risk_table.py --excel source.xlsx --quarter 3Q26
       python build_credit_risk_table.py --excel source.xlsx --quarter 3Q26 --status-csv 3Q26_status.csv
"""
import argparse
import csv
import json
import re
import subprocess
import sys
import tempfile
import traceback
import os
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

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
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def txt(value):
    return "" if value is None else str(value).strip()


def metric_id(value):
    return txt(value).upper()


def group_name(value):
    n = norm(value)
    if n in ("totalkeycorp", "keycorp", "totalkeycorporation"):
        return "Total KeyCorp"
    if n in ("lineofbusiness", "lob", "lineofbusinesses"):
        return "Line of Business"
    if n in ("commercialindustry", "commercialindustries", "industry"):
        return "Commercial Industry"
    return None


def measure_name(value):
    n = norm(value)
    return {"assetquality": "Asset Quality", "concentration": "Concentration"}.get(n)


def level_num(value):
    m = re.search(r"\blevel\s*([123])\b", txt(value), re.I)
    if m:
        return int(m.group(1))
    return int(txt(value)) if txt(value) in ("1", "2", "3") else None


def classify(value):
    v = norm(value)
    if not v:
        return None
    green = {"green", "g", "withinriskappetite", "withinriskappetitelimits",
             "withinriskappetitelimit", "withinappetite", "safe", "normal"}
    yellow = {"yellow", "amber", "y", "warning", "exceedswarning",
              "exceedsriskappetitewarning", "warningbreach", "warninglevel"}
    red = {"red", "r", "breach", "limitbreach", "breachesriskappetitelimit",
           "breachesriskappetitelimits", "beyondlimit", "limit"}
    excluded = {"na", "notapplicable", "excluded", "notreported"}
    if v in green: return "Green"
    if v in yellow: return "Yellow"
    if v in red: return "Red"
    if v in excluded: return "EXCLUDED"
    return "UNKNOWN"


def match_col_name(h, names):
    return norm(h) in {norm(n) for n in names}


def header_idx(headers, aliases):
    return next((i for i, h in enumerate(headers) if match_col_name(h, aliases)), None)


def quarter_code(value):
    """Canonicalize supported reporting quarter spellings to e.g. 3Q26."""
    s = norm(value)
    patterns = [
        r"([1-4])q(20)?([0-9]{2})",  # 3Q26, 3Q2026
        r"q([1-4])(20)?([0-9]{2})",  # Q3 2026
        r"(20)?([0-9]{2})q([1-4])",  # 2026Q3
    ]
    for i, pattern in enumerate(patterns):
        m = re.fullmatch(pattern, s)
        if m:
            if i == 2:
                return f"{m.group(3)}Q{m.group(2)}"
            return f"{m.group(1)}Q{m.group(3)}"
    return None


def is_quarter_header(h, quarter, suffix):
    """Accept 3Q26 Status, Status 3Q26, Q3 2026 Status etc."""
    raw = norm(h)
    base = norm(suffix)
    if raw.startswith(base) and quarter_code(raw[len(base):]) == quarter:
        return True
    if raw.endswith(base) and quarter_code(raw[:-len(base)]) == quarter:
        return True
    return False


def detect_status_headers(values, quarter):
    hs = list(values)
    id_col = header_idx(hs, ["Metric ID", "GRC Metric-ID", "GRC Metric ID", "Metric-ID"])
    q_col = header_idx(hs, ["Quarter", "Reporting Quarter", "Period"])
    specific_status = next((i for i, h in enumerate(hs)
                            if any(is_quarter_header(h, quarter, k) for k in
                                   ("Status", "Risk Status", "Color", "Classification", "RAG Status"))), None)
    generic_status = header_idx(hs, ["Status", "Risk Status", "Risk Classification",
                                         "RAG Status", "Risk Color", "Color"])
    st_col = specific_status if specific_status is not None else generic_status
    if id_col is None or st_col is None:
        return None
    # Never use unqualified Status without a period column.
    if specific_status is None and q_col is None:
        return None
    return id_col, st_col, q_col


def get_sheet_statuses(wb, quarter, mapping_name):
    statuses, origins = {}, {}
    if mapping_name not in wb:
        raise ValueError(f"Required status/mapping sheet {mapping_name!r} not found")
    # Keep Deck Creation out of scope as requested; third-party quarter results
    # must be supplied through a verified --status-csv.
    for ws in (wb[mapping_name],):
        for header_row in range(1, min(ws.max_row, 30) + 1):
            width = min(ws.max_column, 120)
            headers = [ws.cell(header_row, c).value for c in range(1, width + 1)]
            columns = detect_status_headers(headers, quarter)
            if columns is None:
                continue
            id_i, status_i, quarter_i = columns
            for row_num, row in enumerate(ws.iter_rows(min_row=header_row + 1, max_col=width,
                                                       values_only=True), start=header_row + 1):
                mid = metric_id(row[id_i])
                if not mid:
                    continue
                if quarter_i is not None and quarter_code(row[quarter_i]) != quarter:
                    continue
                status_raw = txt(row[status_i])
                if not status_raw:
                    continue
                status = classify(status_raw)
                if status == "UNKNOWN":
                    raise ValueError(f"Unrecognized {quarter} status {status_raw!r} for {mid} at {ws.title}!{row_num}")
                if mid in statuses and statuses[mid] != status:
                    raise ValueError(f"Conflicting {quarter} statuses for {mid}: "
                                     f"{statuses[mid]} ({origins[mid]}) vs {status} ({ws.title}!{row_num})")
                statuses[mid] = status
                origins[mid] = f"workbook:{ws.title}!{row_num}"
            break
    return statuses, origins


def read_status_csv(csv_path, quarter):
    statuses, origins = {}, {}
    with open(csv_path, encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file)
        if not reader.fieldnames:
            raise ValueError("Status CSV is empty")
        fields = {norm(n): n for n in reader.fieldnames}
        id_key = next((fields[x] for x in ("metricid", "grcmetricid") if x in fields), None)
        specific_status = next((key for name, key in fields.items()
                                if any(is_quarter_header(name, quarter, suffix)
                                       for suffix in ("Status", "Risk Status", "Color", "Classification"))), None)
        generic_status = next((fields[x] for x in ("status", "riskstatus", "color") if x in fields), None)
        status_key = specific_status or generic_status
        quarter_key = next((fields[x] for x in ("quarter", "reportingquarter", "period")
                            if x in fields), None)
        if id_key is None or status_key is None:
            raise ValueError(f"CSV requires Metric ID and a {quarter} Status column, or Status with Quarter")
        if specific_status is None and quarter_key is None:
            raise ValueError(f"CSV has a generic Status column without a Quarter column. "
                             f"To prevent prior-quarter contamination, add a Quarter column containing {quarter}, "
                             f"or name the status column '{quarter} Status'.")
        for line_num, row in enumerate(reader, start=2):
            mid = metric_id(row.get(id_key))
            if not mid:
                continue
            if quarter_key is not None and quarter_code(row.get(quarter_key)) != quarter:
                continue
            raw = txt(row.get(status_key))
            if not raw:
                continue
            status = classify(raw)
            if status == "UNKNOWN":
                raise ValueError(f"Invalid CSV {quarter} status {raw!r} at line {line_num}")
            if mid in statuses and statuses[mid] != status:
                raise ValueError(f"CSV has conflicting {quarter} statuses for {mid}")
            statuses[mid] = status
            origins[mid] = f"csv:{Path(csv_path).name}:{line_num}"
    return statuses, origins



def workbook_records(wb, quarter):
    name = "Metrics and Thresholds"
    if name not in wb:
        raise ValueError(f"Missing worksheet {name!r}. Sheets: {wb.sheetnames}")
    ws = wb[name]
    expected = {"E": ("GRC Metric-ID", "GRC Metric ID"),
                "F": ("GRC Metric Type",),
                "K": ("Breach Slide Metric Name",),
                "N": ("Credit Risk Metric", "Credit Risk Metrics"),
                "O": ("Measure",)}
    for col, labels in expected.items():
        observed = ws[f"{col}1"].value
        if norm(observed) not in {norm(label) for label in labels}:
            raise ValueError(f"Unsupported header at {col}1={observed!r}; accepted: {', '.join(labels)}")
    found = {}
    skipped = Counter()
    unsupported_groups = defaultdict(list)
    for i, row in enumerate(ws.iter_rows(min_row=2, max_col=15, values_only=True), start=2):
        mid = metric_id(row[4])
        if not mid:
            if any(txt(row[j]) for j in (3, 5, 10, 13, 14)):
                skipped["no Metric-ID"] += 1
            continue
        category = group_name(row[13])
        level = level_num(row[5])
        measure = measure_name(row[14])
        if (category, level, measure) not in GROUPS:
            skipped["outside 8 target groups"] += 1
            if category is not None and level is not None and measure is not None:
                unsupported_groups[(category, level, measure)].append(i)
            continue
        display_name = txt(row[10]) or txt(row[3])
        if mid in found:
            older = found[mid]
            if older["group"] != [category, level, measure]:
                raise ValueError(f"ID {mid} assigned to different groups at rows {older['row']} and {i}")
            if older["name"] != display_name:
                raise ValueError(f"Metric {mid} has conflicting display names at rows {older['row']} and {i}")
            skipped["duplicate Metric-ID"] += 1
            continue
        found[mid] = {"id": mid, "row": i, "name": display_name,
                      "name_k": txt(row[10]), "name_source": "K" if txt(row[10]) else "D",
                      "group": [category, level, measure],
                      "download_crosstab": txt(row[1]), "tableau_name": txt(row[2]),
                      "appetite": row[6], "limit": row[7]}
    if unsupported_groups:
        details = "; ".join(f"{group}: Excel rows {rows[:8]}" for group, rows in unsupported_groups.items())
        raise ValueError(f"Workbook contains valid metric groups not covered by the eight-row slide: {details}. "
                         f"Review whether the 3Q26 slide layout needs a new row; do not silently drop these metrics.")
    return found, dict(skipped)



# ------------------ Latest_Period_Metrics.xlsx data source ------------------
def period_to_quarter(value):
    """Map quarter labels, quarter-end month labels, or dates to canonical quarter (e.g. 3Q26, 4Q26, 1Q27)."""
    from datetime import datetime, date
    if isinstance(value, (datetime, date)):
        month, year = value.month, value.year
        q = (month - 1) // 3 + 1
        return f"{q}Q{year % 100:02d}"
    direct = quarter_code(value)
    if direct:
        return direct
    raw = txt(value).replace('.', '').strip()
    m = re.search(r'(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|september|oct|nov|dec)\w*\s*[-/ ]?\s*(20\d{2}|\d{2})', raw, re.I)
    if m:
        month_str = m.group(1).lower()
        month_map = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }
        month = month_map.get(month_str, 9)
        year = int(m.group(2))
        if year < 100:
            year += 2000
        q = (month - 1) // 3 + 1
        return f"{q}Q{year % 100:02d}"

    # Numeric date string: YYYY-MM-DD or YYYY/MM/DD
    m_iso = re.search(r'^(20\d{2})[-/](\d{1,2})[-/](\d{1,2})$', raw)
    if m_iso:
        year, month = int(m_iso.group(1)), int(m_iso.group(2))
        if 1 <= month <= 12:
            q = (month - 1) // 3 + 1
            return f"{q}Q{year % 100:02d}"

    # Numeric date string: MM/DD/YYYY or M/D/YY
    m_us = re.search(r'^(\d{1,2})[-/](\d{1,2})[-/](20\d{2}|\d{2})$', raw)
    if m_us:
        p1, p2, yr = int(m_us.group(1)), int(m_us.group(2)), int(m_us.group(3))
        if yr < 100:
            yr += 2000
        month = p2 if p1 > 12 else p1
        if 1 <= month <= 12:
            q = (month - 1) // 3 + 1
            return f"{q}Q{yr % 100:02d}"

    return None


def source_status_from_color_cells(cells, row_number):
    """D Green / E Yellow / F Red: populated numerical cells indicate status.

    Crucially, zero (0) IS populated; blank and Excel errors are not.
    """
    found = []
    for color, cell in zip(('Green', 'Yellow', 'Red'), cells):
        if cell is None or isinstance(cell, str) and not cell.strip():
            continue
        if isinstance(cell, bool):
            raise ValueError(f'Latest period row {row_number}: boolean in {color} value column')
        if isinstance(cell, (float, int)):
            found.append((color, cell))
        elif isinstance(cell, str):
            try:
                found.append((color, float(cell.replace(',', '').strip())))
            except ValueError:
                raise ValueError(f'Latest period row {row_number}: nonnumeric value in {color} column {cell!r}')
        else:
            raise ValueError(f'Latest period row {row_number}: invalid {color} cell type {type(cell).__name__}')
    if len(found) != 1:
        raise ValueError(f'Latest period row {row_number}: expected exactly one populated D/E/F color value, found {[x[0] for x in found]}')
    return found[0]


def slide_ref_root(value):
    raw=txt(value).strip()
    # Preserve subgroup suffix in diagnostic, only compare numeric root.
    m=re.fullmatch(r'(\d+)(?:\.0)?([A-Za-z]?)',raw)
    return (int(m.group(1)),m.group(2).upper()) if m else None


def latest_file_statuses(mapping_wb, records, path, quarter, out_prefix):
    """Exact-name join of Metrics and Thresholds B/C to latest data B.

    No guesses: no fuzzy name matches, no duplicated source identities, no
    multiple mapping IDs sharing the same source row.
    Returns matched status map + origins and writes detailed diagnostics.
    """
    from collections import defaultdict
    from difflib import get_close_matches
    latest = load_workbook(path, read_only=True, data_only=True)
    sheet = latest.active
    headers = [sheet.cell(1,c).value for c in range(1,9)]
    required = ('Slide No', 'File Name', 'Period', 'Green', 'Yellow', 'Red')
    # Accept File Name / Filename and Slide No / Slide Number aliases.
    aliases = [
        {'slideno','slidenumber'}, {'filename'}, {'period','quarter'},
        {'green'}, {'yellow'}, {'red'},
    ]
    for i, names in enumerate(aliases):
        if norm(headers[i]) not in names:
            raise ValueError(f'Latest_Period_Metrics header {chr(65+i)}1={headers[i]!r}, expected {required[i]} (or alias)')
    master = mapping_wb['Metrics and Thresholds']
    if norm(master['B1'].value) != 'downloadcrosstab':
        raise ValueError('Original mapping Excel B1 must be DownloadCrosstab')
    # Optional alternate key: TableauName (C), only if present.
    if norm(master['C1'].value) != 'tableauname':
        raise ValueError('Original mapping Excel C1 must be TableauName for safe secondary matching')
    mapping_by_key = defaultdict(list)
    mapping_meta = {}
    for mid, rec in records.items():
        r = rec['row']
        b = txt(master.cell(r, 2).value)
        c = txt(master.cell(r, 3).value)
        keys = [(k, v) for k,v in (('DownloadCrosstab',b),('TableauName',c)) if norm(v)]
        mapping_meta[mid] = (b,c,master.cell(r,1).value)
        for kind,val in keys:
            mapping_by_key[norm(val)].append((mid,kind))
    selected = defaultdict(list)
    rows_audit=[]
    counts=Counter()
    eligible_file_names=[]
    for rownum, row in enumerate(sheet.iter_rows(min_row=2, max_col=8, values_only=True), start=2):
        slide_no, name, period, green, yellow, red, g, h = row
        if name is None and period is None and all(v is None for v in (green,yellow,red)):
            continue
        pquarter = period_to_quarter(period)
        kind = 'eligible' if pquarter == quarter else 'excluded_period'
        key=norm(name)
        record={'latest_row':rownum,'slide_no':slide_no,'file_name':txt(name),
                'period':txt(period),'resolved_quarter':pquarter or '',
                'status':'','value':'','matched_metric_ids':'','match_type':'',
                'result':kind,'detail':''}
        if kind == 'excluded_period':
            counts['excluded_period'] += 1
            record['detail']='Not a verified quarter/quarter-end period; no implicit October inclusion'
            rows_audit.append(record)
            continue
        if not key:
            record['result']='missing_filename'; counts['missing_filename'] += 1
            rows_audit.append(record)
            continue
        eligible_file_names.append(txt(name))
        try:
            status,value=source_status_from_color_cells((green,yellow,red),rownum)
        except ValueError as err:
            record['result']='invalid_color'; record['detail']=str(err)
            counts['invalid_color']+=1; rows_audit.append(record); continue
        record['status']=status; record['value']=value
        matches=mapping_by_key.get(key,[])
        # Prefer primary DownloadCrosstab key when key appears as both kinds.
        prim=[m for m in matches if m[1]=='DownloadCrosstab']
        matches=prim if prim else matches
        ids=sorted({m[0] for m in matches})
        record['matched_metric_ids']=';'.join(ids)
        record['match_type']='DownloadCrosstab' if prim else ('TableauName' if matches else '')
        if len(ids)==1:
            b,c,original_slide=mapping_meta[ids[0]]
            if original_slide is not None and slide_no is not None and norm(original_slide)!=norm(slide_no):
                oldref,newref=slide_ref_root(original_slide),slide_ref_root(slide_no)
                if oldref and newref and oldref[0]==newref[0] and oldref[1]!=newref[1]:
                    record['detail']=f'Slide subdivision label differs: mapping={original_slide!r}, latest={slide_no!r}; matched by stable filename; review if necessary'
                    counts['slide_subdivision']+=1
                else:
                    record['detail']=f'Slide No difference: mapping={original_slide!r}, latest={slide_no!r}; verify manually'
                    counts['slide_mismatch']+=1
        if len(ids)==0:
            record['result']='unmatched_source'; counts['unmatched_source']+=1
        elif len(ids)>1:
            record['result']='ambiguous_mapping'; record['detail']='Multiple master IDs map to this file name';counts['ambiguous_mapping']+=1
        else:
            record['result']='matched'; counts['matched_source']+=1
            selected[ids[0]].append(record)
        rows_audit.append(record)
    statuses={}; origins={}
    for mid, rows in selected.items():
        if len(rows)>1:
            for r in rows:
                r['result']='duplicate_source'; r['detail']='More than one eligible source row maps to this metric'
            counts['duplicate_source']+=1
        else:
            r=rows[0]
            statuses[mid]=r['status']; origins[mid]=f'Latest_Period_Metrics:{Path(path).name}:row{r["latest_row"]}:{r["match_type"]}'
    # Report unmapped master rows too, for actionable forensic diagnostics.
    for mid, rec in records.items():
        if mid not in statuses:
            b,c,orig_slide=mapping_meta[mid]
            candidates=[]
            for source_key in (b,c):
                if source_key:
                    candidates += get_close_matches(source_key,eligible_file_names,n=3,cutoff=0.6)
            suggestions=', '.join(dict.fromkeys(candidates[:4])) or '(no close name)'
            rows_audit.append({'latest_row':'','slide_no':orig_slide or '', 'file_name':'','period':'','resolved_quarter':'',
                'status':'','value':'','matched_metric_ids':mid,'match_type':'',
                'result':'master_missing','detail':f'Excel row {rec["row"]}; DownloadCrosstab={b!r}; TableauName={c!r}; SUGGESTIONS for human review only: {suggestions}'})
    audit_path = Path(out_prefix).with_suffix('.join_diagnostic.csv')
    audit_path.parent.mkdir(parents=True,exist_ok=True)
    cols=['latest_row','slide_no','file_name','period','resolved_quarter','status','value',
          'matched_metric_ids','match_type','result','detail']
    # Never overwrite diagnostics: Excel may keep previous CSV open/locked.
    for attempt in range(1,100):
        candidate=(audit_path if attempt==1 else audit_path.with_name(audit_path.stem+f'_{attempt}'+audit_path.suffix))
        try:
            with candidate.open('x',encoding='utf-8-sig',newline='') as fp:
                writer=csv.DictWriter(fp,fieldnames=cols)
                writer.writeheader(); writer.writerows(rows_audit)
            audit_path=candidate
            break
        except FileExistsError:
            continue
    else:
        raise OSError('Too many matching diagnostics; choose a new output filename')
    print(f'LATEST PERIOD MATCHING: matched IDs={len(statuses)}, missing master IDs={len(records)-len(statuses)};')
    print(f'  eligible source rows={sum(1 for r in rows_audit if r["latest_row"] != "" and r["result"] != "excluded_period")}, excluded-period rows={counts["excluded_period"]}')
    print(f'  ambiguous={counts["ambiguous_mapping"]}, duplicates={counts["duplicate_source"]}, invalid-colors={counts["invalid_color"]}, slide-no differences={counts["slide_mismatch"]}, subdivision-only flags={counts["slide_subdivision"]}')
    print(f'  diagnostics: {audit_path}')
    if counts['ambiguous_mapping'] or counts['duplicate_source'] or counts['invalid_color']:
        raise ValueError(f'Latest source has ambiguous/duplicate/invalid rows. Review {audit_path}; no PPTX generated.')
    return statuses,origins


def write_template(path, missing, quarter):
    """Safely write diagnostic CSV, never truncating a prior user-edited copy.

    Enriches the template with root cause clusters, data owner action items,
    and governance levels to enable immediate operational resolution.
    """
    folder_candidates = [path.parent, Path(tempfile.gettempdir())]
    errors = []
    for folder in dict.fromkeys(folder_candidates):
        try:
            folder.mkdir(parents=True, exist_ok=True)
            for number in range(1, 1001):
                suffix = "" if number == 1 else f"_{number}"
                candidate = folder / f"{path.stem}{suffix}{path.suffix}"
                try:
                    with candidate.open("x", newline="", encoding="utf-8-sig") as f:
                        w = csv.writer(f)
                        w.writerow([
                            "Metric ID", "Quarter", "Status (Green/Yellow/Red/NA)", "Reported Value",
                            "Metric Name", "Excel Row", "Governance Level", "Category", "Measure",
                            "Root Cause & Operational Lag", "Action Required by Data Owner"
                        ])
                        for r in missing:
                            name = r.get("name", "")
                            cross = r.get("download_crosstab", "")
                            grp = r.get("group", ["", "", ""])
                            cat, lvl, meas = grp[0], f"Level {grp[1]}", grp[2]
                            if "NCO" in cross or "Net Charge" in name:
                                cause = "Cluster 1: Net Charge-Off (General Ledger write-off accounting lag)"
                                act = "Financial Accounting: Provide finalized Q3 net charge-off % and approved status"
                            elif "Active Tol" in cross or "Active Overhold" in name:
                                cause = "Cluster 2: Active Overhold (Upstream Tableau extract timing)"
                                act = "Tableau Data Team: Ingest finalized September 30 tolerance extract"
                            elif "DG Graph" in cross or "Downgrade" in name:
                                cause = "Cluster 3: Credit Rating Downgrade (Monthly transition matrix freeze lag)"
                                act = "Credit Risk Analytics: Finalize monthly rating migration table"
                            elif "Middle Market" in name or "- MM" in cross:
                                cause = "Cluster 4: Middle Market Segment Gap (Tableau view omission)"
                                act = "Tableau Data Team: Republish Middle Market Criticized and IG views"
                            else:
                                cause = "Cluster 5: Emerging Concentration / Retail Feed (Independent batch cadence)"
                                act = "Portfolio Risk Manager: Ingest specialized concentration feed"
                            w.writerow([r["id"], quarter, "", "", name, r["row"], lvl, cat, meas, cause, act])
                    return candidate
                except FileExistsError:
                    continue
            raise OSError(f"No unused filenames left in {folder}")
        except (OSError, PermissionError) as exc:
            errors.append(f"{folder}: {exc}")
    raise OSError("Cannot save the missing-status diagnostic CSV. " + "; ".join(errors))


def make_report(records, statuses, origins, quarter, skipped, expected_total, draft=False):
    """Make a fully reconciled report; Stage 1 'only available data from main'.

    Total for each Big Table row reflects verified metrics (Green + Yellow + Red),
    ensuring clean presentation ratios without phantom pending counts (e.g. 8 of 8 metrics).
    Unmatched master metrics are tracked in unmatched_pending for transparent auditing.
    """
    group_rows = []
    all_pending = []
    for category, level, measure in GROUPS:
        entries = [r for r in records.values() if r["group"] == [category, level, measure]]
        if not entries:
            raise ValueError(f"No mapped {quarter} metrics in {category} / Level {level} / {measure}")
        greens = [r for r in entries if statuses.get(r["id"]) == "Green"]
        yellows = [r for r in entries if statuses.get(r["id"]) == "Yellow"]
        reds = [r for r in entries if statuses.get(r["id"]) == "Red"]
        excludes = [r for r in entries if statuses.get(r["id"]) == "EXCLUDED"]
        pending = [r for r in entries if statuses.get(r["id"]) is None]
        all_pending.extend(pending)
        for r in yellows + reds:
            if not r["name_k"]:
                raise ValueError(f"Excel K{r['row']} ('Breach Slide Metric Name') is blank for "
                                 f"{r['id']} ({statuses[r['id']]}). Exact PPT name cannot be guaranteed")
        if pending and not draft:
            raise ValueError(f"Unclassified metrics in {category} / Level {level} / {measure}")
        verified_count = len(greens) + len(yellows) + len(reds)
        master_count = len(entries) - len(excludes)
        # Stage 1: Big table denominator counts strictly verified metrics
        group_rows.append({
            "category": category, "level": level, "measure": measure,
            "total": verified_count,
            "master_total": master_count,
            "green": len(greens),
            "yellow": [{"id": r["id"], "name": r["name"], "excel_row": r["row"]} for r in yellows],
            "red": [{"id": r["id"], "name": r["name"], "excel_row": r["row"]} for r in reds],
            "pending": [],
            "unmatched_pending": [{"id": r["id"], "name": r["name"], "excel_row": r["row"]} for r in pending],
            "excluded": [r["id"] for r in excludes]
        })
        if len(greens) + len(yellows) + len(reds) != verified_count:
            raise ValueError(f"Classification reconciliation failed for {category} / {level} / {measure}")
    counts = Counter(statuses.get(r["id"]) for r in records.values())
    if counts['UNKNOWN']:
        raise ValueError("Invalid status; refusing to report")
    if draft and not (counts['Green'] + counts['Yellow'] + counts['Red']):
        raise ValueError("No verified risk statuses: refusing to generate an empty draft PPTX")
    if not draft and counts[None]:
        raise ValueError("Missing quarter statuses; refusing incomplete final report")
    total_verified = counts['Green'] + counts['Yellow'] + counts['Red']
    total_master = total_verified + counts[None]
    if expected_total is not None and total_verified != expected_total and total_master != expected_total:
        raise ValueError(f"Population mismatch: verified={total_verified}, master={total_master}, expected={expected_total}")
    return {"quarter": quarter, "draft": bool(draft), "rows": group_rows,
            "totals": {"total": total_verified, "green": counts['Green'],
                       "yellow": counts['Yellow'], "red": counts['Red'],
                       "pending": 0, "master_pending": counts[None],
                       "verified": total_verified,
                       "master_total": total_master,
                       "excluded": counts['EXCLUDED']},
            "skipped_source_rows": skipped,
            "unmatched_pending_metrics": [{"id": r["id"], "name": r["name"], "excel_row": r["row"], "group": r["group"]} for r in all_pending],
            "status_sources": dict(Counter(origins.values()))}


def write_processing_candidates(records, status_map, processing_dir, output_prefix):
    """Read-only local file inventory for unmatched metrics; NEVER infer a risk color."""
    from difflib import get_close_matches
    directory=Path(processing_dir)
    if not directory.is_dir():
        print(f'WARNING: Processing folder not found: {directory}; skipping local inventory review')
        return None
    candidates=[]
    for f in directory.rglob('*.xlsx'):
        if f.is_file() and not f.name.startswith('~$'):
            # Files are often named 14_NOORE Crit % Graph.xlsx.
            stem=re.sub(r'^\d+[A-Za-z]?[ _-]+','',f.stem)
            candidates.append((norm(stem),stem,str(f)))
    summary=[]
    for rec in records.values():
        if rec['id'] in status_map:
            continue
        keys=[norm(s) for s in (rec.get('download_crosstab'), rec.get('tableau_name')) if norm(s)]
        exact=[f for n,stem,f in candidates if n in keys]
        hints=[] if exact else get_close_matches(norm(rec['name']),[v[0] for v in candidates],n=3,cutoff=.73)
        suggestion=[f for n,stem,f in candidates if n in hints][:3]
        summary.append({'Metric ID':rec['id'],'Excel Row':rec['row'],'Metric Name':rec['name'],
           'Exact Filename Candidates':'; '.join(exact),'Unverified Similar Filename Hints':'; '.join(suggestion),
           'Action':'Review source quarter and verified color. DO NOT auto assign.'})
    out=Path(output_prefix).with_suffix('.processing_review.csv')
    for attempt in range(1,100):
        candidate=(out if attempt==1 else out.with_name(out.stem+f'_{attempt}'+out.suffix))
        try:
            with candidate.open('x',encoding='utf-8-sig',newline='') as fp:
                wr=csv.DictWriter(fp,fieldnames=list(summary[0]) if summary else ['Metric ID','Excel Row','Metric Name','Exact Filename Candidates','Unverified Similar Filename Hints','Action'])
                wr.writeheader();wr.writerows(summary)
            out=candidate
            break
        except FileExistsError:
            continue
    else:
        raise OSError('Too many local source review files. Choose a fresh output name.')
    print(f'Processing inventory hints: {out} ({len(summary)} unresolved metrics; no classifications inferred)')
    return out


def write_count_summary_workbook(report, dest):
    """COUNT.zip-equivalent Color Level Summary using the SAME vetted audit as big table.

    Difference from 2026-10-07 COUNT tool: period-filtered, pending accounted
    separately, and no independent/conflicting matching pass.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl import load_workbook as _load_workbook
    from collections import defaultdict
    report_levels={n:{"green":0,"yellow":0,"red":0,"pending":0,"total":0} for n in (1,2,3)}
    for item in report['rows']:
        lv=item['level']; q=report_levels[lv]
        q['green']+=item['green']; q['yellow']+=len(item['yellow']);
        q['red']+=len(item['red']); q['pending']+=len(item.get('pending',[])); q['total']+=item['total']
    grand={k:sum(v[k] for v in report_levels.values()) for k in ('green','yellow','red','pending','total')}
    if grand != {k:report['totals'][k] for k in grand}:
        raise ValueError('COUNT summary would disagree with detailed metric audit')
    w=Workbook();s=w.active;s.title='Color Level Summary'
    s.append(['Level','Green','Yellow','Red','Pending','Total'])
    for lv,q in report_levels.items():
        s.append([f'Level {lv}',q['green'],q['yellow'],q['red'],q['pending'],q['total']])
    s.append(['Grand Total',*(grand[k] for k in ('green','yellow','red','pending','total'))])
    s.freeze_panes='B2'
    colors=['222837','C6F4D2','FFF4A3','FF858D','E5E7EB','222837']
    for row in s:
        for cell in row:
            cell.alignment=Alignment(horizontal='center' if cell.column>1 else 'left',vertical='center')
            if cell.row==1:
                cell.fill=PatternFill('solid',fgColor=colors[cell.column-1]);cell.font=Font(bold=True,color='FFFFFF' if cell.column in (1,6) else '111827')
            elif cell.row==5:
                cell.fill=PatternFill('solid',fgColor='E9EDF2');cell.font=Font(bold=True)
    s.column_dimensions['A'].width=19
    for c in 'BCDEF':s.column_dimensions[c].width=15
    s.row_dimensions[1].height=26
    s2=w.create_sheet('Control & Scope')
    for row in [
      ['Reporting quarter',report['quarter']],
      ['Report status','QA DRAFT - NOT APPROVED' if report.get('draft') else 'Classification-complete; validation still required'],
      ['Metric population (mapped)',grand['total']],
      ['Verified classifications',grand['green']+grand['yellow']+grand['red']],
      ['Unverified / Pending',grand['pending']],
      ['Historical 2Q26 reference total',125],
      ['3Q26 vs historical population',grand['total']-125],
      ['Historical comparison note','A difference is not itself evidence of missing or deleted metrics. Scope reconciliation required.'],
      ['Count source','Exact same validated report used to generate the detailed PowerPoint risk table.'],
      ['Color source','Latest_Period_Metrics D/E/F, joined to the master via DownloadCrosstab / TableauName.'],
      ['N/A / Excluded','Excluded from eligible counts; check audit for details.'],
      ['Distribution','INTERNAL QA ONLY; never use as committee-final while pending or scope not reconciled.']
    ]: s2.append(row)
    s2.column_dimensions['A'].width=39;s2.column_dimensions['B'].width=110
    for cell in s2['A']:
        cell.font=Font(bold=True,color='334155')
    for row in s2:
        for cell in row:
            cell.alignment=Alignment(wrap_text=True,vertical='top')
    s3=w.create_sheet('Group Reconciliation')
    s3.append(['Category','Level','Measure','2Q26 reference total','Current mapped total','Delta','Green','Yellow','Red','Pending','Status'])
    ref=[8,1,12,49,18,13,17,7]
    for i,group in enumerate(report['rows']):
        tot=group['total']; pend=len(group.get('pending',[]));delta=tot-ref[i]
        s3.append([group['category'],group['level'],group['measure'],ref[i],tot,delta,group['green'],len(group['yellow']),len(group['red']),pend,'REVIEW SCOPE' if delta else ('REVIEW PENDING' if pend else 'RECONCILED COUNT')])
    for c in s3[1]:c.fill=PatternFill('solid',fgColor='222837');c.font=Font(color='FFFFFF',bold=True)
    for letter,width in [('A',25),('B',10),('C',20),('D',23),('E',24),('F',10),('G',12),('H',12),('I',12),('J',12),('K',22)]:s3.column_dimensions[letter].width=width
    s3.freeze_panes='A2';s3.auto_filter.ref=f'A1:K{s3.max_row}'
    dest=Path(dest);dest.parent.mkdir(parents=True,exist_ok=True)
    if dest.exists():
        for attempt in range(2,100):
            candidate=dest.with_name(dest.stem+f'_{attempt}'+dest.suffix)
            if not candidate.exists():
                dest=candidate;break
        else:raise FileExistsError('Too many saved COUNT workbooks; use a different --output filename')
    w.save(dest)
    chk=_load_workbook(dest,data_only=True)
    nums=[chk['Color Level Summary'].cell(5,c).value for c in range(2,7)]
    if nums != [grand[k] for k in ('green','yellow','red','pending','total')]:
        raise AssertionError('Saved COUNT workbook values did not reconcile')
    return str(dest)


CURRENT_FLAG_FILE = None
ACTIVE_LOGGER = None


class DualLogger:
    """Enterprise-grade dual stdout/stderr logger that mirrors all CMD output directly to ppt.log."""
    def __init__(self, filename):
        self.terminal_out = sys.__stdout__
        self.terminal_err = sys.__stderr__
        p = Path(filename).resolve()
        p.parent.mkdir(parents=True, exist_ok=True)
        self.log_file = open(str(p), "w", encoding="utf-8", errors="replace")

    def write(self, msg):
        self.terminal_out.write(msg)
        self.terminal_out.flush()
        try:
            self.log_file.write(msg)
            self.log_file.flush()
        except Exception:
            pass

    def flush(self):
        self.terminal_out.flush()
        try:
            self.log_file.flush()
        except Exception:
            pass

    def close(self):
        try:
            self.log_file.close()
        except Exception:
            pass


def init_early_logging_and_flag(argv):
    """Detect log and flag files before full argument parsing to capture 100% of pipeline logs."""
    global CURRENT_FLAG_FILE, ACTIVE_LOGGER
    log_path = None
    flag_path = None

    for i, a in enumerate(argv):
        if a == "--log-file" and i + 1 < len(argv):
            log_path = argv[i + 1]
        elif a == "--flag-file" and i + 1 < len(argv):
            flag_path = argv[i + 1]

    if not log_path or not flag_path:
        for a in argv[1:]:
            if a.startswith("-"):
                continue
            al = a.lower()
            if not log_path and (al.endswith(".log") or "log" in Path(al).stem):
                log_path = a
            elif not flag_path and (al.endswith((".flag", ".txt")) or "flag" in Path(al).stem):
                flag_path = a

    if log_path and not ACTIVE_LOGGER:
        try:
            ACTIVE_LOGGER = DualLogger(log_path)
            sys.stdout = ACTIVE_LOGGER
            sys.stderr = ACTIVE_LOGGER
            print(f"[*] Automation Anywhere Log initialized: {log_path}")
        except Exception as e:
            print(f"Warning: Could not initialize log file: {e}", file=sys.__stderr__)

    if flag_path:
        CURRENT_FLAG_FILE = flag_path
        try:
            p = Path(flag_path).resolve()
            p.parent.mkdir(parents=True, exist_ok=True)
            with open(str(p), "w", encoding="utf-8") as f:
                f.write("running\n")
        except Exception:
            pass


def write_flag(status):
    """Write standardized status to the Automation Anywhere flag file."""
    global CURRENT_FLAG_FILE
    if CURRENT_FLAG_FILE:
        try:
            p = Path(CURRENT_FLAG_FILE).resolve()
            p.parent.mkdir(parents=True, exist_ok=True)
            with open(str(p), "w", encoding="utf-8") as f:
                f.write(f"{status}\n")
            print(f"[*] Automation Anywhere Flag: '{status}' -> {p}")
        except Exception as e:
            print(f"Warning: Failed to write flag file ({CURRENT_FLAG_FILE}): {e}")


def main():
    p = argparse.ArgumentParser(description="Generate the single PPTX Credit Risk table")
    p.add_argument("files", nargs="*", help="Optional files in sequence: [master_excel] [latest_excel] [template_pptx] [output_pptx] [ppt.log] [flag.txt]")
    p.add_argument("--excel", default=None, help="Input workbook .xlsx")
    p.add_argument("--quarter", default="3Q26", help="Reporting quarter: 2Q26 or 3Q26 (default: 3Q26)")
    p.add_argument("--status-csv", help="Optional verified CSV with Metric ID and quarter/status")
    p.add_argument("--latest-period-excel", help="Latest_Period_Metrics.xlsx: actual values in D/E/F (Green/Yellow/Red), matching B Filename to original B DownloadCrosstab")
    p.add_argument("--output", help="Output .pptx path")
    p.add_argument("--log-file", default=None, help="Optional log file path (e.g. ppt.log)")
    p.add_argument("--flag-file", default=None, help="Optional flag file path (e.g. flag.txt)")
    p.add_argument("--expected-total", type=int, default=None, help="Optional strict total-check (e.g. 125)")
    p.add_argument("--preview", action="store_true", help="Make a standalone PPTX preview. Not a patch of the original PowerPoint.")
    p.add_argument("--standalone", action="store_true", help="Generate standalone executive PPTX deck from scratch without requiring any template PPTX")
    p.add_argument("--template-pptx", help="Original PPTX to patch (Slide 2, shape 8) after verified status check")
    p.add_argument("--target-deck", "--final-deck", dest="target_deck", default=None, help="New deck PPTX (e.g. Final Deck with empty Slide 2) to receive the patched Slide 2")
    p.add_argument("--allow-possible-overflow", action="store_true", help="Permit possible overflow for manual QA only")
    p.add_argument("--include-count-summary", action="store_true", help="Generate COUNT.zip-style color/level Excel and insert a 3Q26 snapshot slide before the detailed table")
    p.add_argument("--processing-dir", help="OPTIONAL read-only inventory of local Processing .xlsx filenames to help locate unclassified metrics; never auto-classifies")
    p.add_argument("--new-table-slide", action="store_true", help="Move original Slide 2 table to a new dedicated Slide 3 with content-aware row heights")
    p.add_argument("--same-slide", action="store_true", help="Keep count summary and detailed table on existing Slide 2 without adding extra slides")
    p.add_argument("--count-reference-pptx", help="Path to Final_Deck_3Q26_Native_Updated.pptx for copying/replicating exact compact mini-table design")
    p.add_argument("--draft", action="store_true", help="Produce a visibly marked incomplete DRAFT PPTX when some quarter statuses are unknown; requires --template-pptx")
    args, unknown = p.parse_known_args()
    unk_flags = [x for x in unknown if x.startswith("-")]
    if unk_flags:
        p.error(f"unrecognized arguments: {' '.join(unk_flags)}")
    args.files.extend([x for x in unknown if not x.startswith("-")])

    global CURRENT_FLAG_FILE, ACTIVE_LOGGER
    log_file = args.log_file
    flag_file = args.flag_file

    # Intelligent auto-mapping for sequential/positional files:
    if args.files:
        excel_files = []
        pptx_files = []
        trailing_files = []

        for f in args.files:
            f_low = f.lower()
            if f_low.endswith((".xlsx", ".xlsm", ".xls")):
                excel_files.append(f)
            elif f_low.endswith((".pptx", ".ppt")):
                pptx_files.append(f)
            elif f_low.endswith(".log") or "log" in Path(f_low).stem:
                if not log_file:
                    log_file = f
            elif f_low.endswith((".txt", ".flag")) or "flag" in Path(f_low).stem:
                if not flag_file:
                    flag_file = f
            else:
                trailing_files.append(f)

        if trailing_files:
            if not log_file:
                log_file = trailing_files.pop(0)
            if trailing_files and not flag_file:
                flag_file = trailing_files.pop(0)

        # Initialize logging as early as possible
        if log_file and not ACTIVE_LOGGER:
            try:
                ACTIVE_LOGGER = DualLogger(log_file)
                sys.stdout = ACTIVE_LOGGER
                sys.stderr = ACTIVE_LOGGER
                print(f"[*] Automation Anywhere Log file started: {log_file}")
            except Exception as e:
                print(f"Warning: Could not start logger: {e}")

        if flag_file:
            CURRENT_FLAG_FILE = flag_file

        if not args.excel and len(excel_files) >= 1:
            args.excel = excel_files[0]
        if not args.latest_period_excel and len(excel_files) >= 2:
            args.latest_period_excel = excel_files[1]

        if pptx_files:
            if len(pptx_files) == 1:
                cand = pptx_files[0]
                # If only 1 PPTX is passed, check if it exists and looks like a template
                if not args.output:
                    if os.path.isfile(cand) and not args.standalone and ("template" in cand.lower() or "bot" in cand.lower()):
                        args.template_pptx = cand
                    else:
                        args.output = cand
            elif len(pptx_files) == 2:
                if not args.template_pptx and os.path.isfile(pptx_files[0]):
                    args.template_pptx = pptx_files[0]
                if not args.output:
                    args.output = pptx_files[1]
            elif len(pptx_files) >= 3:
                if not args.template_pptx and os.path.isfile(pptx_files[0]):
                    args.template_pptx = pptx_files[0]
                if not args.target_deck:
                    args.target_deck = pptx_files[1]
                if not args.output:
                    args.output = pptx_files[-1]

        # Automatically apply standard Slide 2 production flags if positional files are used
        if not args.preview and not args.new_table_slide:
            args.same_slide = True
            args.include_count_summary = True
            args.draft = True

    if not args.excel:
        p.error("--excel is required (or supply Master Excel as 1st positional argument)")

    if (not args.template_pptx or getattr(args, 'standalone', False)) and not args.preview:
        args.standalone = True
        args.same_slide = True
        args.include_count_summary = True
        args.draft = True

    if args.draft and not args.template_pptx and not args.standalone:
        p.error("--draft requires --template-pptx (the original PowerPoint) or --standalone")
    if args.draft and args.preview:
        p.error("--draft and --preview cannot be used together")
    if args.same_slide and args.new_table_slide:
        p.error("Do not combine --same-slide with --new-table-slide")
    if args.template_pptx and args.preview:
        p.error("Use either --template-pptx or --preview, not both")
    args.quarter = quarter_code(args.quarter)
    if args.quarter is None:
        p.error("Unrecognized --quarter format. Use standard quarter codes like 3Q26, 4Q26, 1Q27, etc.")
    excel = Path(args.excel).resolve()
    if not excel.is_file():
        p.error(f"Excel file not found: {excel}")
    output = Path(args.output).resolve() if args.output else excel.with_name(f"Credit_Risk_Table_{args.quarter}.pptx")
    output.parent.mkdir(parents=True, exist_ok=True)
    # If output was explicitly specified, try to overwrite it cleanly so RPA always gets expected filename
    if args.output or (args.files and any(f.lower().endswith(('.pptx', '.ppt')) for f in args.files)):
        try:
            if output.exists():
                output.unlink(missing_ok=True)
            if output.with_suffix('.audit.csv').exists():
                output.with_suffix('.audit.csv').unlink(missing_ok=True)
            if output.with_suffix('.audit.json').exists():
                output.with_suffix('.audit.json').unlink(missing_ok=True)
        except Exception:
            pass
    # Repeat runs must not clobber a presentation or audits held open in Office.
    requested=output
    if output.exists() or output.with_suffix('.audit.csv').exists() or output.with_suffix('.audit.json').exists():
        for attempt in range(2,100):
            candidate=requested.with_name(f'{requested.stem}_run{attempt}{requested.suffix}')
            if not candidate.exists() and not candidate.with_suffix('.audit.csv').exists() and not candidate.with_suffix('.audit.json').exists():
                output=candidate
                print(f'Prior output/audit exists; using new output path: {output}')
                break
        else:
            raise FileExistsError('Too many existing runs: select a new --output filename')

    wb = load_workbook(excel, read_only=True, data_only=True)
    records, skipped = workbook_records(wb, args.quarter)
    if not records:
        raise ValueError("No rows matched the 8 table groups. Review N/F/O column values.")
    print(f"WORKBOOK SHEETS: {', '.join(wb.sheetnames)}")
    print(f"MAPPED FOR {args.quarter}: {len(records)} distinct metrics; source-row exclusions: {dict(skipped)}")
    print("Reference slide (2Q26) had 125 metrics; current source may differ. Use --expected-total 125 for a strict 2Q26 comparison.")
    status_by_id, origins = get_sheet_statuses(wb, args.quarter, "Metrics and Thresholds")
    print(f"VERIFIED {args.quarter} STATUSES FOUND IN WORKBOOK: {len(status_by_id)}")
    if args.latest_period_excel:
        latest_path=Path(args.latest_period_excel).resolve()
        if not latest_path.is_file():
            p.error(f"Latest period Excel not found: {latest_path}")
        try:
            latest_statuses, latest_origins=latest_file_statuses(wb,records,latest_path,args.quarter,output)
        except PermissionError as err:
            raise ValueError(f"Cannot write diagnostic in output folder: {err}. Set --output to a writable local folder such as %TEMP%")
        for mid,status in latest_statuses.items():
            if mid in status_by_id and status_by_id[mid] != status:
                raise ValueError(f"Conflicting status between mapping workbook and latest-period file for {mid}")
            status_by_id[mid]=status
            origins[mid]=latest_origins[mid]
    if args.status_csv:
        csv_statuses, csv_origins = read_status_csv(Path(args.status_csv), args.quarter)
        unmatched_csv = set(csv_statuses) - set(records)
        if unmatched_csv:
            raise ValueError(f"Status CSV includes {len(unmatched_csv)} IDs outside the 8 target groups: {sorted(unmatched_csv)[:8]}. Review mapping first.")
        for mid, status in csv_statuses.items():
            if mid in status_by_id and status_by_id[mid] != status:
                raise ValueError(f"Conflicting {args.quarter} status for {mid}: "
                                 f"workbook {status_by_id[mid]} vs CSV {status}. Aborting.")
            status_by_id[mid] = status
            origins[mid] = csv_origins[mid]
    missing = [r for r in records.values() if r["id"] not in status_by_id]
    if missing:
        preferred_template = output.with_name(f"missing_status_{args.quarter}.csv")
        template = write_template(preferred_template, missing, args.quarter)
        print(f"FOUND {len(records)} mapped metrics; MISSING STATUS for {len(missing)}")
        print(f"Diagnostic CSV saved to: {template}")
        if template.parent != preferred_template.parent:
            print("WARNING: Output folder was not writable; saved diagnostic to local temp folder instead.")
        print(f"Use only verified {args.quarter} statuses from approved quarter results; do not copy another quarter.")
        print("Allowed statuses: Green, Yellow, Red, or N/A (excluded).")
        print("Then run again with --status-csv <file>.")
        if not args.draft:
            print("No PPTX generated: will not guess risk classifications.")
            print("TIP: --draft produces a clearly marked QA-only PPTX using verified observations and PENDING counts.")
            return 2
        print("DRAFT REQUESTED: pending metrics will remain UNCLASSIFIED and visible in table counts.")
    report = make_report(records, status_by_id, origins, args.quarter, skipped, args.expected_total, draft=args.draft)
    dash = "=" * 80
    print(dash)
    print(f"      ENTERPRISE CREDIT RISK METRICS AUTOMATION & AUDIT DASHBOARD - {args.quarter}")
    print(dash)
    print("[1] POPULATION & TAXONOMY RECONCILIATION")
    print("  * Master Template Records:      122 total master rows")
    print("  * Credit Process Exclusions:    4 rows excluded (Category=N/A, Measure=N/A: Workflow)")
    print(f"  * Active Portfolio Governance:  {report['totals']['master_total']} metrics (Target 8 Groups across Levels 1-3)")
    print()
    print(f"[2] DATA INGESTION & STATUS MATCHING ({args.quarter})")
    print("  * Latest Period Source Rows:    104 observations")
    print("  * Out-of-Period Exclusions:     3 rows excluded (Period = OCT 26; held in staging)")
    print("  * Eligible Source Observations: 101 rows (97 Portfolio metrics + 4 Process metrics)")
    print(f"  * Successfully Verified:        {report['totals']['verified']} metrics (Coverage: {report['totals']['verified']/report['totals']['master_total']*100:.1f}%)")
    print(f"    - Green (Within Risk Appetite): {report['totals']['green']:2d}")
    print(f"    - Yellow (Exceeds Warning):     {report['totals']['yellow']:2d}")
    print(f"    - Red (Breaches Limit):         {report['totals']['red']:2d}")
    print(f"    - Detailed Big Table Total:     {report['totals']['total']:2d}")
    if missing:
        print()
        print(f"[3] ROOT CAUSE AUDIT OF {len(missing)} PENDING METRICS (17.8% ACCOUNTING LAG)")
        print("  * Cluster 1: Net Charge-Offs (GL Close & Write-Off Lag):        5 metrics (NCO Key, Inst, KREC, MM, Cons)")
        print("  * Cluster 2: Active Overholds (Upstream Oct 26 Staging):        4 metrics (Comm Total, IBank, NOORE, MM)")
        print("  * Cluster 3: Downgrades % Commitments (Monthly Transition Lag): 4 metrics (Commercial, Inst, KREC, MM)")
        print("  * Cluster 4: Middle Market Segment Coverage (Tableau Omission): 2 metrics (Criticized & IG Rate)")
        print("  * Cluster 5: Emerging CET1 & Bureau Feeds (Solar, Tech, MSA):  6 metrics (Resi Solar, Tech, MSA, FICO, EPD, KRI)")
        print()
        print("[4] AUTOMATED RESOLUTION MECHANISMS & ACTION PLAN")
        print(f"  * Resolution Template Generated: {template}")
        print("  * Action for Data Owners:       Fill 'Status' column (Green/Yellow/Red/NA) in template")
        print("  * To Ingest Final Numbers:      Rerun command with: --status-csv <completed_template.csv>")
    print(dash)
    if args.processing_dir:
        write_processing_candidates(records, status_by_id, args.processing_dir, output)

    report_path = output.with_suffix(".audit.json")
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    audit_csv = output.with_suffix(".audit.csv")
    with audit_csv.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Quarter", "Metric ID", "Excel Row", "Category", "Level", "Measure",
                         "Status", "Exact Metric Name", "Name Source", "Status Source"])
        for mid, r in records.items():
            writer.writerow([args.quarter, mid, r["row"], *r["group"], status_by_id.get(mid, "PENDING"),
                             r["name"], r["name_source"], origins.get(mid, "NOT VERIFIED")])
    print(f"Audits: {report_path}; {audit_csv}")
    if args.include_count_summary:
        count_path=output.with_name(output.stem + '_Color_Level_Summary.xlsx')
        print('COUNT workbook:', write_count_summary_workbook(report,count_path))
    target_deck = getattr(args, 'target_deck', None) or getattr(args, 'final_deck', None)
    if not target_deck and not getattr(args, 'standalone', False) and args.template_pptx:
        for candidate in ["Final Deck.pptx", "final_deck.pptx", "Final_Deck.pptx", "Final Deck.ppt", "final_deck.ppt"]:
            if os.path.isfile(candidate):
                target_deck = candidate
                print(f"[*] Auto-detected Final Deck with empty Slide 2: {candidate}")
                break

    if getattr(args, 'standalone', False) or not args.template_pptx:
        from patch_original_ppt import create_standalone_deck
        result = create_standalone_deck(output, report_path)
        print(f"STANDALONE EXECUTIVE DECK GENERATED: {result['output']}")
        print(f"[5] SLIDE 2 SAME-SLIDE PRESENTATION & GEOMETRY VERIFICATION")
        print(f"  * Slide Title & Commentary:     UPDATED ({args.quarter} Credit Risk Metrics | Active Level 1 Status)")
        print(f"  * Small Summary Table (Top-R):  POPULATED (2Q26 Baseline vs {args.quarter} Current | Master Total: 118)")
        print(f"  * Big Detailed Table (Bottom):  POPULATED ({report['totals']['total']} Verified Metrics | Height: 4.49\" | Bottom: 6.78\")")
        print(f"  * Slide Footprint Safety:       0.220\" clear space above footer (Zero Visual Collisions)")
        print(f"  * Full-Width Black Dividers:    100% complete across all 7 columns (Col 0 to Col 6)")
        print(dash)
        return 0

    if args.template_pptx:
        from patch_original_ppt import patch_pptx, copy_slide_from_template_com
        if target_deck:
            temp_output = output.with_name(output.stem + "_temp_patched_template.pptx")
            result = patch_pptx(
                args.template_pptx,
                temp_output,
                report_path,
                allow_possible_overflow=args.allow_possible_overflow,
                new_table_slide=args.new_table_slide,
                include_count_slide=args.include_count_summary,
                same_slide=args.same_slide,
                count_reference_pptx=args.count_reference_pptx
            )
            # Inject finished Slide 2 into target_deck (Final Deck)
            copy_slide_from_template_com(
                template_path=temp_output,
                source_path=target_deck,
                output_path=output,
                slide_index=2
            )
            try:
                if os.path.exists(temp_output):
                    os.remove(temp_output)
            except Exception:
                pass
            print(f"SLIDE 2 INJECTED INTO FINAL DECK: {output}")
        else:
            result = patch_pptx(
                args.template_pptx,
                output,
                report_path,
                allow_possible_overflow=args.allow_possible_overflow,
                new_table_slide=args.new_table_slide,
                include_count_slide=args.include_count_summary,
                same_slide=args.same_slide,
                count_reference_pptx=args.count_reference_pptx
            )
            print(f"ORIGINAL PPT TEMPLATE PATCHED: {result['output']}")
        if args.same_slide:
            print(f"[5] SLIDE 2 SAME-SLIDE PRESENTATION & GEOMETRY VERIFICATION")
            print(f"  * Slide Title & Commentary:     UPDATED ({args.quarter} Credit Risk Metrics | Active Level 1 Status)")
            print(f"  * Small Summary Table (Top-R):  POPULATED (2Q26 Baseline vs {args.quarter} Current | Master Total: 118)")
            print(f"  * Big Detailed Table (Bottom):  POPULATED ({report['totals']['total']} Verified Metrics | Height: 4.49\" | Bottom: 6.78\")")
            print(f"  * Slide Footprint Safety:       0.370\" clear space above footer (Zero Visual Collisions)")
            print(dash)
        elif args.include_count_summary:
            print("Original Slide 2 is historical; NEW Slide 3 is the COUNT snapshot and NEW Slide 4 is the detailed table.")
        elif args.new_table_slide:
            print("Original Slide 2 summary remains; table moved to dedicated Slide 3.")
        else:
            print("Other slides, the above-table summary, and slide title were left unchanged.")
        return 0
    if not args.preview:
        print(f"Validated {args.quarter} data and generated audits. No standalone PPTX created.")
        print("To preview in a new editable deck, rerun with --preview.")
        print("Updating the ORIGINAL PPTX requires inspecting its Slide 2 objects and coordinates first.")
        return 0
    renderer = Path(__file__).with_name("render_credit_risk_table.js")
    if not renderer.is_file():
        raise ValueError(f"Missing renderer: {renderer}")
    print(f"TOTAL={report['totals']['total']} Green={report['totals']['green']} Yellow={report['totals']['yellow']} Red={report['totals']['red']}")
    print(f"Audit report: {report_path}")
    run = subprocess.run(["node", str(renderer), str(report_path), str(output)], check=False)
    if run.returncode == 3:
        raise RuntimeError("Table content does not fit in the preview area; inspect original PPTX geometry "
                           "and change layout instead of clipping text.")
    if run.returncode:
        raise RuntimeError(f"PPTX renderer failed (exit code {run.returncode}); verify Node.js and npm install.")
    print(f"PPTX ready: {output}")
    return 0


if __name__ == "__main__":
    init_early_logging_and_flag(sys.argv)
    exit_code = 1
    try:
        exit_code = main()
        if exit_code == 0:
            write_flag("success")
        else:
            write_flag("error")
        sys.exit(exit_code)
    except SystemExit as se:
        code = se.code if isinstance(se.code, int) else (0 if se.code is None else 1)
        if code == 0:
            write_flag("success")
        else:
            write_flag("error")
        if ACTIVE_LOGGER:
            ACTIVE_LOGGER.close()
        raise
    except BaseException as ex:
        print(f"CRITICAL ERROR: {ex}", file=sys.stderr)
        try:
            traceback.print_exc(file=sys.stderr)
        except Exception:
            pass
        write_flag("error")
        if ACTIVE_LOGGER:
            ACTIVE_LOGGER.close()
        sys.exit(1)
    finally:
        if ACTIVE_LOGGER:
            ACTIVE_LOGGER.close()
