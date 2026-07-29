#!/usr/bin/env python3
"""
Build the pharmacy cashback calculator HTML tool from a merchandising Excel file.

Usage:
    python build_tool.py <path/to/merchandising.xlsx> [output.html]
    python build_tool.py <path/to/merchandising.xlsx> --dump-headers
    python build_tool.py <path/to/merchandising.xlsx> [output.html] --name-col E --q3-dist-col P ...

The Excel file is expected to have one row per pharmacy with columns
(header text match is fuzzy/case-insensitive, order doesn't matter) for:
  - pharmacy name              (e.g. "Ims_Customer_Came", "Customer Name", "Pharmacy Name")
  - customer code              (e.g. "Customer_Code")
  - MR / rep name               (e.g. "MR Name")
  - DM / district manager name (e.g. "DM Name")
  - region                     (e.g. "Region_Name")
  - territory                  (e.g. "Territory_Name")
  - pharmacy status            (e.g. "Pharmacy Status")
  - Q3 actual sales            (header containing "Q3" and "Actual")
  - Q3 distributor sales       (header containing "Q3" and "DIST")
  - Q4 actual sales            (header containing "Q4" and "Actual")
  - Q4 distributor sales       (header containing "Q4" and "DIST")

It scans the first 15 rows of the first sheet to find the header row
(whichever row matches the most expected labels), then reads every row
below it until it hits a run of empty pharmacy-name cells.

WHEN AUTO-DETECTION ISN'T ENOUGH: if a source file uses header wording this
script doesn't recognize, or has no clear headers at all, don't guess at the
mapping. Run with --dump-headers first to see exactly what's in the file
(row-by-row, column letter + text), ask the user which column is which, then
re-run passing the confirmed columns explicitly via --name-col, --code-col,
--mr-col, --dm-col, --region-col, --territory-col, --status-col,
--q3-actual-col, --q3-dist-col, --q4-actual-col, --q4-dist-col (each takes a
column letter like "E"), plus --header-row if the header row itself was
misdetected. Explicit overrides always win over auto-detection for that field.

Output: a single self-contained HTML file (Arabic RTL) with the pharmacy
data embedded, ready to open in any browser. No sales data leaves the
user's machine at build time -- everything is local file I/O.
"""
import sys
import json
import re
import argparse
from pathlib import Path

try:
    import openpyxl
    from openpyxl.utils import column_index_from_string, get_column_letter
except ImportError:
    print("Missing dependency. Run: pip install openpyxl --break-system-packages", file=sys.stderr)
    sys.exit(1)

HEADER_PATTERNS = {
    "n":  [r"ims_customer_came", r"customer name", r"pharmacy name", r"customer_name"],
    "c":  [r"customer_code", r"^code$", r"customer code"],
    "mr": [r"mr name", r"\bmr\b.*name"],
    "dm": [r"dm name", r"\bdm\b.*name"],
    "rg": [r"region_name", r"^region$"],
    "tr": [r"territory_name", r"^territory$"],
    "st": [r"pharmacy status", r"^status$"],
}

# Maps CLI override flag names to the internal column keys used everywhere else.
OVERRIDE_KEYS = {
    "name_col": "n", "code_col": "c", "mr_col": "mr", "dm_col": "dm",
    "region_col": "rg", "territory_col": "tr", "status_col": "st",
    "q3_actual_col": "q3a", "q3_dist_col": "q3d",
    "q4_actual_col": "q4a", "q4_dist_col": "q4d",
}

def find_col(header_row, patterns):
    for cell_val, col_idx in header_row:
        if cell_val is None:
            continue
        text = str(cell_val).strip().lower()
        for pat in patterns:
            if re.search(pat, text):
                return col_idx
    return None

def find_quarter_cols(header_row, quarter_label):
    """Find the Actual and DIST sales columns for a given quarter (e.g. 'q3', 'q4')."""
    actual_col = dist_col = None
    for cell_val, col_idx in header_row:
        if cell_val is None:
            continue
        text = str(cell_val).strip().lower()
        if quarter_label in text and "actual" in text:
            actual_col = col_idx
        elif quarter_label in text and "dist" in text:
            dist_col = col_idx
    return actual_col, dist_col

def _read_all_rows(xlsx_path):
    """Load a workbook's first sheet in read-only mode and materialize every row as a
    tuple via iter_rows (sequential access). This is dramatically faster than repeated
    ws.cell(row, column) random-access lookups on large sheets -- that pattern alone
    was taking 45+ seconds on a ~2,750-row x 100+ column file. Returns (all_rows, max_col)
    where all_rows[r-1] is row r's tuple of values (1-indexed row numbers, 0-indexed tuple)."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    all_rows = [row for row in ws.iter_rows(values_only=True)]
    max_col = ws.max_column
    wb.close()
    return all_rows, max_col

def locate_header_row(all_rows, max_scan_rows=15):
    best_row, best_score = None, -1
    for r in range(1, min(max_scan_rows, len(all_rows)) + 1):
        row_vals = all_rows[r - 1]
        row_cells = list(enumerate(row_vals, start=1))
        row_cells = [(v, c) for c, v in row_cells]
        score = 0
        for patterns in HEADER_PATTERNS.values():
            if find_col(row_cells, patterns) is not None:
                score += 1
        q3a, q3d = find_quarter_cols(row_cells, "q3")
        q4a, q4d = find_quarter_cols(row_cells, "q4")
        score += sum(x is not None for x in [q3a, q3d, q4a, q4d])
        if score > best_score:
            best_score, best_row = score, r
    return best_row

def dump_headers(xlsx_path, max_scan_rows=15):
    """Print every non-empty header cell (row by row, column letter + text) so a
    human can see exactly what's in the file and say which column is which."""
    all_rows, max_col = _read_all_rows(xlsx_path)
    detected = locate_header_row(all_rows, max_scan_rows)

    print(f"(auto-detected header row: {detected})\n")
    for r in range(1, min(max_scan_rows, len(all_rows)) + 1):
        row_vals = all_rows[r - 1]
        nonblank = [(v, c) for c, v in enumerate(row_vals, start=1) if v is not None and str(v).strip() != ""]
        if not nonblank:
            continue
        marker = "  <-- auto-detected header row" if r == detected else ""
        print(f"Row {r}:{marker}")
        for v, c in nonblank:
            print(f"  {get_column_letter(c)}: {v!r}")
        print()

def extract_data(xlsx_path, overrides=None, forced_header_row=None):
    overrides = overrides or {}
    all_rows, max_col = _read_all_rows(xlsx_path)

    header_r = forced_header_row if forced_header_row else locate_header_row(all_rows)
    if header_r is None:
        raise RuntimeError(
            "Could not locate a header row in the first sheet. "
            "Run with --dump-headers to see the raw file contents, then re-run "
            "with --header-row and the --*-col overrides once you know which is which."
        )

    header_vals = all_rows[header_r - 1]
    row_cells = list(zip(header_vals, range(1, len(header_vals) + 1)))
    cols = {key: find_col(row_cells, pats) for key, pats in HEADER_PATTERNS.items()}
    q3a_col, q3d_col = find_quarter_cols(row_cells, "q3")
    q4a_col, q4d_col = find_quarter_cols(row_cells, "q4")
    auto_cols = {**cols, "q3a": q3a_col, "q3d": q3d_col, "q4a": q4a_col, "q4d": q4d_col}

    # Explicit column-letter overrides always win over auto-detection.
    for override_flag, key in OVERRIDE_KEYS.items():
        letter = overrides.get(override_flag)
        if letter:
            auto_cols[key] = column_index_from_string(letter.strip().upper())

    missing = [k for k in ("n", "q3d", "q4d") if auto_cols.get(k) is None]
    if missing:
        header_preview = "\n".join(
            f"  {get_column_letter(c)}: {v!r}" for v, c in row_cells if v is not None and str(v).strip() != ""
        )
        raise RuntimeError(
            f"Could not find required columns for: {missing}.\n"
            f"Header row {header_r} contains:\n{header_preview}\n\n"
            f"Don't guess -- ask the user which column above is which, then re-run with "
            f"the matching --name-col / --q3-dist-col / --q4-dist-col (etc.) flags using "
            f"the column letters shown."
        )

    def val(row_vals, col):
        if col is None or col > len(row_vals):
            return None
        return row_vals[col - 1]

    rows = []
    blank_streak = 0
    r = header_r + 1
    while r <= len(all_rows) and blank_streak < 25:
        row_vals = all_rows[r - 1]
        name = val(row_vals, auto_cols["n"])
        if name is None or str(name).strip() == "":
            blank_streak += 1
            r += 1
            continue
        blank_streak = 0
        rows.append({
            "c":  val(row_vals, auto_cols["c"]),
            "n":  str(name).strip(),
            "mr": str(val(row_vals, auto_cols["mr"]) or "").strip(),
            "dm": str(val(row_vals, auto_cols["dm"]) or "").strip(),
            "rg": str(val(row_vals, auto_cols["rg"]) or "").strip(),
            "tr": str(val(row_vals, auto_cols["tr"]) or "").strip(),
            "st": str(val(row_vals, auto_cols["st"]) or "").strip(),
            "q3d": round(float(val(row_vals, auto_cols["q3d"]) or 0), 2),
            "q3a": round(float(val(row_vals, auto_cols["q3a"]) or 0), 2) if auto_cols["q3a"] else 0,
            "q4d": round(float(val(row_vals, auto_cols["q4d"]) or 0), 2),
            "q4a": round(float(val(row_vals, auto_cols["q4a"]) or 0), 2) if auto_cols["q4a"] else 0,
        })
        r += 1

    return rows

SUPPLEMENT_HEADER_PATTERNS = {
    "code": [r"customer_code", r"^code$", r"customer code"],
    "value": [r"dist_value", r"distributor.*value", r"^value$"],
}

def load_supplement_map(xlsx_path, code_col_override=None, value_col_override=None, header_row=1):
    """Load a secondary, differently-shaped sales file (e.g. a weekly distributor
    extract) and sum its sales column per customer code. Returns {code: total_value}.

    This is intentionally separate from extract_data/HEADER_PATTERNS: a weekly/
    ad-hoc sales file usually won't share the merchandising list's column layout
    (no MR/DM, no Q3/Q4 slabs), so don't try to force it through the same auto-
    detection. If this can't confidently find a code + value column, it raises
    rather than guessing -- dump the file's headers and ask the user which
    columns are which, then pass --supplement-code-col / --supplement-value-col.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]

    header_cells = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
    row_cells = [(v, i + 1) for i, v in enumerate(header_cells)]

    code_col = column_index_from_string(code_col_override.strip().upper()) if code_col_override else find_col(row_cells, SUPPLEMENT_HEADER_PATTERNS["code"])
    value_col = column_index_from_string(value_col_override.strip().upper()) if value_col_override else find_col(row_cells, SUPPLEMENT_HEADER_PATTERNS["value"])

    if code_col is None or value_col is None:
        header_preview = "\n".join(f"  {get_column_letter(c)}: {v!r}" for v, c in row_cells if v is not None and str(v).strip() != "")
        raise RuntimeError(
            f"Could not confidently find a customer-code column and/or a sales-value column in "
            f"{xlsx_path}.\nHeader row {header_row} contains:\n{header_preview}\n\n"
            f"Don't guess -- ask the user which column is which, then re-run with "
            f"--supplement-code-col and --supplement-value-col using the column letters shown."
        )

    totals = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        code = row[code_col - 1]
        value = row[value_col - 1]
        if code is None or str(code).strip() == "":
            continue
        code = str(code).strip()
        totals[code] = totals.get(code, 0) + (float(value) if value else 0)
    wb.close()
    return totals

def merge_supplement(rows, supplement_map, field_key):
    """Attach a supplemental sales figure to each row by matching on customer code.
    Rows with no match in the supplement file get None (not 0), so the tool can
    show "no data for this period" instead of implying zero sales."""
    matched = 0
    for row in rows:
        code = str(row.get("c") or "").strip()
        if code in supplement_map:
            row[field_key] = round(supplement_map[code], 2)
            matched += 1
        else:
            row[field_key] = None
    return matched

def load_monthly_supplement_map(xlsx_path, code_col_override=None, header_row=1):
    """Load a file shaped with one customer-code column plus one column per month
    (header like 202601, 202602, ... -- YYYYMM as an int or string). Returns
    (monthly_map, months) where monthly_map = {code: {month_int: value}} and
    months is the sorted list of month ints found, so the caller/template can
    render a proper multi-point time series instead of a single lump figure."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]

    header_cells = next(ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True))
    row_cells = [(v, i + 1) for i, v in enumerate(header_cells)]

    code_col = column_index_from_string(code_col_override.strip().upper()) if code_col_override else find_col(row_cells, SUPPLEMENT_HEADER_PATTERNS["code"])

    month_cols = {}
    for v, c in row_cells:
        if v is None:
            continue
        text = str(v).strip()
        m = re.match(r"^(20\d{2})(0[1-9]|1[0-2])$", text)
        if m:
            month_cols[c] = int(text)

    if code_col is None or not month_cols:
        header_preview = "\n".join(f"  {get_column_letter(c)}: {v!r}" for v, c in row_cells if v is not None and str(v).strip() != "")
        raise RuntimeError(
            f"Could not confidently find a customer-code column and/or YYYYMM month columns in "
            f"{xlsx_path}.\nHeader row {header_row} contains:\n{header_preview}\n\n"
            f"Don't guess -- ask the user which column is which, then re-run with "
            f"--s1-code-col and/or --s1-header-row."
        )

    monthly_map = {}
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        code = row[code_col - 1]
        if code is None or str(code).strip() == "":
            continue
        code = str(code).strip()
        entry = monthly_map.setdefault(code, {})
        for c, month_int in month_cols.items():
            v = row[c - 1] if c - 1 < len(row) else None
            entry[month_int] = round(entry.get(month_int, 0) + (float(v) if v else 0), 2)
    wb.close()
    return monthly_map, sorted(month_cols.values())

def merge_monthly_supplement(rows, monthly_map, field_key="s1"):
    """Attach a {month_str: value} dict per row under field_key, None if the
    pharmacy's customer code has no match in the monthly file."""
    matched = 0
    for row in rows:
        code = str(row.get("c") or "").strip()
        if code in monthly_map:
            row[field_key] = {str(m): v for m, v in sorted(monthly_map[code].items())}
            matched += 1
        else:
            row[field_key] = None
    return matched

def build_html(rows, template_path, output_path, supplement_label=None, week1_label=None):
    template = Path(template_path).read_text(encoding="utf-8")
    data_json = json.dumps(rows, ensure_ascii=False, separators=(",", ":"))
    label_json = json.dumps(supplement_label, ensure_ascii=False) if supplement_label else "null"
    week1_label_json = json.dumps(week1_label, ensure_ascii=False) if week1_label else "null"
    html = (template
            .replace("__DATA__", data_json)
            .replace("__COUNT__", str(len(rows)))
            .replace("__SUPPLEMENT_LABEL__", label_json)
            .replace("__WEEK1_LABEL__", week1_label_json))
    Path(output_path).write_text(html, encoding="utf-8")

def parse_args():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("xlsx_path")
    p.add_argument("output_path", nargs="?", default="pharmacy_cashback_tool.html")
    p.add_argument("--dump-headers", action="store_true",
                   help="Print every header cell found (row by row, column letter + text) and exit. "
                        "Use this first when a file's structure is unfamiliar, so you can ask the user "
                        "which column is which instead of guessing.")
    p.add_argument("--header-row", type=int, default=None,
                   help="Force which row number is the actual header row, if auto-detection picks the wrong one.")
    p.add_argument("--name-col")
    p.add_argument("--code-col")
    p.add_argument("--mr-col")
    p.add_argument("--dm-col")
    p.add_argument("--region-col")
    p.add_argument("--territory-col")
    p.add_argument("--status-col")
    p.add_argument("--q3-actual-col")
    p.add_argument("--q3-dist-col")
    p.add_argument("--q4-actual-col")
    p.add_argument("--q4-dist-col")
    p.add_argument("--supplement-file",
                   help="Path to a secondary sales file (e.g. a weekly distributor extract) to merge in "
                        "by customer code, as an extra 'wk' figure per pharmacy alongside the Q3/Q4 data. "
                        "Does not affect growth %% or slab/cashback calculations -- purely informational.")
    p.add_argument("--supplement-label", default="بيانات إضافية",
                   help="Label shown in the tool for the supplemental figure (e.g. 'July 2026').")
    p.add_argument("--supplement-code-col")
    p.add_argument("--supplement-value-col")
    p.add_argument("--supplement-header-row", type=int, default=1)
    p.add_argument("--s1-file",
                   help="Path to a monthly-shaped file (customer code + one column per YYYYMM) to merge "
                        "in as each pharmacy's 's1' monthly sales series, shown as extra points on the "
                        "sales-history chart. Purely informational -- does not affect growth/slab/cashback.")
    p.add_argument("--s1-code-col")
    p.add_argument("--s1-header-row", type=int, default=1)
    p.add_argument("--week1-file",
                   help="Path to a week-1-of-month distributor extract (customer code + value column) to "
                        "merge in as each pharmacy's 'wk1' figure, used by the week-1 pacing card to "
                        "project a full-month (and full-quarter) run-rate. Re-run this script with an "
                        "updated --week1-file each time a new week-1 extract is dropped on Google Drive.")
    p.add_argument("--week1-label", default="الأسبوع الأول من الشهر",
                   help="Label shown in the tool for the week-1 pacing figure.")
    p.add_argument("--week1-code-col")
    p.add_argument("--week1-value-col")
    p.add_argument("--week1-header-row", type=int, default=1)
    return p.parse_args()

def main():
    args = parse_args()

    if args.dump_headers:
        dump_headers(args.xlsx_path)
        return

    overrides = {
        "name_col": args.name_col, "code_col": args.code_col,
        "mr_col": args.mr_col, "dm_col": args.dm_col,
        "region_col": args.region_col, "territory_col": args.territory_col,
        "status_col": args.status_col,
        "q3_actual_col": args.q3_actual_col, "q3_dist_col": args.q3_dist_col,
        "q4_actual_col": args.q4_actual_col, "q4_dist_col": args.q4_dist_col,
    }
    template_path = Path(__file__).parent.parent / "assets" / "tool_template.html"

    rows = extract_data(args.xlsx_path, overrides=overrides, forced_header_row=args.header_row)
    if not rows:
        print("No pharmacy rows extracted -- double check the source file structure.", file=sys.stderr)
        sys.exit(1)

    supplement_label = None
    if args.supplement_file:
        supplement_map = load_supplement_map(
            args.supplement_file,
            code_col_override=args.supplement_code_col,
            value_col_override=args.supplement_value_col,
            header_row=args.supplement_header_row,
        )
        matched = merge_supplement(rows, supplement_map, "wk")
        supplement_label = args.supplement_label
        print(f"Merged supplement file: {matched}/{len(rows)} pharmacies matched by customer code.")

    if args.s1_file:
        monthly_map, months = load_monthly_supplement_map(
            args.s1_file,
            code_col_override=args.s1_code_col,
            header_row=args.s1_header_row,
        )
        matched = merge_monthly_supplement(rows, monthly_map, "s1")
        print(f"Merged S1 monthly file: {matched}/{len(rows)} pharmacies matched by customer code. Months found: {months}")

    week1_label = None
    if args.week1_file:
        week1_map = load_supplement_map(
            args.week1_file,
            code_col_override=args.week1_code_col,
            value_col_override=args.week1_value_col,
            header_row=args.week1_header_row,
        )
        matched = merge_supplement(rows, week1_map, "wk1")
        week1_label = args.week1_label
        print(f"Merged week-1 file: {matched}/{len(rows)} pharmacies matched by customer code.")

    build_html(rows, template_path, args.output_path, supplement_label=supplement_label, week1_label=week1_label)
    print(f"Extracted {len(rows)} pharmacies. Wrote {args.output_path}")

if __name__ == "__main__":
    main()
