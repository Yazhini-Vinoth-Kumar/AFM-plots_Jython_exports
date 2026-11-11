#!/usr/bin/env python3
"""Export accepted AFM curves to four condition CSV files and an Excel workbook.

Run from the GitHub repository root, beside the csv_data directory:
    python export_accepted_afm_curves.py

Excel workbook: AFM_accepted_curves.xlsx (if artifact_tool is installed)
CSV files:      AFM_accepted_curves_CSV/Live.csv, 20_mg_L.csv, 320_mg_L.csv, Dead.csv

The exported values and matching logic follow dashboard.ipynb. Unmatched curve
exports and orphan summary rows are excluded. All five acceptance rules apply.
An individual fit-point CSV is required for every exported row.
"""

import argparse
import json
import shutil
import csv
import math
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

CONDITIONS = ("Live", "20 mg/L", "320 mg/L", "Dead")
SHEET_NAMES = {"Live": "Live", "20 mg/L": "20 mg_L", "320 mg/L": "320 mg_L", "Dead": "Dead"}
CSV_NAMES = {"Live": "Live.csv", "20 mg/L": "20_mg_L.csv", "320 mg/L": "320_mg_L.csv", "Dead": "Dead.csv"}
FOLDER_RE = re.compile(r"^(Live|20|320|Dead)_cell(\d+)_crop(\d+)$", re.I)
CONDITION_MAP = {"live": "Live", "20": "20 mg/L", "320": "320 mg/L", "dead": "Dead"}

# Listed in the same sequence as the HTML values table, plus identifiers
# and the contact coordinates used by its fitted-contact marker.
FIELDS = [
    ("Condition", None),
    ("Cell", None),
    ("Crop", None),
    ("Source folder", None),
    ("Curve number", "curve_number"),
    ("Position index", "position_index"),
    ("Linear slope (N/m)", "linear_slope_bruker_N_per_m"),
    ("Bacterial spring constant (nN/µm)", "bacterial_spring_constant_nN_per_um"),
    ("Linear RMSD custom (pN)", "linear_rmsd_custom_pN"),
    ("Linear X minimum (nm)", "linear_x_min_used_nm"),
    ("Linear X maximum (nm)", "linear_x_max_used_nm"),
    ("Linear fit range (nm)", "linear_fit_range_nm"),
    ("Linear/elasticity fit overlap (nm)", "fit_overlap_nm"),
    ("Young's modulus (MPa)", "youngs_modulus_MPa"),
    ("JPK Elasticity RMS (pN)", "elasticity_rms_bruker_pN"),
    ("Elasticity RMSD custom (pN)", "elasticity_rmsd_custom_pN"),
    ("Elasticity RMS percentage difference (%)", "elasticity_rms_percentage_difference"),
    ("Elasticity X minimum (nm)", "elasticity_x_min_used_nm"),
    ("Elasticity X maximum (nm)", "elasticity_x_max_used_nm"),
    ("Elasticity fit range (nm)", "elasticity_fit_range_nm"),
    ("Elasticity range to preprocessing contact (nm)", "elasticity_fit_range_to_preprocessing_contact_nm"),
    ("Elasticity range to fitted contact (nm)", "elasticity_fit_range_to_fitted_contact_nm"),
    ("Fitted contact X (nm)", "contact_point_x_nm"),
    ("Fitted contact Y (nN)", "contact_point_y_fit_nN"),
    ("Source filename", "filename"),
]
HEADERS = [f[0] for f in FIELDS]
TEXT_COLS = {0, 3, len(FIELDS) - 1}
INTEGER_COLS = {1, 2, 4, 5}


def as_number(text):
    try:
        n = float(text)
    except (ValueError, TypeError, OverflowError):
        return None
    return n if math.isfinite(n) else None


def read_summary(path):
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def first_export_row(path):
    # This reads the identifier only; it does not load or refit the force points.
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return next(csv.DictReader(f), None)


def key(row):
    return (row.get("filename", ""), row.get("curve_number", ""), row.get("position_index", ""))


def span(row, lower_key, upper_key):
    a = as_number(row.get(lower_key))
    b = as_number(row.get(upper_key))
    return None if a is None or b is None else abs(b - a)


def elasticity_to_contact(row, contact_x):
    xmin = as_number(row.get("elasticity_x_min_used_nm"))
    xmax = as_number(row.get("elasticity_x_max_used_nm"))
    if xmin is None or xmax is None or contact_x is None:
        return None
    lo, hi = sorted((xmin, xmax))
    boundary = min(max(contact_x, lo), hi)
    return max(0.0, boundary - lo)


def fit_overlap_nm(row):
    keys = ("linear_x_min_used_nm", "linear_x_max_used_nm",
            "elasticity_x_min_used_nm", "elasticity_x_max_used_nm")
    numbers = [as_number(row.get(k)) for k in keys]
    if any(n is None for n in numbers):
        return None
    a, b = sorted(numbers[:2])
    c, d = sorted(numbers[2:])
    return max(0.0, min(b, d) - max(a, c))


def five_rule_failures(row):
    """Matches apply_afm_five_rules.py and the agreed acceptance limits."""
    failures = []
    linear_rms = as_number(row.get("linear_rmsd_custom_pN"))
    elasticity_rms = as_number(row.get("elasticity_rms_bruker_pN"))
    linear_width = span(row, "linear_x_min_used_nm", "linear_x_max_used_nm")
    elastic_width = elasticity_to_contact(row, 0.0)
    overlap = fit_overlap_nm(row)

    if linear_rms is None or linear_rms > 25.0:
        failures.append("linear RMS > 25 pN or missing")
    if elasticity_rms is None or elasticity_rms > 25.0:
        failures.append("JPK elasticity RMS > 25 pN or missing")
    if linear_width is None or linear_width < 5.0:
        failures.append("linear range < 5 nm or missing")
    if elastic_width is None or elastic_width < 10.0:
        failures.append("elasticity range to X=0 < 10 nm or missing")
    if overlap is None or overlap > 5.0:
        failures.append("fit overlap > 5 nm or missing")
    return failures


def display_values(row):
    # Preserve source numeric strings exactly in CSVs; create calculated fields
    # only when their required inputs are available.
    values = dict(row)
    slope = as_number(row.get("linear_slope_bruker_N_per_m"))
    values["bacterial_spring_constant_nN_per_um"] = (
        None if slope is None else abs(slope) * 1000.0
    )
    values["linear_fit_range_nm"] = span(
        row, "linear_x_min_used_nm", "linear_x_max_used_nm"
    )
    values["elasticity_fit_range_nm"] = span(
        row, "elasticity_x_min_used_nm", "elasticity_x_max_used_nm"
    )
    values["fit_overlap_nm"] = fit_overlap_nm(row)
    values["elasticity_fit_range_to_preprocessing_contact_nm"] = elasticity_to_contact(row, 0.0)
    values["elasticity_fit_range_to_fitted_contact_nm"] = elasticity_to_contact(
        row, as_number(row.get("contact_point_x_nm"))
    )
    jpk = as_number(row.get("elasticity_rms_bruker_pN"))
    custom = as_number(row.get("elasticity_rmsd_custom_pN"))
    if jpk is None or custom is None:
        values["elasticity_rms_percentage_difference"] = None
    else:
        mean_abs = (abs(jpk) + abs(custom)) / 2.0
        values["elasticity_rms_percentage_difference"] = (
            0.0 if mean_abs == 0.0 else abs(jpk - custom) / mean_abs * 100.0
        )
    return values


def to_export_row(condition, cell, crop, folder, source_row):
    values = display_values(source_row)
    output = []
    for title, source_col in FIELDS:
        if title == "Condition":
            val = condition
        elif title == "Cell":
            val = cell
        elif title == "Crop":
            val = crop
        elif title == "Source folder":
            val = folder
        else:
            val = values.get(source_col, "")
        output.append("" if val is None else val)
    return output


def read_accepted(root):
    data_dir = root / "csv_data"
    if not data_dir.is_dir():
        raise FileNotFoundError("csv_data folder not found beside script: {}".format(data_dir))

    entries = {condition: [] for condition in CONDITIONS}
    counters = Counter()
    issues = []
    all_folders = sorted(
        (f for f in data_dir.iterdir() if f.is_dir() and FOLDER_RE.match(f.name)),
        key=lambda f: (
            CONDITIONS.index(CONDITION_MAP[FOLDER_RE.match(f.name).group(1).lower()]),
            int(FOLDER_RE.match(f.name).group(2)),
            int(FOLDER_RE.match(f.name).group(3)),
        ),
    )
    for folder in all_folders:
        m = FOLDER_RE.match(folder.name)
        condition = CONDITION_MAP[m.group(1).lower()]
        cell, crop = int(m.group(2)), int(m.group(3))
        curve_csvs = sorted(folder.glob("AFM_fit_points_curve_*.csv"))
        if not curve_csvs:
            continue
        summary_csv = folder / "AFM_jython_summary.csv"
        if not summary_csv.exists():
            for file in curve_csvs:
                issues.append((folder.name, file.name, "summary missing"))
                counters["unmatched"] += 1
            continue
        summary = read_summary(summary_csv)
        by_key = defaultdict(list)
        by_identity = defaultdict(list)
        for row in summary:
            if row.get("filename") and row.get("position_index"):
                by_key[key(row)].append(row)
            by_identity[(row.get("filename", ""), row.get("position_index", ""))].append(row)

        rows_for_crop = []
        counted_keys = set()
        for path in curve_csvs:
            counters["individual_curve_CSVs"] += 1
            first = first_export_row(path)
            if not first:
                counters["unmatched"] += 1
                issues.append((folder.name, path.name, "empty curve CSV"))
                continue
            hits = by_key.get(key(first), [])
            if len(hits) == 1:
                matched = hits[0]
            elif len(hits) > 1:
                counters["unmatched"] += 1
                issues.append((folder.name, path.name, "ambiguous duplicate full identity"))
                continue
            else:
                # Matches dashboard.ipynb: recover by filename + position
                # only when no exact match and the fallback is unique.
                candidates = by_identity.get(
                    (first.get("filename", ""), first.get("position_index", "")), []
                )
                if len(candidates) != 1:
                    counters["unmatched"] += 1
                    issues.append((folder.name, path.name, "no unique summary identity"))
                    continue
                matched = candidates[0]

            matched_key = key(matched)
            if matched_key in counted_keys:
                counters["duplicate_exports"] += 1
                issues.append((folder.name, path.name, "duplicate fit-point export"))
                continue
            counted_keys.add(matched_key)

            failures = five_rule_failures(matched)
            if failures:
                counters["failed_five_rules"] += 1
                for fail in failures:
                    counters[fail] += 1
                issues.append((folder.name, path.name, "; ".join(failures)))
                continue

            kb = as_number(matched.get("linear_slope_bruker_N_per_m"))
            modulus = as_number(matched.get("youngs_modulus_MPa"))
            if kb is None and modulus is None:
                counters["no_mechanical_parameter"] += 1
                issues.append((folder.name, path.name, "both stiffness and modulus missing"))
                continue

            row = to_export_row(condition, cell, crop, folder.name, matched)
            rows_for_crop.append(row)
            counters[condition] += 1
            counters["accepted_total"] += 1

        rows_for_crop.sort(key=lambda r: (int(r[4]), int(r[5]), str(r[-1])))
        if rows_for_crop:
            entries[condition].append(((cell, crop, folder.name), rows_for_crop))

    return entries, counters, issues


def output_rows(entries):
    """Exactly one empty row between nonempty crop groups."""
    result = {}
    for condition in CONDITIONS:
        matrix = [list(HEADERS)]
        for index, (_, crop_rows) in enumerate(entries[condition]):
            if index:
                matrix.append([""] * len(HEADERS))
            matrix.extend(crop_rows)
        result[condition] = matrix
    return result


def write_csvs(root, sheets):
    out = root / "AFM_accepted_curves_CSV"
    out.mkdir(exist_ok=True)
    for condition in CONDITIONS:
        path = out / CSV_NAMES[condition]
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f).writerows(sheets[condition])
    archive = root / "AFM_accepted_curves_CSV.zip"
    with zipfile.ZipFile(str(archive), "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for condition in CONDITIONS:
            name = CSV_NAMES[condition]
            zf.write(str(out / name), arcname=name)
    return out


def number_for_workbook(v, col_index):
    if v == "" or v is None:
        return None
    if col_index in TEXT_COLS:
        return str(v)
    if col_index in INTEGER_COLS:
        return int(v)
    n = as_number(v)
    return n if n is not None else str(v)


def write_excel(root, sheets):
    try:
        from artifact_tool import Workbook, SpreadsheetFile
    except ImportError:
        return None

    workbook = Workbook.create()
    for condition in CONDITIONS:
        sheet = workbook.worksheets.add(SHEET_NAMES[condition])
        rows = [list(sheets[condition][0])] + [
            [number_for_workbook(value, col_idx) for col_idx, value in enumerate(row)]
            for row in sheets[condition][1:]
        ]
        nrows, ncols = len(rows), len(HEADERS)
        sheet.get_range_by_indexes(0, 0, nrows, ncols).values = rows
        head = sheet.get_range_by_indexes(0, 0, 1, ncols)
        head.format = {
            "fill": "#14324C",
            "font": {"bold": True, "color": "#FFFFFF", "size": 10},
            "wrap_text": True,
            "vertical_alignment": "center",
            "row_height": 42,
        }
        sheet.freeze_panes.freeze_rows(1)
        sheet.get_range_by_indexes(0, 0, nrows, 1).format.column_width = 12
        sheet.get_range_by_indexes(0, 1, nrows, 2).format.column_width = 10
        sheet.get_range_by_indexes(0, 3, nrows, 1).format.column_width = 22
        sheet.get_range_by_indexes(0, 4, nrows, 2).format.column_width = 13
        sheet.get_range_by_indexes(0, 6, nrows, ncols - 7).format.column_width = 20
        sheet.get_range_by_indexes(0, ncols - 1, nrows, 1).format.column_width = 36
        if nrows > 1:
            sheet.get_range_by_indexes(1, 6, nrows - 1, ncols - 7).set_number_format("0.###############")
        # Formula-driven derived metrics. They use the same input columns
        # and clipping rules as the HTML dashboard. Leave crop-separator
        # rows completely empty, without formulas or hidden values.
        derived_formulas = {
            7: lambda r: '=IF(ISNUMBER(G{0}),ABS(G{0})*1000,"")'.format(r),
            11: lambda r: '=IF(COUNT(J{0}:K{0})=2,ABS(K{0}-J{0}),"")'.format(r),
            12: lambda r: ('=IF(AND(COUNT(J{0}:K{0})=2,COUNT(R{0}:S{0})=2),'
                           'MAX(0,MIN(MAX(J{0},K{0}),MAX(R{0},S{0}))-'
                           'MAX(MIN(J{0},K{0}),MIN(R{0},S{0}))),"")').format(r),
            16: lambda r: ('=IF(COUNT(O{0}:P{0})=2,'
                           'IF(ABS(O{0})+ABS(P{0})=0,0,'
                           '200*ABS(O{0}-P{0})/(ABS(O{0})+ABS(P{0}))),"")').format(r),
            19: lambda r: '=IF(COUNT(R{0}:S{0})=2,ABS(S{0}-R{0}),"")'.format(r),
            20: lambda r: ('=IF(COUNT(R{0}:S{0})=2,'
                           'MAX(0,MIN(MAX(R{0},S{0}),0)-MIN(R{0},S{0})),"")').format(r),
            21: lambda r: ('=IF(AND(COUNT(R{0}:S{0})=2,ISNUMBER(W{0})),'
                           'MAX(0,MIN(MAX(R{0},S{0}),MAX(W{0},MIN(R{0},S{0})))-'
                           'MIN(R{0},S{0})),"")').format(r),
        }
        blocks = []
        first = None
        for row_index in range(1, nrows + 1):
            is_data = row_index < nrows and rows[row_index][0] is not None
            if is_data and first is None:
                first = row_index
            elif not is_data and first is not None:
                blocks.append((first, row_index - 1))
                first = None
        for start, end in blocks:
            for col_index, formula in derived_formulas.items():
                formulas = [[formula(row_index + 1)]
                            for row_index in range(start, end + 1)]
                sheet.get_range_by_indexes(start, col_index, end - start + 1, 1).formulas = formulas
    path = root / "AFM_accepted_curves.xlsx"
    SpreadsheetFile.export_xlsx(workbook).save(str(path))
    return path


def install_into_dashboard(root):
    """Add a single, idempotent export cell after the dashboard HTML generator."""
    notebook_path = root / "dashboard.ipynb"
    if not notebook_path.exists():
        raise FileNotFoundError("dashboard.ipynb not found: {}".format(notebook_path))
    nb = json.loads(notebook_path.read_text(encoding="utf-8"))
    marker = "# AFM_ACCEPTED_CURVES_EXPORT_AUTOMATIC_V1"
    for cell in nb.get("cells", []):
        if marker in "".join(cell.get("source", [])):
            print("Notebook export cell already installed.")
            return
    code = (marker + "\n"
            "from pathlib import Path\n"
            "import export_accepted_afm_curves as afm_export\n"
            "afm_export.export_all(Path.cwd())\n")
    nb["cells"].append({
        "cell_type": "code", "execution_count": None, "metadata": {},
        "outputs": [], "source": code.splitlines(keepends=True),
    })
    backup = root / "dashboard.before_excel_export.ipynb"
    if not backup.exists():
        shutil.copy2(notebook_path, backup)
    notebook_path.write_text(json.dumps(nb, ensure_ascii=False, indent=1) + "\n",
                             encoding="utf-8")
    print("Installed exporter as the final cell of dashboard.ipynb")
    print("Backup saved:", backup)


def export_all(root):
    root = Path(root).resolve()
    entries, counts, issues = read_accepted(root)
    sheets = output_rows(entries)
    csv_folder = write_csvs(root, sheets)
    try:
        workbook = write_excel(root, sheets)
    except Exception as exc:
        print("Excel export could not be generated: {}".format(exc), file=sys.stderr)
        workbook = None
    print("Exported {} accepted curves.".format(counts["accepted_total"]))
    for condition in CONDITIONS:
        print("  {:<9}: {} curves across {} crops".format(
            condition, counts[condition], len(entries[condition])
        ))
    print("  Rejected by five rules:", counts["failed_five_rules"])
    print("  Unmatched curve exports:", counts["unmatched"])
    print("  Duplicate curve exports:", counts["duplicate_exports"])
    print("CSV files:", csv_folder)
    print("CSV zip:", root / "AFM_accepted_curves_CSV.zip")
    if workbook:
        print("Excel workbook:", workbook)
    else:
        print("Excel workbook not created; artifact_tool is not available. Four CSVs were created.")
    if issues:
        print("\nEXCLUSIONS (not present in accepted files):")
        for folder, path, reason in issues:
            print("  {} / {}: {}".format(folder, path, reason))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent,
                        help="Repository root containing csv_data (default: script folder)")
    parser.add_argument("--install-into-dashboard", action="store_true",
                        help="Add an export cell to dashboard.ipynb so Run All exports again")
    args = parser.parse_args()
    root = args.root.resolve()
    export_all(root)
    if args.install_into_dashboard:
        install_into_dashboard(root)


if __name__ == "__main__":
    main()
