"""
Unit tests for the MyGate Dashboard Tool.

Run from the repo root:
    python3 -m pytest mygate-dashboard-tool/tests/ -v

Two layers:

1. Synthetic unit tests (no external files needed) — parse/normalize/sort rules,
   workbook structure, chart presence, filter dropdowns, self-consistency.

2. GOLDEN TESTS against the real hand-made dashboard
   ("Mygate Complaints Dashboard as on 22.09.26.xlsm" prepared by Nikhil from
   the Dump Tool's master dump). They run whenever the reference files are
   found at $MYGATE_REF_DIR (default: the Downloads folder with the dumps).
   - test_golden_nikhil_data_roundtrip: the input feed is reconstructed from
     the exact rows Nikhil's file contains, so the generated dashboard must
     match his file 1:1 on the data layer (same 852 rows, same Data sheet, same
      engine). The card/chart layout is deliberately our own uniform-grid design
      now, so layout parity is not tested — only our own invariants.
   - test_golden_real_dump_integration: feeding the actual
     MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-22.xlsx must produce a
     Data sheet that agrees with Nikhil's on every row the two snapshots
     share (the hand-made file was snapshotted at a different time of day, so
     a handful of tickets legitimately drifted between the two downloads —
     the test reports that drift explicitly instead of failing on it).
"""

import importlib.util
import os
import re
import subprocess
import sys
import zipfile
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import pytest

TOOL_DIR = Path(__file__).resolve().parents[1]
REPON_ROOT = TOOL_DIR.parent

_ref_dir_env = os.environ.get("MYGATE_REF_DIR")
if _ref_dir_env:
    REF_DIR = Path(_ref_dir_env)
else:
    REF_DIR = Path("/Users/aditya/Downloads/mygate-new-tool data dumps")

NIKHIL_DASH = REF_DIR / "Mygate Complaints Dashboard as on 22.09.26.xlsm"
NIKHIL_DUMP = REF_DIR / "MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-22.xlsx"

HAS_REFS = NIKHIL_DASH.exists() and NIKHIL_DUMP.exists()

# ---------------------------------------------------------------------------
# import the tool module
# ---------------------------------------------------------------------------

spec = importlib.util.spec_from_file_location(
    "make_dashboard", TOOL_DIR / "make_dashboard.py"
)
make_dashboard = importlib.util.module_from_spec(spec)
sys.modules["make_dashboard"] = make_dashboard
spec.loader.exec_module(make_dashboard)


def write_mini_dump(path: Path, rows, from_str="1-1-2024", to_str="22-9-2026"):
    """Write a dump file exactly the way mygate-dump-tool does (ExcelJS layout):
    row1 title, row2 subtitle, row3 headers, rows4+ data."""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.append(["DLF Independent Floors"])
    ws.append([f"Help Desk Report: From {from_str} To {to_str}"])
    ws.append(["Sr No.", "I.D", "Created Date", "Category", "Sub Category", "Flat",
               "Subject", "Status"])
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


MINI_ROWS = [
    [1, 100, "10-08-24, 09:09 PM", "Products  Appliances ", "Air Conditoners",
     "Phase-A-T-10/2-D", "Fan noise ", "In Progress"],
    [2, 101, "01-01-25, 12:05 AM", "FM Common Area Related Issue", "Carpentry",
     "Phase-B-SR/38-C", "Lift lobby", "Reopened"],
    [3, 102, "22-05-24, 09:36 AM", "Construction Or Project Related", "Cracks",
     "Phase-D-W-2/7-C", "Cracks on walls", "New "],
    [4, 103, "12-09-25, 11:00 AM", "Accounts Billing", "None",
     "Flat 42 - odd one", "No subcat", "On Hold"],
]


@pytest.fixture()
def mini(tmp_path):
    dump = write_mini_dump(tmp_path / "dump.xlsx", MINI_ROWS)
    out = tmp_path / "out.xlsx"
    make_dashboard.main([str(dump), "--as-on", "22.09.26",
                         "--out", str(out), "--quiet"])
    from openpyxl import load_workbook

    return load_workbook(out), tmp_path


# ---------------------------------------------------------------------------
# transformation rules
# ---------------------------------------------------------------------------

def test_collapse_spaces():
    assert make_dashboard.collapse_spaces("Products  Appliances ") == "Products Appliances"
    assert make_dashboard.collapse_spaces("IT   WIFI    Network") == "IT WIFI Network"
    assert make_dashboard.collapse_spaces(None) == ""
    assert make_dashboard.collapse_spaces("  Estate  Infra   Outer ") == "Estate Infra Outer"


def test_derive_phase():
    assert make_dashboard.derive_phase("Phase-A-T-10/2-D") == "Phase A"
    assert make_dashboard.derive_phase("Phase-P-M-11/5-B") == "Phase P"
    assert make_dashboard.derive_phase("Flat 42") == "Unknown"
    assert make_dashboard.derive_phase(None) == "Unknown"


def test_parse_created_date():
    assert make_dashboard.parse_created_date("22-05-24, 09:36 AM") == date(2024, 5, 22)
    assert make_dashboard.parse_created_date("10-08-24, 09:09 PM") == date(2024, 8, 10)
    assert make_dashboard.parse_created_date("01-01-25, 12:05 AM") == date(2025, 1, 1)
    assert make_dashboard.parse_created_date(datetime(2024, 3, 3)) == date(2024, 3, 3)
    assert make_dashboard.parse_created_date("garbage") is None
    assert make_dashboard.parse_created_date(None) is None


def test_ordinal():
    assert make_dashboard.ordinal(1) == "1st"
    assert make_dashboard.ordinal(2) == "2nd"
    assert make_dashboard.ordinal(3) == "3rd"
    assert make_dashboard.ordinal(22) == "22nd"
    assert make_dashboard.ordinal(11) == "11th"
    assert make_dashboard.ordinal(30) == "30th"


def test_month_span_zero_fills():
    span = make_dashboard.month_span(date(2024, 5, 1), date(2024, 9, 1))
    assert span == [(2024, 5), (2024, 6), (2024, 7), (2024, 8), (2024, 9)]


def test_as_on_from_filename(tmp_path):
    dump = write_mini_dump(
        tmp_path / "MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-07.xlsx",
        MINI_ROWS)
    rows, _ = make_dashboard.read_dump_rows(dump)
    assert make_dashboard.derive_as_on(dump, rows) == date(2026, 9, 7)


# ---------------------------------------------------------------------------
# mini end-to-end
# ---------------------------------------------------------------------------

def test_mini_rows_sorted_and_normalized(mini):
    _, tmp = mini
    from openpyxl import load_workbook

    wb = load_workbook(tmp / "out.xlsx")
    ws = wb["Data"]
    # order asc by date: 22-05-24 (id102), 10-08-24 (100), 01-01-25 (101), 12-09-25 (103)
    ids = [ws.cell(r, 2).value for r in range(2, 6)]
    assert ids == [102, 100, 101, 103]
    sr = [ws.cell(r, 1).value for r in range(2, 6)]
    assert sr == [1, 2, 3, 4]
    # normalization: collapse + trim
    # row2 = id102, row3 = id100, row4 = id101, row5 = id103
    assert ws.cell(2, 10).value == "New"          # 'New ' trimmed
    assert ws.cell(3, 5).value == "Products Appliances"  # double-space collapsed
    assert ws.cell(2, 7).value == "Phase D"       # id102: 'Phase-D-W-2/7-C'
    assert ws.cell(5, 7).value == "Unknown"       # id103: 'Flat 42 - odd one'
    assert ws.cell(5, 8).value == "Flat 42 - odd one"


def test_mini_dashboards_structure(mini):
    wb, _ = mini
    assert wb.sheetnames == ["Interactive", "Overview (static)", "Data",
                             "Analysis", "Calc"]
    assert wb["Calc"].sheet_state == "hidden"
    assert wb["Data"].freeze_panes == "A2"
    # 5 charts on each dashboard sheet
    assert len(wb["Interactive"]._charts) == 5
    assert len(wb["Overview (static)"]._charts) == 5
    chart_types = sorted(type(c).__name__ for c in wb["Interactive"]._charts)
    assert chart_types == ["BarChart", "BarChart", "BarChart", "DoughnutChart", "LineChart"]
    # 4 dropdown validations on Interactive
    dvs = wb["Interactive"].data_validations.dataValidation
    assert len(dvs) == 4
    formulas = " | ".join(str(dv.formula1) for dv in dvs)
    assert "SubCategoryFilterList" in formulas
    assert "Phase A" in formulas and "Phase B" in formulas and "Unknown" in formulas
    # statuses dropdown must contain the real statuses found in data
    for dv in dvs:
        f = str(dv.formula1)
        if "In Progress" in f:
            assert "On Hold" in f and "New" in f and "Reopened" in f
    # named range defined
    assert "SubCategoryFilterList" in wb.defined_names
    # KPI engine formulas
    ws = wb["Interactive"]
    # KPI engine formulas (cards live on the uniform grid B:C E:F H:I K:L N:O)
    assert ws["B7"].value == "FILTERED TOTAL" and ws["B8"].value == "=Calc!$V$1"
    assert ws["E7"].value == "% OF ALL" and ws["E8"].value == "=Calc!$V$1/Calc!$V$7"
    assert ws["H7"].value == "IN PROGRESS" and ws["H8"].value == "=Calc!$V$2"
    assert ws["K7"].value == "REOPENED" and ws["K8"].value == "=Calc!$V$3"
    assert ws["N7"].value == "TOP ISSUE COUNT" and ws["N8"].value == "=Calc!$V$5"
    # uniform card grid: 13-wide card columns, 3-wide spacers, 2-wide margins
    dims = ws.column_dimensions
    assert all(dims[c].width == 13.0 for c in "BCEFHIKLNO")
    assert all(dims[c].width == 3.0 for c in "DGJM")
    assert dims["A"].width == 2.0 and dims["P"].width == 2.0
    assert {"B2:O2", "B3:O3"} <= {str(r) for r in ws.merged_cells.ranges}
    calc = wb["Calc"]
    assert calc["N1"].value.startswith('=IF(Interactive!$C$5="All"')
    assert "COUNTIFS" in calc["V1"].value
    assert "In Progress" in calc["V2"].value
    assert 'DATE(2024,5,1)' in calc["T2"].value  # first trend month formula
    # trend zero-fills every month in range (mini data spans May-24..Sep-25)
    months = [calc.cell(r, 19).value for r in range(2, 30) if calc.cell(r, 19).value]
    assert len(months) == 17  # May-24 .. Sep-25 inclusive (8 + 9)


def test_mini_self_consistency(mini):
    """The Overview cards must be live COUNTA/COUNTIF formulas against the Data
    sheet, exactly like the hand-made dashboard, and the status literals used
    must match the statuses actually written into the Data sheet."""
    wb, tmp = mini
    rows = []
    ws = wb["Data"]
    for r in range(2, ws.max_row + 1):
        rows.append(dict(cat=ws.cell(r, 5).value, sub=ws.cell(r, 6).value,
                         phase=ws.cell(r, 7).value, status=ws.cell(r, 10).value,
                         date=ws.cell(r, 4).value))
    counts = Counter(x["status"] for x in rows)
    ov = wb["Overview (static)"]
    fe = make_dashboard.FORMULA_ROW_END
    assert ov["B6"].value == f"=COUNTA(Data!$B$2:$B${fe})"
    assert ov["E6"].value == f'=COUNTIF(Data!$J$2:$J${fe},"In Progress")'
    assert ov["H6"].value == f'=COUNTIF(Data!$J$2:$J${fe},"Reopened")'
    assert ov["K6"].value == f'=COUNTIF(Data!$J$2:$J${fe},"On Hold")'
    assert ov["B5"].value == "TOTAL COMPLAINTS"
    assert ov["E5"].value == "OPEN / IN PROGRESS"
    assert ov["H5"].value == "REOPENED"
    assert ov["K5"].value == "ON HOLD"
    # uniform card grid: 16-wide card columns, 3-wide spacers, 2-wide margins
    dims = ov.column_dimensions
    assert all(dims[c].width == 16.0 for c in "BCEFHIKL")
    assert all(dims[c].width == 3.0 for c in "DGJ")
    assert dims["A"].width == 2.0 and dims["M"].width == 2.0
    assert {"B2:L2", "B3:L3"} <= {str(r) for r in ov.merged_cells.ranges}
    # The literals inside the card formulas must be exactly the status values
    # the Data sheet holds — otherwise a COUNTIF card would silently read 0.
    assert set(counts) <= {"In Progress", "Reopened", "On Hold", "New"}
    # bundle + cli smoke
    result = subprocess.run(
        [sys.executable, str(TOOL_DIR / "make_dashboard.py"),
         str(tmp / "dump.xlsx"), "--as-on", "22.09.26",
         "--out", str(tmp / "cli.xlsx")],
        capture_output=True, text=True,
    )
    assert result.returncode == 0, result.stderr
    assert (tmp / "cli.xlsx").exists()


def test_analysis_band_grid(mini):
    """Design invariants for the Analysis sheet: every section shares one
    identical column treatment (tinted label column, centred counts, thin grey
    boxes, uniform widths), and the Phase/Monthly band starts below the top
    band with its titles intact — the old fixed rows 12-14 let band-1 rows
    overwrite them whenever the sub-category list filled up."""
    wb, _ = mini
    ar = wb["Analysis"]
    # band 1: three tables, headers on row 2
    assert ar["A2"].value == "Category" and ar["B2"].value == "Count"
    assert ar["D2"].value == "Sub-Category" and ar["E2"].value == "Count"
    assert ar["G2"].value == "Status" and ar["H2"].value == "Count"
    # band 2 starts below the longest band-1 body, one empty row between
    t2 = next(r for r in range(3, 40)
              if ar.cell(r, 1).value == "Complaints by Phase / Block")
    assert ar.cell(t2, 4).value == "Monthly Complaint Trend"
    assert ar.cell(t2 + 1, 1).value == "Phase" and ar.cell(t2 + 1, 2).value == "Count"
    assert ar.cell(t2 + 1, 4).value == "Month" and ar.cell(t2 + 1, 5).value == "Count"
    gap_rows = [r for r in range(3, t2)
                if all(ar.cell(r, c).value is None for c in (1, 2, 4, 5, 7, 8))]
    assert gap_rows == [t2 - 1]
    # DLP section: restored as a manual-entry block — the dump carries no
    # purchase dates, so counts stay deliberate dashes (never dead formulas)
    assert ar["J1"].value == "Defect Liability Period (DLP) Status"
    assert ar["J2"].value == "DLP Status" and ar["K2"].value == "Count"
    assert [ar[f"J{i}"].value for i in (3, 4, 5)] == \
        ["Out Of DLP", "Under DLP", "Not Specified"]
    assert [ar[f"K{i}"].value for i in (3, 4, 5)] == ["—", "—", "—"]
    assert all(ar[f"K{i}"].border.top.style == "thin" for i in (3, 4, 5))
    # uniform column rhythm across the three table slots
    dims = ar.column_dimensions
    assert all(dims[c].width == 26.0 for c in "ADG")
    assert all(dims[c].width == 10.0 for c in "BEH")
    assert all(dims[c].width == 3.0 for c in "CFI")
    # same cell treatment in every section: tinted label, boxed count,
    # labels left with an indent, counts centred
    for row, cols in ((3, (1, 4, 7)), (t2 + 2, (1, 4))):
        for col in cols:
            lc, cc = ar.cell(row, col), ar.cell(row, col + 1)
            assert lc.fill.fgColor.rgb == "FFF2F6FC"
            assert lc.alignment.horizontal == "left"
            assert lc.border.left.style == "thin" and cc.border.top.style == "thin"
            assert cc.alignment.horizontal == "center"


def test_fullcalc_on_load(mini):
    wb, _ = mini
    assert wb.calculation.fullCalcOnLoad is True


def test_output_filename_convention():
    assert (make_dashboard.output_name(date(2026, 9, 22))
            == "Mygate Complaints Dashboard as on 22.09.26.xlsx")
    assert (make_dashboard.output_name(date(2026, 10, 1))
            == "Mygate Complaints Dashboard as on 01.10.26.xlsx")


# ---------------------------------------------------------------------------
# GOLDEN TESTS — real hand-made dashboard ("Nikhil's file")
# ---------------------------------------------------------------------------

def _normalize_formula(text, ref_sheet):
    """Make formulas comparable: freeze all Data range ends to $10000 and
    reconcile the intentionally-fixed status labels."""
    text = re.sub(
        rf"({ref_sheet}!\$[A-Z]+\$?2?:\$[A-Z]+)\$(?:9752|793|10000)",
        r"\1$10000",
        text,
    )
    text = text.replace('"HOLD"', '"On Hold"')
    return text


@pytest.mark.skipif(not HAS_REFS, reason="reference files not found")
def test_golden_nikhil_data_roundtrip(tmp_path):
    """Feed the tool THE SAME DATA which Nikhil's file has: his dashboard Data
    sheet rows are converted back into a dump-shaped input file, the tool runs
    on it, and the generated dashboard must match his workbook 1:1 on the
    data layer (the card/chart layout is our own design now)."""
    from openpyxl import load_workbook

    # data_only=True replaces '=Calc!$V$1' with its cached '852', so formulas
    # and values must come from two different loads of the same file.
    ref_f = load_workbook(NIKHIL_DASH)                 # formula strings
    ref_v = load_workbook(NIKHIL_DASH, data_only=True)  # cached values
    ref_ws = ref_v["Data"]

    # rebuild a dump-shaped input from his rows (the reverse of what the tool does)
    dump_rows = []
    for r in range(2, ref_ws.max_row + 1):
        if ref_ws.cell(r, 2).value is None:
            continue
        dump_rows.append([
            None, ref_ws.cell(r, 2).value,
            str(ref_ws.cell(r, 3).value),
            str(ref_ws.cell(r, 5).value),
            str(ref_ws.cell(r, 6).value),
            str(ref_ws.cell(r, 8).value),
            str(ref_ws.cell(r, 9).value),
            str(ref_ws.cell(r, 10).value),
        ])
    dump = write_mini_dump(tmp_path / "nikhil_roundtrip.xlsx", dump_rows)
    out = tmp_path / "generated.xlsx"
    make_dashboard.main([str(dump), "--as-on", "22.09.26",
                         "--out", str(out), "--quiet"])

    gen = load_workbook(out)
    gen_ws = gen["Data"]

    # --- 1) Data sheet must be identical row-for-row, cell-for-cell ---------
    mismatches = []
    gen_r = 2
    for ref_r in range(2, ref_ws.max_row + 1):
        if ref_ws.cell(ref_r, 2).value is None:
            continue
        for c in range(1, 11):
            a = ref_ws.cell(ref_r, c).value
            b = gen_ws.cell(gen_r, c).value
            if isinstance(a, datetime) and isinstance(b, datetime):
                a, b = a.date(), b.date()
            if a != b:
                mismatches.append((gen_r, c, a, b))
        gen_r += 1
    assert gen_ws.max_row >= gen_r - 1
    assert not mismatches, f"{len(mismatches)} Data mismatches: {mismatches[:10]}"

    # --- 2) header + text fixtures ------------------------------------------
    assert ref_ws["B1"].value == gen_ws["B1"].value == "ID"
    assert ref_ws["J1"].value == gen_ws["J1"].value == "Status"

    # --- 3) Interactive KPI cards on OUR card grid ----------------------------
    # The dashboard layout is deliberately our own now (uniform grid), so the
    # expected labels/engine formulas are checked at the new cells instead of
    # being compared against his geometry.
    gi = gen["Interactive"]
    assert gi["B7"].value == "FILTERED TOTAL"
    assert gi["E7"].value == "% OF ALL"
    assert gi["H7"].value == "IN PROGRESS"
    assert gi["K7"].value == "REOPENED"
    assert gi["N7"].value == "TOP ISSUE COUNT"
    assert gi["B8"].value == "=Calc!$V$1"
    assert gi["E8"].value == "=Calc!$V$1/Calc!$V$7"
    assert gi["H8"].value == "=Calc!$V$2"
    assert gi["K8"].value == "=Calc!$V$3"
    assert gi["N8"].value == "=Calc!$V$5"
    assert gi["N9"].value == ("=INDEX(Calc!$D$2:$D$13,"
                              "MATCH(Calc!$V$5,Calc!$E$2:$E$13,0))")

    # --- 4) Calc criteria cells identical ------------------------------------
    gc, rcf = gen["Calc"], ref_f["Calc"]
    for ref_cell in ["N1", "N2", "N3", "N4"]:
        assert _normalize_formula(rcf[ref_cell].value, "Data") == \
            _normalize_formula(gc[ref_cell].value, "Data"), ref_cell
    # KPI engine formulas identical (modulo range length + HOLD->On Hold fix)
    for ref_cell in ["V1", "V2", "V3", "V4", "V5", "V7"]:
        a = _normalize_formula(str(rcf[ref_cell].value), "Data")
        b = _normalize_formula(str(gc[ref_cell].value), "Data")
        assert a == b, (ref_cell, a, b)

    # --- 5) counts agree ------------------------------------------------------
    # python-recomputed category totals equal Nikhil's cached Calc values
    rcv = ref_v["Calc"]
    cats = Counter(
        str(ref_ws.cell(r, 5).value).strip()
        for r in range(2, ref_ws.max_row + 1)
        if ref_ws.cell(r, 2).value is not None
    )
    assert sum(cats.values()) == 852  # Nikhil's own snapshot total
    golden_cat_rows = {str(rcv.cell(r, 1).value).strip(): rcv.cell(r, 2).value
                       for r in range(2, 10) if rcv.cell(r, 1).value is not None}
    for cat, cached in golden_cat_rows.items():
        assert cached == cats[cat], (cat, cached, cats[cat])
    # statuses from Nikhil's Data agree with his Calc cached doughnut values
    stats = Counter(
        str(ref_ws.cell(r, 10).value).strip()
        for r in range(2, ref_ws.max_row + 1)
        if ref_ws.cell(r, 2).value is not None
    )
    for r in range(2, 7):
        label = rcv.cell(r, 7).value
        if label is None:
            continue
        label = str(label).strip()
        assert rcv.cell(r, 8).value == stats.get(label, 0), (label, stats)

    # --- 6) structure: sheets, charts, validations, defined name --------------
    assert gen.sheetnames == ref_f.sheetnames == [
        "Interactive", "Overview (static)", "Data", "Analysis", "Calc"]
    for sheet in ("Interactive", "Overview (static)"):
        gt = sorted(type(c).__name__ for c in gen[sheet]._charts)
        rt = sorted(type(c).__name__ for c in ref_f[sheet]._charts)
        assert gt == rt == ["BarChart", "BarChart", "BarChart",
                            "DoughnutChart", "LineChart"], sheet
        g_titles = sorted(_chart_title(c) for c in gen[sheet]._charts)
        r_titles = sorted(_chart_title(c) for c in ref_f[sheet]._charts)
        assert g_titles == r_titles, (sheet, g_titles, r_titles)
    assert "SubCategoryFilterList" in gen.defined_names
    assert len(gen["Interactive"].data_validations.dataValidation) == 4
    assert gen["Calc"].sheet_state == "hidden"

    # --- 7) Overview KPI cards on OUR card grid -------------------------------
    go = gen["Overview (static)"]
    assert go["B5"].value == "TOTAL COMPLAINTS"
    assert go["E5"].value == "OPEN / IN PROGRESS"
    assert go["H5"].value == "REOPENED"
    assert go["K5"].value == "ON HOLD"
    # Card values are live COUNTA/COUNTIF formulas over the Data sheet (range
    # ends normalized to $10000). The counts match his snapshot because the
    # Data sheets are row-identical — asserted above.
    expected_cards = {
        "B6": f"=COUNTA(Data!$B$2:$B${make_dashboard.FORMULA_ROW_END})",
        "E6": f'=COUNTIF(Data!$J$2:$J${make_dashboard.FORMULA_ROW_END},"In Progress")',
        "H6": f'=COUNTIF(Data!$J$2:$J${make_dashboard.FORMULA_ROW_END},"Reopened")',
        "K6": f'=COUNTIF(Data!$J$2:$J${make_dashboard.FORMULA_ROW_END},"On Hold")',
    }
    for cell, our_formula in expected_cards.items():
        assert (
            _normalize_formula(str(go[cell].value), "Data")
            == _normalize_formula(our_formula, "Data")
        ), cell


def _chart_title(chart):
    try:
        return chart.title.tx.rich.p[0].r[0].t
    except Exception:
        return ""


@pytest.mark.skipif(not HAS_REFS, reason="reference files not found")
def test_golden_real_dump_integration(tmp_path):
    """Run the tool on the ACTUAL Dump-Tool file for 22.09.26. Nikhil's
    dashboard and that dump were captured at different times of day, so the
    Data sheets legitimately differ by a handful of drifted tickets; every row
    the two snapshots share must still be identical."""
    from openpyxl import load_workbook

    out = tmp_path / "generated.xlsx"
    make_dashboard.main([str(NIKHIL_DUMP), "--as-on", "22.09.26",
                         "--out", str(out), "--quiet"])

    ref = load_workbook(NIKHIL_DASH, data_only=True)
    ref_ws = ref["Data"]
    gen = load_workbook(out)
    gen_ws = gen["Data"]

    ref_rows = {}
    for r in range(2, ref_ws.max_row + 1):
        tid = ref_ws.cell(r, 2).value
        if tid is None:
            continue
        ref_rows[int(tid)] = [
            ref_ws.cell(r, c).value for c in range(1, 11)
        ]
    gen_rows = {}
    for r in range(2, gen_ws.max_row + 1):
        tid = gen_ws.cell(r, 2).value
        if tid is None:
            continue
        gen_rows[int(tid)] = [
            gen_ws.cell(r, c).value for c in range(1, 11)
        ]

    common = set(ref_rows) & set(gen_rows)
    assert len(common) > 800, "unexpectedly small overlap with Nikhil's snapshot"

    # Sr No (col 1) is excluded: the snapshots contain different row counts, so
    # numbering shifts wherever membership differs. ID (col 2) is the merge key.
    # MyGate lets users re-classify tickets (category col 5, sub col 6) and move
    # them between statuses (col 10) between the morning snapshot Nikhil worked
    # from and the later dump download — measured drift is 13 cells on 10 of
    # 819 shared rows. Every IMMUTABLE field must nonetheless agree exactly:
    strict_bad = []
    drift = {}
    for tid in sorted(common):
        for c in (3, 4, 7, 8, 9):
            a, b = ref_rows[tid][c - 1], gen_rows[tid][c - 1]
            if isinstance(a, datetime) and isinstance(b, datetime):
                a, b = a.date(), b.date()
            if a != b:
                strict_bad.append((tid, c, a, b))
        moved = [(c, str(ref_rows[tid][c - 1]), str(gen_rows[tid][c - 1]))
                 for c in (5, 6, 10) if ref_rows[tid][c - 1] != gen_rows[tid][c - 1]]
        if moved:
            drift[str(tid)] = moved
    assert not strict_bad, (
        f"{len(strict_bad)} mismatches on immutable columns: {strict_bad[:10]}")
    assert len(drift) <= 12, \
        f"unexpected classification drift between snapshots: {drift}"

    # Nikhil's 852-row snapshot vs the dump's 829: the union difference is the
    # tickets created/closed/reclassified between the two downloads (measured:
    # 33 + 10).
    only_ref = set(ref_rows) - set(gen_rows)
    only_gen = set(gen_rows) - set(ref_rows)
    assert len(only_ref) + len(only_gen) <= 60

    # dashboard self-consistency on the generated file: live card formulas whose
    # literals match the statuses actually written to the Data sheet
    counts = Counter(x[9] for x in gen_rows.values())
    ov = gen["Overview (static)"]
    fe = make_dashboard.FORMULA_ROW_END
    assert ov["B6"].value == f"=COUNTA(Data!$B$2:$B${fe})"
    assert ov["E6"].value == f'=COUNTIF(Data!$J$2:$J${fe},"In Progress")'
    assert ov["H6"].value == f'=COUNTIF(Data!$J$2:$J${fe},"Reopened")'
    assert ov["K6"].value == f'=COUNTIF(Data!$J$2:$J${fe},"On Hold")'
    assert set(counts) <= {"In Progress", "Reopened", "On Hold", "New"}
