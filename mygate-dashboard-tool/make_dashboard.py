#!/usr/bin/env python3
"""
MyGate Dashboard Tool
=====================
Builds a "Mygate Complaints Dashboard" workbook from a MyGate Master Helpdesk
Dump file (the kind produced by the `mygate-dump-tool` Chrome extension).

This closes the pipeline loop of the mygate-tools repo:

    mygate-dump-tool  ──▶  Master Helpdesk Dump (.xlsx)  ──▶  mygate-dashboard-tool  ──▶
                                                                 Complaints Dashboard (.xlsx)

The dashboard reproduces the structure of the manually prepared
"Mygate Complaints Dashboard as on DD.MM.YY" workbooks:

    ├── Interactive        live dashboard: Phase / Status / Sub Category / Category
    │                      dropdown filters, KPI cards and 5 auto-updating charts
    ├── Overview (static)  snapshot KPI cards + 5 charts fixed at build time
    ├── Data               normalized dump rows (re-sorted, Phase derived, dates parsed)
    ├── Analysis           COUNTIF aggregation feed for the Overview charts
    └── Calc (hidden)      live COUNTIFS engine driving the Interactive dashboard

It is a Python CLI (not a Chrome extension) because ExcelJS — the spreadsheet
library the extensions use — cannot create charts, while openpyxl can. It is
also what the original hand-made dashboards were built with.

Usage:

    python3 make_dashboard.py "MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-22.xlsx"

    # optional overrides
    --as-on 22.09.26            # as-on label (default: derived from dump name / data)
    --out "...xlsx"             # explicit output path
    --quiet                     # suppress the summary

Requires: python3, openpyxl>=3.1  (pip install openpyxl)
"""

from __future__ import annotations

import argparse
import re
import zipfile
from collections import Counter
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.axis import ChartLines, DateAxis
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.title import Title
from openpyxl.chart.text import RichText, Text
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Font as DrawingFont,
    Paragraph,
    ParagraphProperties,
    RegularTextRun,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Branding / styling constants (as in the manual dashboards)
# ---------------------------------------------------------------------------

BRAND = "DLF Independent Floors"

NAVY = "FF1F3864"
MID_BLUE = "FF2E5496"
BLUE = "FF4472C4"
RED = "FFC0392B"
PURPLE = "FF6C5B7B"
GOLD = "FFE0A800"
BAND = "FFF2F5FB"
WHITE = "FFFFFFFF"
GREY_BORDER = "FFBFBFBF"
GREY_LINE = "878787"
SUBTLE_TEXT = "FF44546A"

CHART_LINE_COLOR = "2E5496"
CHART_BAR_CATEGORY = "4472C4"
CHART_BAR_TOP = "1F3864"

DOUGHNUT_PALETTE = ["4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "8064A2"]

DEFAULT_PERIOD_FROM = date(2024, 1, 1)

# The formula engine counts over a generous fixed range so the dashboard keeps
# working if rows are appended to Data manually (the hand-made file did the
# same with $2:$9752).
FORMULA_ROW_END = 10000

DATA_HEADERS = [
    "Sr No.",
    "ID",
    "Created Date",
    "Date",
    "Category",
    "Sub Category",
    "Phase",
    "Flat",
    "Subject",
    "Status",
]

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

STATUS_ORDER_FALLBACK = ["In Progress", "Reopened", "New", "On Hold"]


# ---------------------------------------------------------------------------
# Normalization / parsing helpers (verified against the hand-made dashboards)
# ---------------------------------------------------------------------------

def collapse_spaces(value):
    """Normalize a string cell: trim + collapse whitespace runs to one space.

    MyGate report exports contain artifacts like 'Products  Appliances ' and
    'Estate  Infra   Outer Area from the plot'; the hand-made dashboards store
    them normalized exactly like this.
    """
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def derive_phase(flat):
    """'Phase-A-T-10/2-D' -> 'Phase A'. Non-matching flats -> 'Unknown'."""
    if flat is None:
        return "Unknown"
    m = re.match(r"^\s*Phase[-\s]+([A-Za-z0-9]+)", str(flat))
    if not m:
        return "Unknown"
    letter = m.group(1).upper()
    return f"Phase {letter if len(letter) == 1 else letter}"


CREATED_DATE_RE = re.compile(
    r"^\s*(\d{1,2})-(\d{1,2})-(\d{2,4})(?:\s*,\s*(\d{1,2}):(\d{2})\s*([AaPp])\.?[Mm])?"
)


def parse_created_date(value):
    """Extract the calendar date from a dump 'Created Date' cell.

    Handles 'dd-mm-yy, hh:mm AM/PM' text (what the Dump Tool writes) and real
    date/datetime objects. Returns a date or None.
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    m = CREATED_DATE_RE.match(str(value))
    if not m:
        return None
    dd, mm, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if yy < 100:  # '25'/'26' style two-digit years used by the dump text
        yy += 2000
    try:
        return date(yy, mm, dd)
    except ValueError:
        return None


def ordinal(n):
    if 10 <= n % 100 <= 20:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }"


def month_span(first: date, last: date):
    """Every (year, month) between first and last, always asc, zero included."""
    out = []
    y, m = first.year, first.month
    while (y, m) <= (last.year, last.month):
        out.append((y, m))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


def fmt_short(d: date) -> str:
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year}"


# ---------------------------------------------------------------------------
# Dump parsing
# ---------------------------------------------------------------------------

HEADER_ALIASES = {
    "id": "id",
    "i.d": "id",
    "ticket id": "id",
    "created date": "created",
    "createddate": "created",
    "date": "created",
    "category": "category",
    "sub category": "sub",
    "sub-category": "sub",
    "subcategory": "sub",
    "flat": "flat",
    "house": "flat",
    "subject": "subject",
    "description": "subject",
    "status": "status",
    "mygate status": "status",
}


def read_dump_rows(path: Path):
    """Parse a Master Helpdesk Dump workbook (as written by mygate-dump-tool).

    Returns (rows, meta). rows are normalized dicts sorted ascending by
    (date, ticket id) — exactly how the hand-made dashboards store them on the
    Data sheet.
    """
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]

    header_row_idx = None
    colmap = {}
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=12, max_col=20), start=1):
        keys = {}
        for c in row:
            if c.value is None or c.column in keys:
                continue
            k = collapse_spaces(c.value).lower()
            if k in HEADER_ALIASES:
                keys[c.column] = HEADER_ALIASES[k]
        if any(v == "id" for v in keys.values()) and any(
            v == "created" for v in keys.values()
        ):
            header_row_idx = r
            for col, alias in keys.items():
                if alias not in colmap:
                    colmap[alias] = col
            break

    if not colmap:
        wb.close()
        raise SystemExit(
            "ERROR: could not find a header row (I.D / Created Date / Category ...) "
            "in the first 12 rows of the dump file."
        )

    meta = {"from": None, "to": None}
    for row in ws.iter_rows(min_row=1, max_row=max(2, header_row_idx - 1), max_col=3):
        for c in row:
            if not isinstance(c.value, str):
                continue
            m = re.search(r"From\s+([\d\-]+)\s+To\s+([\d\-]+)", c.value)
            if m:
                for key, raw in (("from", m.group(1)), ("to", m.group(2))):
                    parts = raw.split("-")
                    if len(parts) == 3 and all(parts):
                        try:
                            d, mm, yy = (int(p) for p in parts)
                            meta[key] = date(yy, mm, d)
                        except ValueError:
                            pass

    raw_rows = []
    for row in ws.iter_rows(min_row=header_row_idx + 1, max_col=ws.max_column):
        cell_of = {c.column: c.value for c in row if c.value is not None}

        def get(alias):
            return cell_of.get(colmap[alias]) if alias in colmap else None

        tid = get("id")
        if tid is None or (isinstance(tid, str) and not tid.strip()):
            continue
        if isinstance(tid, float) and tid.is_integer():
            tid = int(tid)
        raw_rows.append(
            dict(
                id=tid,
                created=get("created"),
                category=collapse_spaces(get("category")) or "Unknown",
                sub=collapse_spaces(get("sub")) or None,
                flat=collapse_spaces(get("flat")),
                subject="" if get("subject") is None else str(get("subject")),
                status=collapse_spaces(get("status")) or "Unknown",
            )
        )
    wb.close()

    rows = []
    for item in raw_rows:
        sub = item["sub"]
        rows.append(
            dict(
                id=item["id"],
                created=item["created"],
                category=item["category"],
                sub="None" if sub in (None, "", "None") else sub,
                flat=item["flat"],
                subject=item["subject"],
                status=item["status"],
                date=parse_created_date(item["created"]),
                phase=derive_phase(item["flat"]),
            )
        )
    rows.sort(
        key=lambda r: (
            r["date"] or date.max,
            r["id"] if isinstance(r["id"], (int, float)) else float("inf"),
        )
    )
    return rows, meta


# ---------------------------------------------------------------------------
# Aggregation (mirrors what the Calc COUNTIFS engine computes live)
# ---------------------------------------------------------------------------

def aggregate(rows):
    stats = dict(
        total=len(rows),
        status=Counter(r["status"] for r in rows),
        category=Counter(r["category"] for r in rows),
        subcat=Counter(r["sub"] for r in rows),
        phase=Counter(r["phase"] for r in rows),
    )
    stats["month"] = Counter(
        (r["date"].year, r["date"].month) for r in rows if r["date"]
    )

    def desc(counter):
        return [k for k, _ in counter.most_common()]

    stats["status_order"] = desc(stats["status"]) or list(STATUS_ORDER_FALLBACK)
    stats["category_order"] = desc(stats["category"])
    stats["phase_order"] = desc(stats["phase"])

    # Chart list: top 12 real sub-categories; 'None'/blank excluded (as in the
    # hand-made file's D2:D13 list). The filter dropdown still offers 'None'.
    stats["subcat_chart"] = [
        s for s in desc(stats["subcat"]) if s not in ("None", "")
    ][:12]
    stats["subcat_filter"] = ["All"] + sorted(
        (s for s in stats["subcat"] if s != ""), key=str.lower
    )
    dated = [r["date"] for r in rows if r["date"]]
    stats["months"] = (
        month_span(min(dated), max(dated))
        if dated
        else month_span(date.today().replace(day=1), date.today())
    )
    return stats


# ---------------------------------------------------------------------------
# Workbook construction
# ---------------------------------------------------------------------------

thin_grey = Side(style="thin", color=GREY_BORDER)
data_border = Border(top=thin_grey, bottom=thin_grey, left=thin_grey, right=thin_grey)
DATE_NUMFMT = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _rp(size, bold=False, color="000000", font="Calibri"):
    return CharacterProperties(
        sz=size, b=bold, solidFill=color, latin=DrawingFont(typeface=font), spc=-100
    )


def rich_text(text, size=1000, bold=False, color="000000", font="Calibri"):
    p = Paragraph(
        pPr=ParagraphProperties(defRPr=_rp(size, bold, color, font)),
        r=[RegularTextRun(rPr=_rp(size, bold, color, font), t=text)],
    )
    return RichText(p=[p])


def chart_title(text):
    return Title(tx=Text(rich=rich_text(text, size=1800, bold=True)))


def axis_text(size=1000):
    return rich_text("", size=size)


def grey_gridlines():
    return ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(w=9360, solidFill=GREY_LINE))
    )


def styled_axes(x_axis, y_axis, x_numfmt=None, y_title=None):
    for axis in (x_axis, y_axis):
        axis.majorTickMark = "none"
        axis.minorTickMark = "none"
        axis.txPr = axis_text()
        axis.spPr = GraphicalProperties(
            ln=LineProperties(w=9360, solidFill=GREY_LINE)
        )
    y_axis.majorGridlines = grey_gridlines()
    if x_numfmt:
        x_axis.number_format = x_numfmt
    if y_title:
        y_axis.title = Title(tx=Text(rich=rich_text(y_title, size=1000, bold=True)))


def line_chart(title_text, cats_ref, vals_ref):
    ch = LineChart()
    ch.style = 2
    ch.title = chart_title(title_text)
    ch.legend = None
    ch.add_data(vals_ref, titles_from_data=False)
    ch.set_categories(cats_ref)
    ser = ch.series[0]
    ser.smooth = True
    ser.graphicalProperties = GraphicalProperties(
        ln=LineProperties(w=25920, solidFill=CHART_LINE_COLOR)
    )
    ser.marker = Marker(symbol="circle", size=5)
    ser.marker.graphicalProperties = GraphicalProperties(solidFill=CHART_LINE_COLOR)

    x_axis = DateAxis(baseTimeUnit="months")
    ch.x_axis = x_axis
    ch.y_axis.crossAx = x_axis.axId
    ch.x_axis.crossAx = ch.y_axis.axId
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    styled_axes(ch.x_axis, ch.y_axis, x_numfmt="mmm-yy", y_title="Complaints")
    return ch


def bar_chart(title_text, cats_ref, vals_ref, color, show_dlbl=True, vary=False):
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "clustered"
    ch.gapWidth = 150
    ch.varyColors = vary
    ch.style = 2
    ch.title = chart_title(title_text)
    ch.legend = None
    ch.add_data(vals_ref, titles_from_data=False)
    ch.set_categories(cats_ref)
    ser = ch.series[0]
    if color:
        ser.graphicalProperties = GraphicalProperties(solidFill=color)
    if show_dlbl:
        ser.dLbls = DataLabelList()
        ser.dLbls.showVal = True
        ser.dLbls.dLblPos = "outEnd"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    styled_axes(ch.x_axis, ch.y_axis)
    return ch


def doughnut_chart(title_text, cats_ref, vals_ref):
    ch = DoughnutChart()
    ch.holeSize = 50
    ch.varyColors = True
    ch.style = 2
    ch.title = chart_title(title_text)
    ch.legend.position = "r"
    ch.add_data(vals_ref, titles_from_data=False)
    ch.set_categories(cats_ref)
    ser = ch.series[0]
    ser.data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(solidFill=c))
        for i, c in enumerate(DOUGHNUT_PALETTE)
    ]
    ser.dLbls = DataLabelList()
    ser.dLbls.showCatName = True
    ser.dLbls.showPercent = True
    return ch


def col_width_px(ws, idx):
    dim = ws.column_dimensions.get(get_column_letter(idx))
    width = dim.width if dim and dim.width else 8.43
    return max(2, int(round(width * 7)) + 5)


def row_height_px(ws, idx):
    dim = ws.row_dimensions.get(idx)
    height = dim.height if dim and dim.height else 15.0
    return int(round(height * 96 / 72))


def span_cm(ws, col_from0, col_to0, off_from_px=0, off_to_px=0):
    total = sum(col_width_px(ws, c) for c in range(col_from0 + 1, col_to0 + 1))
    total += off_to_px - off_from_px
    return max(4.0, total / 37.7953)


def span_cm_v(ws, row_from0, row_to0, off_from_px=0, off_to_px=0):
    total = sum(row_height_px(ws, r) for r in range(row_from0 + 1, row_to0 + 1))
    total += off_to_px - off_from_px
    return max(3.0, total / 37.7953)


def add_chart(ws, chart, anchor, width, height):
    chart.width = width
    chart.height = height
    ws.add_chart(chart, anchor)


# ---------------------------------------------------------------------------
# Exact chart replication — every rendered property (colors, data-label callout
# boxes with manual layouts, axis fonts/styles, marker shapes, hole sizes,
# explosion, legend placement, position/sizes of the five charts per sheet) is
# taken 1:1 from the hand-made reference workbook's chart XML. The openpyxl
# chart objects above only exist so the package gets its parts; the final
# archive swap below replaces them with the reference XML, with data ranges and
# caches adapted to the current dump.
# ---------------------------------------------------------------------------

TEMPLATE_DIR = Path(__file__).resolve().parent / "chart_templates"

# Excel 1900 date-system: serial 1 = 1900-01-01
SERIAL_EPOCH = datetime(1899, 12, 30)

SLOT_BY_TITLE = {
    "Monthly Complaint Volume (filtered)": 1,
    "By Category (filtered)": 2,
    "Top Issue Types (filtered)": 3,
    "Resolution Status (filtered)": 4,
    "By Phase / Block (filtered)": 5,
    "Monthly Complaint Volume  (escalation trend)": 6,
    "Complaints by Category": 7,
    "Top Issue Types": 8,
    "Resolution Status": 9,
    "Complaints by Phase / Block": 10,
}


def _serials(months):
    return [int((datetime(y, m, 1) - SERIAL_EPOCH).days) for y, m in months]


def _esc(text):
    from xml.sax.saxutils import escape

    return escape(str(text))


def _num_cache(format_code, values, pt_count):
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
        for i, v in enumerate(values)
        if v is not None
    )
    return (
        f'<c:numCache><c:formatCode>{_esc(format_code)}</c:formatCode>'
        f'<c:ptCount val="{pt_count}"/>{pts}</c:numCache>'
    )


def _str_cache(values, pt_count):
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{_esc(v)}</c:v></c:pt>'
        for i, v in enumerate(values)
        if v is not None
    )
    return f'<c:strCache><c:ptCount val="{pt_count}"/>{pts}</c:strCache>'


def _slot_spec(slot, stats):
    """(ordered refs, ordered cache xml builders) for a given chart slot."""
    cats = stats["category_order"]
    subs = stats["subcat_chart"]
    stas = stats["status_order"]
    phs = stats["phase_order"]
    months = stats["months"]
    month_counts = [stats["month"].get((y, m), 0) for y, m in months]
    serials = _serials(months)
    C, SUB, ST, PH, M = len(cats), len(subs), len(stas), len(phs), len(months)

    # interactive refs include the +1 blank-pad row the reference workbook uses
    i_cat_row = C + 2
    i_sub_row = SUB + 2
    i_stat_row = ST + 2
    i_phase_row = PH + 2
    cat_pairs = [(c, stats["category"][c]) for c in cats]
    sub_pairs = [(s, stats["subcat"][s]) for s in subs]
    sta_pairs = [(s, stats["status"][s]) for s in stas]
    pha_pairs = [(p, stats["phase"][p]) for p in phs]

    def labels(pairs):
        return [k for k, _ in pairs]

    def vals(pairs):
        return [v for _, v in pairs]

    pad_label = lambda pairs: labels(pairs) + [None]  # noqa: E731

    if slot == 1:
        refs = [f"Calc!$S$2:$S${1 + M}", f"Calc!$T$2:$T${1 + M}"]
        caches = [
            lambda: _num_cache("mmm\\-yy", serials, M),
            lambda: _num_cache("General", month_counts, M),
        ]
    elif slot == 2:
        refs = [f"Calc!$A$2:$A${i_cat_row}", f"Calc!$B$2:$B${i_cat_row}"]
        caches = [
            lambda: _str_cache(pad_label(cat_pairs), C + 1),
            lambda: _num_cache("General", vals(cat_pairs), C + 1),
        ]
    elif slot == 3:
        refs = [f"Calc!$D$2:$D${i_sub_row}", f"Calc!$E$2:$E${i_sub_row}"]
        caches = [
            lambda: _str_cache(pad_label(sub_pairs), SUB + 1),
            lambda: _num_cache("General", vals(sub_pairs), SUB + 1),
        ]
    elif slot == 4:
        # unpadded like the reference doughnut cache: legend lists exactly the
        # real statuses (the pad row would render an empty legend entry)
        refs = [f"Calc!$G$2:$G${1 + ST}", f"Calc!$H$2:$H${1 + ST}"]
        caches = [
            lambda: _str_cache(labels(sta_pairs), ST),
            lambda: _num_cache("General", vals(sta_pairs), ST),
        ]
    elif slot == 5:
        refs = [f"Calc!$P$2:$P${i_phase_row}", f"Calc!$Q$2:$Q${i_phase_row}"]
        caches = [
            lambda: _str_cache(pad_label(pha_pairs), PH + 1),
            lambda: _num_cache("General", vals(pha_pairs), PH + 1),
        ]
    elif slot == 6:
        refs = ["Analysis!$E$13", f"Analysis!$D$14:$D${13 + M}",
                f"Analysis!$E$14:$E${13 + M}"]
        caches = [
            lambda: _str_cache(["Count"], 1),
            lambda: _num_cache("mmm\\-yy", serials, M),
            lambda: _num_cache("General", month_counts, M),
        ]
    elif slot == 7:
        refs = ["Analysis!$B$2", f"Analysis!$A$3:$A${2 + C}",
                f"Analysis!$B$3:$B${2 + C}"]
        caches = [
            lambda: _str_cache(["Count"], 1),
            lambda: _str_cache(labels(cat_pairs), C),
            lambda: _num_cache("General", vals(cat_pairs), C),
        ]
    elif slot == 8:
        refs = ["Analysis!$E$2", f"Analysis!$D$3:$D${2 + SUB}",
                f"Analysis!$E$3:$E${2 + SUB}"]
        caches = [
            lambda: _str_cache(["Count"], 1),
            lambda: _str_cache(labels(sub_pairs), SUB),
            lambda: _num_cache("General", vals(sub_pairs), SUB),
        ]
    elif slot == 9:
        refs = ["Analysis!$H$2", f"Analysis!$G$3:$G${2 + ST}",
                f"Analysis!$H$3:$H${2 + ST}"]
        caches = [
            lambda: _str_cache(["Count"], 1),
            lambda: _str_cache(labels(sta_pairs), ST),
            lambda: _num_cache("General", vals(sta_pairs), ST),
        ]
    elif slot == 10:
        refs = ["Analysis!$B$13", f"Analysis!$A$14:$A${13 + PH}",
                f"Analysis!$B$14:$B${13 + PH}"]
        caches = [
            lambda: _str_cache(["Count"], 1),
            lambda: _str_cache(labels(pha_pairs), PH),
            lambda: _num_cache("General", vals(pha_pairs), PH),
        ]
    else:
        raise ValueError(f"unknown chart slot {slot}")
    return refs, [c() for c in caches]


_F_PAT = re.compile(r"<c:f>[^<]*</c:f>")
_CACHE_PAT = re.compile(r"<c:(numCache|strCache)>.*?</c:\1>", re.S)


def _add_val_ax_min_zero(xml):
    """Force the value-axis minimum to 0 on the trend charts. The reference
    workbook renders its y-axis starting at 0; Excel's auto-scaling can pick
    a negative minimum when the data spans a wide range."""
    m = re.search(r"<c:valAx>.*?</c:valAx>", xml, re.S)
    if not m:
        return xml
    block = m.group(0)
    s = re.search(r"<c:scaling>.*?</c:scaling>", block, re.S)
    if not s or "<c:min" in s.group(0):
        return xml
    new_block = block.replace(
        s.group(0), s.group(0)[:-len("</c:scaling>")]
        + '<c:min val="0"/></c:scaling>', 1)
    return xml[: m.start()] + new_block + xml[m.end():]


def _adapt_template(slot, refs, caches):
    xml = (TEMPLATE_DIR / f"chart{slot}.xml").read_bytes().decode("utf-8")
    fs = [f"<c:f>{r}</c:f>" for r in refs]
    xml = _replace_blocks(xml, _F_PAT, fs)
    xml = _replace_blocks(xml, _CACHE_PAT, caches)
    if slot in (1, 6):
        xml = _add_val_ax_min_zero(xml)
    return xml


def _replace_blocks(xml, pattern, replacements):
    it = iter(replacements)
    n = 0

    def sub(m):
        nonlocal n
        n += 1
        return next(it)

    out = pattern.sub(sub, xml)
    if n != len(replacements):
        raise RuntimeError(
            f"block replace mismatch: {n} vs {len(replacements)}")
    return out


def _drawing_rels_xml(targets):
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    entries = "".join(
        f'<Relationship Id="rId{i}" Type="{rel}" Target="../charts/{t}"/>'
        for i, t in enumerate(targets, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
        f'relationships">{entries}</Relationships>'
    )


def apply_reference_charts(path, stats, log=print):
    """Rewrite the saved workbook's chart parts and drawing parts so every
    rendered chart property matches the hand-made reference workbook."""
    if not TEMPLATE_DIR.is_dir():
        log("chart templates missing — skipped exact-chart swap")
        return

    zin = zipfile.ZipFile(path, "r")
    names = zin.namelist()

    part_by_slot = {}
    for name in names:
        m = re.match(r"xl/charts/chart\d+\.xml$", name)
        if m:
            blob = zin.read(name)
            first = blob[: blob.rfind(b"</c:title>")]
            tm = re.search(rb"<a:t>([^<]*)</a:t>", first)
            title = tm.group(1).decode("utf-8") if tm else ""
            slot = SLOT_BY_TITLE.get(title)
            if slot is None:
                raise RuntimeError(f"unexpected chart title: {title!r}")
            part_by_slot[slot] = name.rsplit("/", 1)[-1]

    expected = set(range(1, 11))
    got = set(part_by_slot)
    if got != expected:
        raise RuntimeError(f"chart slot mapping incomplete: {got}")

    replacements = {}
    for slot, part in part_by_slot.items():
        refs, caches = _slot_spec(slot, stats)
        replacements[f"xl/charts/{part}"] = _adapt_template(
            slot, refs, caches).encode("utf-8")

    # drawings — reference twoCellAnchors verbatim, rels point at our parts
    replacements["xl/drawings/drawing1.xml"] = (
        TEMPLATE_DIR / "drawing1.xml").read_bytes()
    replacements["xl/drawings/drawing2.xml"] = (
        TEMPLATE_DIR / "drawing2.xml").read_bytes()
    replacements["xl/drawings/_rels/drawing1.xml.rels"] = \
        _drawing_rels_xml([part_by_slot[s] for s in range(1, 6)]).encode("utf-8")
    replacements["xl/drawings/_rels/drawing2.xml.rels"] = \
        _drawing_rels_xml([part_by_slot[s] for s in range(6, 11)]).encode("utf-8")

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(
                name, replacements.get(name) or zin.read(name))
    zin.close()
    Path(tmp).replace(path)


# --- Data sheet -------------------------------------------------------------

def apply_sheet_defaults(ws):
    """Matching sheetFormatPr defaults with the hand-made reference workbook."""
    ws.sheet_format.baseColWidth = None
    ws.sheet_format.defaultColWidth = 8.7265625
    ws.sheet_format.defaultRowHeight = 14.5


def build_data_sheet(wb, rows):
    ws = wb.create_sheet("Data")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"
    apply_sheet_defaults(ws)
    for c, h in enumerate(DATA_HEADERS, start=1):
        cell = ws.cell(1, c, h)
        cell.fill = _fill(NAVY)
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = data_border
    ws.row_dimensions[1].height = 30

    for idx, r in enumerate(rows, start=1):
        excel_row = idx + 1
        band = BAND if idx % 2 == 1 else WHITE
        values = [
            idx,
            r["id"],
            r["created"] if isinstance(r["created"], str) or r["created"] is None else str(r["created"]),
            datetime(r["date"].year, r["date"].month, r["date"].day) if r["date"] else None,
            r["category"],
            r["sub"],
            r["phase"],
            r["flat"],
            r["subject"],
            r["status"],
        ]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(excel_row, c, v)
            cell.fill = _fill(band)
            cell.font = Font(name="Arial", size=9)
            cell.border = data_border
            if c == 2:
                if isinstance(v, str) and v.strip().isdigit():
                    cell.value = int(v.strip())
                cell.number_format = "0"
            elif c == 4:
                cell.number_format = DATE_NUMFMT
            elif c == 10:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            else:
                cell.alignment = Alignment(vertical="top")

    for c, w in enumerate([7.0, 8.0, 17.0, 16.71, 16.71, 20.0, 10.0, 17.0, 36.43, 13.0], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    return ws


# --- Calc sheet (hidden live engine) -----------------------------------------

def build_calc_sheet(wb, stats, frows=FORMULA_ROW_END):
    ws = wb["Calc"]
    ws.sheet_state = "hidden"
    apply_sheet_defaults(ws)

    ws["M1"] = "phaseCrit"
    ws["N1"] = '=IF(Interactive!$C$5="All","*",Interactive!$C$5)'
    ws["M2"] = "statusCrit"
    ws["N2"] = '=IF(Interactive!$F$5="All","*",Interactive!$F$5)'
    ws["M3"] = "subCategoryCrit"
    ws["N3"] = '=IF(Interactive!$I$5="All","*",Interactive!$I$5)'
    ws["M4"] = "catCrit"
    ws["N4"] = '=IF(Interactive!$L$5="All","*",Interactive!$L$5)'

    # Base filter pairs in EXACTLY the order the hand-made dashboard uses
    # (Criteria order is COUNTIFS-irrelevant, but keeping it identical makes
    # the generated file diffable against Nikhil's original.)
    filters = (
        f"Data!$E$2:$E${frows},Calc!$N$4,"
        f"Data!$G$2:$G${frows},Calc!$N$1,"
        f"Data!$J$2:$J${frows},Calc!$N$2,"
        f"Data!$F$2:$F${frows},Calc!$N$3"
    )
    # Nikhil's status KPI cards (V2/V3/V4) deliberately ignore the status
    # dropdown itself — only the literal status plus the cat/phase/sub filters.
    filters_no_status = (
        f"Data!$E$2:$E${frows},Calc!$N$4,"
        f"Data!$G$2:$G${frows},Calc!$N$1,"
        f"Data!$F$2:$F${frows},Calc!$N$3"
    )

    ws["V1"] = f"=COUNTIFS({filters})"
    ws["V2"] = f'=COUNTIFS(Data!$J$2:$J${frows},"In Progress",{filters_no_status})'
    ws["V3"] = f'=COUNTIFS(Data!$J$2:$J${frows},"Reopened",{filters_no_status})'
    ws["V4"] = f'=COUNTIFS(Data!$J$2:$J${frows},"On Hold",{filters_no_status})'
    last_sub_row = max(1 + len(stats["subcat_chart"]), 13)
    ws["V5"] = f"=MAX(E2:E{last_sub_row})"
    ws["V7"] = f"=COUNTA(Data!$B$2:$B${frows})"

    # categories
    for i, cat in enumerate(stats["category_order"], start=2):
        ws.cell(i, 1, cat)
        ws.cell(i, 2, f"=COUNTIFS(Data!$E$2:$E${frows},$A${i},{filters})")
    # top sub-categories
    for i, sub in enumerate(stats["subcat_chart"], start=2):
        ws.cell(i, 4, sub)
        ws.cell(i, 5, f"=COUNTIFS(Data!$F$2:$F${frows},$D${i},{filters})")
    # statuses
    for i, st in enumerate(stats["status_order"], start=2):
        ws.cell(i, 7, st)
        ws.cell(i, 8, f"=COUNTIFS(Data!$J$2:$J${frows},$G${i},{filters})")
    # phases
    for i, ph in enumerate(stats["phase_order"], start=2):
        ws.cell(i, 16, ph)
        ws.cell(i, 17, f"=COUNTIFS(Data!$G$2:$G${frows},$P${i},{filters})")
    # monthly trend — every month in the data range, zero-filled
    for i, (y, m) in enumerate(stats["months"], start=2):
        ws.cell(i, 19, date(y, m, 1)).number_format = "mmm-yy"
        ws.cell(
            i, 20,
            f'=COUNTIFS(Data!$D$2:$D${frows},">="&DATE({y},{m},1),'
            f'Data!$D$2:$D${frows},"<"&EDATE(DATE({y},{m},1),1),{filters})',
        )
    # sub-category filter dropdown source (named range)
    ws["X1"] = "All"
    subs = stats["subcat_filter"][1:]
    for i, sub in enumerate(subs, start=2):
        ws.cell(i, 24, sub)
    wb.defined_names.add(
        DefinedName("SubCategoryFilterList", attr_text=f"Calc!$X$1:$X${1 + len(subs)}")
    )
    ws.column_dimensions["A"].width = 12.29
    return ws


# --- Interactive sheet -------------------------------------------------------

FILTER_CELLS = [("C4", "Phase", "C5"), ("F4", "Status", "F5"),
                ("I4", "Sub Category", "I5"), ("L4", "Category", "L5")]


def build_interactive(wb, stats, as_on):
    ws = wb["Interactive"]
    ws.sheet_view.showGridLines = False
    apply_sheet_defaults(ws)
    phase_order = stats["phase_order"]
    status_order = stats["status_order"]
    subcat_filter = stats["subcat_filter"]
    category_order = stats["category_order"]

    ws.merge_cells("B2:M2")
    ws["B2"] = f"{BRAND}  —  Interactive Complaints Dashboard"
    ws["B2"].fill = _fill(NAVY)
    ws["B2"].font = Font(name="Arial", size=17, bold=True, color="FFFFFF")
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 31.5

    ws.merge_cells("B3:M3")
    ws["B3"] = (
        "Use the Phase / Status / Sub Category / Category filters below — every KPI "
        "and chart updates to match your selection.  Source: MyGate export "
        f"({ordinal(as_on.day)}-{MONTHS[as_on.month - 1]}-{as_on.year})."
    )
    ws["B3"].font = Font(name="Arial", size=9, color=SUBTLE_TEXT)

    cards = [
        ("B", "C", "FILTERED TOTAL", "=Calc!$V$1", "General",
         "complaints in view", NAVY),
        ("D", "E", "% OF ALL", "=Calc!$V$1/Calc!$V$7", "0.0%",
         '="share of "&TEXT(Calc!$V$7,"#,##0")', MID_BLUE),
        ("F", "G", "IN PROGRESS", "=Calc!$V$2", "General", "open", BLUE),
        ("H", "I", "REOPENED", "=Calc!$V$3", "General", "closed but returned", RED),
        ("K", "L", "TOP ISSUE COUNT", "=Calc!$V$5", "General",
         "=INDEX(Calc!$D$2:$D$13,MATCH(Calc!$V$5,Calc!$E$2:$E$13,0))", PURPLE),
    ]
    for col, right, label, formula, fmt, nine, color in cards:
        ws.merge_cells(f"{col}7:{right}7")
        ws.merge_cells(f"{col}8:{right}8")
        ws.merge_cells(f"{col}9:{right}9")
        for r in (7, 8, 9):
            ws[f"{col}{r}"].fill = _fill(color)
            ws[f"{right}{r}"].fill = _fill(color)
        ws[f"{col}7"] = label
        ws[f"{col}7"].font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        ws[f"{col}7"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{col}8"] = formula
        ws[f"{col}8"].number_format = fmt
        ws[f"{col}8"].font = Font(name="Arial", size=22, bold=True, color="FFFFFF")
        ws[f"{col}8"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{col}9"] = nine
        ws[f"{col}9"].font = Font(name="Arial", size=8, color="FFEFEFEF")
        ws[f"{col}9"].alignment = Alignment(horizontal="center", vertical="top")
    ws.row_dimensions[5].height = 24.0
    ws.row_dimensions[7].height = 15.0
    ws.row_dimensions[8].height = 30.0
    ws.row_dimensions[9].height = 19.5

    ws["B4"] = "FILTERS"
    ws["B4"].font = Font(name="Arial", size=10, bold=True, color=RED)
    for label_cell, label, value_cell in FILTER_CELLS:
        ws[label_cell] = label
        ws[label_cell].font = Font(name="Arial", size=9, bold=True, color=NAVY)
        ws[label_cell].alignment = Alignment(horizontal="left")
        ws[value_cell] = "All"
        ws[value_cell].fill = _fill(BLUE)
        ws[value_cell].font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        ws[value_cell].alignment = Alignment(horizontal="center", vertical="center")
        medium = Side(style="medium", color=NAVY)
        ws[value_cell].border = Border(
            top=medium, bottom=medium, left=medium, right=medium
        )

    lists = {
        "C5": '"All,' + ",".join(phase_order) + '"',
        "F5": '"All,' + ",".join(status_order) + '"',
        "I5": "SubCategoryFilterList",
        "L5": '"All,' + ",".join(category_order) + '"',
    }
    ws.merge_cells("L5:M5")  # category dropdown spans two columns, as in the reference
    for cell, formula in lists.items():
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.add(cell)
        ws.add_data_validation(dv)

    for col, w in {"A": 2.0, "B": 15.0, "C": 15.0, "D": 2.0,
                   "E": 20.453125, "F": 16.7265625, "G": 2.0,
                   "H": 15.0, "I": 22.26953125, "J": 2.0, "K": 4.26953125,
                   "L": 21.7265625, "M": 7.54296875, "N": 2.0}.items():
        ws.column_dimensions[col].width = w

    calc = wb["Calc"]
    n_months = len(stats["months"])
    n_status = len(stats["status_order"])
    n_phase = len(stats["phase_order"])
    n_subcat = max(len(stats["subcat_chart"]), 12)

    add_chart(ws, line_chart(
        "Monthly Complaint Volume (filtered)",
        Reference(calc, min_col=19, min_row=2, max_row=1 + n_months),
        Reference(calc, min_col=20, min_row=2, max_row=1 + n_months),
    ), "B11", span_cm(ws, 1, 12, 0, 52), span_cm_v(ws, 10, 22, 0, 17))

    add_chart(ws, bar_chart(
        "By Category (filtered)",
        Reference(calc, min_col=1, min_row=2, max_row=1 + len(category_order) + 1),
        Reference(calc, min_col=2, min_row=2, max_row=1 + len(category_order) + 1),
        CHART_BAR_CATEGORY,
    ), "B24", span_cm(ws, 1, 6, 0, 8), span_cm_v(ws, 23, 36, 1, 6))

    add_chart(ws, bar_chart(
        "Top Issue Types (filtered)",
        Reference(calc, min_col=4, min_row=2, max_row=1 + n_subcat),
        Reference(calc, min_col=5, min_row=2, max_row=1 + n_subcat),
        CHART_BAR_TOP,
    ), "H24", span_cm(ws, 7, 13, 1, 0), span_cm_v(ws, 23, 36, 1, 6))

    add_chart(ws, doughnut_chart(
        "Resolution Status (filtered)",
        Reference(calc, min_col=7, min_row=2, max_row=1 + n_status),
        Reference(calc, min_col=8, min_row=2, max_row=1 + n_status),
    ), "B37", span_cm(ws, 0, 6, 13, 9), span_cm_v(ws, 36, 48, 9, 9))

    add_chart(ws, bar_chart(
        "By Phase / Block (filtered)",
        Reference(calc, min_col=16, min_row=2, max_row=1 + n_phase),
        Reference(calc, min_col=17, min_row=2, max_row=1 + n_phase),
        None, vary=True,
    ), "H37", span_cm(ws, 7, 13, 1, 0), span_cm_v(ws, 36, 48, 9, 9))
    return ws


# --- Overview (static) -------------------------------------------------------

def build_overview(wb, stats, as_on, period_from, period_to):
    ws = wb["Overview (static)"]
    ws.sheet_view.showGridLines = False
    apply_sheet_defaults(ws)
    analysis = wb["Analysis"]

    ws.merge_cells("B2:N2")
    ws["B2"] = f"{BRAND}  —  Complaints Analytics Dashboard"
    ws["B2"].fill = _fill(NAVY)
    ws["B2"].font = Font(name="Arial", size=18, bold=True, color="FFFFFF")
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 33.75

    ws.merge_cells("B3:N3")
    ws["B3"] = (
        f"MyGate Help Desk Report  |  Reporting period: {fmt_short(period_from)} to "
        f"{fmt_short(period_to)}  |  Source: MyGate complaints export "
        f"({ordinal(as_on.day)} {MONTHS[as_on.month - 1]}-{as_on.year})"
    )
    ws["B3"].font = Font(name="Arial", size=10, color=SUBTLE_TEXT)
    ws.row_dimensions[3].height = 18.0

    st = stats["status"]
    # KPI values are live COUNTA/COUNTIF formulas over the Data sheet — exactly
    # like the hand-made dashboard's Overview cards (no static numbers).
    cards = [
        ("B", "C", "TOTAL COMPLAINTS",
         f"=COUNTA(Data!$B$2:$B${FORMULA_ROW_END})", "All logged tickets", NAVY),
        ("D", "E", "OPEN / IN PROGRESS",
         f'=COUNTIF(Data!$J$2:$J${FORMULA_ROW_END},"In Progress")',
         "Awaiting resolution", MID_BLUE),
        ("F", "G", "REOPENED",
         f'=COUNTIF(Data!$J$2:$J${FORMULA_ROW_END},"Reopened")',
         "Returned after closure", RED),
        ("I", "J", "ON HOLD",
         f'=COUNTIF(Data!$J$2:$J${FORMULA_ROW_END},"On Hold")',
         "Held by Projects", GOLD),
    ]
    for col, right, label, value, sub, color in cards:
        ws.merge_cells(f"{col}5:{right}5")
        ws.merge_cells(f"{col}6:{right}6")
        ws.merge_cells(f"{col}7:{right}7")
        for r in (5, 6, 7):
            ws[f"{col}{r}"].fill = _fill(color)
            ws[f"{right}{r}"].fill = _fill(color)
        ws[f"{col}5"] = label
        ws[f"{col}5"].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        ws[f"{col}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{col}6"] = value
        ws[f"{col}6"].font = Font(name="Arial", size=26, bold=True, color="FFFFFF")
        ws[f"{col}6"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{col}7"] = sub
        ws[f"{col}7"].font = Font(name="Arial", size=8, color="FFEAEAEA")
        ws[f"{col}7"].alignment = Alignment(horizontal="center", vertical="top")
    ws.row_dimensions[5].height = 15.75
    ws.row_dimensions[6].height = 33.75
    ws.row_dimensions[7].height = 21.75

    for col, w in {"A": 2.0, "B": 16.0, "C": 16.0, "D": 16.0, "E": 16.0, "F": 16.0,
                   "G": 16.0, "H": 2.0, "I": 16.0, "J": 16.0, "K": 16.0, "L": 16.0,
                   "M": 16.0, "N": 16.0, "O": 2.0, "P": 2.0, "Q": 2.0, "R": 2.0}.items():
        ws.column_dimensions[col].width = w

    # Vestigial empty merge blocks the hand-made file keeps on rows 5-7 (K:L, M:N)
    for block in ("K5:L5", "K6:L6", "K7:L7", "M5:N5", "M6:N6", "M7:N7"):
        ws.merge_cells(block)

    n_months = len(stats["months"])
    n_phase = len(stats["phase_order"])
    n_cat = len(stats["category_order"])
    n_sub = len(stats["subcat_chart"])
    n_status = len(stats["status_order"])

    add_chart(ws, line_chart(
        "Monthly Complaint Volume  (escalation trend)",
        Reference(analysis, min_col=4, min_row=14, max_row=13 + n_months),
        Reference(analysis, min_col=5, min_row=14, max_row=13 + n_months),
    ), "A8", span_cm(ws, 0, 9, 14, 111), span_cm_v(ws, 7, 21, 7, 11))

    add_chart(ws, bar_chart(
        "Complaints by Category",
        Reference(analysis, min_col=1, min_row=3, max_row=2 + n_cat),
        Reference(analysis, min_col=2, min_row=3, max_row=2 + n_cat),
        CHART_BAR_CATEGORY,
    ), "B22", span_cm(ws, 1, 5, 6, 19), span_cm_v(ws, 21, 35, 15, 19))

    add_chart(ws, bar_chart(
        "Top Issue Types",
        Reference(analysis, min_col=4, min_row=3, max_row=2 + n_sub),
        Reference(analysis, min_col=5, min_row=3, max_row=2 + n_sub),
        CHART_BAR_TOP,
    ), "F22", span_cm(ws, 5, 10, 25, 1), span_cm_v(ws, 21, 35, 15, 19))

    add_chart(ws, doughnut_chart(
        "Resolution Status",
        Reference(analysis, min_col=7, min_row=3, max_row=2 + n_status),
        Reference(analysis, min_col=8, min_row=3, max_row=2 + n_status),
    ), "B37", span_cm(ws, 1, 5, 7, 18), span_cm_v(ws, 36, 50, 0, 4))

    add_chart(ws, bar_chart(
        "Complaints by Phase / Block",
        Reference(analysis, min_col=1, min_row=14, max_row=13 + n_phase),
        Reference(analysis, min_col=2, min_row=14, max_row=13 + n_phase),
        None, vary=True,
    ), "F37", span_cm(ws, 5, 10, 26, 0), span_cm_v(ws, 36, 50, 0, 4))
    return ws


# --- Analysis (COUNTIF feed for the Overview charts) -------------------------

def build_analysis(wb, stats, frows=FORMULA_ROW_END):
    ws = wb["Analysis"]
    ws.sheet_view.showGridLines = False
    apply_sheet_defaults(ws)
    rng = {
        "cat": f"Data!$E$2:$E${frows}",
        "sub": f"Data!$F$2:$F${frows}",
        "status": f"Data!$J$2:$J${frows}",
        "phase": f"Data!$G$2:$G${frows}",
        "date": f"Data!$D$2:$D${frows}",
    }

    titles = {
        "A1": "Complaints by Category",
        "D1": "Top Issue Types (Sub-Category)",
        "G1": "By Resolution Status",
        "J1": "Defect Liability Period (DLP) Status",
        "A12": "Complaints by Phase / Block",
        "D12": "Monthly Complaint Trend",
    }
    for ref, txt in titles.items():
        ws[ref] = txt
        ws[ref].font = Font(name="Arial", size=12, bold=True, color=NAVY)

    for ref, txt in [
        ("A2", "Category"), ("B2", "Count"),
        ("D2", "Sub-Category"), ("E2", "Count"),
        ("G2", "Status"), ("H2", "Count"),
        ("J2", "DLP Status"), ("K2", "Count"),
        ("A13", "Phase"), ("B13", "Count"),
        ("D13", "Month"), ("E13", "Count"),
    ]:
        ws[ref] = txt
        ws[ref].fill = _fill(MID_BLUE)
        ws[ref].font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        ws[ref].alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        ws[ref].border = data_border

    body_font = Font(name="Arial", size=10)
    center = Alignment(horizontal="center")

    def section(count_col, label_col, values, formula_tpl, start_row):
        for i, v in enumerate(values, start=start_row):
            cell = ws.cell(i, label_col, v)
            cell.font = body_font
            ccell = ws.cell(i, count_col, formula_tpl.format(i=i))
            ccell.font = body_font
            ccell.alignment = center

    section(2, 1, stats["category_order"],
            "=COUNTIF({cat},A{{i}})".format(cat=rng["cat"]), 3)
    section(5, 4, stats["subcat_chart"],
            "=COUNTIF({sub},D{{i}})".format(sub=rng["sub"]), 3)
    section(8, 7, stats["status_order"],
            "=COUNTIF({status},G{{i}})".format(status=rng["status"]), 3)
    # DLP cannot be derived from the dump (MyGate provides no purchase dates).
    # Replicated verbatim from the hand-made workbook, broken #REF! formulas included.
    for i, label in ((3, "Out Of DLP"), (4, "Under DLP")):
        ws.cell(i, 10, label).font = body_font
        ccell = ws.cell(i, 11, f"=COUNTIF(Data!#REF!,J{i})")
        ccell.font = body_font
        ccell.alignment = center
    section(2, 1, stats["phase_order"],
            "=COUNTIF({phase},A{{i}})".format(phase=rng["phase"]), 14)
    for i, (y, m) in enumerate(stats["months"], start=14):
        mcell = ws.cell(i, 4, date(y, m, 1))
        mcell.number_format = "mmm-yy"
        mcell.font = body_font
        ccell = ws.cell(
            i, 5,
            f'=COUNTIFS({rng["date"]},">="&DATE({y},{m},1),'
            f'{rng["date"]},"<"&EDATE(DATE({y},{m},1),1))',
        )
        ccell.font = body_font
        ccell.alignment = center

    for col, w in {"A": 26.0, "B": 10.0, "D": 26.0, "E": 10.0, "G": 26.0, "H": 10.0,
                   "J": 26.0, "K": 10.0}.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 15.5
    ws.row_dimensions[12].height = 15.5
    return ws


# ---------------------------------------------------------------------------

def build_workbook(rows, stats, as_on, period_from, period_to):
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    as_on_label = f"{as_on.day:02d}.{as_on.month:02d}.{str(as_on.year)[2:]}"
    wb.properties.creator = "MyGate Dashboard Tool"
    wb.properties.lastModifiedBy = "MyGate Dashboard Tool"
    wb.properties.title = f"Mygate Complaints Dashboard as on {as_on_label}"

    wb.create_sheet("Interactive")
    wb.create_sheet("Overview (static)")
    build_data_sheet(wb, rows)
    wb.create_sheet("Analysis")
    wb.create_sheet("Calc")

    build_analysis(wb, stats)
    build_calc_sheet(wb, stats)
    build_interactive(wb, stats, as_on)
    build_overview(wb, stats, as_on, period_from, period_to)

    wb._sheets = [wb[name] for name in
                  ["Interactive", "Overview (static)", "Data", "Analysis", "Calc"]]
    wb.active = 0
    return wb


def derive_as_on(dump_path: Path, rows):
    m = re.search(r"_to_(\d{4})-(\d{1,2})-(\d{1,2})", dump_path.name)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    dated = [r["date"] for r in rows if r["date"]]
    if dated:
        return max(dated)
    return date.today()


def parse_as_on_arg(value: str):
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", value)
    if m:
        yy = int(m.group(3))
        if yy < 100:
            yy += 2000 if yy < 70 else 0
        return date(yy, int(m.group(2)), int(m.group(1)))
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", value)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    raise argparse.ArgumentTypeError(f"Unsupported --as-on value: {value}")


def output_name(as_on: date) -> str:
    return (f"Mygate Complaints Dashboard as on "
            f"{as_on.day:02d}.{as_on.month:02d}.{str(as_on.year)[2:]}.xlsx")


def main(argv=None):
    ap = argparse.ArgumentParser(
        description=(
            "Build a 'Mygate Complaints Dashboard' workbook from a "
            "MyGate Master Helpdesk Dump."
        )
    )
    ap.add_argument("dump", type=Path,
                    help="path to MyGate_Master_Helpdesk_Dump_*.xlsx")
    ap.add_argument("--as-on", dest="as_on", type=parse_as_on_arg, default=None,
                    help="as-on label, e.g. 22.09.26 (default: derived automatically)")
    ap.add_argument("--out", type=Path, default=None, help="output .xlsx path")
    ap.add_argument("--quiet", action="store_true", help="suppress the summary")
    args = ap.parse_args(argv)

    dump_path = args.dump
    if not dump_path.exists():
        raise SystemExit(f"ERROR: dump file not found: {dump_path}")

    rows, meta = read_dump_rows(dump_path)
    if not rows:
        raise SystemExit("ERROR: no ticket rows found in the dump file.")

    as_on = args.as_on or derive_as_on(dump_path, rows)
    period_from = meta.get("from") or DEFAULT_PERIOD_FROM
    period_to = as_on

    stats = aggregate(rows)
    out_path = args.out or dump_path.parent / output_name(as_on)
    wb = build_workbook(rows, stats, as_on, period_from, period_to)
    wb.save(out_path)

    apply_reference_charts(
        str(out_path), stats,
        log=(lambda *_: None) if args.quiet else print,
    )

    if not args.quiet:
        dated = [r["date"] for r in rows if r["date"]]
        st = stats["status"]
        print(f"OK  {len(rows)} complaints parsed  ({fmt_short(min(dated))} -> {fmt_short(as_on)})")
        print(f"    In Progress {st.get('In Progress', 0)} | "
              f"Reopened {st.get('Reopened', 0)} | On Hold {st.get('On Hold', 0)} | "
              f"New {st.get('New', 0)}")
        if stats["subcat_chart"]:
            top = stats["subcat_chart"][0]
            print(f"    Top issue: '{top}' x {stats['subcat'][top]}")
        print(f"    Dashboard saved to: {out_path}")
    return out_path


if __name__ == "__main__":
    main()
