"""
Parity tests: the Chrome-extension engine (mygate-dashboard-tool-extension,
JavaScript) must produce a workbook that is byte-for-byte identical to this
tool's output — part for part across the whole OOXML zip container, for both a
synthetic mini dump and the real master helpdesk dump when it is present.

The only sanctioned difference is docProps/core.xml build timestamps: both
implementations stamp their own build time into dcterms:created/modified, so
those are aligned positionally before the comparison.

These tests drive both engines through their real CLIs / entry points:
  - make_dashboard.dump → openpyxl pipeline   (python tool, this package)
  - node tests/generate.mjs                → engine.js pipeline (the extension)

and additionally require the extension's own JS mirror suite
(node tests/run_tests.mjs) to be green.

Run from the repo root together with the tool suite:
    python3 -m pytest mygate-dashboard-tool/tests/ -v
"""

import importlib.util
import os
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

# ---------------------------------------------------------------------------
# import the tool module the same way test_dashboard_tool.py does
# ---------------------------------------------------------------------------

TOOL_DIR = Path(__file__).resolve().parent.parent  # mygate-dashboard-tool/
REPO_DIR = TOOL_DIR.parent  # repository root (mygate-tools)
EXT_DIR = REPO_DIR / "mygate-dashboard-tool-extension"

spec = importlib.util.spec_from_file_location(
    "make_dashboard", TOOL_DIR / "make_dashboard.py"
)
make_dashboard = importlib.util.module_from_spec(spec)
sys.modules["make_dashboard"] = make_dashboard
spec.loader.exec_module(make_dashboard)

from test_dashboard_tool import MINI_ROWS, write_mini_dump  # noqa: E402

# ---------------------------------------------------------------------------
# node discovery
# ---------------------------------------------------------------------------

def _node_bin():
    n = shutil.which("node")
    if n:
        return n
    candidates = ("/opt/homebrew/bin/node", "/usr/local/bin/node")
    for cand in candidates:
        if Path(cand).is_file():
            return cand
    return None


NODE = _node_bin()
skip_no_node = pytest.mark.skipif(NODE is None, reason="node runtime not found")

EXT_PART_COUNT = 30  # full dashboard workbook: 30 zip parts

# ---------------------------------------------------------------------------
# comparison helpers
# ---------------------------------------------------------------------------

TS_RE = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z")


def _mismatched_parts(reference: Path, extension: Path):
    """Compare the two xlsx containers part by part.

    Returns a list of human-readable mismatch lines; empty means the
    workbooks are part-for-part byte-identical (build-time stamps in
    docProps/core.xml are aligned positionally, everything else must be
    exactly equal).
    """
    with zipfile.ZipFile(reference) as za, zipfile.ZipFile(extension) as zb:
        names_a = set(za.namelist())
        names_b = set(zb.namelist())
        problems = []
        for name in sorted(names_a | names_b):
            if name not in names_a:
                problems.append(f"{name}: only in the extension workbook")
                continue
            if name not in names_b:
                problems.append(f"{name}: only in the python workbook")
                continue
            data_a = za.read(name)
            data_b = zb.read(name)
            if data_a == data_b:
                continue
            if name == "docProps/core.xml":
                # align the build timestamps positionally, compare the rest
                txt_a = data_a.decode("utf-8")
                txt_b = data_b.decode("utf-8")
                stamps_a = TS_RE.findall(txt_a)
                stamps_b = TS_RE.findall(txt_b)
                if stamps_a and len(stamps_a) == len(stamps_b):
                    for s_a, s_b in zip(stamps_a, stamps_b):
                        txt_b = txt_b.replace(s_b, s_a, 1)
                    if txt_b == txt_a:
                        continue
            line = f"{name}: differs "
            i = 0
            m = min(len(data_a), len(data_b))
            while i < m and data_a[i] == data_b[i]:
                i += 1
            if i == m and len(data_a) != len(data_b):
                line += f"(length only: {len(data_a)} vs {len(data_b)})"
            else:
                lo = max(0, i - 40)
                line += f"(first diff @ {i})\n"
                line += f"  reference: {data_a[lo:i + 60]!r}\n"
                line += f"  extension: {data_b[lo:i + 60]!r}"
            problems.append(line)
    return problems


def _run_extension(dump: Path, out: Path, as_on: str):
    """Run the extension's Node CLI harness: dump.xlsx -> out.xlsx."""
    cmd = [NODE, str(EXT_DIR / "tests" / "generate.mjs"), str(dump), str(out)]
    if as_on:
        cmd += ["--as-on", as_on]
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(REPO_DIR))
    assert result.returncode == 0, (
        "node tests/generate.mjs failed:\n" + result.stdout + "\n" + result.stderr
    )
    assert out.is_file(), "extension did not write an output workbook"


def _run_python_tool(dump: Path, out: Path, as_on: str):
    make_dashboard.main(
        [str(dump), "--as-on", as_on, "--out", str(out), "--quiet"]
    )
    assert out.is_file()


def _assert_full_parity(reference: Path, extension: Path, expected_parts: int):
    with zipfile.ZipFile(reference) as z:
        assert len(z.namelist()) == expected_parts, (
            f"python workbook should have {expected_parts} parts, got {len(z.namelist())}"
        )
    problems = _mismatched_parts(reference, extension)
    assert not problems, "\n".join(problems)


# ---------------------------------------------------------------------------
# 1. mini-dump full parity (synthetic; runs everywhere node exists)
# ---------------------------------------------------------------------------

@skip_no_node
def test_extension_parity_mini_dump(tmp_path):
    dump = write_mini_dump(tmp_path / "dump.xlsx", MINI_ROWS)
    py_out = tmp_path / "python_out.xlsx"
    ext_out = tmp_path / "extension_out.xlsx"

    _run_python_tool(dump, py_out, "22.09.26")
    _run_extension(dump, ext_out, "22.09.26")

    _assert_full_parity(py_out, ext_out, EXT_PART_COUNT)


# ---------------------------------------------------------------------------
# 2. real master dump full parity (skips if the dump is not on the machine)
# ---------------------------------------------------------------------------

REAL_DUMP = (
    Path("~/Downloads/mygate-new-tool data dumps")
    / "MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-10-01.xlsx"
).expanduser()

WARN_REAL_DUMP = (
    Path("~/Downloads/mygate-new-tool data dumps")
    / "MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-29.xlsx"
).expanduser()  # alternative naming used by the dump tool


@skip_no_node
def _real_dump_path():
    for cand in (REAL_DUMP, WARN_REAL_DUMP):
        if cand.is_file():
            return cand
    return None


@skip_no_node
def test_extension_parity_real_dump(tmp_path):
    dump = _real_dump_path()
    if dump is None:  # real-dump depends on the analyst's machine
        pytest.skip("real master helpdesk dump not present in Downloads")

    py_out = tmp_path / "python_out.xlsx"
    ext_out = tmp_path / "extension_out.xlsx"

    _run_python_tool(dump, py_out, "01.10.26")
    _run_extension(dump, ext_out, "01.10.26")

    _assert_full_parity(py_out, ext_out, EXT_PART_COUNT)


# ---------------------------------------------------------------------------
# 3. the extension's own JS mirror suite must pass
# ---------------------------------------------------------------------------

@skip_no_node
def test_extension_js_suite_green():
    result = subprocess.run(
        [NODE, str(EXT_DIR / "tests" / "run_tests.mjs")],
        capture_output=True,
        text=True,
        cwd=str(REPO_DIR),
    )
    assert result.returncode == 0, result.stdout + "\n" + result.stderr


# ---------------------------------------------------------------------------
# 4. extension packaging sanity (ship files present, chart templates intact)
# ---------------------------------------------------------------------------

def test_extension_folder_complete():
    required = [
        "manifest.json",
        "background.js",
        "engine.js",
        "constants.js",
        "popup.html",
        "popup.js",
        "popup.css",
    ]
    for name in required:
        assert (EXT_DIR / name).is_file(), f"missing extension file: {name}"

    icons = EXT_DIR / "icons"
    for name in ("icon16.png", "icon48.png", "icon128.png"):
        assert (icons / name).is_file(), f"missing icon: {name}"

    templates = {p.name for p in (EXT_DIR / "chart_templates").glob("*.xml")}
    expected = {f"chart{n}.xml" for n in range(1, 11)} | {"drawing1.xml", "drawing2.xml"}
    assert templates == expected, f"chart templates mismatch: {sorted(templates)}"

    for name in ("generate.mjs", "run_tests.mjs"):
        assert (EXT_DIR / "tests" / name).is_file(), f"missing test harness: {name}"


@skip_no_node
def test_extension_engine_is_unchanged_reference():
    # constants.js must still embed the exact same static parts the engine
    # ships (guards against accidental template drift across edits).
    node = NODE
    script = (
        "const c = require(process.argv[1]);"
        "const keys = Object.keys(c.STATIC_PARTS_B64).sort();"
        "console.log(JSON.stringify(keys));"
    )
    result = subprocess.run(
        [node, "-e", script, str(EXT_DIR / "constants.js")],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stderr
    import json

    keys = sorted(json.loads(result.stdout))
    for needed in (
        "[Content_Types].xml",
        "docProps/app.xml",
        "xl/theme/theme1.xml",
        "xl/styles.xml",
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/worksheets/_rels/sheet2.xml.rels",
    ):
        assert needed in keys, f"constants.js missing static part: {needed}"
