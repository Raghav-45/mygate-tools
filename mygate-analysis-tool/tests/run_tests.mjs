/**
 * Regression suite for the standalone JS engine. Transformation rules and
 * workbook invariants mirror the reference generator's behavior (the
 * original python suite was retired when this became a JS-only tool).
 * Run from anywhere:
 *
 *     node tests/run_tests.mjs
 *
 * Layers:
 *   1. transformation-rule unit tests (collapse/phase/date parsing/ordinal/
 *      month-span/filename conventions/as-on derivation)
 *   2. mini end-to-end: a mini dump is built in JS the same way the suite's
 *      write_mini_dump lays it out (title / report line / headers / rows),
 *      then readDumpRows + generate assert the same invariants.
 *
 * Byte-identity with the frozen reference build was verified on a mini and
 * the full 813-row master dump during development; these structural checks
 * keep guarding it.
 */
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const eng = require(path.join(here, "..", "engine.js"));

/* -------------------------------------------------------------------------- */
/* mini-dump writer (js twin of write_mini_dump in the Python suite)           */
/* -------------------------------------------------------------------------- */

const te = new TextEncoder();
const esc = (s) =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function sheetCell(ref, value) {
  if (typeof value === "number") {
    return `<c r="${ref}" t="n"><v>${value}</v></c>`;
  }
  const txt = value === null || value === undefined ? "" : String(value);
  const preserve = txt !== txt.trim() ? ' xml:space="preserve"' : "";
  return `<c r="${ref}" t="inlineStr"><is><t${preserve}>${esc(txt)}</t></is></c>`;
}

function rowXml(number, cells) {
  return `<row r="${number}">` +
    cells.map((c) => sheetCell(c.col + number, c.v)).join("") + "</row>";
}

/**
 * Build a dump xlsx in-memory the way mygate-dump-tool writes it:
 * row1 title, row2 subtitle, row3 headers, rows4+ data. Python's suite writes
 * 8-tuples [sr, id, created, cat, sub, flat, subj, status] straight into
 * columns A..H under those same headers.
 */
async function buildMiniDump(tuples, from = "1-1-2024", to = "22-9-2026") {
  const body = [];
  body.push(`<row r="1">${sheetCell("A1", "DLF Independent Floors")}</row>`);
  body.push(`<row r="2">${sheetCell("A2", "Help Desk Report: From " + from + " To " + to)}</row>`);
  body.push(rowXml(3, [
    { col: "A", v: "Sr No." }, { col: "B", v: "I.D" },
    { col: "C", v: "Created Date" }, { col: "D", v: "Category" },
    { col: "E", v: "Sub Category" }, { col: "F", v: "Flat" },
    { col: "G", v: "Subject" }, { col: "H", v: "Status" },
  ]));
  tuples.forEach((r, i) => {
    body.push(rowXml(4 + i, [
      { col: "A", v: r[0] }, { col: "B", v: r[1] }, { col: "C", v: r[2] },
      { col: "D", v: r[3] }, { col: "E", v: r[4] },
      { col: "F", v: r[5] }, { col: "G", v: r[6] }, { col: "H", v: r[7] },
    ]));
  });

  const sheet1 =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>` +
    `<dimension ref="A1:H8" /><sheetViews><sheetView workbookViewId="0" /></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="15" /><sheetData>${body.join("")}</sheetData>` +
    `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />` +
    `</worksheet>`;

  const parts = [
    ["[Content_Types].xml", te.encode(_ctXml())],
    ["_rels/.rels", te.encode(_rootRels())],
    ["xl/workbook.xml", te.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
      ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheets><sheet name="Master Dump" sheetId="1" state="visible" r:id="rId1"/></sheets>` +
      `</workbook>`)],
    ["xl/_rels/workbook.xml.rels", te.encode(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
      `</Relationships>`)],
    ["xl/worksheets/sheet1.xml", te.encode(sheet1)],
    ["xl/styles.xml", te.encode(_miniStyles())],
  ];
  return eng.buildZip(parts);

  function _ctXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
      `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
      `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
      `</Types>`;
  }
  function _rootRels() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
      `</Relationships>`;
  }
  function _miniStyles() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></_fonts>`.replace("_fonts", "fonts") +
      `<fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
      `<borders count="1"><border><left /><right /><top /><bottom /><diagonal /></border></borders>` +
      `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
      `<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>` +
      `</styleSheet>`;
  }
}

/* -------------------------------------------------------------------------- */
/* tiny test runner                                                            */
/* -------------------------------------------------------------------------- */

let passed = 0;
const failures = [];
function ok(cond, name, detail) {
  if (cond) {
    passed++;
  } else {
    failures.push(name + (detail !== undefined ? " :: " + detail : ""));
    console.error("FAIL", name, detail !== undefined ? detail : "");
  }
}
function eq(a, b, name) {
  const ja = JSON.stringify(a);
  const jb = JSON.stringify(b);
  ok(ja === jb, name, ja === jb ? "" : `got ${ja} want ${jb}`);
}
const D = (y, m, d) => ({ y, m, d });

async function main() {
  /* ---------------------------------------------------------------------- */
  /* 1. transformation rules (mirrors the Python unit tests)                 */
  /* ---------------------------------------------------------------------- */

  // test_collapse_spaces
  eq(eng.collapseSpaces("Products  Appliances "), "Products Appliances", "collapse: doubles");
  eq(eng.collapseSpaces("IT   WIFI    Network"), "IT WIFI Network", "collapse: triples");
  eq(eng.collapseSpaces(null), "", "collapse: null");
  eq(eng.collapseSpaces("  Estate  Infra   Outer "), "Estate Infra Outer", "collapse: outer");

  // test_derive_phase
  eq(eng.derivePhase("Phase-A-T-10/2-D"), "Phase A", "phase: A");
  eq(eng.derivePhase("Phase-P-M-11/5-B"), "Phase P", "phase: P");
  eq(eng.derivePhase("Flat 42"), "Unknown", "phase: flat");
  eq(eng.derivePhase(null), "Unknown", "phase: null");

  // test_parse_created_date
  eq(eng.parseCreatedDate("22-05-24, 09:36 AM"), D(2024, 5, 22), "date: am");
  eq(eng.parseCreatedDate("10-08-24, 09:09 PM"), D(2024, 8, 10), "date: pm");
  eq(eng.parseCreatedDate("01-01-25, 12:05 AM"), D(2025, 1, 1), "date: midnight");
  eq(eng.parseCreatedDate({ y: 2024, m: 3, d: 3 }), D(2024, 3, 3), "date: struct passthrough");
  eq(eng.parseCreatedDate("garbage"), null, "date: garbage");
  eq(eng.parseCreatedDate(null), null, "date: null");

  // test_ordinal
  eq(eng.ordinal(1), "1st", "ordinal 1");
  eq(eng.ordinal(2), "2nd", "ordinal 2");
  eq(eng.ordinal(3), "3rd", "ordinal 3");
  eq(eng.ordinal(22), "22nd", "ordinal 22");
  eq(eng.ordinal(11), "11th", "ordinal 11");
  eq(eng.ordinal(30), "30th", "ordinal 30");

  // test_month_span_zero_fills
  eq(eng.monthSpan(D(2024, 5, 1), D(2024, 9, 1)), [
    { y: 2024, m: 5 }, { y: 2024, m: 6 }, { y: 2024, m: 7 },
    { y: 2024, m: 8 }, { y: 2024, m: 9 },
  ], "month span zero fill");

  // test_output_filename_convention
  eq(eng.outputName(D(2026, 9, 22)),
    "Mygate Complaints Dashboard as on 22.09.26.xlsx", "filename 22.09");
  eq(eng.outputName(D(2026, 10, 1)),
    "Mygate Complaints Dashboard as on 01.10.26.xlsx", "filename 01.10");

  /* ---------------------------------------------------------------------- */
  /* 2. mini end-to-end (mirrors test_mini_rows_sorted_and_normalized /      */
  /*    test_mini_dashboards_structure, minus diagram-level internals that   */
  /*    the byte-parity pytest covers against make_dashboard.py)             */
  /* ---------------------------------------------------------------------- */

  const miniRows = [
    [1, 100, "10-08-24, 09:09 PM", "Products  Appliances ", "Air Conditoners",
      "Phase-A-T-10/2-D", "Fan noise ", "In Progress"],
    [2, 101, "01-01-25, 12:05 AM", "FM Common Area Related Issue", "Carpentry",
      "Phase-B-SR/38-C", "Lift lobby", "Reopened"],
    [3, 102, "22-05-24, 09:36 AM", "Construction Or Project Related", "Cracks",
      "Phase-D-W-2/7-C", "Cracks on walls", "New "],
    [4, 103, "12-09-25, 11:00 AM", "Accounts Billing", "None",
      "Flat 42 - odd one", "No subcat", "On Hold"],
  ];

  // test_as_on_from_filename (same naming rule)
  {
    const dump = await buildMiniDump(miniRows);
    const { rows, meta } = await eng.readDumpRows(dump);
    eq(rows.length, 4, "mini dump: 4 rows parsed");
    eq(eng.deriveAsOn(meta, rows, "MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-09-07.xlsx"),
      D(2026, 9, 7), "as-on derived from filename");

    // sorted + normalized (ids asc by date: 102, 100, 101, 103)
    eq(rows.map((r) => r.id), [102, 100, 101, 103], "rows sorted by (date, id)");
    eq(rows.map((r) => r.status), ["New", "In Progress", "Reopened", "On Hold"],
      "status trimmed");
    eq(rows[1].category, "Products Appliances", "category collapsed");
    eq(rows[0].phase, "Phase D", "phase derived D");
    eq(rows[3].phase, "Unknown", "phase derived Unknown");
    eq(rows[3].flat, "Flat 42 - odd one", "flat preserved");
    eq(rows[1].created, "10-08-24, 09:09 PM", "created kept");

    // generate the workbook
    const generated = await eng.generate(dump, {
      asOn: D(2026, 9, 22),
      dumpFileName: "dump.xlsx",
    });
    ok(generated.bytes.length > 20000, "workbook bytes written", generated.bytes.length);
    eq(generated.info.filename,
      "Mygate Complaints Dashboard as on 22.09.26.xlsx", "generated filename");
    eq(generated.info.months, 17, "trend zero-fills 17 months (May-24..Sep-25)");

    // structural checks against the actual parts
    const zip = await eng.zipRead(generated.bytes);
    const body = new TextDecoder().decode(await zip.read("xl/workbook.xml"));
    const wbTextRe = /<sheet name="([^"]+)" sheetId="\d+" state="([^"]+)" r:id="rId\d+"/g;
    const names = [];
    let m;
    while ((m = wbTextRe.exec(body))) names.push([m[1], m[2]]);
    eq(names, [
      ["Interactive", "visible"], ["Overview (static)", "visible"], ["Data", "visible"],
      ["Analysis", "visible"], ["Calc", "hidden"],
    ], "sheet order + Calc hidden");

    const sheet1 = new TextDecoder().decode(await zip.read("xl/worksheets/sheet1.xml"));
    ok((sheet1.match(/<dataValidation /g) || []).length === 4,
      "interactive has 4 dropdown validations");
    ok(sheet1.includes('<formula1>"All,Phase D,Phase A,Phase B,Unknown"</formula1>'),
      "phase dv at C5 lists phases with All (count order)");
    ok(sheet1.includes('<formula1>"All,New,In Progress,Reopened,On Hold"</formula1>'),
      "status dv at F5 lists the real statuses");
    ok(sheet1.includes("<formula1>SubCategoryFilterList</formula1>"),
      "sub category dv at I5 references defined name");
    ok(sheet1.includes('<c r="B8" s="16"><f>Calc!$V$1</f>'),
      "Total card formula (B8)");
    ok(sheet1.includes('<c r="E8" s="17"><f>Calc!$V$1/Calc!$V$7</f>'),
      "percent card formula (E8)");
    ok(sheet1.includes("<f>Calc!$V$2</f>") && sheet1.includes("<f>Calc!$V$3</f>"),
      "in-progress / reopened engines");
    ok(sheet1.includes("<f>INDEX(Calc!$D$2:$D$13,MATCH(Calc!$V$5,Calc!$E$2:$E$13,0))</f>"),
      "top issue INDEX card");
    ok(body.includes('<definedName name="SubCategoryFilterList">Calc!$X$1:$X$5</definedName>'),
      "named range defined (4 subcats + header)");
    // freeze panes + charts
    const sheet3 = new TextDecoder().decode(await zip.read("xl/worksheets/sheet3.xml"));
    ok(sheet3.includes('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" />'),
      "Data frozen at A2");
    for (let c = 1; c <= 10; c++) {
      const bytes = await zip.read(`xl/charts/chart${c}.xml`);
      ok(bytes && bytes.length > 200, "chart part exists: chart" + c);
    }
    // Calc KPI + trend formulas (mirrors test_mini_dashboards_structure)
    const sheet5 = new TextDecoder().decode(await zip.read("xl/worksheets/sheet5.xml"));
    ok(sheet5.includes('<f>IF(Interactive!$C$5="All","*",Interactive!$C$5)</f>'),
      "N1 criteria formula");
    ok(/<c r="V1"><f>COUNTIFS/.test(sheet5), "V1 is a COUNTIFS engine");
    ok(sheet5.includes('<f>COUNTIFS(Data!$J$2:$J$9752,"In Progress",'), "V2 counts In Progress");
    ok(sheet5.includes("DATE(2024,5,1)"), "first trend month formula is DATE(2024,5,1)");
    eq((sheet5.match(/<c r="S\d+" s="55"/g) || []).length, 17,
      "Calc trend has 17 month rows");
  }

  /* ---------------------------------------------------------------------- */
  /* 3. zip container self-check (round-trip through zipRead/buildZip)        */
  /* ---------------------------------------------------------------------- */
  {
    const parts = [
      ["a.txt", te.encode("hello zip")],
      ["dir/b.xml", te.encode("<x/>&".repeat(2000))],
    ];
    const bytes = await eng.buildZip(parts);
    const z = await eng.zipRead(bytes);
    const a = await z.read("a.txt");
    eq(new TextDecoder().decode(a), "hello zip", "zip roundtrip a.txt");
    const b = await z.read("dir/b.xml");
    ok(b.length === "<x/>&".repeat(2000).length, "zip roundtrip b.xml size");
  }

  /* ---------------------------------------------------------------------- */
  /* report                                                                  */
  /* ---------------------------------------------------------------------- */
  const total = passed + failures.length;
  console.log(`\n${passed}/${total} passed`);
  if (failures.length) {
    console.error(failures.map((f) => "  ✗ " + f).join("\n"));
    process.exit(1);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(2);
});
