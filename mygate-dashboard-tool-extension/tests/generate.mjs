/**
 * Node CLI harness for the workbook engine:
 *   node tests/generate.mjs <dump.xlsx> <out.xlsx> [--as-on DD.MM.YY]
 *
 * Mirrors the Python tool's CLI (`make_dashboard.py dump.xlsx out.xlsx
 * [--as-on ...]`) so parity runs can drive both identically.
 */
import { createRequire } from "node:module";
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const engine = require(path.join(here, "..", "engine.js"));

const args = process.argv.slice(2);
if (args.length < 2) {
  console.error("usage: node tests/generate.mjs <dump.xlsx> <out.xlsx> [--as-on DD.MM.YY]");
  process.exit(2);
}
const [dumpPath, outPath] = args;
let asOn = null;
const aIdx = process.argv.indexOf("--as-on");
if (aIdx !== -1 && process.argv[aIdx + 1]) {
  const m = /^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/.exec(process.argv[aIdx + 1]);
  if (!m) {
    console.error("--as-on must look like 01.10.26");
    process.exit(2);
  }
  let y = +m[3];
  if (y < 100) y += 2000;
  asOn = { y, m: +m[2], d: +m[1] };
}

const dumpBytes = readFileSync(dumpPath);
const result = await engine.generate(dumpBytes, {
  asOn,
  dumpFileName: path.basename(dumpPath),
});
writeFileSync(outPath, Buffer.from(result.bytes));
console.log(
  `${path.basename(dumpPath)} -> ${result.info.filename} ` +
  `(${result.info.total} rows, ${result.info.months} months); written ${outPath}`);
