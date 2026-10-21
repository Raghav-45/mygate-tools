/**
 * MyGate Analysis Tool — workbook generation engine.
 *
 * A faithful JavaScript port of the frozen reference generator (openpyxl
 * 3.1.5 semantics). Produces a workbook that is part-for-part byte-identical
 * to that reference build (except
 * docProps/core.xml build timestamps and the zip container timestamps,
 * which reflect each build's wall-clock time — the Python tool does the
 * same, so behavior matches too).
 *
 * Runs inside the MV3 service worker (importScripts) and under Node
 * (tests + parity harness). No DOM APIs; async only for zip streams.
 */
(function (root) {
"use strict";

/* ---------------------------------------------------------------------------
 * 0. Ground-truth constants (generated file constants.js)
 * ------------------------------------------------------------------------- */

let CN;
if (typeof DASH_CONSTANTS !== "undefined") {
  CN = DASH_CONSTANTS;
} else if (typeof globalThis !== "undefined" && globalThis.DASH_CONSTANTS) {
  CN = globalThis.DASH_CONSTANTS;
} else if (typeof require === "function") {
  CN = require("./constants.js");
}
const STATIC_B64 = CN.STATIC_PARTS_B64;
const CT = CN.CT;

const BRAND = "DLF Independent Floors";
const SUBTLE_TEXT = "FF44546A";

const CHART_BAR_CATEGORY = "4472C4";
const CHART_BAR_TOP = "1F3864";

const DEFAULT_PERIOD_FROM = { y: 2024, m: 1, d: 1 };
const FORMULA_ROW_END = 9752;
const SUBCAT_CHART_ROWS = 12;

const DATA_HEADERS = [
  "Sr No.", "ID", "Created Date", "Date", "Category",
  "Sub Category", "Phase", "Flat", "Subject", "Status",
];

const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

const STATUS_ORDER_FALLBACK = ["In Progress", "Reopened", "New", "On Hold"];
const STATUS_PIE_ROWS = ["In Progress", "Reopened", "On Hold", "Repeated"];

const SLOT_BY_PART = { /* fixed verified mapping: chartN.xml  <->  slot N */ };

/* ---------------------------------------------------------------------------
 * 1. Small helpers
 * ------------------------------------------------------------------------- */

function pad2(x) { return x < 10 ? "0" + x : "" + x; }

function collapseSpaces(v) {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\s+/g, " ").trim();
}

function derivePhase(flat) {
  if (flat === null || flat === undefined) return "Unknown";
  const m = /^\s*Phase[-\s]+([A-Za-z0-9]+)/.exec(String(flat));
  if (!m) return "Unknown";
  return "Phase " + m[1].toUpperCase();
}

const CREATED_DATE_RE =
  /^\s*(\d{1,2})-(\d{1,2})-(\d{2,4})(?:\s*,\s*(\d{1,2}):(\d{2})\s*([AaPp])\.?[Mm])?/;

function daysInMonth(y, m) {
  const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const dim = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  return dim[m - 1];
}

function parseCreatedDate(value) {
  // value: string, {y,m,d}(+h/min) struct, or null → {y,m,d} | null
  if (value === null || value === undefined) return null;
  if (typeof value === "object") return { y: value.y, m: value.m, d: value.d };
  const m = CREATED_DATE_RE.exec(String(value));
  if (!m) return null;
  const dd = +m[1], mm = +m[2];
  let yy = +m[3];
  if (yy < 100) yy += 2000;
  if (mm < 1 || mm > 12 || dd < 1 || dd > daysInMonth(yy, mm)) return null;
  return { y: yy, m: mm, d: dd };
}

function ordinal(n) {
  if (n % 100 >= 10 && n % 100 <= 20) return n + "th";
  const s = n % 10;
  return n + (s === 1 ? "st" : s === 2 ? "nd" : s === 3 ? "rd" : "th");
}

function monthSpan(first, last) {
  const out = [];
  let y = first.y, m = first.m;
  while (y < last.y || (y === last.y && m <= last.m)) {
    out.push({ y, m });
    m += 1;
    if (m === 13) { y += 1; m = 1; }
  }
  return out;
}

function fmtShort(d) {
  return pad2(d.d) + "-" + MONTHS[d.m - 1] + "-" + d.y;
}

function asOnLabel(d) {
  return pad2(d.d) + "." + pad2(d.m) + "." + String(d.y).slice(2);
}

function todayStruct() {
  const n = new Date();
  return { y: n.getFullYear(), m: n.getMonth() + 1, d: n.getDate() };
}

/* --- Excel serial dates (1900 system; python epoch 1899-12-30) ------------- */

// days since 0000-03-01 (Howard Hinnant civil algorithm)
function daysFromCivil(y, m, d) {
  const yy = y - (m <= 2 ? 1 : 0);
  const era = Math.floor((yy >= 0 ? yy : yy - 399) / 400);
  const yoe = yy - era * 400;
  const doy = Math.floor((153 * (m + (m > 2 ? -3 : 9)) + 2) / 5) + d - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe;
}
function civilFromDays(z) {
  z += 305;                    // shift to days since 0000-03-01 at z'=0
  const era = Math.floor((z >= 0 ? z : z - 146096) / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) -
    Math.floor(doe / 146096)) / 365);
  const y = yoe + era * 400;
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const d = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const m = mp + (mp < 10 ? 3 : -9);
  return { y: y + (m <= 2 ? 1 : 0), m, d };
}
const EPOCH_Z = daysFromCivil(1899, 12, 30);
function dateSerial(y, m, d) { return daysFromCivil(y, m, d) - EPOCH_Z; }
function serialToDate(serial) {
  return civilFromDays(serial + EPOCH_Z);
}
function pyDateStr(dt) {   // str(datetime(...)) replication
  const pad = (x, n) => {
    let s = String(Math.abs(x | 0));
    while (s.length < (n || 2)) s = "0" + s;
    return s;
  };
  return dt.y + "-" + pad(dt.m) + "-" + pad(dt.d) + " " +
    pad(dt.h || 0) + ":" + pad(dt.min || 0) + ":" + pad(dt.sec || 0);
}

/* --- misc ------------------------------------------------------------------ */

function xmlEscape(t) {
  return String(t)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function unescapeXml(s) {
  return String(s)
    .replace(/&(?:lt|gt|quot|apos|#x[0-9A-Fa-f]+|#x?[0-9A-Fa-f]*|amp);/g, (m) => {
      if (m === "&lt;") return "<";
      if (m === "&gt;") return ">";
      if (m === "&quot;") return "\"";
      if (m === "&apos;") return "'";
      if (m === "&amp;") return "&";
      const hex = /^&#x([0-9A-Fa-f]+);$/.exec(m);
      if (hex) return String.fromCodePoint(parseInt(hex[1], 16));
      return String.fromCodePoint(parseInt(m.slice(3, -1), 10));
    });
}

/** python str(float/int) for <v> and width/height attributes. */
function numStr(v) {
  return String(v);
}

function utf8(s) {
  return new TextEncoder().encode(s);
}
function utf8Read(u8) {
  return new TextDecoder("utf-8").decode(u8);
}

function b64Decode(b64) {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let clean = b64.replace(/[^A-Za-z0-9+/=]/g, "");
  const pad = clean.endsWith("==") ? 2 : clean.endsWith("=") ? 1 : 0;
  const out = new Uint8Array(Math.floor(clean.length / 4) * 3 - pad);
  let o = 0, buf = 0, bits = 0;
  for (let i = 0; i < clean.length; i++) {
    const ch = clean[i];
    if (ch === "=") break;
    const v = chars.indexOf(ch);
    if (v < 0) continue;
    buf = (buf << 6) | v;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      out[o++] = (buf >> bits) & 0xff;
    }
  }
  return out;
}

/* ---------------------------------------------------------------------------
 * 2. ZIP reader (dumps + tests) and writer (final workbook)
 * ------------------------------------------------------------------------- */

function rl_u16(b, o) { return b[o] | (b[o + 1] << 8); }
function rl_u32(b, o) {
  return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
}

async function inflateRaw(u8) {
  const ds = new DecompressionStream("deflate-raw");
  const resp = new Response(new Blob([u8]).stream().pipeThrough(ds));
  return new Uint8Array(await resp.arrayBuffer());
}

async function zipRead(u8) {
  let i = u8.length - 22;
  for (; i >= 0; i--) if (rl_u32(u8, i) === 0x06054b50) break;
  if (i < 0) throw new Error("not a zip file");
  const count = rl_u16(u8, i + 10);
  let p = rl_u32(u8, i + 16);
  const entries = new Map();
  for (let n = 0; n < count; n++) {
    if (rl_u32(u8, p) !== 0x02014b50) throw new Error("bad central directory");
    const method = rl_u16(u8, p + 10);
    const csize = rl_u32(u8, p + 20);
    const nameLen = rl_u16(u8, p + 28);
    const extraLen = rl_u16(u8, p + 30);
    const commentLen = rl_u16(u8, p + 32);
    const localOfs = rl_u32(u8, p + 42);
    const name = utf8Read(u8.subarray(p + 46, p + 46 + nameLen));
    entries.set(name, { method, csize, localOfs });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return {
    has: (n) => entries.has(n),
    names: () => [...entries.keys()],
    async read(name) {
      const e = entries.get(name);
      if (!e) return null;
      const lo = e.localOfs;
      if (rl_u32(u8, lo) !== 0x04034b50) throw new Error("bad local header " + name);
      const nlen = rl_u16(u8, lo + 26), elen = rl_u16(u8, lo + 28);
      const start = lo + 30 + nlen + elen;
      const comp = u8.subarray(start, start + e.csize);
      if (e.method === 0) return comp.slice();
      return await inflateRaw(comp);
    },
  };
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(u8) {
  let c = 0xffffffff;
  for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

async function deflateRaw(u8) {
  const cs = new CompressionStream("deflate-raw");
  const resp = new Response(new Blob([u8]).stream().pipeThrough(cs));
  return new Uint8Array(await resp.arrayBuffer());
}

/** python zipfile-like container: DEFLATE entries, now-stamped local times. */
async function buildZip(parts) {
  const chunks = [];
  const central = [];
  const now = new Date();
  const dTime = ((now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1)) & 0xffff;
  const dDate = (((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate()) & 0xffff;
  let offset = 0;
  for (const [name, data] of parts) {
    const nb = utf8(name);
    const comp = await deflateRaw(data);
    const crc = crc32(data);
    const lh = new Uint8Array(30 + nb.length);
    const dv = new DataView(lh.buffer);
    dv.setUint16(0, 0x4b50, true);           // "PK" (bytes 0-1 = 0x50,0x4b little-endian)
    lh[0] = 0x50; lh[1] = 0x4b; lh[2] = 0x03; lh[3] = 0x04;
    dv.setUint16(4, 20, true);               // version needed
    dv.setUint16(6, 0, true);                // flags
    dv.setUint16(8, 8, true);                // method = deflate
    dv.setUint16(10, dTime, true);
    dv.setUint16(12, dDate, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, comp.length, true);
    dv.setUint32(22, data.length, true);
    dv.setUint16(26, nb.length, true);
    dv.setUint16(28, 0, true);
    lh.set(nb, 30);
    chunks.push(lh, comp);

    const cd = new Uint8Array(46 + nb.length);
    const dv2 = new DataView(cd.buffer);
    cd[0] = 0x50; cd[1] = 0x4b; cd[2] = 0x01; cd[3] = 0x02;
    dv2.setUint16(4, 20 | (3 << 8), true);   // version made by: unix/20
    dv2.setUint16(6, 20, true);
    dv2.setUint16(8, 0, true);               // flags
    dv2.setUint16(10, 8, true);              // method
    dv2.setUint16(12, dTime, true);
    dv2.setUint16(14, dDate, true);
    dv2.setUint32(16, crc, true);
    dv2.setUint32(20, comp.length, true);
    dv2.setUint32(24, data.length, true);
    dv2.setUint16(28, nb.length, true);
    dv2.setUint16(30, 0, true);              // extra len
    dv2.setUint16(32, 0, true);              // comment len
    dv2.setUint16(34, 0, true);              // disk start
    dv2.setUint16(36, 0, true);              // internal attrs
    dv2.setUint32(38, 0x01800000, true);     // external attrs (0o600 << 16)
    dv2.setUint32(42, offset, true);
    cd.set(nb, 46);
    central.push(cd);
    offset += 30 + nb.length + comp.length;
  }
  const cdStart = offset;
  let cdSize = 0;
  for (const cd of central) cdSize += cd.length;
  for (const cd of central) chunks.push(cd);
  const eocd = new Uint8Array(22);
  const dv3 = new DataView(eocd.buffer);
  eocd[0] = 0x50; eocd[1] = 0x4b; eocd[2] = 0x05; eocd[3] = 0x06;
  dv3.setUint16(8, central.length, true);
  dv3.setUint16(10, central.length, true);
  dv3.setUint32(12, cdSize, true);
  dv3.setUint32(16, cdStart, true);
  chunks.push(eocd);
  const total = chunks.reduce((a, c) => a + c.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return out;
}

/* ---------------------------------------------------------------------------
 * 3. Dump parsing (port of read_dump_rows)
 * ------------------------------------------------------------------------- */

const HEADER_ALIASES = {
  "id": "id", "i.d": "id", "ticket id": "id",
  "created date": "created", "createddate": "created", "date": "created",
  "category": "category",
  "sub category": "sub", "sub-category": "sub", "subcategory": "sub",
  "flat": "flat", "house": "flat",
  "subject": "subject", "description": "subject",
  "status": "status", "mygate status": "status",
};

function colIndexOf(ref) {
  let n = 0;
  for (const ch of ref) {
    if (ch >= "A" && ch <= "Z") n = n * 26 + (ch.charCodeAt(0) - 64);
    else break;
  }
  return n;
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - 1 - m) / 26;
  }
  return s;
}

/** Scan xl/styles.xml → is the (sheet, cellStyleIndex) a date format? */
async function dumpDateStyleSet(zip) {
  const set = new Set([0]);
  const styles = await zip.read("xl/styles.xml");
  if (styles) {
    const txt = utf8Read(styles);
    const custom = new Map();
    const nfRe = /<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/g;
    let m;
    while ((m = nfRe.exec(txt))) custom.set(+m[1], m[2]);
    const builtinDate = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22,
      27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54,
      55, 56, 57, 58]);
    const isDateFmt = (id) => {
      if (builtinDate.has(id)) return true;
      const code = custom.get(id);
      if (!code) return false;
      const bare = code.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "");
      return /[dyhs]/.test(bare.toLowerCase());
    };
    const xfs = txt.match(/<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/);
    if (xfs) {
      const xfRe = /<xf ([^>]*?)(\/>|>)/g;
      let xm;
      let i = 0;
      while ((xm = xfRe.exec(xfs[1]))) {
        const idm = /numFmtId="(\d+)"/.exec(xm[1]);
        const id = idm ? +idm[1] : 0;
        if (isDateFmt(id)) set.add(i);
        i++;
      }
    }
  }
  return set;
}

async function readDumpRows(dumpBytes) {
  const zip = await zipRead(dumpBytes);

  // Locate the first worksheet (mirrors openpyxl wb.worksheets[0])
  let sheetPart = null;
  const wbXml = await zip.read("xl/workbook.xml");
  if (wbXml) {
    const txt = utf8Read(wbXml);
    const sm = txt.match(/<sheet [^>]*?r:id="([^"]+)"[^>]*?\/>/) ||
               txt.match(/<sheet [^>]*?r:id="([^"]+)"/);
    const rid = sm && sm[1];
    if (rid) {
      const relsXml = await zip.read("xl/_rels/workbook.xml.rels");
      if (relsXml) {
        const rels = utf8Read(relsXml);
        let target = null;
        const rm1 = new RegExp('<Relationship[^>]*?Id="' + rid + '"[^>]*?Target="([^"]+)"')
          .exec(rels);
        const rm2 = new RegExp('<Relationship[^>]*?Target="([^"]+)"[^>]*?Id="' + rid + '"')
          .exec(rels);
        if (rm1) target = rm1[1];
        else if (rm2) target = rm2[1];
        if (target) {
          target = target.replace(/^\//, "");
          if (!/^xl\//.test(target)) target = "xl/" + target;
          sheetPart = target;
        }
      }
    }
  }
  if (!sheetPart) {
    const names = zip.names();
    sheetPart =
      names.find((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)) ||
      names.find((n) => /worksheets\/.*\.xml/.test(n));
  }
  if (!sheetPart) throw new Error("dump has no worksheet");

  // shared strings (if any)
  const sst = [];
  const sstPart = await zip.read("xl/sharedStrings.xml");
  if (sstPart) {
    const sstRe = /<si(?:\s[^>]*)?>([\s\S]*?)<\/si>/g;
    const body = utf8Read(sstPart);
    let m;
    while ((m = sstRe.exec(body))) {
      const runs = [];
      const tRe = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
      let t;
      while ((t = tRe.exec(m[1]))) runs.push(unescapeXml(t[1]));
      sst.push(runs.join(""));
    }
  }

  // date-formatted style indexes (for numeric Created Date cells)
  const dateStyles = await dumpDateStyleSet(zip);

  const xml = utf8Read(await zip.read(sheetPart));
  const grid = new Map();                    // rowNum → Map(colNum → valueInfo)
  const rowRe = /<row\b[^>]*?\br="(\d+)"[^>]*?>([\s\S]*?)<\/row>/g;
  let rm;
  while ((rm = rowRe.exec(xml))) {
    const rowNum = +rm[1];
    const cells = new Map();
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cm;
    while ((cm = cellRe.exec(rm[2]))) {
      const attrs = cm[1];
      const body = cm[2] === undefined ? "" : cm[2];
      const rmo = /\br="([A-Z]+)(\d+)"/.exec(attrs);
      if (!rmo) continue;
      let value = null;
      let raw = null;
      const tm = body.match(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/);
      const vm = body.match(/<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/);
      const type = (/\bt="([a-zA-Z]+)"/.exec(attrs) || [])[1] || "n";
      if (type === "inlineStr") {
        value = tm ? unescapeXml(tm[1]) : "";
      } else if (type === "s") {
        const idx = vm ? +vm[1] : -1;
        value = sst[idx] !== undefined ? sst[idx] : null;
      } else if (type === "str") {
        value = vm ? unescapeXml(vm[1]) : null;
      } else if (type === "b") {
        value = vm ? (vm[1] !== "0" ? "TRUE" : "FALSE") : null;
      } else if (vm && vm[1] !== "") {
        raw = vm[1];
        value = +vm[1];
        // openpyxl-style numeric typing: int when no '.', otherwise float
        if (!/[.eE]/.test(vm[1])) value = Math.trunc(value);
      }
      cells.set(colIndexOf(rmo[1]), { value, raw, s: +((/\bs="(\d+)"/.exec(attrs) || [])[1] || 0) });
    }
    grid.set(rowNum, cells);
  }

  // header row detection (rows 1..12, cols 1..20; needs id + created)
  let headerRowIdx = null;
  let colmap = {};
  for (let r = 1; r <= 12; r++) {
    const cells = grid.get(r);
    if (!cells) continue;
    const keys = {};
    for (const [col, info] of cells) {
      if (info.value === null || info.value === undefined || col in keys || col > 20) continue;
      const k = collapseSpaces(info.value).toLowerCase();
      if (k in HEADER_ALIASES) keys[col] = HEADER_ALIASES[k];
    }
    const valsFound = Object.values(keys);
    if (valsFound.includes("id") && valsFound.includes("created")) {
      headerRowIdx = r;
      for (const [col, alias] of Object.entries(keys)) {
        if (!(alias in colmap)) colmap[alias] = +col;
      }
      break;
    }
  }
  if (!Object.keys(colmap).length) {
    throw new Error(
      "Could not find a header row (I.D / Created Date / Category ...) " +
      "in the first 12 rows of the dump file.");
  }

  // From/To meta scan before the header (rows 1..max(2, header-1), cols 1..3)
  const meta = { from: null, to: null };
  const metaTop = Math.max(2, headerRowIdx - 1);
  for (let r = 1; r <= metaTop; r++) {
    const cells = grid.get(r);
    if (!cells) continue;
    for (const [col, info] of cells) {
      if (typeof info.value !== "string" || col > 3) continue;
      const mm = /From\s+([\d\-]+)\s+To\s+([\d\-]+)/.exec(info.value);
      if (!mm) continue;
      const pairs = [["from", mm[1]], ["to", mm[2]]];
      for (const [key, rawDate] of pairs) {
        const parts = rawDate.split("-");
        if (parts.length !== 3 || !parts.every((x) => x !== "")) continue;
        const d = +parts[0], m2 = +parts[1], y2 = +parts[2];
        if (m2 >= 1 && m2 <= 12 && d >= 1 && d <= daysInMonth(y2, m2)) {
          meta[key] = { y: y2, m: m2, d };
        }
      }
    }
  }

  // raw rows
  const rawRows = [];
  for (const [rowNum, cells] of grid) {
    if (rowNum <= headerRowIdx) continue;
    const get = (alias) => {
      if (!(alias in colmap)) return null;
      const info = cells.get(colmap[alias]);
      return info === undefined ? null : info;
    };
    const tidInfo = get("id");
    let tid = tidInfo === null ? null : tidInfo.value;
    if (tid === null || (typeof tid === "string" && !tid.trim())) continue;
    const createdInfo = get("created");
    let created = createdInfo === null ? null : createdInfo.value;
    // openpyxl converts serial+date-numFmt to datetime; python str() renders it
    if (createdInfo !== null && createdInfo.raw !== null &&
        typeof created === "number" && dateStyles.has(createdInfo.s)) {
      const d2 = serialToDate(created);
      created = {
        y: d2.y, m: d2.m, d: d2.d,
        h: Math.floor((created % 1) * 24),
        min: Math.floor(((created % 1) * 24 % 1) * 60),
      };
    }
    const subjInfo = get("subject");
    let subject;
    if (subjInfo === null || subjInfo.value === null) subject = "";
    else if (typeof subjInfo.value === "number") subject = subjInfo.raw;
    else subject = String(subjInfo.value);
    // float ids that are whole numbers become ints (python does the same)
    if (typeof tid === "number" && Number.isInteger(tid)) tid = Math.trunc(tid);
    rawRows.push({
      id: tid,
      created,
      category: collapseSpaces(getSafe(get("category"), true)) || "Unknown",
      sub: collapseSpaces(getSafe(get("sub"), true)) || null,
      flat: collapseSpaces(getSafe(get("flat"), true)),
      subject,
      status: collapseSpaces(getSafe(get("status"), true)) || "Unknown",
    });
  }
  function getSafe(info, collapse) {
    if (info === null) return null;
    if (info.value === null || info.value === undefined) return null;
    if (typeof info.value === "number") return String(info.raw);
    return info.value;
  }

  const rows = rawRows.map((item) => ({
    id: item.id,
    created: item.created,
    category: item.category,
    sub: item.sub === null || item.sub === "" || item.sub === "None" ? "None" : item.sub,
    flat: item.flat,
    subject: item.subject,
    status: item.status,
    date: parseCreatedDate(item.created),
    phase: derivePhase(item.flat),
  }));

  rows.sort((a, b) => {
    const da = a.date ? (a.date.y * 10000 + a.date.m * 100 + a.date.d) : Infinity;
    const db = b.date ? (b.date.y * 10000 + b.date.m * 100 + b.date.d) : Infinity;
    const ia = typeof a.id === "number" ? a.id : Infinity;
    const ib = typeof b.id === "number" ? b.id : Infinity;
    if (da !== db) return da - db;
    if (ia !== ib) return ia - ib;
    return 0;
  });

  return { rows, meta };
}

function deriveAsOn(meta, rows, dumpFileName) {
  if (dumpFileName) {
    const m = /_to_(\d{4})-(\d{1,2})-(\d{1,2})/.exec(dumpFileName);
    if (m) {
      const y = +m[1], mo = +m[2], d = +m[3];
      if (mo >= 1 && mo <= 12 && d >= 1 && d <= daysInMonth(y, mo) &&
          y > 1200) return { y, m: mo, d };
    }
  }
  let maxD = null;
  for (const r of rows) {
    if (!r.date) continue;
    const k = r.date.y * 10000 + r.date.m * 100 + r.date.d;
    if (!maxD || k > (maxD.y * 10000 + maxD.m * 100 + maxD.d)) maxD = r.date;
  }
  return maxD || todayStruct();
}

function outputName(asOn) {
  return "Mygate Complaints Dashboard as on " + asOnLabel(asOn) + ".xlsx";
}

/* ---------------------------------------------------------------------------
 * 4. Aggregation (port of aggregate; python Counter.most_common ties keep
 *    first-seen order — replicated with insertion-index tiebreaks)
 * ------------------------------------------------------------------------- */

function FirstSeenCounter(items) {
  const counts = new Map();
  const first = new Map();
  let n = 0;
  this.add = function (k) {
    if (!counts.has(k)) { counts.set(k, 0); first.set(k, n++); }
    counts.set(k, counts.get(k) + 1);
  };
  this.mostCommon = function () {
    return [...counts.entries()]
      .sort((a, b) => (b[1] !== a[1] ? b[1] - a[1] : first.get(a[0]) - first.get(b[0])))
      .map((e) => [e[0], e[1]]);
  };
  this.get = function (k, dflt) { return counts.has(k) ? counts.get(k) : dflt; };
  this.has = function (k) { return counts.has(k); };
}

function aggregate(rows) {
  const status = new FirstSeenCounter();
  const category = new FirstSeenCounter();
  const subcat = new FirstSeenCounter();
  const phase = new FirstSeenCounter();
  for (const r of rows) {
    status.add(r.status);
    category.add(r.category);
    subcat.add(r.sub);
    phase.add(r.phase);
  }
  const month = new Map();
  for (const r of rows) {
    if (!r.date) continue;
    const key = r.date.y + ":" + r.date.m;
    const rec = month.get(key);
    if (rec) rec.count += 1;
    else month.set(key, { count: 1, y: r.date.y, m: r.date.m });
  }

  const stats = {
    total: rows.length,
    status, category, subcat, phase, month,
  };

  const statusDesc = status.mostCommon().map((e) => e[0]);
  stats.status_order = statusDesc.length ? statusDesc : STATUS_ORDER_FALLBACK.slice();
  const pie = STATUS_PIE_ROWS.slice();
  for (const s of stats.status_order) if (!pie.includes(s)) pie.push(s);
  stats.status_table = pie;
  stats.category_order = category.mostCommon().map((e) => e[0]);
  stats.phase_order = phase.mostCommon().map((e) => e[0]);

  stats.subcat_chart = subcat.mostCommon()
    .map((e) => e[0])
    .filter((s) => s !== "None" && s !== "")
    .slice(0, 12);

  // ["All"] + sorted(first-seen subs, by lowercase; stable like python sorted)
  const seen = [];
  for (const r of rows) {
    const s = r.sub;
    if (s === undefined) continue;
    if (s === "" || seen.indexOf(s) !== -1) continue;
    seen.push(s);
  }
  stats.subcat_filter = ["All"].concat(
    seen.sort((a, b) => {
      const la = a.toLowerCase(), lb = b.toLowerCase();
      return la < lb ? -1 : la > lb ? 1 : 0;
    })
  );

  let minD = null, maxD = null;
  for (const r of rows) {
    if (!r.date) continue;
    const k = r.date.y * 10000 + r.date.m * 100 + r.date.d;
    if (!minD || k < (minD.y * 10000 + minD.m * 100 + minD.d)) minD = r.date;
    if (!maxD || k > (maxD.y * 10000 + maxD.m * 100 + maxD.d)) maxD = r.date;
  }
  const t = todayStruct();
  stats.months = minD
    ? monthSpan(minD, maxD)
    : monthSpan({ y: t.y, m: t.m, d: 1 }, { y: t.y, m: t.m, d: t.d });
  return stats;
}

/* ---------------------------------------------------------------------------
 * 5. Worksheet model + openpyxl 3.1.5-shaped serializer
 * ------------------------------------------------------------------------- */

function SheetModel(opts) {
  opts = opts || {};
  this.showGridLines = opts.showGridLines;          // true/unset → no attr
  this.freeze = opts.freeze || null;                // e.g. "A2" on Data
  this.hasDrawing = !!opts.drawing;
  this.cols = [];                                   // [letter, width]
  this.rowHeights = {};                             // row → number
  this.rows = new Map();                            // row → [cell,...]
  this.merges = opts.merges || [];
  this.dvs = null;
  this.minRow = Infinity; this.maxRow = -Infinity;
  this.minCol = Infinity; this.maxCol = -Infinity;
}

SheetModel.prototype.cell = function (row, col, value, styleId, opts) {
  opts = opts || {};
  if (!this.rows.has(row)) this.rows.set(row, []);
  this.rows.get(row).push({
    col, value,
    s: styleId || 0,
    isFormula: !!opts.isFormula,
  });
  if (row < this.minRow) this.minRow = row;
  if (row > this.maxRow) this.maxRow = row;
  if (col < this.minCol) this.minCol = col;
  if (col > this.maxCol) this.maxCol = col;
};

SheetModel.prototype.setColWidths = function (pairs) {
  for (const [letter, width] of pairs) this.cols.push([letter, width]);
};

SheetModel.prototype.setHeights = function (map) {
  Object.assign(this.rowHeights, map);
};

function SstStore() {
  this.index = new Map();
  this.entries = [];
  this.refs = 0;
  this.intern = function (text) {
    this.refs += 1;
    let idx = this.index.get(text);
    if (idx === undefined) {
      idx = this.entries.length;
      this.index.set(text, idx);
      this.entries.push(text);
    }
    return idx;
  };
  this.doc = function () {
    let out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'count="' + this.refs + '" uniqueCount="' + this.entries.length + '">';
    for (const text of this.entries) {
      const p = text !== text.trim() ? ' xml:space="preserve"' : "";
      // python's restore captures the xml-escaped text out of the sheet XML,
      // so <si><t> must carry the escaped form too.
      out += "<si><t" + p + ">" + xmlEscape(text) + "</t></si>";
    }
    return out + "</sst>";
  };
}

/** Serialize one worksheet — byte-faithful to openpyxl 3.1.5's incremental
 *  writer for exactly the feature subset the analysis sheets use. */
function renderSheet(mod, sstParam) {
  const sst = sstParam || new SstStore();
  const out = [];
  out.push('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
  out.push('<sheetPr><outlinePr summaryBelow="1" summaryRight="1" /><pageSetUpPr /></sheetPr>');
  out.push('<dimension ref="' +
    colLetter(mod.minCol) + mod.minRow + ":" + colLetter(mod.maxCol) + mod.maxRow + '" />');
  if (mod.freeze) {
    // python: ws.freeze_panes = "A2" → ySplit=1 pane at A2
    const m = /^([A-Z]+)(\d+)$/.exec(mod.freeze);
    const colSplit = colIndexOf(m[1]), rowSplit = +m[2];
    let pane = "";
    if (rowSplit > 1) pane += 'ySplit="' + (rowSplit - 1) + '" ';
    if (colSplit > 1) pane += 'xSplit="' + (colSplit - 1) + '" ';
    pane += 'topLeftCell="' + mod.freeze + '" activePane="bottomLeft" state="frozen"';
    const isBottom = (rowSplit > 1 && colSplit > 1) ? "bottomRight" : colSplit > 1 ? "topRight" : "bottomLeft";
    out.push('<sheetViews><sheetView showGridLines="' + (mod.showGridLines === false ? "0" : "1") +
      '" workbookViewId="0"><pane ' + pane + ' /><selection pane="' + isBottom +
      '" activeCell="A1" sqref="A1" /></sheetView></sheetViews>');
  } else if (mod.showGridLines === false) {
    out.push('<sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="A1" sqref="A1" /></sheetView></sheetViews>');
  } else {
    out.push('<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1" /></sheetView></sheetViews>');
  }
  out.push('<sheetFormatPr defaultColWidth="8.7265625" defaultRowHeight="14.5" />');
  if (mod.cols.length) {
    out.push("<cols>");
    for (const [letter, width] of mod.cols) {
      const n = colIndexOf(letter);
      out.push('<col width="' + numStr(width) + '" customWidth="1" min="' + n +
        '" max="' + n + '" />');
    }
    out.push("</cols>");
  }
  out.push("<sheetData>");
  const rowsSorted = [...mod.rows.keys()].sort((a, b) => a - b);
  for (const r of rowsSorted) {
    const cells = mod.rows.get(r).slice().sort((a, b) => a.col - b.col);
    const ht = mod.rowHeights[r];
    out.push("<row " + (ht === undefined
      ? 'r="' + r + '"'
      : 'r="' + r + '" ht="' + numStr(ht) + '" customHeight="1"') + ">");
    for (const c of cells) out.push(emitCell(r, c, sst));
    out.push("</row>");
  }
  out.push("</sheetData>");
  if (mod.merges.length) {
    out.push('<mergeCells count="' + mod.merges.length + '">');
    for (const mr of mod.merges) out.push('<mergeCell ref="' + mr + '" />');
    out.push("</mergeCells>");
  }
  if (mod.dvs && mod.dvs.length) {
    out.push('<dataValidations count="' + mod.dvs.length + '">');
    for (const dv of mod.dvs) {
      out.push('<dataValidation sqref="' + dv.sqref +
        '" showDropDown="0" showInputMessage="0" showErrorMessage="0" allowBlank="1" type="list">' +
        "<formula1>" + xmlEscape(dv.formula1) + "</formula1></dataValidation>");
    }
    out.push("</dataValidations>");
  }
  out.push('<pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5" />');
  if (mod.hasDrawing) {
    out.push('<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" />');
  }
  out.push("</worksheet>");
  // Final form replicates python's SST restoration: inline strings were
  // captured in document order across sheet1..5 and deduped.
  return { xml: out.join(""), sst };
}

function emitCell(r, c, sst) {
  const ref = colLetter(c.col) + r;
  const sAttr = c.s ? ' s="' + c.s + '"' : "";
  if (c.isFormula) {
    // openpyxl drops the leading "=" at bind time; <f> text never has it
    const fText = c.value.charAt(0) === "=" ? c.value.slice(1) : c.value;
    return '<c r="' + ref + '"' + sAttr + '><f>' + xmlEscape(fText) + "</f><v /></c>";
  }
  if (c.value === null || c.value === undefined) {
    return '<c r="' + ref + '"' + sAttr + ' t="n" />';
  }
  if (typeof c.value === "number") {
    return '<c r="' + ref + '"' + sAttr + ' t="n"><v>' + numStr(c.value) + "</v></c>";
  }
  if (typeof c.value === "object" && c.value.serial !== undefined) {
    return '<c r="' + ref + '"' + sAttr + ' t="n"><v>' + numStr(c.value.serial) + "</v></c>";
  }
  if (c.value === "") {
    return '<c r="' + ref + '"' + sAttr + ' t="inlineStr" />';
  }
  const idx = sst.intern(c.value);
  return '<c r="' + ref + '"' + sAttr + ' t="s"><v>' + idx + "</v></c>";
}

/* ---------------------------------------------------------------------------
 * 6. Sheet builders
 * ------------------------------------------------------------------------- */

const DATA_COLS = CT.DATA_COLS;

function buildDataSheet(rows) {
  const mod = new SheetModel({ showGridLines: false, freeze: "A2" });
  const widths = [7.0, 8.0, 17.0, 16.71, 16.71, 20.0, 10.0, 17.0, 36.43, 13.0];
  for (let i = 0; i < 10; i++) {
    mod.cell(1, i + 1, DATA_HEADERS[i], DATA_COLS.hdr[colLetter(i + 1)]);
  }
  mod.rowHeights[1] = 30;
  let idx = 1;
  for (const r of rows) {
    const row = idx + 1;
    const band = idx % 2 === 1 ? DATA_COLS.odd : DATA_COLS.even;
    const createdVal = r.created === null || r.created === undefined
      ? null
      : typeof r.created === "string"
        ? r.created
        : typeof r.created === "object" && r.created.y
          ? pyDateStr(r.created)
          : String(r.created);
    const idVal = typeof r.id === "string" && /^\s*\d+\s*$/.test(r.id)
      ? parseInt(r.id.trim(), 10)
      : r.id;
    const values = [
      idx,
      idVal,
      createdVal,
      r.date ? { serial: dateSerial(r.date.y, r.date.m, r.date.d) } : null,
      r.category, r.sub, r.phase, r.flat, r.subject, r.status,
    ];
    for (let i = 0; i < 10; i++) mod.cell(row, i + 1, values[i], band[colLetter(i + 1)]);
    idx++;
  }
  const letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J"];
  for (let i = 0; i < 10; i++) mod.setColWidths([[letters[i], widths[i]]]);
  return mod;
}

function buildCalcSheet(stats) {
  const frows = FORMULA_ROW_END;
  const mod = new SheetModel({});
  const filters =
    "Data!$E$2:$E$" + frows + ",Calc!$N$4," +
    "Data!$G$2:$G$" + frows + ",Calc!$N$1," +
    "Data!$J$2:$J$" + frows + ",Calc!$N$2," +
    "Data!$F$2:$F$" + frows + ",Calc!$N$3";
  const filtersNoStatus =
    "Data!$E$2:$E$" + frows + ",Calc!$N$4," +
    "Data!$G$2:$G$" + frows + ",Calc!$N$1," +
    "Data!$F$2:$F$" + frows + ",Calc!$N$3";

  mod.cell(1, 13, "phaseCrit");
  mod.cell(1, 14, '=IF(Interactive!$C$5="All","*",Interactive!$C$5)', 0, { isFormula: true });
  mod.cell(2, 13, "statusCrit");
  mod.cell(2, 14, '=IF(Interactive!$F$5="All","*",Interactive!$F$5)', 0, { isFormula: true });
  mod.cell(3, 13, "subCategoryCrit");
  mod.cell(3, 14, '=IF(Interactive!$I$5="All","*",Interactive!$I$5)', 0, { isFormula: true });
  mod.cell(4, 13, "catCrit");
  mod.cell(4, 14, '=IF(Interactive!$L$5="All","*",Interactive!$L$5)', 0, { isFormula: true });

  mod.cell(1, 22, "COUNTIFS(" + filters + ")", 0, { isFormula: true });
  mod.cell(2, 22, 'COUNTIFS(Data!$J$2:$J$' + frows + ',"In Progress",' + filtersNoStatus + ")", 0, { isFormula: true });
  mod.cell(3, 22, 'COUNTIFS(Data!$J$2:$J$' + frows + ',"Reopened",' + filtersNoStatus + ")", 0, { isFormula: true });
  mod.cell(4, 22, 'COUNTIFS(Data!$J$2:$J$' + frows + ',"On Hold",' + filtersNoStatus + ")", 0, { isFormula: true });
  const lastSubRow = Math.max(1 + stats.subcat_chart.length, 13);
  mod.cell(5, 22, "MAX(E2:E" + lastSubRow + ")", 0, { isFormula: true });
  mod.cell(7, 22, "COUNTA(Data!$B$2:$B$" + frows + ")", 0, { isFormula: true });

  let i = 2;
  for (const cat of stats.category_order) {
    mod.cell(i, 1, cat);
    mod.cell(i, 2, "COUNTIFS(Data!$E$2:$E$" + frows + ",$A$" + i + "," + filters + ")", 0, { isFormula: true });
    i++;
  }
  i = 2;
  for (const sub of stats.subcat_chart) {
    mod.cell(i, 4, sub);
    mod.cell(i, 5, "COUNTIFS(Data!$F$2:$F$" + frows + ",$D$" + i + "," + filters + ")", 0, { isFormula: true });
    i++;
  }
  i = 2;
  for (const st of stats.status_table) {
    mod.cell(i, 7, st);
    mod.cell(i, 8, "COUNTIFS(Data!$J$2:$J$" + frows + ",$G$" + i + "," + filters + ")", 0, { isFormula: true });
    i++;
  }
  i = 2;
  for (const ph of stats.phase_order) {
    mod.cell(i, 16, ph);
    mod.cell(i, 17, "COUNTIFS(Data!$G$2:$G$" + frows + ",$P$" + i + "," + filters + ")", 0, { isFormula: true });
    i++;
  }
  i = 2;
  for (const mo of stats.months) {
    mod.cell(i, 19, { serial: dateSerial(mo.y, mo.m, 1) }, CT.CALC_DATE);
    mod.cell(i, 20,
      'COUNTIFS(Data!$D$2:$D$' + frows + ',">="&DATE(' + mo.y + "," + mo.m + ",1)," +
      "Data!$D$2:$D$" + frows + ',"<"&EDATE(DATE(' + mo.y + "," + mo.m + ",1),1)," +
      filters + ")", 0, { isFormula: true });
    i++;
  }
  mod.cell(1, 24, "All");
  i = 2;
  for (const sub of stats.subcat_filter.slice(1)) {
    mod.cell(i, 24, sub);
    i++;
  }
  mod.setColWidths([["A", 12.29]]);
  return mod;
}

const FILTER_CELLS = [["C4", "Phase", "C5"], ["F4", "Status", "F5"],
  ["I4", "Sub Category", "I5"], ["L4", "Category", "L5"]];

function buildInteractive(stats, asOn) {
  const mod = new SheetModel({ showGridLines: false, drawing: true });
  const S = CT.S.interactive;

  mod.merges = CT.MERGES1.slice();

  mod.cell(2, 2, BRAND + "  \u2014  Interactive Complaints Dashboard", S.B2);
  mod.cell(3, 2,
    "Use the Phase / Status / Sub Category / Category filters below \u2014 every KPI " +
    "and chart updates to match your selection.  Source: MyGate export " +
    "(" + ordinal(asOn.d) + "-" + MONTHS[asOn.m - 1] + "-" + asOn.y + ").", S.B3);
  mod.cell(4, 2, "FILTERS", S.B4);
  for (const [labelCell, label, valueCell] of FILTER_CELLS) {
    const lm = /^([A-Z]+)(\d+)$/.exec(labelCell);
    const vm = /^([A-Z]+)(\d+)$/.exec(valueCell);
    mod.cell(+lm[2], colIndexOf(lm[1]), label, S[labelCell]);
    mod.cell(+vm[2], colIndexOf(vm[1]), "All", S[valueCell]);
  }

  const cards = [
    ["B", "C", "FILTERED TOTAL", "Calc!$V$1", "General", "complaints in view"],
    ["E", "F", "% OF ALL", "Calc!$V$1/Calc!$V$7", "0.0%",
      '="share of "&TEXT(Calc!$V$7,"#,##0")'],
    ["H", "I", "IN PROGRESS", "Calc!$V$2", "General", "open"],
    ["K", "L", "REOPENED", "Calc!$V$3", "General", "closed but returned"],
    ["N", "O", "TOP ISSUE COUNT", "Calc!$V$5", "General",
      "=INDEX(Calc!$D$2:$D$13,MATCH(Calc!$V$5,Calc!$E$2:$E$13,0))"],
  ];
  for (const [col, right, label, formula, fmt, nine] of cards) {
    const cI = colIndexOf(col), rI = colIndexOf(right);
    mod.cell(7, cI, label, S[col + "7"]);
    mod.cell(7, rI, null, S[right + "7"]);
    mod.cell(8, cI, formula, S[col + "8"], { isFormula: true });
    mod.cell(8, rI, null, S[right + "8"]);
    if (nine.startsWith("=")) {
      mod.cell(9, cI, nine.slice(1), S[col + "9"], { isFormula: true });
    } else {
      mod.cell(9, cI, nine, S[col + "9"]);
    }
    mod.cell(9, rI, null, S[right + "9"]);
  }

  mod.dvs = [
    { sqref: "C5", formula1: '"All,' + stats.phase_order.join(",") + '"' },
    { sqref: "F5", formula1: '"All,' + stats.status_order.join(",") + '"' },
    { sqref: "I5", formula1: "SubCategoryFilterList" },
    { sqref: "L5", formula1: '"All,' + stats.category_order.join(",") + '"' },
  ];

  mod.setHeights(CT.HEIGHTS.interactive);
  mod.setColWidths([
    ["A", 2.0], ["B", 13.0], ["C", 13.0], ["D", 3.0], ["E", 13.0], ["F", 13.0],
    ["G", 3.0], ["H", 13.0], ["I", 13.0], ["J", 3.0], ["K", 13.0], ["L", 13.0],
    ["M", 3.0], ["N", 13.0], ["O", 13.0], ["P", 2.0],
  ]);
  return mod;
}

function buildOverview(stats, asOn, periodFrom, periodTo) {
  const mod = new SheetModel({ showGridLines: false, drawing: true });
  const S = CT.S.overview;
  mod.merges = CT.MERGES2.slice();

  mod.cell(2, 2, BRAND + "  \u2014  Complaints Analytics Dashboard", S.B2);
  mod.cell(3, 2,
    "MyGate Help Desk Report  |  Reporting period: " + fmtShort(periodFrom) + " to " +
    fmtShort(periodTo) + "  |  Source: MyGate complaints export " +
    "(" + ordinal(asOn.d) + " " + MONTHS[asOn.m - 1] + "-" + asOn.y + ")", S.B3);

  const cards = [
    ["B", "C", "TOTAL COMPLAINTS", "COUNTA(Data!$B$2:$B$" + FORMULA_ROW_END + ")", "All logged tickets"],
    ["E", "F", "OPEN / IN PROGRESS", 'COUNTIF(Data!$J$2:$J$' + FORMULA_ROW_END + ',"In Progress")', "Awaiting resolution"],
    ["H", "I", "REOPENED", 'COUNTIF(Data!$J$2:$J$' + FORMULA_ROW_END + ',"Reopened")', "Returned after closure"],
    ["K", "L", "ON HOLD", 'COUNTIF(Data!$J$2:$J$' + FORMULA_ROW_END + ',"On Hold")', "Held by Projects"],
  ];
  for (const [col, right, label, formula, sub] of cards) {
    const cI = colIndexOf(col), rI = colIndexOf(right);
    mod.cell(5, cI, label, S[col + "5"]);
    mod.cell(5, rI, null, S[right + "5"]);
    mod.cell(6, cI, formula, S[col + "6"], { isFormula: true });
    mod.cell(6, rI, null, S[right + "6"]);
    mod.cell(7, cI, sub, S[col + "7"]);
    mod.cell(7, rI, null, S[right + "7"]);
  }

  mod.setHeights(CT.HEIGHTS.overview);
  mod.setColWidths([
    ["A", 2.0], ["B", 16.0], ["C", 16.0], ["D", 3.0], ["E", 16.0], ["F", 16.0],
    ["G", 3.0], ["H", 16.0], ["I", 16.0], ["J", 3.0], ["K", 16.0], ["L", 16.0],
    ["M", 2.0],
  ]);
  return mod;
}

function buildAnalysis(stats) {
  const frows = FORMULA_ROW_END;
  const mod = new SheetModel({ showGridLines: false });
  const A = CT.ANALYSIS;
  const rng = {
    cat: "Data!$E$2:$E$" + frows,
    sub: "Data!$F$2:$F$" + frows,
    status: "Data!$J$2:$J$" + frows,
    phase: "Data!$G$2:$G$" + frows,
    date: "Data!$D$2:$D$" + frows,
  };

  const cats = stats.category_order;
  const subs = stats.subcat_chart.slice(0, SUBCAT_CHART_ROWS);
  const stas = stats.status_table;
  const phs = stats.phase_order;
  const months = stats.months;

  const band1End = Math.max(2 + cats.length, 2 + subs.length, 2 + stas.length, 5);
  const title2 = band1End + 2;
  const header2 = title2 + 1;
  const body2 = header2 + 1;
  stats.analysis_rows = [title2, header2, body2];

  const title = (ref, txt) => {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    mod.cell(+m[2], colIndexOf(m[1]), txt, A.TITLE);
  };
  title("A1", "Complaints by Category");
  title("D1", "Top Issue Types (Sub-Category)");
  title("G1", "By Resolution Status");
  title("J1", "Defect Liability Period (DLP) Status");
  title("A" + title2, "Complaints by Phase / Block");
  title("D" + title2, "Monthly Complaint Trend");
  mod.rowHeights[1] = 15.5;
  mod.rowHeights[title2] = 15.5;

  const headers = [
    ["A2", "Category"], ["B2", "Count"], ["D2", "Sub-Category"], ["E2", "Count"],
    ["G2", "Status"], ["H2", "Count"], ["J2", "DLP Status"], ["K2", "Count"],
    ["A" + header2, "Phase"], ["B" + header2, "Count"],
    ["D" + header2, "Month"], ["E" + header2, "Count"],
  ];
  for (const [ref, txt] of headers) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    mod.cell(+m[2], colIndexOf(m[1]), txt, A.HEADER);
  }

  const styled = (row, labelCol, labelValue, countValue, opts) => {
    opts = opts || {};
    mod.cell(row, labelCol, labelValue, opts.month ? A.MONTH_LABEL : A.LABEL);
    mod.cell(row, labelCol + 1, countValue, A.COUNT,
      opts.formula ? { isFormula: true } : undefined);
  };

  let i = 3;
  for (const v of cats) {
    styled(i, 1, v, "COUNTIF(" + rng.cat + ",A" + i + ")", { formula: true });
    i++;
  }
  i = 3;
  for (const v of subs) {
    styled(i, 4, v, "COUNTIF(" + rng.sub + ",D" + i + ")", { formula: true });
    i++;
  }
  i = 3;
  for (const v of stas) {
    styled(i, 7, v, "COUNTIF(" + rng.status + ",G" + i + ")", { formula: true });
    i++;
  }
  i = 3;
  for (const label of ["Out Of DLP", "Under DLP", "Not Specified"]) {
    styled(i, 10, label, "\u2014");             // DLP: dashes, not formulas
    i++;
  }
  i = body2;
  for (const v of phs) {
    styled(i, 1, v, "COUNTIF(" + rng.phase + ",A" + i + ")", { formula: true });
    i++;
  }
  i = body2;
  for (const mo of months) {
    mod.cell(i, 4, { serial: dateSerial(mo.y, mo.m, 1) }, A.MONTH_LABEL);
    mod.cell(i, 5,
      'COUNTIFS(' + rng.date + ',">="&DATE(' + mo.y + "," + mo.m + ",1)," +
      rng.date + ',"<"&EDATE(DATE(' + mo.y + "," + mo.m + ",1),1))",
      A.COUNT, { isFormula: true });
    i++;
  }

  mod.setColWidths([
    ["A", 26.0], ["B", 10.0], ["C", 3.0], ["D", 26.0], ["E", 10.0], ["F", 3.0],
    ["G", 26.0], ["H", 10.0], ["I", 3.0], ["J", 26.0], ["K", 10.0],
  ]);
  return mod;
}

/* ---------------------------------------------------------------------------
 * 7. Chart adaptation (1:1 port of the python surgery helpers)
 * ------------------------------------------------------------------------- */

function serials(months) {
  return months.map((mo) => dateSerial(mo.y, mo.m, 1));
}

function numCache(formatCode, values, ptCount) {
  let pts = "";
  for (let i = 0; i < values.length; i++) {
    if (values[i] !== null && values[i] !== undefined) {
      pts += '<c:pt idx="' + i + '"><c:v>' + xmlEscape(values[i]) + "</c:v></c:pt>";
    }
  }
  return '<c:numCache><c:formatCode>' + xmlEscape(formatCode) +
    '</c:formatCode><c:ptCount val="' + ptCount + '"/>' + pts + "</c:numCache>";
}

function strCache(values, ptCount) {
  let pts = "";
  for (let i = 0; i < values.length; i++) {
    if (values[i] !== null && values[i] !== undefined) {
      pts += '<c:pt idx="' + i + '"><c:v>' + xmlEscape(values[i]) + "</c:v></c:pt>";
    }
  }
  return '<c:strCache><c:ptCount val="' + ptCount + '"/>' + pts + "</c:strCache>";
}

function slotSpec(slot, stats) {
  const cats = stats.category_order;
  const subs = stats.subcat_chart;
  const stas = stats.status_table;
  const phs = stats.phase_order;
  const months = stats.months;
  const monthCounts = months.map((mo) => {
    const rec = stats.month.get(mo.y + ":" + mo.m);
    return rec ? rec.count : 0;
  });
  const sers = serials(months);
  const C = cats.length, SUB = subs.length, ST = stas.length,
    PH = phs.length, M = months.length;

  const ana = stats.analysis_rows || [12, 13, 14];
  const aHdr2 = ana[1], aBody2 = ana[2];

  const iCatRow = C + 2, iSubRow = SUB + 2, iStatRow = ST + 2, iPhaseRow = PH + 2;
  const catPairs = cats.map((c) => [c, stats.category.get(c) || 0]);
  const subPairs = subs.map((s) => [s, stats.subcat.get(s) || 0]);
  const staPairs = stas.map((s) => [s, stats.status.get(s, 0)]);
  const phaPairs = phs.map((p) => [p, stats.phase.get(p) || 0]);
  const labels = (pairs) => pairs.map((p) => p[0]);
  const vals = (pairs) => pairs.map((p) => p[1]);
  const padLabel = (pairs) => labels(pairs).concat([null]);

  switch (slot) {
    case 1:
      return {
        refs: ["Calc!$S$2:$S$" + (1 + M), "Calc!$T$2:$T$" + (1 + M)],
        caches: [numCache("mmm\\-yy", sers, M), numCache("General", monthCounts, M)],
      };
    case 2:
      return {
        refs: ["Calc!$A$2:$A$" + iCatRow, "Calc!$B$2:$B$" + iCatRow],
        caches: [strCache(padLabel(catPairs), C + 1),
                 numCache("General", vals(catPairs), C + 1)],
      };
    case 3:
      return {
        refs: ["Calc!$D$2:$D$" + iSubRow, "Calc!$E$2:$E$" + iSubRow],
        caches: [strCache(padLabel(subPairs), SUB + 1),
                 numCache("General", vals(subPairs), SUB + 1)],
      };
    case 4:
      // Doughnut spans the status table INCLUDING its trailing blank row;
      // Excel caches the range asymmetrically (strCache drops the blank
      // point, numCache keeps the empty idx-ST slot). Replicate exactly.
      return {
        refs: ["Calc!$G$2:$G$" + iStatRow, "Calc!$H$2:$H$" + iStatRow],
        caches: [strCache(labels(staPairs), ST),
                 numCache("General", vals(staPairs), ST + 1)],
      };
    case 5:
      return {
        refs: ["Calc!$P$2:$P$" + iPhaseRow, "Calc!$Q$2:$Q$" + iPhaseRow],
        caches: [strCache(padLabel(phaPairs), PH + 1),
                 numCache("General", vals(phaPairs), PH + 1)],
      };
    case 6:
      return {
        refs: ["Analysis!$E$" + aHdr2,
               "Analysis!$D$" + aBody2 + ":$D$" + (aHdr2 + M),
               "Analysis!$E$" + aBody2 + ":$E$" + (aHdr2 + M)],
        caches: [strCache(["Count"], 1), numCache("mmm\\-yy", sers, M),
                 numCache("General", monthCounts, M)],
      };
    case 7:
      return {
        refs: ["Analysis!$B$2", "Analysis!$A$3:$A$" + (2 + C), "Analysis!$B$3:$B$" + (2 + C)],
        caches: [strCache(["Count"], 1), strCache(labels(catPairs), C),
                 numCache("General", vals(catPairs), C)],
      };
    case 8:
      return {
        refs: ["Analysis!$E$2", "Analysis!$D$3:$D$" + (2 + SUB), "Analysis!$E$3:$E$" + (2 + SUB)],
        caches: [strCache(["Count"], 1), strCache(labels(subPairs), SUB),
                 numCache("General", vals(subPairs), SUB)],
      };
    case 9:
      return {
        refs: ["Analysis!$H$2", "Analysis!$G$3:$G$" + (2 + ST), "Analysis!$H$3:$H$" + (2 + ST)],
        caches: [strCache(["Count"], 1), strCache(labels(staPairs), ST),
                 numCache("General", vals(staPairs), ST)],
      };
    case 10:
      return {
        refs: ["Analysis!$B$" + aHdr2,
               "Analysis!$A$" + aBody2 + ":$A$" + (aHdr2 + PH),
               "Analysis!$B$" + aBody2 + ":$B$" + (aHdr2 + PH)],
        caches: [strCache(["Count"], 1), strCache(labels(phaPairs), PH),
                 numCache("General", vals(phaPairs), PH)],
      };
    default:
      throw new Error("unknown chart slot " + slot);
  }
}

function addValAxMinZero(xml) {
  const m = /<c:valAx>[\s\S]*?<\/c:valAx>/.exec(xml);
  if (!m) return xml;
  const block = m[0];
  const s = /<c:scaling>[\s\S]*?<\/c:scaling>/.exec(block);
  if (!s || s[0].indexOf("<c:min") !== -1) return xml;
  const newBlock = block.replace(
    s[0],
    s[0].slice(0, s[0].length - "</c:scaling>".length) + '<c:min val="0"/></c:scaling>');
  return xml.slice(0, m.index) + newBlock + xml.slice(m.index + m[0].length);
}

const F_PAT = /<c:f>[^<]*<\/c:f>/;
const CACHE_PAT = /<c:(numCache|strCache)>[\s\S]*?<\/c:\1>/;

function replaceBlocks(xml, pattern, replacements) {
  let n = 0;
  const re = new RegExp(pattern.source, "g");
  const out = xml.replace(re, () => replacements[n++] || "");
  if (n !== replacements.length) {
    throw new Error("block replace mismatch: " + n + " vs " + replacements.length);
  }
  return out;
}

function adaptTemplate(slot, refs, caches, templateText) {
  let xml = templateText;
  xml = replaceBlocks(xml, F_PAT, refs.map((r) => "<c:f>" + r + "</c:f>"));
  xml = replaceBlocks(xml, CACHE_PAT, caches);
  if (slot === 1 || slot === 6) xml = addValAxMinZero(xml);
  return xml;
}

function drawingRelsXml(targets) {
  const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
  let entries = "";
  for (let i = 0; i < targets.length; i++) {
    entries += '<Relationship Id="rId' + (i + 1) + '" Type="' + rel +
      '" Target="../charts/' + targets[i] + '"/>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/' +
    'relationships">' + entries + "</Relationships>";
}

/* ---------------------------------------------------------------------------
 * 8. Template provider (async; channeled in from the host environment)
 * ------------------------------------------------------------------------- */

function defaultTemplateProvider(kind) {
  if (typeof importScripts === "function" && typeof self !== "undefined") {
    const base = self.location.href.replace(/[^/]*$/, "");
    return fetch(base + "chart_templates/" + kind + ".xml").then((r) => r.text());
  }
  if (typeof require === "function") {
    const fs = require("fs");
    const path = require("path");
    const here = typeof __dirname !== "undefined" ? __dirname : ".";
    return new Promise((res, rej) =>
      fs.readFile(path.join(here, "chart_templates", kind + ".xml"), "utf8",
        (e, d) => (e ? rej(e) : res(d))));
  }
  return Promise.reject(new Error("no template provider for " + kind));
}

/* ---------------------------------------------------------------------------
 * 9. generate() — full pipeline, final zip in python's exact part order
 * ------------------------------------------------------------------------- */

async function generate(dumpBytes, options) {
  options = options || {};
  const templates = options.templates || defaultTemplateProvider;
  const dumpFileName = options.dumpFileName || "";

  const { rows, meta } = await readDumpRows(dumpBytes);
  if (!rows.length) throw new Error("No ticket rows found in the dump file.");

  const asOn = options.asOn || deriveAsOn(meta, rows, dumpFileName);
  const periodFrom = (meta && meta.from) || DEFAULT_PERIOD_FROM;
  const periodTo = asOn;

  const stats = aggregate(rows);

  const workbookXml = utf8Read(b64Decode(STATIC_B64.workbookTemplate))
    .replace("@@SUBCAT_RANGE@@",
      "Calc!$X$1:$X$" + (1 + stats.subcat_filter.slice(1).length));

  const sheets = [
    buildInteractive(stats, asOn),
    buildOverview(stats, asOn, periodFrom, periodTo),
    buildDataSheet(rows),
    buildAnalysis(stats),
    buildCalcSheet(stats),
  ];

  // Single-pass emission with ONE master sst (python's restore walks the
  // sheets in zip order — sheet1..sheet5 — and interns in document order).
  const master = new SstStore();
  const perSheet = sheets.map((mod) => renderSheet(mod, master).xml);
  // perSheet[k] = xml with master indices already (each sheet rendered in
  // order against the same store, exactly like the python restore pass).
  const xml1 = perSheet[0], xml2 = perSheet[1], xml3 = perSheet[2],
        xml4 = perSheet[3], xml5 = perSheet[4];

  const chartParts = [];
  for (let slot = 1; slot <= 10; slot++) {
    const spec = slotSpec(slot, stats);
    const tpl = await templates("chart" + slot);
    chartParts.push(adaptTemplate(slot, spec.refs, spec.caches, tpl));
  }
  const drawing1 = await templates("drawing1");
  const drawing2 = await templates("drawing2");
  const dr1Rels = drawingRelsXml(chartParts.slice(0, 5).map((_, i) => "chart" + (i + 1) + ".xml"));
  const dr2Rels = drawingRelsXml(chartParts.slice(5).map((_, i) => "chart" + (i + 6) + ".xml"));

  const coreNow = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const coreXml =
    "<cp:coreProperties" +
    ' xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/"' +
    ' xmlns:dcterms="http://purl.org/dc/terms/"' +
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    "<dc:creator>MyGate Analysis Tool</dc:creator>" +
    "<dc:title>Mygate Complaints Dashboard as on " + asOnLabel(asOn) + "</dc:title>" +
    '<dcterms:created xsi:type="dcterms:W3CDTF">' + coreNow + "</dcterms:created>" +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">' + coreNow + "</dcterms:modified>" +
    "<cp:lastModifiedBy>MyGate Analysis Tool</cp:lastModifiedBy>" +
    "</cp:coreProperties>";

  let contentTypes = utf8Read(b64Decode(STATIC_B64["[Content_Types].xml"]));
  contentTypes = contentTypes.replace(
    "</Types>",
    '<Override PartName="/xl/sharedStrings.xml" ' +
    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
    "</Types>");

  let wbRels = utf8Read(b64Decode(STATIC_B64["xl/_rels/workbook.xml.rels"]));
  const ids = (wbRels.match(/Id="rId(\d+)"/g) || []).map((s) => Number(/(\d+)/.exec(s)[1]));
  const rid = ids.length ? Math.max.apply(null, ids) + 1 : 1;
  wbRels = wbRels.replace(
    "</Relationships>",
    '<Relationship Id="rId' + rid + '" Type="http://schemas.openxmlformats' +
    '.org/officeDocument/2006/relationships/sharedStrings" ' +
    'Target="sharedStrings.xml"/></Relationships>');

  const parts = [
    ["docProps/app.xml", b64Decode(STATIC_B64["docProps/app.xml"])],
    ["docProps/core.xml", utf8(coreXml)],
    ["xl/theme/theme1.xml", b64Decode(STATIC_B64["xl/theme/theme1.xml"])],
    ["xl/worksheets/sheet1.xml", utf8(xml1)],
    ["xl/drawings/drawing1.xml", utf8(drawing1)],
    ["xl/drawings/_rels/drawing1.xml.rels", utf8(dr1Rels)],
    ["xl/worksheets/_rels/sheet1.xml.rels", b64Decode(STATIC_B64["xl/worksheets/_rels/sheet1.xml.rels"])],
    ["xl/worksheets/sheet2.xml", utf8(xml2)],
    ["xl/drawings/drawing2.xml", utf8(drawing2)],
    ["xl/drawings/_rels/drawing2.xml.rels", utf8(dr2Rels)],
    ["xl/worksheets/_rels/sheet2.xml.rels", b64Decode(STATIC_B64["xl/worksheets/_rels/sheet2.xml.rels"])],
    ["xl/worksheets/sheet3.xml", utf8(xml3)],
    ["xl/worksheets/sheet4.xml", utf8(xml4)],
    ["xl/worksheets/sheet5.xml", utf8(xml5)],
    ["xl/charts/chart1.xml", utf8(chartParts[0])],
    ["xl/charts/chart2.xml", utf8(chartParts[1])],
    ["xl/charts/chart3.xml", utf8(chartParts[2])],
    ["xl/charts/chart4.xml", utf8(chartParts[3])],
    ["xl/charts/chart5.xml", utf8(chartParts[4])],
    ["xl/charts/chart6.xml", utf8(chartParts[5])],
    ["xl/charts/chart7.xml", utf8(chartParts[6])],
    ["xl/charts/chart8.xml", utf8(chartParts[7])],
    ["xl/charts/chart9.xml", utf8(chartParts[8])],
    ["xl/charts/chart10.xml", utf8(chartParts[9])],
    ["xl/styles.xml", b64Decode(STATIC_B64["xl/styles.xml"])],
    ["_rels/.rels", b64Decode(STATIC_B64["_rels/.rels"])],
    ["xl/workbook.xml", utf8(workbookXml)],
    ["xl/_rels/workbook.xml.rels", utf8(wbRels)],
    ["[Content_Types].xml", utf8(contentTypes)],
    ["xl/sharedStrings.xml", utf8(master.doc())],
  ];

  const bytes = await buildZip(parts);
  const st = stats.status;
  const info = {
    filename: outputName(asOn),
    asOn,
    total: stats.total,
    counts: {
      "In Progress": st.get("In Progress", 0),
      "Reopened": st.get("Reopened", 0),
      "On Hold": st.get("On Hold", 0),
      "New": st.get("New", 0),
    },
    months: stats.months.length,
    topIssue: stats.subcat_chart[0] || null,
    topIssueCount: stats.subcat_chart[0]
      ? stats.subcat.get(stats.subcat_chart[0], 0) : 0,
  };
  return { bytes, info, rows, stats };
}

/* ---------------------------------------------------------------------------
 * 10. Export
 * ------------------------------------------------------------------------- */

const api = {
  // helpers (unit tests)
  collapseSpaces, derivePhase, parseCreatedDate, ordinal, monthSpan, fmtShort,
  asOnLabel, outputName, dateSerial, serialToDate, deriveAsOn, aggregate,
  slotSpec, strCache, numCache, adaptTemplate, addValAxMinZero,
  drawingRelsXml, colIndexOf, colLetter, xmlEscape, unescapeXml,
  readDumpRows, zipRead, crc32, buildZip,
  generate,
};

if (typeof self !== "undefined" && typeof importScripts === "function") {
  self.AnalysisEngine = api;
}
if (typeof module !== "undefined" && module.exports) {
  module.exports = api;
}
root.AnalysisEngine = api;

})(typeof globalThis !== "undefined" ? globalThis : this);
