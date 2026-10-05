// ==================== TEMPLATE PATCHER ====================
// Writes consolidated rows into the customer's template .xlsx by editing only
// the worksheet XML of the sheets we own ("<Plant> Daily" + "Analyze").
//
// Why not ExcelJS?  Re-serialising the whole workbook through ExcelJS silently
// drops pivot tables / pivot caches / comments and rewrites every style.  The
// template relies on pivot tables ("Summary", "By Plant", "Analyze"), so we
// patch the zip in place instead: everything we do not touch survives 1:1.
//
// Rules:
//   - cells that contain a formula are never overwritten or cleared
//   - a written cell keeps the style of the cell it replaces (or of the first
//     data row in the template) and gets a light-blue fill + blue border
//   - rows below the data are cleared (values only, styles kept)
//   - dimension / autoFilter / _FilterDatabase / pivot source ranges are
//     extended so they cover the new rows
//   - fullCalcOnLoad + refreshOnLoad are set so Excel recalculates formulas
//     and refreshes every pivot table on open
import JSZip from 'jszip';
import { getColumnIndex, getColumnLetter, safeSheetName } from '../../utils/format.js';
import { ANALYZE_SHEET, TEMPLATE } from '../../constants.js';

// ---------- XML helpers ----------
const ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };
const escapeXml = (s) =>
  String(s)
    .replace(/[&<>"]/g, (c) => ESCAPES[c])
    // control characters are not allowed in XML 1.0
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
const decodeXml = (s) =>
  String(s)
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(Number(d)))
    .replace(/&amp;/g, '&');
const attr = (attrs, name) => attrs.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
// String.replace() with a function so "$&"-style patterns in formulas are not interpreted
const spliceAt = (xml, match, replacement) =>
  xml.slice(0, match.index) + replacement + xml.slice(match.index + match[0].length);

const ROW_RE = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;

// ---------- workbook parts ----------
const readText = async (zip, path) => {
  const f = zip.file(path);
  return f ? f.async('string') : null;
};
// write without adding synthetic folder entries to the archive
const writeText = (zip, path, content) => zip.file(path, content, { createFolders: false });

const resolveTarget = (target) => (target.startsWith('/') ? target.slice(1) : `xl/${target}`);

// [{ name, rId, path, index }] in workbook order
export const listSheets = async (zip) => {
  const wb = await readText(zip, 'xl/workbook.xml');
  const rels = await readText(zip, 'xl/_rels/workbook.xml.rels');
  if (!wb || !rels) throw new Error('ไม่พบ xl/workbook.xml (ไฟล์หลักต้องเป็น .xlsx)');
  const relMap = {};
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = attr(m[0], 'Id');
    const target = attr(m[0], 'Target');
    if (id && target) relMap[id] = resolveTarget(target);
  }
  return [...wb.matchAll(/<sheet\b[^>]*>/g)].map((m, index) => {
    const rId = attr(m[0], 'r:id');
    return { name: decodeXml(attr(m[0], 'name') ?? ''), rId, path: relMap[rId], index, hidden: /state="hidden"/.test(m[0]) };
  });
};

// ArrayBuffer -> sheet names (cheap: only workbook.xml is parsed)
export const listSheetNames = async (buffer) => {
  const zip = await JSZip.loadAsync(buffer);
  return (await listSheets(zip)).map((s) => s.name);
};

const loadSharedStrings = async (zip) => {
  const xml = await readText(zip, 'xl/sharedStrings.xml');
  if (!xml) return [];
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) =>
    [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decodeXml(t[1])).join(''),
  );
};

// ---------- sheet parsing ----------
const parseCell = (match) => {
  const attrs = match[1] || '';
  const inner = match[2] || '';
  const ref = attr(attrs, 'r');
  if (!ref) return null;
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  if (!m) return null;
  return {
    col: getColumnIndex(m[1]),
    xml: match[0],
    inner,
    s: attr(attrs, 's'),
    t: attr(attrs, 't'),
    hasFormula: /<f\b/.test(inner),
  };
};

const parseRow = (match) => {
  const attrs = match[1] || '';
  const inner = match[2] || '';
  const cells = new Map();
  for (const c of inner.matchAll(CELL_RE)) {
    const cell = parseCell(c);
    if (cell) cells.set(cell.col, cell);
  }
  return { r: Number(attr(attrs, 'r')), attrs, cells, xml: match[0] };
};

export const parseSheet = (xml) => {
  const sheetData = xml.match(/<sheetData\b[^>]*?(?:\/>|>([\s\S]*?)<\/sheetData>)/);
  if (!sheetData) throw new Error('ไม่พบ <sheetData>');
  const rows = new Map();
  for (const m of (sheetData[1] || '').matchAll(ROW_RE)) {
    const row = parseRow(m);
    if (Number.isFinite(row.r)) rows.set(row.r, row);
  }
  return { sheetData, rows, lastRow: Math.max(0, ...rows.keys()) };
};

const cellText = (cell, sst) => {
  if (!cell) return '';
  if (cell.t === 'inlineStr') {
    return [...cell.inner.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1])).join('');
  }
  const v = cell.inner.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  if (v === undefined) return '';
  return cell.t === 's' ? sst[Number(v)] ?? '' : decodeXml(v);
};

const norm = (s) => String(s ?? '').trim().toUpperCase();

// First row (1-based) whose cell in `col` reads `text`.
const findRowWithText = (rows, sst, col, text, limit = 80) => {
  for (const row of [...rows.values()].sort((a, b) => a.r - b.r)) {
    if (row.r > limit) break;
    if (norm(cellText(row.cells.get(col), sst)) === text) return row.r;
  }
  return null;
};

// Right-most column in a row whose cell reads `text`.
const lastColWithText = (row, sst, text) => {
  let found = null;
  if (!row) return null;
  for (const [col, cell] of row.cells) {
    if (norm(cellText(cell, sst)) === text && (found === null || col > found)) found = col;
  }
  return found;
};

// ---------- cell building ----------
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const buildCell = (col, row, value, s) => {
  const ref = `${getColumnLetter(col)}${row}`;
  const sAttr = s !== undefined && s !== null && s !== '' ? ` s="${s}"` : '';
  if (value === null || value === undefined || value === '') return `<c r="${ref}"${sAttr}/>`;
  if (typeof value === 'number') {
    return Number.isFinite(value) ? `<c r="${ref}"${sAttr}><v>${value}</v></c>` : `<c r="${ref}"${sAttr}/>`;
  }
  if (typeof value === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${value ? 1 : 0}</v></c>`;
  if (value instanceof Date) {
    return `<c r="${ref}"${sAttr}><v>${(value.getTime() - EXCEL_EPOCH) / 86400000}</v></c>`;
  }
  const text = String(value);
  const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : '';
  return `<c r="${ref}"${sAttr} t="inlineStr"><is><t${space}>${escapeXml(text)}</t></is></c>`;
};

// ---------- styles.xml: derived "highlighted" cell formats ----------
export class StyleRegistry {
  constructor(xml, { fill = 'FFE6F7FF', border = 'FF1E40AF' } = {}) {
    this.xml = xml;
    this.fillRgb = fill;
    this.borderRgb = border;
    this.map = new Map();
    this.fillId = null;
    this.borderId = null;
    this.xfs = null;
    this.newXfs = [];
  }

  ensureBase() {
    if (this.fillId !== null) return;
    if (!this.xml) throw new Error('ไม่พบ xl/styles.xml');
    const fills = this.xml.match(/<fills count="(\d+)">([\s\S]*?)<\/fills>/);
    const borders = this.xml.match(/<borders count="(\d+)">([\s\S]*?)<\/borders>/);
    if (!fills || !borders) throw new Error('styles.xml: ไม่พบ <fills>/<borders>');
    this.fillId = Number(fills[1]);
    this.xml = spliceAt(this.xml, fills,
      `<fills count="${this.fillId + 1}">${fills[2]}<fill><patternFill patternType="solid"><fgColor rgb="${this.fillRgb}"/><bgColor indexed="64"/></patternFill></fill></fills>`);
    // re-match because indexes shifted
    const borders2 = this.xml.match(/<borders count="(\d+)">([\s\S]*?)<\/borders>/);
    this.borderId = Number(borders2[1]);
    const side = (name) => `<${name} style="thin"><color rgb="${this.borderRgb}"/></${name}>`;
    this.xml = spliceAt(this.xml, borders2,
      `<borders count="${this.borderId + 1}">${borders2[2]}<border>${side('left')}${side('right')}${side('top')}${side('bottom')}<diagonal/></border></borders>`);
    const xfs = this.xml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
    if (!xfs) throw new Error('styles.xml: ไม่พบ <cellXfs>');
    this.xfs = [...xfs[2].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]);
  }

  // index of a cellXf identical to `base` but with the highlight fill + border
  highlighted(base) {
    this.ensureBase();
    const b = Number(base) || 0;
    if (this.map.has(b)) return this.map.get(b);
    const setAttr = (xml, name, val) =>
      new RegExp(`\\b${name}=`).test(xml)
        ? xml.replace(new RegExp(`\\b${name}="[^"]*"`), `${name}="${val}"`)
        : xml.replace(/^<xf\b/, `<xf ${name}="${val}"`);
    let xf = this.xfs[b] || this.xfs[0] || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    xf = setAttr(xf, 'fillId', this.fillId);
    xf = setAttr(xf, 'borderId', this.borderId);
    xf = setAttr(xf, 'applyFill', 1);
    xf = setAttr(xf, 'applyBorder', 1);
    const idx = this.xfs.length + this.newXfs.length;
    this.newXfs.push(xf);
    this.map.set(b, idx);
    return idx;
  }

  // returns the new styles.xml, or null when nothing changed
  commit() {
    if (!this.newXfs.length) return null;
    const xfs = this.xml.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
    return spliceAt(this.xml, xfs,
      `<cellXfs count="${this.xfs.length + this.newXfs.length}">${xfs[2]}${this.newXfs.join('')}</cellXfs>`);
  }
}

// ---------- the per-sheet patch ----------
/**
 * @param {string} xml            worksheet xml
 * @param {object} parsed         result of parseSheet(xml)
 * @param {object} opts
 * @param {number} opts.startRow  first data row (1-based)
 * @param {Map<number, any>[]} opts.dataRows  one Map(col -> value) per output row
 * @param {number} opts.clearThroughRow  rows startRow..clearThroughRow lose their values
 * @param {StyleRegistry} opts.styles
 * @param {boolean} opts.highlight
 */
export const patchSheetXml = (xml, parsed, { startRow, dataRows, clearThroughRow, styles, highlight }) => {
  const { sheetData, rows } = parsed;
  const templateRow = rows.get(startRow);
  const lastDataRow = startRow + dataRows.length - 1;
  const endClear = Math.max(clearThroughRow ?? 0, lastDataRow);
  let maxCol = 0;
  let written = 0;

  for (let r = startRow; r <= endClear; r++) {
    const existing = rows.get(r);
    const data = dataRows[r - startRow];
    if (!existing && !data) continue;

    const cells = new Map();
    if (existing) {
      for (const [col, cell] of existing.cells) {
        // formulas are sacred; everything else is cleared (style kept)
        cells.set(col, cell.hasFormula ? cell.xml : buildCell(col, r, null, cell.s));
      }
    }
    if (data) {
      for (const [col, value] of data) {
        const ex = existing?.cells.get(col);
        if (ex?.hasFormula) continue;
        if (value === null || value === undefined || value === '') continue;
        let s = ex?.s ?? templateRow?.cells.get(col)?.s;
        if (highlight && styles) s = styles.highlighted(s);
        cells.set(col, buildCell(col, r, value, s));
        if (col > maxCol) maxCol = col;
        written++;
      }
    }

    const cols = [...cells.keys()].sort((a, b) => a - b);
    let attrs = existing?.attrs ?? templateRow?.attrs ?? '';
    attrs = /\br="\d+"/.test(attrs) ? attrs.replace(/\br="\d+"/, `r="${r}"`) : ` r="${r}"${attrs}`;
    if (cols.length && /\bspans="[^"]*"/.test(attrs)) {
      attrs = attrs.replace(/\bspans="[^"]*"/, `spans="${cols[0]}:${cols[cols.length - 1]}"`);
    }
    const rowXml = cols.length ? `<row${attrs}>${cols.map((c) => cells.get(c)).join('')}</row>` : `<row${attrs}/>`;
    rows.set(r, { r, attrs, cells: existing?.cells ?? new Map(), xml: rowXml });
  }

  const body = [...rows.values()].sort((a, b) => a.r - b.r).map((row) => row.xml).join('');
  let out = spliceAt(xml, sheetData, `<sheetData>${body}</sheetData>`);

  // <dimension ref="A1:EL26"/>
  const dim = out.match(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\/>/);
  if (dim) {
    const endRow = Math.max(Number(dim[4]), lastDataRow, 1);
    const endCol = Math.max(getColumnIndex(dim[3]), maxCol, 1);
    out = spliceAt(out, dim, `<dimension ref="${dim[1]}${dim[2]}:${getColumnLetter(endCol)}${endRow}"/>`);
  }
  // <autoFilter ref="A13:EF21" .../>
  const af = out.match(/<autoFilter ref="([A-Z]+\d+):([A-Z]+)(\d+)"/);
  if (af && lastDataRow >= startRow) {
    out = spliceAt(out, af, `<autoFilter ref="${af[1]}:${af[2]}${Math.max(Number(af[3]), lastDataRow)}"`);
  }

  return { xml: out, lastDataRow, maxCol, written };
};

// rawRow (array, index 0 = source column A) -> Map(col -> value)
const rowToMap = (rawRow, offset, maxCols = Infinity) => {
  const map = new Map();
  if (!Array.isArray(rawRow)) return map;
  const limit = Math.min(rawRow.length, maxCols);
  for (let i = 0; i < limit; i++) {
    const v = rawRow[i];
    if (v !== undefined && v !== null && v !== '') map.set(i + offset, v);
  }
  return map;
};

// ---------- workbook-level range fix-ups ----------
const extendDefinedNames = async (zip, patched) => {
  let wb = await readText(zip, 'xl/workbook.xml');
  if (!wb) return;
  for (const { index, lastDataRow } of patched) {
    const re = new RegExp(`(<definedName name="_xlnm\\._FilterDatabase" localSheetId="${index}"[^>]*>)([^<]*)(</definedName>)`);
    wb = wb.replace(re, (m, open, range, close) => {
      const rm = range.match(/^(.*\$[A-Z]+\$)(\d+)$/);
      if (!rm) return m;
      return `${open}${rm[1]}${Math.max(Number(rm[2]), lastDataRow)}${close}`;
    });
  }
  writeText(zip, 'xl/workbook.xml', wb);
};

const extendPivotSources = async (zip, patched) => {
  const byName = new Map(patched.map((p) => [p.name, p]));
  const paths = Object.keys(zip.files).filter((p) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(p));
  for (const path of paths) {
    let xml = await readText(zip, path);
    xml = xml.replace(/<worksheetSource\b[^>]*\/>/g, (tag) => {
      const sheet = decodeXml(attr(tag, 'sheet') ?? '');
      const ref = attr(tag, 'ref');
      const p = byName.get(sheet);
      if (!p || !ref) return tag;
      const rm = ref.match(/^([A-Z]+\d+:[A-Z]+)(\d+)$/);
      if (!rm) return tag;
      return tag.replace(/\bref="[^"]*"/, `ref="${rm[1]}${Math.max(Number(rm[2]), p.lastDataRow)}"`);
    });
    writeText(zip, path, xml);
  }
};

// Force Excel to recalc every formula and refresh every pivot cache on open.
export const applyRecalcFlags = async (zip) => {
  let wb = await readText(zip, 'xl/workbook.xml');
  if (wb) {
    const calcPr = wb.match(/<calcPr\b([^>]*?)(\/?)>/);
    if (calcPr) {
      let attrs = calcPr[1];
      attrs = /fullCalcOnLoad=/.test(attrs)
        ? attrs.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"')
        : `${attrs} fullCalcOnLoad="1"`;
      wb = spliceAt(wb, calcPr, `<calcPr${attrs}${calcPr[2]}>`);
    } else {
      wb = wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    }
    writeText(zip, 'xl/workbook.xml', wb);
  }
  const paths = Object.keys(zip.files).filter((p) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(p));
  for (const path of paths) {
    let xml = await readText(zip, path);
    xml = xml.replace(/<pivotCacheDefinition\b([^>]*)>/, (m, attrs) =>
      /refreshOnLoad=/.test(attrs)
        ? `<pivotCacheDefinition${attrs.replace(/refreshOnLoad="[^"]*"/, 'refreshOnLoad="1"')}>`
        : `<pivotCacheDefinition${attrs} refreshOnLoad="1">`);
    writeText(zip, path, xml);
  }
};

// Backwards-compatible helper: ArrayBuffer in, patched ArrayBuffer out.
export const forceRecalcOnOpen = async (buffer) => {
  const zip = await JSZip.loadAsync(buffer);
  await applyRecalcFlags(zip);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
};


// ---------- extra sheets (QA report, changes, ...) ----------
const WORKSHEET_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const WORKSHEET_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

const aoaToDataRows = (aoa) =>
  (aoa || []).map((row) => {
    const map = new Map();
    (Array.isArray(row) ? row : []).forEach((v, i) => {
      if (v !== undefined && v !== null && v !== '') map.set(i + 1, v);
    });
    return map;
  });

const blankSheetXml = () =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<dimension ref="A1:A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
  '<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';

// best-effort: keep docProps/app.xml's worksheet list in sync (cosmetic, Excel does not validate it)
const registerInAppProps = async (zip, name) => {
  const app = await readText(zip, 'docProps/app.xml');
  if (!app) return;
  const count = app.match(/<vt:lpstr>Worksheets<\/vt:lpstr><\/vt:variant><vt:variant><vt:i4>(\d+)<\/vt:i4>/);
  const titles = app.match(/<TitlesOfParts><vt:vector size="(\d+)" baseType="lpstr">([\s\S]*?)<\/vt:vector><\/TitlesOfParts>/);
  if (!count || !titles) return;
  const n = Number(count[1]);
  const entries = [...titles[2].matchAll(/<vt:lpstr>[\s\S]*?<\/vt:lpstr>/g)].map((m) => m[0]);
  if (entries.length < n) return;
  entries.splice(n, 0, `<vt:lpstr>${escapeXml(name)}</vt:lpstr>`);
  let out = spliceAt(app, titles, `<TitlesOfParts><vt:vector size="${entries.length}" baseType="lpstr">${entries.join('')}</vt:vector></TitlesOfParts>`);
  const count2 = out.match(/<vt:lpstr>Worksheets<\/vt:lpstr><\/vt:variant><vt:variant><vt:i4>(\d+)<\/vt:i4>/);
  out = spliceAt(out, count2, `<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${n + 1}</vt:i4>`);
  writeText(zip, 'docProps/app.xml', out);
};

/**
 * Create (or fully replace the contents of) a plain values-only sheet.
 * Existing sheets keep their formulas/formatting rules (patchSheetXml semantics).
 * @returns {{ name: string, created: boolean, rows: number }}
 */
export const upsertSheet = async (zip, rawName, aoa) => {
  const name = safeSheetName(rawName);
  const sheets = await listSheets(zip);
  let sheet = sheets.find((s) => s.name === name);
  let created = false;

  if (!sheet) {
    created = true;
    const used = Object.keys(zip.files).map((p) => Number(p.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)?.[1] || 0));
    const num = Math.max(0, ...used) + 1;
    const path = `xl/worksheets/sheet${num}.xml`;
    let wb = await readText(zip, 'xl/workbook.xml');
    let rels = await readText(zip, 'xl/_rels/workbook.xml.rels');
    let ct = await readText(zip, '[Content_Types].xml');
    if (!wb || !rels || !ct) throw new Error('ไม่พบส่วนประกอบหลักของไฟล์ xlsx');
    const sheetId = Math.max(0, ...[...wb.matchAll(/sheetId="(\d+)"/g)].map((m) => Number(m[1]))) + 1;
    const rId = `rId${Math.max(0, ...[...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]))) + 1}`;
    wb = wb.replace('</sheets>', () => `<sheet name="${escapeXml(name)}" sheetId="${sheetId}" r:id="${rId}"/></sheets>`);
    rels = rels.replace('</Relationships>', () => `<Relationship Id="${rId}" Type="${WORKSHEET_REL}" Target="worksheets/sheet${num}.xml"/></Relationships>`);
    if (!ct.includes(`PartName="/${path}"`)) {
      ct = ct.replace('</Types>', () => `<Override PartName="/${path}" ContentType="${WORKSHEET_CT}"/></Types>`);
    }
    writeText(zip, 'xl/workbook.xml', wb);
    writeText(zip, 'xl/_rels/workbook.xml.rels', rels);
    writeText(zip, '[Content_Types].xml', ct);
    writeText(zip, path, blankSheetXml());
    await registerInAppProps(zip, name);
    sheet = { name, path, index: sheets.length };
  }

  const xml = await readText(zip, sheet.path);
  const parsed = parseSheet(xml);
  const dataRows = aoaToDataRows(aoa);
  const res = patchSheetXml(xml, parsed, { startRow: 1, dataRows, clearThroughRow: parsed.lastRow, styles: null, highlight: false });
  writeText(zip, sheet.path, res.xml);
  return { name, created, rows: dataRows.length, cells: res.written };
};

// ---------- entry point ----------
/**
 * @param {ArrayBuffer} templateBuffer  the customer's .xlsx
 * @param {object} opts
 * @param {Record<string, object[]>} opts.processedData  { 'BP Daily': rows, ... }
 * @param {object[]} opts.analyzeRows  [{ plant, rawRow, ... }] already sorted
 * @param {boolean} [opts.highlight=true]
 * @param {{ name: string, aoa: any[][] }[]} [opts.extraSheets]  values-only sheets to add/replace
 * @returns {Promise<{ buffer: ArrayBuffer, report: object }>}
 */
export const patchTemplate = async (templateBuffer, { processedData, analyzeRows = [], highlight = true, extraSheets = [] }) => {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheets = await listSheets(zip);
  const sst = await loadSharedStrings(zip);
  const styles = new StyleRegistry(await readText(zip, 'xl/styles.xml'));
  const report = { patched: [], skipped: [], warnings: [] };
  const patched = [];

  // 1) "<Plant> Daily" sheets: 1:1 copy of the source rows from the header row + 1
  for (const [sheetName, rows] of Object.entries(processedData || {})) {
    const sheet = sheets.find((s) => s.name === sheetName);
    if (!sheet?.path || !zip.file(sheet.path)) {
      report.skipped.push(sheetName);
      continue;
    }
    const xml = await readText(zip, sheet.path);
    const parsed = parseSheet(xml);
    const headerRow = findRowWithText(parsed.rows, sst, 1, 'PART NUMBER') ?? TEMPLATE.dailyHeaderRow;
    const startRow = headerRow + 1;
    const dataRows = rows.map((row) => rowToMap(row.rawRow, 1));
    const res = patchSheetXml(xml, parsed, { startRow, dataRows, clearThroughRow: parsed.lastRow, styles, highlight });
    writeText(zip, sheet.path, res.xml);
    patched.push({ ...sheet, lastDataRow: res.lastDataRow });
    report.patched.push({ sheet: sheetName, rows: rows.length, startRow, cells: res.written });
  }

  // 2) "Analyze": column B.. = source columns, plant code in the column after N+3
  const analyze = sheets.find((s) => s.name === ANALYZE_SHEET);
  if (analyze?.path && zip.file(analyze.path)) {
    const xml = await readText(zip, analyze.path);
    const parsed = parseSheet(xml);
    const offset = TEMPLATE.analyzeSourceOffset;
    const headerRow = findRowWithText(parsed.rows, sst, offset, 'PART NUMBER') ?? TEMPLATE.analyzeHeaderRow;
    const startRow = headerRow + 1;
    const n3Col = lastColWithText(parsed.rows.get(headerRow), sst, 'N+3');
    const plantCol = n3Col ? n3Col + 1 : TEMPLATE.analyzePlantCol;
    const maxSourceCols = plantCol - offset;
    const af = xml.match(/<autoFilter ref="[A-Z]+\d+:[A-Z]+(\d+)"/);
    const clearThroughRow = af ? Number(af[1]) : startRow + 300;
    const capacity = clearThroughRow - startRow + 1;
    if (analyzeRows.length > capacity) {
      report.warnings.push(
        `Analyze: ข้อมูล ${analyzeRows.length} แถว มากกว่าพื้นที่ตารางในเทมเพลต (${capacity} แถว) ` +
        'แถวที่เกินจะไม่มีสูตร key (คอลัมน์ A) ให้ขยายสูตรในเทมเพลตก่อนใช้งาน',
      );
    }
    const dataRows = analyzeRows.map((row) => {
      const map = rowToMap(row.rawRow, offset, maxSourceCols);
      map.set(plantCol, row.plant);
      return map;
    });
    const res = patchSheetXml(xml, parsed, { startRow, dataRows, clearThroughRow, styles, highlight });
    writeText(zip, analyze.path, res.xml);
    patched.push({ ...analyze, lastDataRow: res.lastDataRow });
    report.patched.push({ sheet: ANALYZE_SHEET, rows: analyzeRows.length, startRow, plantCol: getColumnLetter(plantCol), cells: res.written });
  } else {
    report.skipped.push(ANALYZE_SHEET);
  }

  // 3) extra values-only sheets (QA report, change list, ...)
  for (const extra of extraSheets) {
    if (!extra?.name) continue;
    const res = await upsertSheet(zip, extra.name, extra.aoa || []);
    report.patched.push({ sheet: res.name, rows: res.rows, startRow: 1, cells: res.cells, created: res.created });
  }

  // 4) styles, ranges, recalc flags
  const newStyles = styles.commit();
  if (newStyles) writeText(zip, 'xl/styles.xml', newStyles);
  await extendDefinedNames(zip, patched);
  await extendPivotSources(zip, patched);
  await applyRecalcFlags(zip);

  const buffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
  return { buffer, report };
};
