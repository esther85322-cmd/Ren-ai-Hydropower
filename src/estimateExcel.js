// Reads a signed-estimate workbook (one sheet per trade: 電氣 / 給排水 / 弱電 …) into flat, editable rows.
// Everything here is a *suggestion*: the UI shows the rows for review before anything is saved.
import * as XLSX from "xlsx";

const SKIP_SHEET = /^(表頭|封面|總表|彙總)/;
const TOP_RE = /^[壹貳參肆伍陸柒捌玖拾]$/;
const SECTION_DOT_RE = /^[一二三四五六七八九十]+\s*[.、]$/;
const SECTION_PAREN_RE = /^[(（][一二三四五六七八九十]+[)）]$/;
const SECTION_BARE_RE = /^[一二三四五六七八九十]+$/;
const NUMBERED_RE = /^\d+$/;
const LETTER_RE = /^[a-zA-Z]$/;

const clean = (v) => String(v == null ? "" : v).replace(/[　\s]+/g, " ").trim();
const num = (v) => {
  if (v == null || v === "") return null;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : null;
};
const sheetLabel = (name) => name.replace(/^[一二三四五六七八九十]+\s*[.、]?\s*/, "");

// Category hints (matched against the site's own category names, so renamed categories still work).
const CATEGORY_RULES = [
  [/運\s*什\s*費|雜費|其他費用/, ["其他", "雜"]],
  [/工資|安裝費|施工費|人工|組立|試車|測試/, ["人工"]],
  [/馬桶|洗臉|面盆|浴缸|淋浴|龍頭|水龍頭|衛浴|蓮蓬頭|小便斗|淨水|濾水|熱水器|防霧鏡/, ["衛浴"]],
  [/電線|電纜|導線|網路線|同軸|光纖|光纜|控制線|耐燃線|耐熱線|\bWIRE\b|\bCABLE\b|THHN|XLPE|\bRG\d|\bCAT\d/i, ["電線", "線"]],
  [/\bCASE\b|配電箱|開關箱|電表箱|箱體/i, ["箱體", "箱"]],
  [/螺絲|螺栓|五金|吊筋|膨脹|夾具|支撐/, ["五金"]],
];
const findCat = (categories, hints) => {
  for (const h of hints) {
    const c = categories.find((x) => (x.name || "").includes(h));
    if (c) return c.id;
  }
  return "";
};
const WATER_WORDS = /管|閥|凡而|止水|逆止|泵|馬達|幫浦|水栓|消防栓|水塔|水箱|水錶|水表|落水頭|濾|法蘭|陰井|排水|給水/;
// Names don't always say what they are, so fall back on which trade (sheet) the line came from.
export function guessCategory(name, categories, sheetLabel = "") {
  const other = findCat(categories, ["其他", "雜"]) || categories[0]?.id || "";
  for (const [re, hints] of CATEGORY_RULES) {
    if (re.test(name)) return findCat(categories, hints) || other;
  }
  if (/給排水/.test(sheetLabel)) return findCat(categories, ["水", "管"]) || other;
  if (/消防/.test(sheetLabel)) return findCat(categories, ["消防"]) || other;
  if (/公共設備/.test(sheetLabel)) return WATER_WORDS.test(name) ? findCat(categories, ["水", "管"]) || other : other;
  return findCat(categories, ["配電", "電"]) || other;
}

let seq = 0;

/** @returns {{sheets: {name, label, count, noPrice, total}[], rows: object[]}} */
export function parseEstimateWorkbook(arrayBuffer, categories) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheets = [];
  const rows = [];
  for (const sheetName of wb.SheetNames) {
    if (SKIP_SHEET.test(sheetName)) continue;
    const ws = wb.Sheets[sheetName];
    const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const label = sheetLabel(sheetName);
    let sections = [];
    let itemHeading = "";
    let count = 0;
    let noPrice = 0;
    let total = 0;
    for (let i = 3; i < grid.length; i++) {
      const r = grid[i] || [];
      const code = clean(r[0]);
      const name = clean(r[1]);
      if (!name) continue;
      const unit = clean(r[2]);
      const qty = num(r[3]);
      const price = num(r[4]);
      const amt = num(r[5]);
      const brand = clean(r[6]);
      const isLeaf = unit !== "" && qty != null && qty > 0;
      if (!isLeaf) {
        if (TOP_RE.test(code)) { sections = []; itemHeading = ""; }
        else if (SECTION_DOT_RE.test(code)) { sections = [name]; itemHeading = ""; }
        else if (SECTION_PAREN_RE.test(code) || SECTION_BARE_RE.test(code)) { sections = [sections[0], name].filter(Boolean); itemHeading = ""; }
        else if (NUMBERED_RE.test(code)) { itemHeading = name; }
        continue;
      }
      // "a / b / c" sub-items only make sense together with their numbered heading ("高籠型不銹鋼落水頭 φ3"…).
      const full = LETTER_RE.test(code) && itemHeading ? `${itemHeading} ${name}` : name;
      const lineTotal = amt != null ? amt : price != null ? price * qty : null;
      const unitPrice = lineTotal != null ? Math.round((lineTotal / qty) * 10000) / 10000 : 0;
      const hasPrice = lineTotal != null && lineTotal > 0;
      const noteParts = [label, ...sections].filter(Boolean).join(" › ");
      rows.push({
        id: "x" + ++seq,
        sheet: sheetName,
        name: full,
        unit,
        quantity: qty,
        unitPrice,
        hasPrice,
        categoryId: guessCategory(full, categories, label),
        note: brand ? `${noteParts}（${brand}）` : noteParts,
        include: true,
      });
      count++;
      if (!hasPrice) noPrice++;
      else total += lineTotal;
    }
    if (count) sheets.push({ name: sheetName, label, count, noPrice, total });
  }
  return { sheets, rows };
}
