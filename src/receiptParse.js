// Turns raw OCR text of a delivery slip / price list into editable rows.
// Everything here is a *suggestion*: the UI always shows the rows for review before saving.

const UNITS = [
  "公尺", "pcs", "PCS", "Pcs", "m2", "M2", "kg", "KG", "cm", "EA", "ea",
  "支", "米", "個", "顆", "組", "卷", "只", "條", "片", "袋", "箱", "包", "桶", "盒", "台", "才", "式",
  "付", "副", "把", "座", "捲", "罐", "瓶", "塊", "張", "根", "件", "對", "套", "盤", "坪", "呎", "尺", "M", "m",
];
const SKIP_RE = /(品名|品項|規格|數量|單價|金額|合計|總計|小計|稅額|統一編號|統編|電話|傳真|地址|日期|簽收|收貨人|送貨|客戶|備註|營業稅|未稅|含稅|應收|訂單|編號|聯|負責|製單|經辦|匯款|帳號|銀行)/;
const CJK = /[㐀-鿿]/;

export function toHalfWidth(s) {
  return String(s || "")
    .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .replace(/　/g, " ");
}

export function cleanLine(raw) {
  return toHalfWidth(raw)
    .replace(/(?<=[㐀-鿿])[ \t]+(?=[㐀-鿿])/g, "")
    .replace(/[|｜¦]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// Keep only digits, latin letters, '.', and CJK; unify multiplication signs. Used for fuzzy comparison.
export function normKey(s) {
  return toHalfWidth(s)
    .toLowerCase()
    .replace(/[×✕✖*]/g, "x")
    .replace(/[^0-9a-z.x㐀-鿿]/g, "");
}

function bigrams(s) {
  const set = new Set();
  if (s.length < 2) { if (s) set.add(s); return set; }
  for (let i = 0; i < s.length - 1; i++) set.add(s.slice(i, i + 2));
  return set;
}

function matchCatalog(lineNorm, catalog) {
  const lineGrams = bigrams(lineNorm);
  let best = null;
  for (const item of catalog) {
    const nameNorm = normKey(item.name);
    if (!nameNorm) continue;
    const grams = bigrams(nameNorm);
    let hit = 0;
    grams.forEach((g) => { if (lineGrams.has(g)) hit++; });
    const recall = hit / grams.size;
    const need = grams.size <= 2 ? 1 : 0.75;
    if (recall < need) continue;
    const score = recall + Math.min(grams.size, 12) / 100;
    if (!best || score > best.score) best = { item, score, recall };
  }
  return best;
}

function unitAt(text, idx) {
  const rest = text.slice(idx).trimStart();
  return UNITS.find((u) => rest.startsWith(u)) || "";
}

function numberTokens(line) {
  const out = [];
  const re = /\d+(?:\.\d+)?/g;
  let m;
  while ((m = re.exec(line.replace(/(\d),(\d{3})/g, "$1$2")))) {
    out.push({ text: m[0], value: parseFloat(m[0]), end: m.index + m[0].length });
  }
  return out;
}

function pickQuantity(tokens, line) {
  if (!tokens.length) return { quantity: "", unit: "", price: "" };
  const withUnit = tokens.find((t) => t.unit && !/^(m|M)$/.test(t.unit) && t.value > 0);
  if (withUnit) {
    const after = tokens.slice(tokens.indexOf(withUnit) + 1);
    return { quantity: withUnit.value, unit: withUnit.unit, price: after.length ? after[0].value : "" };
  }
  for (let i = 0; i < tokens.length; i++) {
    for (let j = i + 1; j < tokens.length; j++) {
      for (let l = j + 1; l < tokens.length; l++) {
        const prod = tokens[i].value * tokens[j].value;
        if (prod > 0 && Math.abs(prod - tokens[l].value) <= Math.max(1, tokens[l].value * 0.01)) {
          return { quantity: tokens[i].value, unit: tokens[i].unit || "", price: tokens[j].value };
        }
      }
    }
  }
  const first = tokens[0];
  return { quantity: first.value, unit: first.unit || "", price: tokens.length > 1 ? tokens[1].value : "" };
}

// Tokens of the line that are *not* part of the (matched) catalog item name.
function tokensOutsideName(line, itemName) {
  const nameNums = (normKey(itemName).match(/\d+(?:\.\d+)?/g) || []).slice();
  const tokens = numberTokens(line);
  const rest = [];
  const normalized = line.replace(/(\d),(\d{3})/g, "$1$2");
  for (const t of tokens) {
    const k = nameNums.indexOf(t.text);
    if (k === 0 || (k > 0 && rest.length === 0)) { nameNums.splice(k, 1); continue; }
    rest.push({ ...t, unit: unitAt(normalized, t.end) });
  }
  return rest;
}

// Splits "name  qty unit price amount" into the name part and the numeric tail.
function splitNameAndTail(line) {
  const chunks = line.split(" ");
  const numRe = new RegExp("^[$NT]*\\d[\\d.,]*(?:" + UNITS.join("|") + ")?$");
  const numericish = (c) => numRe.test(c) || UNITS.includes(c);
  let list = chunks;
  if (list.length > 2 && /^\d{1,3}[.)、]?$/.test(list[0])) list = list.slice(1);
  const k = list.findIndex((c, i) => i >= 1 && numericish(c));
  if (k < 0) return { name: list.join(" "), tail: "" };
  return { name: list.slice(0, k).join(" "), tail: list.slice(k).join(" ") };
}

let rowSeq = 0;
const nextId = () => "r" + ++rowSeq;

/** Delivery slip lines -> order rows. */
export function parseReceiptText(text, catalog) {
  const rows = [];
  const lines = String(text || "").split("\n").map(cleanLine).filter(Boolean);
  for (const line of lines) {
    const lineNorm = normKey(line);
    if (lineNorm.length < 2 || !(CJK.test(line) || /[a-zA-Z]{2,}/.test(line))) continue;
    const match = matchCatalog(lineNorm, catalog);
    if (!match && SKIP_RE.test(line)) continue;
    let tokens;
    let name;
    if (match) {
      name = match.item.name;
      tokens = tokensOutsideName(line, name);
    } else {
      const parts = splitNameAndTail(line);
      if (!parts.tail) continue; // no quantity-looking tail: ignore non-item chatter
      name = parts.name;
      const normalizedTail = parts.tail.replace(/(\d),(\d{3})/g, "$1$2");
      tokens = numberTokens(parts.tail).map((t) => ({ ...t, unit: unitAt(normalizedTail, t.end) }));
      if (!tokens.length) continue;
    }
    const pick = pickQuantity(tokens, line);
    const price = pick.price !== "" ? pick.price : (match && match.item.defaultPrice !== "" && match.item.defaultPrice != null ? Number(match.item.defaultPrice) : "");
    rows.push({
      id: nextId(),
      raw: line,
      name,
      matchedId: match ? match.item.id : "",
      categoryId: match ? match.item.categoryId || "" : "",
      quantity: pick.quantity > 100000 ? "" : pick.quantity,
      unit: pick.unit || (match ? match.item.unit || "" : ""),
      unitPrice: price,
    });
  }
  return rows;
}

/** Price list / catalog photo lines -> new catalog item rows. */
export function parseItemListText(text, catalog) {
  const rows = [];
  const seen = new Set();
  const lines = String(text || "").split("\n").map(cleanLine).filter(Boolean);
  for (const line of lines) {
    const lineNorm = normKey(line);
    if (lineNorm.length < 2 || !(CJK.test(line) || /[a-zA-Z]{2,}/.test(line))) continue;
    if (SKIP_RE.test(line)) continue;
    const { name, tail } = splitNameAndTail(line);
    const key = normKey(name);
    if (key.length < 2 || seen.has(key)) continue;
    seen.add(key);
    const normalizedTail = tail.replace(/(\d),(\d{3})/g, "$1$2");
    const tokens = numberTokens(tail).map((t) => ({ ...t, unit: unitAt(normalizedTail, t.end) }));
    const unitFromTokens = tokens.find((t) => t.unit);
    let standalone = tail.split(" ").find((c) => UNITS.includes(c)) || "";
    let finalName = name;
    if (!unitFromTokens && !standalone && tail) {
      const last = name.slice(-1);
      if (name.length > 2 && UNITS.includes(last)) { standalone = last; finalName = name.slice(0, -1); }
    }
    const priceToken = tokens.length ? tokens[tokens.length - 1] : null;
    const exists = catalog.some((c) => normKey(c.name) === key);
    rows.push({
      id: nextId(),
      raw: line,
      name: finalName,
      unit: (unitFromTokens && unitFromTokens.unit) || standalone || "",
      defaultPrice: priceToken && !priceToken.unit ? priceToken.value : "",
      exists,
    });
  }
  return rows;
}
