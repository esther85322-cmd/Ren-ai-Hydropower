// Reads a 付款比例表 (payment-ratio schedule) workbook into rows for 甲方樓層請款項目:
//   - a floor × milestone matrix (each milestone has a "%" and a "金額" column), and
//   - extra milestone rows below the 小計 line (name, %, 金額; two groups side by side).
// The result is only a suggestion: the UI previews it before saving.
import * as XLSX from "xlsx";

const clean = (v) => String(v == null ? "" : v).replace(/[\s　]+/g, " ").trim();
const isNum = (v) => typeof v === "number" && Number.isFinite(v);
const round1 = (n) => Math.round(n * 100) / 100;

export function parsePaymentSchedule(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const g = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const title = clean(g[0] && g[0].find((v) => typeof v === "string"));

  // header: the row whose first cell is 樓別
  const head = g.findIndex((r) => /樓\s*別/.test(clean(r && r[0])));
  if (head < 0) throw new Error("no header");
  const names = [];
  (g[head] || []).forEach((v, c) => {
    if (c === 0 || typeof v !== "string" || !clean(v)) return;
    const sub = clean(g[head + 1] && g[head + 1][c]);
    names.push({ col: c, name: clean(v) + sub });
  });
  if (!names.length) throw new Error("no milestones");

  const items = [];
  let r = head + 1;
  // skip the sub-title / "%｜金額" header rows
  while (r < g.length && !clean(g[r] && g[r][0]) ) r++;
  for (; r < g.length; r++) {
    const row = g[r] || [];
    const label = clean(row[0]);
    if (!label) continue;
    if (/^小計/.test(label)) { r++; break; }
    names.forEach((m) => {
      const pct = row[m.col];
      const amt = row[m.col + 1];
      if (isNum(pct) && pct !== 0) items.push({ itemName: m.name, floor: label, percent: round1(pct), amount: isNum(amt) ? amt : 0 });
    });
  }

  // extra milestones below 小計: "name … % 金額 [name … % 金額]"
  let total = 0;
  for (; r < g.length; r++) {
    const row = g[r] || [];
    if (/承攬總價|總價/.test(clean(row[0]))) {
      const nums = row.filter(isNum);
      total = nums.find((n) => n > 100000) || 0;
      break;
    }
    let cur = null;
    row.forEach((v) => {
      if (typeof v === "string" && clean(v) && !/^[：:％%（(]/.test(clean(v))) {
        cur = { itemName: clean(v), nums: [] };
        items.push(cur);
      } else if (isNum(v) && cur) cur.nums.push(v);
    });
  }
  const extra = [];
  for (let i = items.length - 1; i >= 0; i--) {
    if (items[i].nums) {
      const it = items.splice(i, 1)[0];
      if (it.nums.length >= 1) extra.unshift({ itemName: it.itemName, floor: "", percent: round1(it.nums[0]), amount: it.nums[1] || 0 });
    }
  }
  items.push(...extra);

  if (!total) {
    const top = (g[0] || []).filter(isNum).find((n) => n > 100000);
    total = top || 0;
  }
  const pctSum = items.reduce((s, it) => s + it.percent, 0);
  return { title, total, items, pctSum: round1(pctSum), amountSum: items.reduce((s, it) => s + it.amount, 0) };
}
