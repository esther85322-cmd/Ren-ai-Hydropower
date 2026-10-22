import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Cell
} from "recharts";
import {
  Wrench, Zap, Package, TrendingUp, TrendingDown, Plus, Trash2, AlertTriangle,
  Gauge, LayoutDashboard, ClipboardList, Boxes, Tags, Loader2, X, Droplets,
  Users, CalendarRange, HardHat, Building2, Layers, Handshake, Banknote,
  ChevronDown, ChevronRight, Copy, ListChecks, Wallet, ListOrdered, UserCheck,
  MessageSquare, CheckCircle2, Circle, FileSpreadsheet, FileText, PiggyBank, Hammer, Receipt, Pencil, NotebookPen, Paperclip, Upload, Download, Camera, ScanLine, ImagePlus, ChevronUp, Workflow, Calculator, Scale
} from "lucide-react";
import { storageGet, storageSet, storageGetAll, fileStore, fileLoad, fileRemove } from "./storage";
import { parseReceiptText, parseItemListText, suggestNames } from "./receiptParse";
import { lockApp } from "./PasswordGate.jsx";

const DEFAULT_CATEGORIES = [
  { id: "c1", name: "給水排水管路", color: "#4A90A4" },
  { id: "c2", name: "電線電纜", color: "#E8A33D" },
  { id: "c3", name: "配電材料", color: "#C97B4A" },
  { id: "c4", name: "衛浴設備", color: "#6B9B6E" },
  { id: "c5", name: "五金零件", color: "#8B95A1" },
  { id: "c6", name: "人工費用", color: "#C1543C" },
  { id: "c7", name: "其他", color: "#7A8A99" },
];

const DEFAULT_SITE = { id: "site-default", name: "預設案場", address: "" };

const WAGE_OPTIONS = [];
for (let v = 1500; v <= 3500; v += 100) WAGE_OPTIONS.push(v);

const uid = () => Math.random().toString(36).slice(2, 10);
const todayStr = () => new Date().toISOString().slice(0, 10);
const monthKey = (d) => (d || "").slice(0, 7);
const fmtMoney = (n) =>
  "NT$ " + Math.round(Number(n) || 0).toLocaleString("zh-TW");
const fmtNum = (n) => {
  const v = Number(n) || 0;
  return Number.isInteger(v) ? v.toLocaleString("zh-TW") : v.toLocaleString("zh-TW", { maximumFractionDigits: 2 });
};

// ---- contract file helpers ----
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
const fmtSize = (n) => (n >= 1024 * 1024 ? (n / 1024 / 1024).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB");
async function prepareUploadFile(file) {
  if (!file.type.startsWith("image/") || file.size <= 600 * 1024) return file;
  try {
    const bitmap = await createImageBitmap(file);
    const scale = Math.min(1, 2000 / Math.max(bitmap.width, bitmap.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(bitmap.width * scale);
    canvas.height = Math.round(bitmap.height * scale);
    canvas.getContext("2d").drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.8));
    return blob ? new File([blob], file.name.replace(/\.[^.]+$/, "") + ".jpg", { type: "image/jpeg" }) : file;
  } catch (e) {
    return file;
  }
}
const blobToBase64 = (blob) => new Promise((resolve, reject) => {
  const reader = new FileReader();
  reader.onload = () => resolve(String(reader.result).split(",")[1] || "");
  reader.onerror = reject;
  reader.readAsDataURL(blob);
});
function base64ToBlob(b64, type) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type });
}

// ---- Excel / Word export helpers ----
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// sheets: [{ name, rows }] where rows is an array of plain objects (keys become column headers)
function exportExcel(filename, sheets) {
  try {
    const wb = XLSX.utils.book_new();
    sheets.forEach(({ name, rows }) => {
      const data = rows && rows.length ? rows : [{ "（無資料）": "" }];
      const ws = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, (name || "工作表").slice(0, 31));
    });
    XLSX.writeFile(wb, filename.endsWith(".xlsx") ? filename : `${filename}.xlsx`);
    return true;
  } catch (e) {
    console.error("excel export failed", e);
    return false;
  }
}

// sections: [{ title, headers, rows }] where rows is an array of arrays (matching headers order)
function buildWordHtml(docTitle, sections) {
  const esc = (v) => String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const sectionsHtml = sections.map((sec) => {
    const rowsHtml = (sec.rows || []).map((r) => `<tr>${r.map((c) => `<td>${esc(c).replace(/\r?\n/g, "<br>")}</td>`).join("")}</tr>`).join("");
    const headHtml = sec.headers ? `<tr>${sec.headers.map((h) => `<th>${esc(h)}</th>`).join("")}</tr>` : "";
    const table = sec.headers || (sec.rows && sec.rows.length)
      ? `<table>${headHtml}${rowsHtml || `<tr><td colspan="${(sec.headers || []).length || 1}">（無資料）</td></tr>`}</table>`
      : "";
    return `<h2>${esc(sec.title)}</h2>${sec.note ? `<p class="note">${esc(sec.note)}</p>` : ""}${table}`;
  }).join("");
  return `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset="utf-8"><title>${esc(docTitle)}</title>
<style>
  body { font-family: 'Microsoft JhengHei', 'PMingLiU', Arial, sans-serif; font-size: 11pt; color: #222; }
  h1 { font-size: 18pt; margin-bottom: 4pt; }
  .subtitle { color: #666; font-size: 10pt; margin-bottom: 18pt; }
  h2 { font-size: 13pt; margin-top: 20pt; margin-bottom: 6pt; border-bottom: 1pt solid #999; padding-bottom: 3pt; }
  p.note { color: #666; font-size: 9.5pt; margin: 2pt 0 6pt; }
  table { border-collapse: collapse; width: 100%; margin-bottom: 10pt; }
  th, td { border: 1pt solid #999; padding: 4pt 7pt; font-size: 10pt; text-align: left; }
  th { background: #eee; font-weight: bold; }
</style></head>
<body>
<h1>${esc(docTitle)}</h1>
<div class="subtitle">匯出時間：${new Date().toLocaleString("zh-TW")}</div>
${sectionsHtml}
</body></html>`;
}
function exportWord(filename, docTitle, sections) {
  try {
    const html = buildWordHtml(docTitle, sections);
    const blob = new Blob(["\ufeff", html], { type: "application/msword" });
    triggerDownload(blob, filename.endsWith(".doc") ? filename : `${filename}.doc`);
    return true;
  } catch (e) {
    console.error("word export failed", e);
    return false;
  }
}

export default function WaterElectricLedger() {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [toast, setToast] = useState(null);
  const toastTimerRef = useRef(null);
  const showToast = useCallback((message) => {
    setToast(message);
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => setToast(null), 3200);
  }, []);
  const [tab, setTab] = useState("labor");
  const [projectName, setProjectName] = useState("水電工程記帳");
  const [editingName, setEditingName] = useState(false);
  const [categories, setCategories] = useState(DEFAULT_CATEGORIES);
  const [materialItems, setMaterialItems] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [orders, setOrders] = useState([]);
  const [usages, setUsages] = useState([]);
  const [monthFilter, setMonthFilter] = useState("all");
  const [workers, setWorkers] = useState([]);
  const [workLogs, setWorkLogs] = useState([]);
  const [laborRange, setLaborRange] = useState({ start: "", end: "" });
  const [sites, setSites] = useState([DEFAULT_SITE]);
  const [currentSiteId, setCurrentSiteId] = useState(DEFAULT_SITE.id);
  const [view, setView] = useState("sites"); // "sites" (landing page) | "workspace"
  const [pendingDrill, setPendingDrill] = useState(null);
  const enterSite = useCallback((siteId, drill) => {
    setCurrentSiteId(siteId);
    setTab(siteId === "all" ? "dashboard" : "labor");
    setPendingDrill(drill || null);
    setView("workspace");
  }, []);
  const returnToSites = useCallback(() => setView("sites"), []);
  const [contracts, setContracts] = useState([]);
  const contractsRef = useRef(contracts);
  contractsRef.current = contracts;
  const [payments, setPayments] = useState([]);
  const [floorItems, setFloorItems] = useState([]);
  const [clientPayments, setClientPayments] = useState([]);
  const [clientTemplateRows, setClientTemplateRows] = useState([]);
  const [clientFloorItems, setClientFloorItems] = useState([]);
  const [contractWorkLogs, setContractWorkLogs] = useState([]);
  const [contractItems, setContractItems] = useState([]);
  const [discussionItems, setDiscussionItems] = useState([]);
  const [notebookItems, setNotebookItems] = useState([]);
  const [workflow, setWorkflow] = useState([]);
  const [estimates, setEstimates] = useState([]);
  const [estimateFilesBySite, setEstimateFilesBySite] = useState({});
  const estimateFilesRef = useRef(estimateFilesBySite);
  estimateFilesRef.current = estimateFilesBySite;
  const [receipts, setReceipts] = useState([]);
  const receiptsRef = useRef(receipts);
  receiptsRef.current = receipts;
  const [capitalContributions, setCapitalContributions] = useState([]);
  const [fixedAssets, setFixedAssets] = useState([]);
  const [otherExpenses, setOtherExpenses] = useState([]);

  // ---- load ----
  useEffect(() => {
    (async () => {
      const { data, ok } = await storageGetAll();
      const meta = data.meta ?? { projectName: "水電工程記帳", categories: DEFAULT_CATEGORIES };
      const ord = data.orders ?? [];
      const use = data.usages ?? [];
      const wk = data.workers ?? [];
      const wl = data.worklogs ?? [];
      const st = data.sites ?? [DEFAULT_SITE];
      const ct = data.contracts ?? [];
      const pay = data.payments ?? [];
      const fi = data.flooritems ?? [];
      const cp = data.clientpayments ?? [];
      const ctr = data.clienttemplate ?? [];
      const cfi = data.clientflooritems ?? [];
      const cwl = data.contractworklogs ?? [];
      const di = data.discussionitems ?? [];
      const nb = data.notebookitems ?? [];
      const rc = data.receipts ?? [];
      const wf = data.workflow ?? [];
      const est = data.estimates ?? [];
      const estf = data.estimatefiles ?? {};
      const ci = data.contractitems ?? [];
      const mi = data.materialitems ?? [];
      const sup = data.suppliers ?? [];
      const cc = data.capitalcontributions ?? [];
      const fa = data.fixedassets ?? [];
      const oe = data.otherexpenses ?? [];
      const siteList = st && st.length ? st : [DEFAULT_SITE];
      const fallbackSiteId = siteList[0].id;
      // migrate legacy records created before 案場 (site) existed
      const migOrd = ord.map((o) => (o.siteId ? o : { ...o, siteId: fallbackSiteId }));
      const migUse = use.map((u) => (u.siteId ? u : { ...u, siteId: fallbackSiteId }));
      // migrate legacy work logs created before each day had its own 日薪 snapshot:
      // freeze in the worker's rate *as it is right now*, so this backfill doesn't
      // change anyone's already-recorded wage, but future rate edits no longer
      // reach back and alter it either.
      const workerRateById = {};
      wk.forEach((w) => (workerRateById[w.id] = Number(w.dailyRate) || 0));
      const migWl0 = wl.map((l) => (l.siteId ? l : { ...l, siteId: fallbackSiteId }));
      const migWl = migWl0.map((l) => (l.dailyRate != null ? l : { ...l, dailyRate: workerRateById[l.workerId] ?? 0 }));
      setProjectName(meta.projectName || "水電工程記帳");
      setCategories(meta.categories && meta.categories.length ? meta.categories : DEFAULT_CATEGORIES);
      setSites(siteList);
      setCurrentSiteId(fallbackSiteId);
      setOrders(migOrd);
      setUsages(migUse);
      setWorkers(wk);
      setWorkLogs(migWl);
      setContracts(ct.map((c) => (c.templateItems ? c : { ...c, templateItems: [] })));
      setPayments(pay);
      setFloorItems(fi);
      setClientPayments(cp);
      setClientTemplateRows(ctr);
      setClientFloorItems(cfi);
      setContractWorkLogs(cwl);
      setDiscussionItems(di);
      setNotebookItems(nb);
      setReceipts(rc);
      setWorkflow(wf);
      setEstimates(est);
      setEstimateFilesBySite(estf);
      setContractItems(ci);
      setMaterialItems(mi);
      setSuppliers(sup);
      setCapitalContributions(cc);
      setFixedAssets(fa);
      setOtherExpenses(oe);
      setLoading(false);
      if (!ok) {
        // The read itself failed (network/auth hiccup) — do NOT run any of the
        // auto-migration writes below. An empty/partial `data` here means
        // "we don't know what's in the database", not "the database is empty",
        // and writing defaults in that case would overwrite real data.
        showToast("資料讀取失敗，請重新整理頁面再試一次");
        return;
      }
      if (JSON.stringify(migOrd) !== JSON.stringify(ord)) storageSet("orders", migOrd);
      if (JSON.stringify(migUse) !== JSON.stringify(use)) storageSet("usages", migUse);
      if (JSON.stringify(migWl) !== JSON.stringify(wl)) storageSet("worklogs", migWl);
      if (!st || !st.length) storageSet("sites", siteList);
    })();
  }, []);

  const persistMeta = useCallback(async (name, cats) => {
    setSaving(true);
    await storageSet("meta", { projectName: name, categories: cats });
    setSaving(false);
  }, []);
  const persistOrders = useCallback(async (list) => {
    setSaving(true);
    await storageSet("orders", list);
    setSaving(false);
  }, []);
  const persistUsages = useCallback(async (list) => {
    setSaving(true);
    await storageSet("usages", list);
    setSaving(false);
  }, []);
  const persistWorkers = useCallback(async (list) => {
    setSaving(true);
    await storageSet("workers", list);
    setSaving(false);
  }, []);
  const persistWorkLogs = useCallback(async (list) => {
    setSaving(true);
    await storageSet("worklogs", list);
    setSaving(false);
  }, []);
  const persistSites = useCallback(async (list) => {
    setSaving(true);
    await storageSet("sites", list);
    setSaving(false);
  }, []);
  const debounceTimersRef = useRef({});
  const debouncedPersist = useCallback((key, persistFn, list) => {
    if (debounceTimersRef.current[key]) clearTimeout(debounceTimersRef.current[key]);
    debounceTimersRef.current[key] = setTimeout(() => {
      persistFn(list);
    }, 700);
  }, []);
  const persistContracts = useCallback(async (list) => {
    setSaving(true);
    await storageSet("contracts", list);
    setSaving(false);
  }, []);
  const persistPayments = useCallback(async (list) => {
    setSaving(true);
    await storageSet("payments", list);
    setSaving(false);
  }, []);
  const persistFloorItems = useCallback(async (list) => {
    setSaving(true);
    await storageSet("flooritems", list);
    setSaving(false);
  }, []);
  const persistClientPayments = useCallback(async (list) => {
    setSaving(true);
    await storageSet("clientpayments", list);
    setSaving(false);
  }, []);
  const persistClientTemplateRows = useCallback(async (list) => {
    setSaving(true);
    await storageSet("clienttemplate", list);
    setSaving(false);
  }, []);
  const persistClientFloorItems = useCallback(async (list) => {
    setSaving(true);
    await storageSet("clientflooritems", list);
    setSaving(false);
  }, []);
  const persistContractWorkLogs = useCallback(async (list) => {
    setSaving(true);
    await storageSet("contractworklogs", list);
    setSaving(false);
  }, []);
  const persistDiscussionItems = useCallback(async (list) => {
    setSaving(true);
    await storageSet("discussionitems", list);
    setSaving(false);
  }, []);
  const persistEstimates = useCallback(async (list) => {
    setSaving(true);
    await storageSet("estimates", list);
    setSaving(false);
  }, []);
  const persistEstimateFiles = useCallback(async (map) => {
    setSaving(true);
    await storageSet("estimatefiles", map);
    setSaving(false);
  }, []);
  const persistWorkflow = useCallback(async (list) => {
    setSaving(true);
    await storageSet("workflow", list);
    setSaving(false);
  }, []);
  const saveWorkflow = (next) => {
    setWorkflow(next);
    persistWorkflow(next);
  };
  const persistReceipts = useCallback(async (list) => {
    setSaving(true);
    await storageSet("receipts", list);
    setSaving(false);
  }, []);
  const persistNotebookItems = useCallback(async (list) => {
    setSaving(true);
    await storageSet("notebookitems", list);
    setSaving(false);
  }, []);
  const persistCapitalContributions = useCallback(async (list) => {
    setSaving(true);
    await storageSet("capitalcontributions", list);
    setSaving(false);
  }, []);
  const persistFixedAssets = useCallback(async (list) => {
    setSaving(true);
    await storageSet("fixedassets", list);
    setSaving(false);
  }, []);
  const persistOtherExpenses = useCallback(async (list) => {
    setSaving(true);
    await storageSet("otherexpenses", list);
    setSaving(false);
  }, []);
  const persistContractItems = useCallback(async (list) => {
    setSaving(true);
    await storageSet("contractitems", list);
    setSaving(false);
  }, []);
  const persistMaterialItems = useCallback(async (list) => {
    setSaving(true);
    await storageSet("materialitems", list);
    setSaving(false);
  }, []);
  const persistSuppliers = useCallback(async (list) => {
    setSaving(true);
    await storageSet("suppliers", list);
    setSaving(false);
  }, []);

  const catById = useMemo(() => {
    const m = {};
    categories.forEach((c) => (m[c.id] = c));
    return m;
  }, [categories]);

  const siteById = useMemo(() => {
    const m = {};
    sites.forEach((s) => (m[s.id] = s));
    return m;
  }, [sites]);

  const isAllSites = currentSiteId === "all";
  const siteOrders = useMemo(
    () => (isAllSites ? orders : orders.filter((o) => o.siteId === currentSiteId)),
    [orders, currentSiteId, isAllSites]
  );
  const siteUsages = useMemo(
    () => (isAllSites ? usages : usages.filter((u) => u.siteId === currentSiteId)),
    [usages, currentSiteId, isAllSites]
  );
  const siteWorkLogsAll = useMemo(
    () => (isAllSites ? workLogs : workLogs.filter((l) => l.siteId === currentSiteId)),
    [workLogs, currentSiteId, isAllSites]
  );
  const siteContracts = useMemo(
    () => (isAllSites ? contracts : contracts.filter((c) => c.siteId === currentSiteId)),
    [contracts, currentSiteId, isAllSites]
  );
  const sitePayments = useMemo(
    () => (isAllSites ? payments : payments.filter((p) => p.siteId === currentSiteId)),
    [payments, currentSiteId, isAllSites]
  );
  const siteFloorItems = useMemo(
    () => (isAllSites ? floorItems : floorItems.filter((f) => f.siteId === currentSiteId)),
    [floorItems, currentSiteId, isAllSites]
  );
  const siteClientPayments = useMemo(
    () => (isAllSites ? clientPayments : clientPayments.filter((p) => p.siteId === currentSiteId)),
    [clientPayments, currentSiteId, isAllSites]
  );
  const clientPaidTotal = useMemo(
    () => siteClientPayments.reduce((s, p) => s + (Number(p.amount) || 0), 0),
    [siteClientPayments]
  );
  const siteClientTemplateRows = useMemo(
    () => clientTemplateRows.filter((r) => r.siteId === currentSiteId),
    [clientTemplateRows, currentSiteId]
  );
  const siteClientFloorItems = useMemo(
    () => (isAllSites ? clientFloorItems : clientFloorItems.filter((f) => f.siteId === currentSiteId)),
    [clientFloorItems, currentSiteId, isAllSites]
  );
  const clientFloorTotal = useMemo(
    () => siteClientFloorItems.reduce((s, f) => s + (Number(f.amount) || 0), 0),
    [siteClientFloorItems]
  );
  const siteContractWorkLogs = useMemo(
    () => (isAllSites ? contractWorkLogs : contractWorkLogs.filter((l) => l.siteId === currentSiteId)),
    [contractWorkLogs, currentSiteId, isAllSites]
  );
  const siteContractItems = useMemo(
    () => (isAllSites ? contractItems : contractItems.filter((it) => it.siteId === currentSiteId)),
    [contractItems, currentSiteId, isAllSites]
  );
  const siteDiscussionItems = useMemo(
    () => (isAllSites ? discussionItems : discussionItems.filter((d) => d.siteId === currentSiteId)),
    [discussionItems, currentSiteId, isAllSites]
  );
  const openDiscussionCount = useMemo(
    () => siteDiscussionItems.filter((d) => !d.resolved).length,
    [siteDiscussionItems]
  );
  const siteEstimates = useMemo(
    () => (isAllSites ? estimates : estimates.filter((l) => l.siteId === currentSiteId)),
    [estimates, currentSiteId, isAllSites]
  );
  const siteReceipts = useMemo(
    () => (isAllSites ? receipts : receipts.filter((r) => r.siteId === currentSiteId)),
    [receipts, currentSiteId, isAllSites]
  );
  const siteNotebookItems = useMemo(
    () => (isAllSites ? notebookItems : notebookItems.filter((n) => n.siteId === currentSiteId)),
    [notebookItems, currentSiteId, isAllSites]
  );
  const siteCapitalContributions = useMemo(
    () => (isAllSites ? capitalContributions : capitalContributions.filter((c) => c.siteId === currentSiteId)),
    [capitalContributions, currentSiteId, isAllSites]
  );
  const capitalTotal = useMemo(
    () => siteCapitalContributions.reduce((s, c) => s + (Number(c.amount) || 0), 0),
    [siteCapitalContributions]
  );
  const siteFixedAssets = useMemo(
    () => (isAllSites ? fixedAssets : fixedAssets.filter((f) => f.siteId === currentSiteId)),
    [fixedAssets, currentSiteId, isAllSites]
  );
  const fixedAssetsTotal = useMemo(
    () => siteFixedAssets.reduce((s, f) => s + (Number(f.amount) || 0) * (Number(f.quantity) || 1), 0),
    [siteFixedAssets]
  );
  const attendanceByContract = useMemo(() => {
    const map = {};
    contractWorkLogs.forEach((l) => {
      if (!map[l.contractId]) map[l.contractId] = { days: 0, manDays: 0 };
      map[l.contractId].days += 1;
      map[l.contractId].manDays += Number(l.headcount) || 0;
    });
    return map;
  }, [contractWorkLogs]);
  const floorTotalByContract = useMemo(() => {
    const map = {};
    floorItems.forEach((f) => {
      map[f.contractId] = (map[f.contractId] || 0) + (Number(f.amount) || 0);
    });
    return map;
  }, [floorItems]);
  const itemsTotalByContract = useMemo(() => {
    const map = {};
    contractItems.forEach((it) => {
      map[it.contractId] = (map[it.contractId] || 0) + (Number(it.amount) || 0);
    });
    return map;
  }, [contractItems]);

  const months = useMemo(() => {
    const s = new Set(siteOrders.map((o) => monthKey(o.date)));
    return Array.from(s).filter(Boolean).sort().reverse();
  }, [siteOrders]);

  const filteredOrders = useMemo(
    () => (monthFilter === "all" ? siteOrders : siteOrders.filter((o) => monthKey(o.date) === monthFilter)),
    [siteOrders, monthFilter]
  );

  const totalCost = useMemo(() => siteOrders.reduce((s, o) => s + Number(o.amount || 0), 0), [siteOrders]);
  const thisMonthCost = useMemo(() => {
    const mk = monthKey(todayStr());
    return siteOrders.filter((o) => monthKey(o.date) === mk).reduce((s, o) => s + Number(o.amount || 0), 0);
  }, [siteOrders]);

  const categoryBreakdown = useMemo(() => {
    const map = {};
    filteredOrders.forEach((o) => {
      map[o.categoryId] = (map[o.categoryId] || 0) + Number(o.amount || 0);
    });
    return categories
      .map((c) => ({ id: c.id, name: c.name, color: c.color, value: map[c.id] || 0 }))
      .filter((d) => d.value > 0)
      .sort((a, b) => b.value - a.value);
  }, [filteredOrders, categories]);

  const monthlyTrend = useMemo(() => {
    const map = {};
    siteOrders.forEach((o) => {
      const mk = monthKey(o.date);
      if (!mk) return;
      map[mk] = (map[mk] || 0) + Number(o.amount || 0);
    });
    return Object.entries(map)
      .sort(([a], [b]) => (a > b ? 1 : -1))
      .map(([month, total]) => ({ month, total }));
  }, [siteOrders]);

  const inventory = useMemo(() => {
    const map = {};
    siteOrders.forEach((o) => {
      const key = o.categoryId + "|" + o.itemName;
      if (!map[key])
        map[key] = { key, itemName: o.itemName, categoryId: o.categoryId, unit: o.unit, ordered: 0, used: 0, cost: 0 };
      map[key].ordered += Number(o.quantity) || 0;
      map[key].cost += Number(o.amount) || 0;
    });
    siteUsages.forEach((u) => {
      const key = u.categoryId + "|" + u.itemName;
      if (!map[key])
        map[key] = { key, itemName: u.itemName, categoryId: u.categoryId, unit: u.unit, ordered: 0, used: 0, cost: 0 };
      map[key].used += (u.type === "return" ? -1 : 1) * (Number(u.quantity) || 0);
    });
    return Object.values(map)
      .map((it) => {
        const avg = it.ordered ? it.cost / it.ordered : 0;
        return { ...it, remaining: it.ordered - it.used, avgPrice: avg, value: (it.ordered - it.used) * avg };
      })
      .sort((a, b) => a.remaining - b.remaining);
  }, [siteOrders, siteUsages]);

  const workerById = useMemo(() => {
    const m = {};
    workers.forEach((w) => (m[w.id] = w));
    return m;
  }, [workers]);

  const workLogRate = useCallback(
    (l) => Number(l.dailyRate ?? workerById[l.workerId]?.dailyRate) || 0,
    [workerById]
  );

  const laborTotalAllTime = useMemo(
    () => siteWorkLogsAll.reduce((s, l) => s + (Number(l.days) || 0) * workLogRate(l), 0),
    [siteWorkLogsAll, workLogRate]
  );

  const rangeFilteredLogs = useMemo(() => {
    return siteWorkLogsAll.filter((l) => {
      if (laborRange.start && l.date < laborRange.start) return false;
      if (laborRange.end && l.date > laborRange.end) return false;
      return true;
    });
  }, [siteWorkLogsAll, laborRange]);

  const laborSummaryByWorker = useMemo(() => {
    const map = {};
    rangeFilteredLogs.forEach((l) => {
      if (!map[l.workerId]) map[l.workerId] = { workerId: l.workerId, days: 0, wage: 0 };
      map[l.workerId].days += Number(l.days) || 0;
      map[l.workerId].wage += (Number(l.days) || 0) * workLogRate(l);
    });
    return workers
      .map((w) => map[w.id] || { workerId: w.id, days: 0, wage: 0 })
      .sort((a, b) => b.wage - a.wage);
  }, [rangeFilteredLogs, workers, workLogRate]);

  const laborRangeTotal = useMemo(
    () => laborSummaryByWorker.reduce((s, r) => s + r.wage, 0),
    [laborSummaryByWorker]
  );
  const laborRangeDays = useMemo(
    () => laborSummaryByWorker.reduce((s, r) => s + r.days, 0),
    [laborSummaryByWorker]
  );

  const grandTotal = totalCost + laborTotalAllTime;

  const contractById = useMemo(() => {
    const m = {};
    contracts.forEach((c) => (m[c.id] = c));
    return m;
  }, [contracts]);

  const paidByContract = useMemo(() => {
    const map = {};
    payments.forEach((p) => {
      map[p.contractId] = (map[p.contractId] || 0) + (Number(p.amount) || 0);
    });
    return map;
  }, [payments]);

  const contractRows = useMemo(() => {
    return siteContracts
      .map((c) => {
        const paid = paidByContract[c.id] || 0;
        const att = attendanceByContract[c.id] || { days: 0, manDays: 0 };
        return {
          ...c, paid, remaining: (Number(c.totalPrice) || 0) - paid,
          floorTotal: floorTotalByContract[c.id] || 0,
          itemsTotal: itemsTotalByContract[c.id] || 0,
          attendanceDays: att.days, attendanceManDays: att.manDays,
        };
      })
      .sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [siteContracts, paidByContract, floorTotalByContract, itemsTotalByContract, attendanceByContract]);

  const contractPaidTotal = useMemo(
    () => sitePayments.reduce((s, p) => s + (Number(p.amount) || 0), 0),
    [sitePayments]
  );
  const contractTotalPrice = useMemo(
    () => siteContracts.reduce((s, c) => s + (Number(c.totalPrice) || 0), 0),
    [siteContracts]
  );
  const siteOtherExpenses = useMemo(
    () => (isAllSites ? otherExpenses : otherExpenses.filter((e) => e.siteId === currentSiteId)),
    [otherExpenses, currentSiteId, isAllSites]
  );
  const otherExpensesTotal = useMemo(
    () => siteOtherExpenses.reduce((s, e) => s + (Number(e.amount) || 0), 0),
    [siteOtherExpenses]
  );

  const projectGrandTotal = grandTotal + contractPaidTotal + otherExpensesTotal + fixedAssetsTotal;
  const profitLoss = clientPaidTotal - projectGrandTotal;

  const siteBreakdown = useMemo(() => {
    const blank = (id, name) => ({ id, name, material: 0, labor: 0, contract: 0, collected: 0, capital: 0, other: 0, asset: 0 });
    const map = {};
    sites.forEach((s) => (map[s.id] = blank(s.id, s.name)));
    // 紀錄所屬的案場已被刪除時，獨立列成一列，避免總數對不起來
    const bucket = (siteId) => map[siteId] || map.__orphan || (map.__orphan = blank("__orphan", "（已刪除案場的紀錄）"));
    orders.forEach((o) => { bucket(o.siteId).material += Number(o.amount) || 0; });
    workLogs.forEach((l) => { bucket(l.siteId).labor += (Number(l.days) || 0) * workLogRate(l); });
    payments.forEach((p) => { bucket(p.siteId).contract += Number(p.amount) || 0; });
    clientPayments.forEach((p) => { bucket(p.siteId).collected += Number(p.amount) || 0; });
    capitalContributions.forEach((c) => { bucket(c.siteId).capital += Number(c.amount) || 0; });
    otherExpenses.forEach((e) => { bucket(e.siteId).other += Number(e.amount) || 0; });
    fixedAssets.forEach((f) => { bucket(f.siteId).asset += (Number(f.amount) || 0) * (Number(f.quantity) || 1); });
    const o = map.__orphan;
    if (o && !(o.material || o.labor || o.contract || o.collected || o.capital || o.other || o.asset)) delete map.__orphan;
    return Object.values(map)
      .map((s) => ({ ...s, total: s.material + s.labor + s.contract + s.other + s.asset, profit: s.collected - (s.material + s.labor + s.contract + s.other + s.asset) }))
      .sort((a, b) => b.total - a.total);
  }, [sites, orders, workLogs, workLogRate, payments, clientPayments, capitalContributions, otherExpenses, fixedAssets]);

  // ---- forms ----
  const emptyOrder = { date: todayStr(), categoryId: categories[0]?.id || "", itemName: "", supplier: "", quantity: "", unit: "", unitPrice: "", note: "", hasInvoice: true, applyTax: false, taxRate: 5 };
  const [orderForm, setOrderForm] = useState(emptyOrder);
  const [newCatName, setNewCatName] = useState("");
  const [newWorker, setNewWorker] = useState({ name: "", dailyRate: "" });
  const emptyWorkLog = { workerId: "", date: todayStr(), days: 1, note: "", dailyRate: "" };
  const [workLogForm, setWorkLogForm] = useState(emptyWorkLog);
  const emptyContract = { name: "", contractor: "", totalPrice: "", date: todayStr(), note: "" };
  const [contractForm, setContractForm] = useState(emptyContract);
  const emptyPayment = { contractId: "", date: todayStr(), amount: "", note: "" };
  const [paymentForm, setPaymentForm] = useState(emptyPayment);
  const emptyClientPayment = { date: todayStr(), amount: "", item: "", note: "" };
  const [clientPaymentForm, setClientPaymentForm] = useState(emptyClientPayment);
  const emptyDiscussion = { date: todayStr(), topic: "", note: "", resolved: false, result: "" };
  const [discussionForm, setDiscussionForm] = useState(emptyDiscussion);
  const emptyNotebookItem = { date: todayStr(), title: "", content: "" };
  const [notebookForm, setNotebookForm] = useState(emptyNotebookItem);
  const emptyCapitalContribution = { date: todayStr(), amount: "", note: "" };
  const [capitalForm, setCapitalForm] = useState(emptyCapitalContribution);
  const emptyFixedAsset = { name: "", quantity: 1, amount: "", date: todayStr(), hasInvoice: true, applyTax: false, taxRate: 5 };
  const [fixedAssetForm, setFixedAssetForm] = useState(emptyFixedAsset);
  const emptyOtherExpense = { date: todayStr(), itemName: "", supplier: "", amount: "", note: "", hasInvoice: true, applyTax: false, taxRate: 5 };
  const [otherExpenseForm, setOtherExpenseForm] = useState(emptyOtherExpense);

  useEffect(() => {
    setWorkLogForm((f) => {
      const workerId = f.workerId || workers[0]?.id || "";
      const worker = workers.find((w) => w.id === workerId);
      const dailyRate = f.dailyRate === "" || f.dailyRate == null ? (worker?.dailyRate ?? "") : f.dailyRate;
      return { ...f, workerId, dailyRate };
    });
  }, [workers]);

  useEffect(() => {
    setPaymentForm((f) => ({ ...f, contractId: siteContracts.some((c) => c.id === f.contractId) ? f.contractId : (siteContracts[0]?.id || "") }));
  }, [siteContracts]);

  useEffect(() => {
    setOrderForm((f) => ({ ...f, categoryId: f.categoryId || categories[0]?.id || "" }));
  }, [categories]);

  const submitOrder = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增叫貨紀錄"); return; }
    if (!orderForm.itemName.trim() || !orderForm.quantity || !orderForm.unitPrice) {
      showToast("請填寫「品項名稱」「數量」「單價」後再新增叫貨紀錄");
      return;
    }
    const amount = Number(orderForm.quantity) * Number(orderForm.unitPrice);
    const rec = { id: uid(), ...orderForm, siteId: currentSiteId, amount };
    const next = [rec, ...orders];
    setOrders(next);
    persistOrders(next);
    setOrderForm({ ...emptyOrder, categoryId: orderForm.categoryId });
  };
  const deleteOrder = (id) => {
    const next = orders.filter((o) => o.id !== id);
    setOrders(next);
    persistOrders(next);
  };
  const updateOrder = (id, patch) => {
    const next = orders.map((o) => (o.id === id ? { ...o, ...patch } : o));
    setOrders(next);
    persistOrders(next);
  };

  const deleteUsage = (id) => {
    const next = usages.filter((u) => u.id !== id);
    setUsages(next);
    persistUsages(next);
  };
  const consumeFromInventory = ({ date, categoryId, itemName, quantity, unit, location, type }) => {
    const rec = { id: uid(), date, categoryId, itemName, quantity, unit, location: location || "", type: type === "return" ? "return" : "out", siteId: currentSiteId };
    const next = [rec, ...usages];
    setUsages(next);
    persistUsages(next);
  };
  // 校正庫存負數：消貨紀錄的類別/品項與叫貨對不上時，併入正確的品項，或直接刪除。
  const inUsageScope = (u) => isAllSites || u.siteId === currentSiteId;
  const reassignUsages = (fromCategoryId, itemName, toCategoryId, toItemName) => {
    let n = 0;
    const next = usages.map((u) => {
      if (inUsageScope(u) && u.categoryId === fromCategoryId && u.itemName === itemName) { n++; return { ...u, categoryId: toCategoryId, itemName: toItemName }; }
      return u;
    });
    setUsages(next);
    persistUsages(next);
    showToast(`已併入「${toItemName}」，共更新 ${n} 筆消貨紀錄`);
  };
  const deleteUsagesByKey = (categoryId, itemName) => {
    const next = usages.filter((u) => !(inUsageScope(u) && u.categoryId === categoryId && u.itemName === itemName));
    const n = usages.length - next.length;
    setUsages(next);
    persistUsages(next);
    showToast(`已刪除 ${n} 筆消貨紀錄`);
  };

  const addCategory = () => {
    const name = newCatName.trim();
    if (!name) return;
    const palette = ["#4A90A4", "#E8A33D", "#C97B4A", "#6B9B6E", "#8B95A1", "#C1543C", "#7A8A99", "#B98BC9"];
    const color = palette[categories.length % palette.length];
    const next = [...categories, { id: uid(), name, color }];
    setCategories(next);
    persistMeta(projectName, next);
    setNewCatName("");
  };
  const removeCategory = (id) => {
    if (orders.some((o) => o.categoryId === id) || usages.some((u) => u.categoryId === id)) {
      showToast("此類別已有紀錄，無法刪除");
      return;
    }
    const next = categories.filter((c) => c.id !== id);
    setCategories(next);
    persistMeta(projectName, next);
  };
  const renameCategory = (id, name) => {
    const nm = (name || "").trim();
    if (!nm) return;
    const next = categories.map((c) => (c.id === id ? { ...c, name: nm } : c));
    setCategories(next);
    persistMeta(projectName, next);
  };

  // ---- 品項清單 (managed material catalog, used by the wheel-style picker in 叫貨；名稱與單價可隨時調整，改名會連動更新過去紀錄) ----
  const addMaterialItem = (name, unit, defaultPrice, categoryId) => {
    const nm = (name || "").trim();
    if (!nm) { showToast("請輸入品項名稱"); return; }
    if (materialItems.some((m) => m.name === nm)) { showToast("此品項已存在"); return; }
    const next = [...materialItems, { id: uid(), name: nm, unit: (unit || "").trim(), defaultPrice: defaultPrice === "" || defaultPrice == null ? "" : Number(defaultPrice), categoryId: categoryId || "" }];
    setMaterialItems(next);
    persistMaterialItems(next);
  };
  const updateMaterialItem = (id, patch) => {
    const item = materialItems.find((m) => m.id === id);
    if (!item) return;
    const nextItems = materialItems.map((m) => (m.id === id ? { ...m, ...patch } : m));
    setMaterialItems(nextItems);
    persistMaterialItems(nextItems);
    // renaming cascades to every past order/usage record so history stays consistent
    if (patch.name && patch.name.trim() && patch.name.trim() !== item.name) {
      const oldName = item.name;
      const newName = patch.name.trim();
      const nextOrders = orders.map((o) => (o.itemName === oldName ? { ...o, itemName: newName } : o));
      const nextUsages = usages.map((u) => (u.itemName === oldName ? { ...u, itemName: newName } : u));
      setOrders(nextOrders);
      setUsages(nextUsages);
      persistOrders(nextOrders);
      persistUsages(nextUsages);
      showToast(`已將「${oldName}」更名為「${newName}」，並同步更新過去的叫貨／領用紀錄`);
    }
  };
  const removeMaterialItem = (id) => {
    const next = materialItems.filter((m) => m.id !== id);
    setMaterialItems(next);
    persistMaterialItems(next);
  };

  // ---- 廠商 / 供應商清單 (managed supplier catalog for 叫貨；改名同樣連動更新過去紀錄) ----
  const addSupplier = (name) => {
    const nm = (name || "").trim();
    if (!nm) { showToast("請輸入廠商名稱"); return; }
    if (suppliers.some((s) => s.name === nm)) { showToast("此廠商已存在"); return; }
    const next = [...suppliers, { id: uid(), name: nm }];
    setSuppliers(next);
    persistSuppliers(next);
  };
  const updateSupplier = (id, patch) => {
    const sup = suppliers.find((s) => s.id === id);
    if (!sup) return;
    const nextSuppliers = suppliers.map((s) => (s.id === id ? { ...s, ...patch } : s));
    setSuppliers(nextSuppliers);
    persistSuppliers(nextSuppliers);
    if (patch.name && patch.name.trim() && patch.name.trim() !== sup.name) {
      const oldName = sup.name;
      const newName = patch.name.trim();
      const nextOrders = orders.map((o) => (o.supplier === oldName ? { ...o, supplier: newName } : o));
      const nextOtherExpenses = otherExpenses.map((oe) => (oe.supplier === oldName ? { ...oe, supplier: newName } : oe));
      setOrders(nextOrders);
      persistOrders(nextOrders);
      setOtherExpenses(nextOtherExpenses);
      persistOtherExpenses(nextOtherExpenses);
      showToast(`已將廠商「${oldName}」更名為「${newName}」，並同步更新過去的叫貨／其他支出紀錄`);
    }
  };
  const removeSupplier = (id) => {
    const next = suppliers.filter((s) => s.id !== id);
    setSuppliers(next);
    persistSuppliers(next);
  };

  const addWorker = () => {
    const name = newWorker.name.trim();
    if (!name) return;
    const next = [...workers, { id: uid(), name, dailyRate: Number(newWorker.dailyRate) || 0 }];
    setWorkers(next);
    persistWorkers(next);
    setNewWorker({ name: "", dailyRate: "" });
  };
  const updateWorkerRate = (id, dailyRate) => {
    const next = workers.map((w) => (w.id === id ? { ...w, dailyRate: Number(dailyRate) || 0 } : w));
    setWorkers(next);
    persistWorkers(next);
  };
  const removeWorker = (id) => {
    if (workLogs.some((l) => l.workerId === id)) {
      showToast("此師傅已有出工紀錄，無法刪除");
      return;
    }
    const next = workers.filter((w) => w.id !== id);
    setWorkers(next);
    persistWorkers(next);
  };

  const submitWorkLog = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增出工紀錄"); return; }
    if (!workLogForm.workerId || !workLogForm.days) {
      showToast("請選擇「師傅」並填寫「出工天數」後再新增出工紀錄");
      return;
    }
    const rec = {
      id: uid(), ...workLogForm, siteId: currentSiteId, days: Number(workLogForm.days),
      dailyRate: Number(workLogForm.dailyRate) || Number(workers.find((w) => w.id === workLogForm.workerId)?.dailyRate) || 0,
    };
    const next = [rec, ...workLogs];
    setWorkLogs(next);
    persistWorkLogs(next);
    setWorkLogForm({ ...emptyWorkLog, workerId: workLogForm.workerId, dailyRate: workers.find((w) => w.id === workLogForm.workerId)?.dailyRate ?? "" });
  };
  const deleteWorkLog = (id) => {
    const next = workLogs.filter((l) => l.id !== id);
    setWorkLogs(next);
    persistWorkLogs(next);
  };
  const updateWorkLog = (id, patch) => {
    const next = workLogs.map((l) => (l.id === id ? { ...l, ...patch } : l));
    setWorkLogs(next);
    persistWorkLogs(next);
  };

  const submitContract = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增發包項目"); return; }
    if (!contractForm.name.trim() || !contractForm.totalPrice) {
      showToast("請填寫「工程總稱」與「發包總價」後再新增");
      return;
    }
    const rec = { id: uid(), ...contractForm, siteId: currentSiteId, totalPrice: Number(contractForm.totalPrice) };
    const next = [rec, ...contracts];
    setContracts(next);
    persistContracts(next);
    setContractForm(emptyContract);
  };
  const updateContract = (id, patch) => {
    const next = contracts.map((c) => (c.id === id ? { ...c, ...patch } : c));
    setContracts(next);
    persistContracts(next);
  };
  const renameVendor = (oldName, newName) => {
    const nm = (newName || "").trim();
    if (!nm || nm === oldName) return;
    const next = contracts.map((c) => (((c.contractor || "").trim() || "未指定廠商") === oldName ? { ...c, contractor: nm } : c));
    setContracts(next);
    persistContracts(next);
    showToast(`已將廠商「${oldName}」改名為「${nm}」，底下的發包項目都已一併更新`);
  };
  const deleteContract = (id) => {
    if (payments.some((p) => p.contractId === id)) {
      showToast("此發包項目已有領款紀錄，無法刪除");
      return;
    }
    if (floorItems.some((f) => f.contractId === id)) {
      showToast("此發包項目已有樓層工作項目，請先清除樓層項目再刪除");
      return;
    }
    if (contractWorkLogs.some((l) => l.contractId === id)) {
      showToast("此發包項目已有出工紀錄，請先清除出工紀錄再刪除");
      return;
    }
    if (contractItems.some((it) => it.contractId === id)) {
      showToast("此發包項目已有項目明細，請先清除項目明細再刪除");
      return;
    }
    (contracts.find((c) => c.id === id)?.files || []).forEach((f) => fileRemove(f.id, f.chunks));
    const next = contracts.filter((c) => c.id !== id);
    setContracts(next);
    persistContracts(next);
  };

  // ---- 發包合約檔案 (files are chunked into separate Firestore docs; only metadata lives on the contract) ----
  const uploadFileMeta = async (file) => {
    try {
      const prepared = await prepareUploadFile(file);
      if (prepared.size > MAX_UPLOAD_BYTES) { showToast("檔案超過 5 MB，請壓縮後再上傳（照片會自動縮小，PDF 請先壓縮）"); return null; }
      const b64 = await blobToBase64(prepared);
      const fileId = uid() + uid();
      const chunks = await fileStore(fileId, b64);
      if (!chunks) { showToast("上傳失敗，請檢查網路後再試一次"); return null; }
      return { id: fileId, name: prepared.name, size: prepared.size, type: prepared.type || "application/octet-stream", chunks, uploadedAt: todayStr() };
    } catch (e) {
      console.error("upload file failed", e);
      showToast("上傳失敗，請檢查網路後再試一次");
      return null;
    }
  };
  const addContractFile = async (contractId, file) => {
    const meta = await uploadFileMeta(file);
    if (!meta) return false;
    const next = contractsRef.current.map((c) => (c.id === contractId ? { ...c, files: [...(c.files || []), meta] } : c));
    setContracts(next);
    persistContracts(next);
    showToast(`已上傳：${meta.name}`);
    return true;
  };

  // ---- 收貨單相簿 (receipt photos, same chunked file storage as contracts) ----
  const addReceipt = async ({ date, title, note, files }) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增收貨單"); return false; }
    const metas = [];
    for (const f of files) {
      const meta = await uploadFileMeta(f);
      if (!meta) { metas.forEach((m) => fileRemove(m.id, m.chunks)); return false; }
      metas.push(meta);
    }
    const rec = { id: uid(), siteId: currentSiteId, date: date || todayStr(), title: (title || "").trim(), note: (note || "").trim(), files: metas };
    const next = [rec, ...receiptsRef.current];
    setReceipts(next);
    persistReceipts(next);
    showToast(`已儲存收貨單（${metas.length} 張照片）`);
    return true;
  };
  const updateReceipt = (id, patch) => {
    const next = receiptsRef.current.map((r) => (r.id === id ? { ...r, ...patch } : r));
    setReceipts(next);
    persistReceipts(next);
  };
  const deleteReceipt = (id) => {
    const rec = receiptsRef.current.find((r) => r.id === id);
    if (!rec || !window.confirm(`確定要刪除這張收貨單（${rec.files.length} 張照片）嗎？刪除後無法復原。`)) return;
    rec.files.forEach((f) => fileRemove(f.id, f.chunks));
    const next = receiptsRef.current.filter((r) => r.id !== id);
    setReceipts(next);
    persistReceipts(next);
  };

  // ---- 估價單（僅供對照，不計入任何總結算） ----
  const addEstimateLines = (rows) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增估價項目"); return; }
    const recs = rows.map((r) => ({
      id: uid(), siteId: currentSiteId, categoryId: r.categoryId || categories[0]?.id || "", name: (r.name || "").trim(),
      quantity: Number(r.quantity) || 0, unit: (r.unit || "").trim(), unitPrice: Number(r.unitPrice) || 0, note: (r.note || "").trim(),
    }));
    const next = [...estimates, ...recs];
    setEstimates(next);
    persistEstimates(next);
    return recs.length;
  };
  const addEstimateLine = (line) => { if (addEstimateLines([line])) showToast("已新增估價項目"); };
  const confirmScannedEstimate = async (rows) => {
    const n = addEstimateLines(rows.map((r) => ({ ...r, note: "辨識匯入" })));
    if (n) showToast(`已加入 ${n} 個估價項目`);
  };
  const updateEstimateLine = (id, patch) => {
    const next = estimates.map((l) => (l.id === id ? { ...l, ...patch } : l));
    setEstimates(next);
    persistEstimates(next);
  };
  const deleteEstimateLine = (id) => {
    const next = estimates.filter((l) => l.id !== id);
    setEstimates(next);
    persistEstimates(next);
  };
  const addEstimateFile = async (file) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能上傳估價單"); return false; }
    const siteId = currentSiteId;
    const meta = await uploadFileMeta(file);
    if (!meta) return false;
    const cur = estimateFilesRef.current;
    const next = { ...cur, [siteId]: [...(cur[siteId] || []), meta] };
    setEstimateFilesBySite(next);
    persistEstimateFiles(next);
    showToast(`已上傳：${meta.name}`);
    return true;
  };
  const deleteEstimateFile = (file) => {
    if (!window.confirm(`確定要刪除估價單檔案「${file.name}」嗎？刪除後無法復原。`)) return;
    const cur = estimateFilesRef.current;
    const next = { ...cur, [currentSiteId]: (cur[currentSiteId] || []).filter((f) => f.id !== file.id) };
    setEstimateFilesBySite(next);
    persistEstimateFiles(next);
    fileRemove(file.id, file.chunks);
  };

  // ---- 掃描校對後批次加入 ----
  const addMaterialItemsBatch = (items) => {
    const have = new Set(materialItems.map((m) => m.name));
    const fresh = [];
    items.forEach((it) => {
      const nm = (it.name || "").trim();
      if (!nm || have.has(nm)) return;
      have.add(nm);
      fresh.push({ id: uid(), name: nm, unit: (it.unit || "").trim(), defaultPrice: it.defaultPrice === "" || it.defaultPrice == null ? "" : Number(it.defaultPrice), categoryId: it.categoryId || "" });
    });
    if (!fresh.length) return 0;
    const next = [...materialItems, ...fresh];
    setMaterialItems(next);
    persistMaterialItems(next);
    return fresh.length;
  };
  const confirmScannedItems = (rows) => {
    const n = addMaterialItemsBatch(rows.map((r) => ({ name: r.name, unit: r.unit, defaultPrice: r.defaultPrice, categoryId: r.categoryId })));
    showToast(n ? `已加入 ${n} 個品項到品項清單` : "這些品項都已經在品項清單裡了");
  };
  const confirmScannedOrders = async (rows, meta, photoFile) => {
    const recs = rows.map((r) => {
      const qty = Number(r.quantity);
      const price = Number(r.unitPrice);
      return {
        id: uid(), date: meta.date, categoryId: r.categoryId || categories[0]?.id || "", itemName: r.name.trim(), supplier: meta.supplier || "",
        quantity: qty, unit: r.unit || "", unitPrice: price, note: meta.note ?? "收貨單掃描", hasInvoice: !!meta.hasInvoice, applyTax: !!(meta.hasInvoice && meta.applyTax), taxRate: Number(meta.taxRate) || 0,
        siteId: currentSiteId, amount: qty * price,
      };
    });
    const next = [...recs, ...orders];
    setOrders(next);
    persistOrders(next);
    let added = 0;
    if (meta.addNewItems) {
      added = addMaterialItemsBatch(rows.map((r) => ({ name: r.name, unit: r.unit, defaultPrice: r.unitPrice, categoryId: r.categoryId })));
    }
    showToast(`已加入叫貨 ${recs.length} 筆${added ? `，並新增 ${added} 個品項到品項清單` : ""}`);
    if (meta.saveToAlbum && photoFile) {
      await addReceipt({ date: meta.date, title: meta.supplier ? `${meta.supplier} 叫貨` : "叫貨收貨單", note: `掃描入帳 ${recs.length} 筆`, files: [photoFile] });
    }
  };

  const deleteContractFile = (contractId, file) => {
    if (!window.confirm(`確定要刪除合約檔案「${file.name}」嗎？刪除後無法復原。`)) return;
    const next = contractsRef.current.map((c) => (c.id === contractId ? { ...c, files: (c.files || []).filter((f) => f.id !== file.id) } : c));
    setContracts(next);
    persistContracts(next);
    fileRemove(file.id, file.chunks);
  };
  const openContractFile = async (file) => {
    try {
      showToast("讀取檔案中…");
      const blob = base64ToBlob(await fileLoad(file.id, file.chunks), file.type);
      const viewable = file.type === "application/pdf" || file.type.startsWith("image/");
      const url = URL.createObjectURL(blob);
      const win = viewable ? window.open(url, "_blank") : null;
      if (!win) triggerDownload(blob, file.name);
      setTimeout(() => URL.revokeObjectURL(url), 60000);
    } catch (e) {
      console.error("open contract file failed", e);
      showToast("讀取檔案失敗，請稍後再試一次");
    }
  };

  const submitPayment = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能登記領款"); return; }
    if (!paymentForm.contractId || !paymentForm.amount) {
      showToast("請選擇「發包項目」並填寫「金額」後再登記領款");
      return;
    }
    const rec = { id: uid(), ...paymentForm, siteId: currentSiteId, amount: Number(paymentForm.amount) };
    const next = [rec, ...payments];
    setPayments(next);
    persistPayments(next);
    setPaymentForm({ ...emptyPayment, contractId: paymentForm.contractId });
  };
  const deletePayment = (id) => {
    const next = payments.filter((p) => p.id !== id);
    setPayments(next);
    persistPayments(next);
  };
  const updatePayment = (id, patch) => {
    const next = payments.map((p) => (p.id === id ? { ...p, ...patch } : p));
    setPayments(next);
    persistPayments(next);
  };
  const addPaymentFromItem = (contractId, date, amount, note, itemId) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能登記領款"); return; }
    if (!amount) { showToast("這個項目金額為 0，請先填寫金額再匯入領款"); return; }
    const rec = { id: uid(), contractId, siteId: currentSiteId, date: date || todayStr(), amount: Number(amount), note: note || "" };
    const next = [rec, ...payments];
    setPayments(next);
    persistPayments(next);
    if (itemId) {
      const nextItems = contractItems.map((it) => (it.id === itemId ? { ...it, paidImported: true } : it));
      setContractItems(nextItems);
      persistContractItems(nextItems);
    }
    showToast(`已匯入領款紀錄：${note || "（未命名項目）"} ${fmtMoney(amount)}`);
  };

  // ---- floor work-item templates (標準層項目範本) ----
  const addTemplateRow = (contractId, name, amount) => {
    const nm = (name || "").trim();
    if (!nm) { showToast("請先輸入範本項目名稱"); return; }
    const next = contracts.map((c) =>
      c.id === contractId ? { ...c, templateItems: [...(c.templateItems || []), { id: uid(), name: nm, amount: amount === "" || amount == null ? "" : Number(amount) }] } : c
    );
    setContracts(next);
    persistContracts(next);
  };
  const updateTemplateRow = (contractId, rowId, patch) => {
    const next = contracts.map((c) =>
      c.id === contractId
        ? { ...c, templateItems: (c.templateItems || []).map((r) => (r.id === rowId ? { ...r, ...patch } : r)) }
        : c
    );
    setContracts(next);
    persistContracts(next);
  };
  const removeTemplateRow = (contractId, rowId) => {
    const next = contracts.map((c) =>
      c.id === contractId ? { ...c, templateItems: (c.templateItems || []).filter((r) => r.id !== rowId) } : c
    );
    setContracts(next);
    persistContracts(next);
  };

  const applyTemplateToFloors = (contractId, floorList, overwrite) => {
    const contract = contracts.find((c) => c.id === contractId);
    if (!contract) return;
    const template = (contract.templateItems || []).filter((r) => r.name.trim() && r.amount !== "");
    if (!template.length || !floorList.length) return;
    let next = floorItems;
    if (overwrite) {
      next = next.filter((f) => !(f.contractId === contractId && floorList.includes(f.floor)));
    }
    const additions = [];
    let skipped = 0;
    floorList.forEach((floor) => {
      const already = next.some((f) => f.contractId === contractId && f.floor === floor);
      if (already && !overwrite) { skipped += 1; return; }
      template.forEach((t) => {
        additions.push({ id: uid(), contractId, siteId: contract.siteId, floor, itemName: t.name, amount: Number(t.amount) || 0, note: "" });
      });
    });
    const finalList = [...next, ...additions];
    setFloorItems(finalList);
    persistFloorItems(finalList);
    const appliedFloors = floorList.length - skipped;
    if (appliedFloors > 0) {
      showToast(`已套用範本到 ${appliedFloors} 個樓層（共新增 ${additions.length} 筆項目）${skipped > 0 ? `，${skipped} 個樓層已有項目故略過` : ""}`);
    } else {
      showToast("所選樓層都已有項目，未新增任何項目（可勾選「覆蓋既有項目」強制套用）");
    }
  };

  const addFloorItem = (contractId, floor, itemName, amount) => {
    const contract = contracts.find((c) => c.id === contractId);
    if (!contract || !floor.trim() || !itemName.trim()) return;
    const rec = { id: uid(), contractId, siteId: contract.siteId, floor: floor.trim(), itemName: itemName.trim(), amount: Number(amount) || 0, note: "" };
    const next = [...floorItems, rec];
    setFloorItems(next);
    persistFloorItems(next);
  };
  const updateFloorItem = (id, patch) => {
    const next = floorItems.map((f) => (f.id === id ? { ...f, ...patch } : f));
    setFloorItems(next);
    persistFloorItems(next);
  };
  const deleteFloorItem = (id) => {
    const next = floorItems.filter((f) => f.id !== id);
    setFloorItems(next);
    persistFloorItems(next);
  };
  const deleteFloorGroup = (contractId, floor) => {
    const next = floorItems.filter((f) => !(f.contractId === contractId && f.floor === floor));
    setFloorItems(next);
    persistFloorItems(next);
  };

  // ---- 甲方收款 (client billing / collections) ----
  const submitClientPayment = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增甲方收款紀錄"); return; }
    if (!clientPaymentForm.amount) {
      showToast("請填寫「金額」後再新增甲方收款紀錄");
      return;
    }
    const rec = { id: uid(), ...clientPaymentForm, siteId: currentSiteId, amount: Number(clientPaymentForm.amount) };
    const next = [rec, ...clientPayments];
    setClientPayments(next);
    persistClientPayments(next);
    setClientPaymentForm(emptyClientPayment);
  };
  const deleteClientPayment = (id) => {
    const next = clientPayments.filter((p) => p.id !== id);
    setClientPayments(next);
    persistClientPayments(next);
  };
  const updateClientPayment = (id, patch) => {
    const next = clientPayments.map((p) => (p.id === id ? { ...p, ...patch } : p));
    setClientPayments(next);
    persistClientPayments(next);
  };

  // ---- 投入成本 (capital contributions) ----
  const submitCapitalContribution = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增投入成本紀錄"); return; }
    if (!capitalForm.amount) {
      showToast("請填寫「金額」後再新增投入成本紀錄");
      return;
    }
    const rec = { id: uid(), ...capitalForm, siteId: currentSiteId, amount: Number(capitalForm.amount) };
    const next = [rec, ...capitalContributions];
    setCapitalContributions(next);
    persistCapitalContributions(next);
    setCapitalForm(emptyCapitalContribution);
  };
  const deleteCapitalContribution = (id) => {
    const next = capitalContributions.filter((c) => c.id !== id);
    setCapitalContributions(next);
    persistCapitalContributions(next);
  };
  const updateCapitalContribution = (id, patch) => {
    const next = capitalContributions.map((c) => (c.id === id ? { ...c, ...patch } : c));
    setCapitalContributions(next);
    persistCapitalContributions(next);
  };

  // ---- 固定資產 (fixed assets / tools inventory) ----
  const submitFixedAsset = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增固定資產"); return; }
    if (!fixedAssetForm.name.trim()) {
      showToast("請填寫「工具名稱」後再新增固定資產");
      return;
    }
    const rec = {
      id: uid(), ...fixedAssetForm, siteId: currentSiteId,
      quantity: Number(fixedAssetForm.quantity) || 1, amount: Number(fixedAssetForm.amount) || 0,
    };
    const next = [rec, ...fixedAssets];
    setFixedAssets(next);
    persistFixedAssets(next);
    setFixedAssetForm(emptyFixedAsset);
  };
  const deleteFixedAsset = (id) => {
    const next = fixedAssets.filter((f) => f.id !== id);
    setFixedAssets(next);
    persistFixedAssets(next);
  };
  const updateFixedAsset = (id, patch) => {
    const next = fixedAssets.map((f) => (f.id === id ? { ...f, ...patch } : f));
    setFixedAssets(next);
    persistFixedAssets(next);
  };

  // ---- 其他支出 (other expenses — counts toward project cost, but doesn't
  // touch 領料/庫存: item name can be picked from the material catalog for
  // convenience but is stored as plain text, not linked to orders/usages) ----
  const submitOtherExpense = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增其他支出"); return; }
    if (!otherExpenseForm.itemName.trim() || !otherExpenseForm.amount) {
      showToast("請填寫「項目名稱」與「金額」後再新增其他支出");
      return;
    }
    const rec = { id: uid(), ...otherExpenseForm, siteId: currentSiteId, amount: Number(otherExpenseForm.amount) };
    const next = [rec, ...otherExpenses];
    setOtherExpenses(next);
    persistOtherExpenses(next);
    setOtherExpenseForm(emptyOtherExpense);
  };
  const deleteOtherExpense = (id) => {
    const next = otherExpenses.filter((e) => e.id !== id);
    setOtherExpenses(next);
    persistOtherExpenses(next);
  };
  const updateOtherExpense = (id, patch) => {
    const next = otherExpenses.map((e) => (e.id === id ? { ...e, ...patch } : e));
    setOtherExpenses(next);
    persistOtherExpenses(next);
  };

  // ---- 甲方樓層請款項目 (client-side billing items — same design as contractor: item -> floor -> percent) ----
  const addClientTemplateRow = (siteId, name, amount) => {
    const nm = (name || "").trim();
    if (!nm) { showToast("請先輸入項目名稱"); return; }
    const next = [...clientTemplateRows, { id: uid(), siteId, name: nm, amount: amount === "" || amount == null ? "" : Number(amount) }];
    setClientTemplateRows(next);
    persistClientTemplateRows(next);
  };
  const addClientFloorItem = (siteId, itemName, amount, date, note, percent, floor) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增項目"); return; }
    if (!itemName.trim()) { showToast("請填寫「項目名稱」後再新增"); return; }
    const rec = {
      id: uid(), siteId, itemName: itemName.trim(), amount: Number(amount) || 0,
      percent: percent === "" || percent == null ? undefined : Number(percent),
      date: date || todayStr(), note: note || "", floor: (floor || "").trim(),
    };
    const next = [...clientFloorItems, rec];
    setClientFloorItems(next);
    persistClientFloorItems(next);
  };
  const addClientFloorItemsBatch = (siteId, itemName, amount, date, note, percent, floors) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增項目"); return; }
    if (!itemName.trim() || !floors.length) return;
    const additions = floors.map((floor) => ({
      id: uid(), siteId, itemName: itemName.trim(), amount: Number(amount) || 0,
      percent: percent === "" || percent == null ? undefined : Number(percent),
      date: date || todayStr(), note: note || "", floor,
    }));
    const next = [...clientFloorItems, ...additions];
    setClientFloorItems(next);
    persistClientFloorItems(next);
    showToast(`已將「${itemName.trim()}」套用到 ${floors.length} 個樓層`);
  };
  const updateClientFloorItem = (id, patch) => {
    const next = clientFloorItems.map((f) => (f.id === id ? { ...f, ...patch } : f));
    setClientFloorItems(next);
    persistClientFloorItems(next);
  };
  const deleteClientFloorItem = (id) => {
    const next = clientFloorItems.filter((f) => f.id !== id);
    setClientFloorItems(next);
    persistClientFloorItems(next);
  };
  const deleteClientFloorItemsByFloor = (siteId, itemName, floor) => {
    const next = clientFloorItems.filter((f) => !(f.siteId === siteId && f.itemName === itemName && (f.floor || "") === floor));
    setClientFloorItems(next);
    persistClientFloorItems(next);
    showToast(`已刪除「${itemName}」在 ${floor || "一般項目"} 的所有紀錄`);
  };
  const deleteClientFloorItemsByName = (siteId, itemName) => {
    const next = clientFloorItems.filter((f) => !(f.siteId === siteId && f.itemName === itemName));
    setClientFloorItems(next);
    persistClientFloorItems(next);
    showToast(`已刪除「${itemName}」的所有樓層紀錄`);
  };
  const addClientPaymentFromItem = (siteId, date, amount, note, itemId) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增甲方收款紀錄"); return; }
    if (!amount) { showToast("這個項目金額為 0，請先填寫金額再匯入收款"); return; }
    const rec = { id: uid(), siteId, date: date || todayStr(), amount: Number(amount), item: note || "", note: "" };
    const next = [rec, ...clientPayments];
    setClientPayments(next);
    persistClientPayments(next);
    if (itemId) {
      const nextItems = clientFloorItems.map((f) => (f.id === itemId ? { ...f, paidImported: true } : f));
      setClientFloorItems(nextItems);
      persistClientFloorItems(nextItems);
    }
    showToast(`已匯入甲方收款：${note || "（未命名項目）"} ${fmtMoney(amount)}`);
  };

  // ---- 發包廠商出工紀錄 (subcontractor attendance) ----
  const addAttendance = (contractId, date, headcount, note) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能登記發包廠商出工"); return; }
    if (!contractId || !date) {
      showToast("請確認發包項目與日期已填寫");
      return;
    }
    const rec = { id: uid(), contractId, siteId: currentSiteId, date, headcount: Number(headcount) || 1, note: note || "" };
    const next = [rec, ...contractWorkLogs];
    setContractWorkLogs(next);
    persistContractWorkLogs(next);
  };
  const deleteAttendance = (id) => {
    const next = contractWorkLogs.filter((l) => l.id !== id);
    setContractWorkLogs(next);
    persistContractWorkLogs(next);
  };
  const updateAttendance = (id, patch) => {
    const next = contractWorkLogs.map((l) => (l.id === id ? { ...l, ...patch } : l));
    setContractWorkLogs(next);
    persistContractWorkLogs(next);
  };

  // ---- 發包項目明細 (a single contractor/vendor can carry multiple priced items, each with its own billing date) ----
  const addContractItem = (contractId, name, amount, date, note, percent, floor) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增項目"); return; }
    const contract = contracts.find((c) => c.id === contractId);
    if (!contract || !name.trim()) { showToast("請填寫「項目名稱」後再新增"); return; }
    const rec = {
      id: uid(), contractId, siteId: contract.siteId, name: name.trim(), amount: Number(amount) || 0,
      percent: percent === "" || percent == null ? undefined : Number(percent),
      date: date || todayStr(), note: note || "", floor: (floor || "").trim(),
    };
    const next = [...contractItems, rec];
    setContractItems(next);
    persistContractItems(next);
  };
  const addContractItemsBatch = (contractId, name, amount, date, note, percent, floors) => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增項目"); return; }
    const contract = contracts.find((c) => c.id === contractId);
    if (!contract || !name.trim() || !floors.length) return;
    const additions = floors.map((floor) => ({
      id: uid(), contractId, siteId: contract.siteId, name: name.trim(), amount: Number(amount) || 0,
      percent: percent === "" || percent == null ? undefined : Number(percent),
      date: date || todayStr(), note: note || "", floor,
    }));
    const next = [...contractItems, ...additions];
    setContractItems(next);
    persistContractItems(next);
    showToast(`已將「${name.trim()}」套用到 ${floors.length} 個樓層`);
  };
  const updateContractItem = (id, patch) => {
    const next = contractItems.map((it) => (it.id === id ? { ...it, ...patch } : it));
    setContractItems(next);
    persistContractItems(next);
  };
  const deleteContractItem = (id) => {
    const next = contractItems.filter((it) => it.id !== id);
    setContractItems(next);
    persistContractItems(next);
  };
  const deleteContractItemsByFloor = (contractId, name, floor) => {
    const next = contractItems.filter((it) => !(it.contractId === contractId && it.name === name && (it.floor || "") === floor));
    setContractItems(next);
    persistContractItems(next);
    showToast(`已刪除「${name}」在 ${floor || "一般項目"} 的所有紀錄`);
  };
  const deleteContractItemsByName = (contractId, name) => {
    const next = contractItems.filter((it) => !(it.contractId === contractId && it.name === name));
    setContractItems(next);
    persistContractItems(next);
    showToast(`已刪除「${name}」的所有樓層紀錄`);
  };

  // ---- 備註討論項目 (discussion / open-issue notes) ----
  const submitDiscussion = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增討論項目"); return; }
    if (!discussionForm.topic.trim()) {
      showToast("請填寫「討論項目」後再新增");
      return;
    }
    const rec = { id: uid(), ...discussionForm, siteId: currentSiteId };
    const next = [rec, ...discussionItems];
    setDiscussionItems(next);
    persistDiscussionItems(next);
    setDiscussionForm(emptyDiscussion);
  };
  const toggleDiscussionResolved = (id) => {
    const next = discussionItems.map((d) => (d.id === id ? { ...d, resolved: !d.resolved } : d));
    setDiscussionItems(next);
    persistDiscussionItems(next);
  };
  const updateDiscussionResult = (id, result) => {
    const next = discussionItems.map((d) => (d.id === id ? { ...d, result } : d));
    setDiscussionItems(next);
    persistDiscussionItems(next);
  };
  const updateDiscussion = (id, patch) => {
    const next = discussionItems.map((d) => (d.id === id ? { ...d, ...patch } : d));
    setDiscussionItems(next);
    persistDiscussionItems(next);
  };
  const deleteDiscussion = (id) => {
    const next = discussionItems.filter((d) => d.id !== id);
    setDiscussionItems(next);
    persistDiscussionItems(next);
  };

  // ---- 大項記事本 (freeform notebook entries) ----
  const submitNotebookItem = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增記事本項目"); return; }
    if (!notebookForm.title.trim()) {
      showToast("請填寫「標題」後再新增記事本項目");
      return;
    }
    const rec = { id: uid(), ...notebookForm, siteId: currentSiteId };
    const next = [rec, ...notebookItems];
    setNotebookItems(next);
    persistNotebookItems(next);
    setNotebookForm(emptyNotebookItem);
  };
  const updateNotebookItem = (id, patch) => {
    const next = notebookItems.map((n) => (n.id === id ? { ...n, ...patch } : n));
    setNotebookItems(next);
    persistNotebookItems(next);
  };
  const deleteNotebookItem = (id) => {
    const next = notebookItems.filter((n) => n.id !== id);
    setNotebookItems(next);
    persistNotebookItems(next);
  };

  const addSite = (name, address, copyFromSiteId) => {
    const nm = (name || "").trim();
    if (!nm) return;
    const rec = { id: uid(), name: nm, address: (address || "").trim() };
    const next = [...sites, rec];
    setSites(next);
    persistSites(next);
    setCurrentSiteId(rec.id);
    if (copyFromSiteId) {
      copySiteVendorSetup(copyFromSiteId, rec.id);
    }
  };
  const copySiteVendorSetup = (sourceSiteId, targetSiteId) => {
    const sourceContracts = contracts.filter((c) => c.siteId === sourceSiteId);
    if (!sourceContracts.length) { showToast("來源案場尚無發包項目可複製"); return; }
    const idMap = {};
    const newContracts = sourceContracts.map((c) => {
      const newId = uid();
      idMap[c.id] = newId;
      return {
        ...c, id: newId, siteId: targetSiteId, date: todayStr(), files: [],
        templateItems: (c.templateItems || []).map((t) => ({ ...t, id: uid() })),
      };
    });
    const sourceItems = contractItems.filter((it) => it.siteId === sourceSiteId && idMap[it.contractId]);
    const newItems = sourceItems.map((it) => ({ ...it, id: uid(), contractId: idMap[it.contractId], siteId: targetSiteId }));
    const nextContracts = [...contracts, ...newContracts];
    const nextItems = [...contractItems, ...newItems];
    setContracts(nextContracts);
    setContractItems(nextItems);
    persistContracts(nextContracts);
    persistContractItems(nextItems);
    showToast(`已從來源案場複製 ${newContracts.length} 個發包項目、${newItems.length} 筆請款項目明細（付款與出工紀錄不會複製，是全新案場）`);
  };
  const renameSite = (id, name, address) => {
    const next = sites.map((s) => (s.id === id ? { ...s, name, address } : s));
    setSites(next);
    persistSites(next);
  };
  const updateSiteClientPricing = (id, patch) => {
    const next = sites.map((s) => {
      if (s.id !== id) return s;
      const merged = { ...s, ...patch };
      const unitPrice = Number(merged.clientUnitPrice) || 0;
      const unitCount = Number(merged.clientUnitCount) || 0;
      // if unit price and count are being edited (not the total itself), auto-compute the total
      if (("clientUnitPrice" in patch || "clientUnitCount" in patch) && unitPrice > 0 && unitCount > 0) {
        merged.clientTotalPrice = unitPrice * unitCount;
      }
      return merged;
    });
    setSites(next);
    debouncedPersist("sites", persistSites, next);
  };
  const removeSite = (id) => {
    if (sites.length <= 1) { showToast("至少需保留一個案場"); return; }
    const hasData =
      orders.some((o) => o.siteId === id) ||
      usages.some((u) => u.siteId === id) ||
      workLogs.some((l) => l.siteId === id) ||
      contracts.some((c) => c.siteId === id) ||
      payments.some((p) => p.siteId === id) ||
      floorItems.some((f) => f.siteId === id) ||
      clientPayments.some((p) => p.siteId === id) ||
      clientTemplateRows.some((r) => r.siteId === id) ||
      clientFloorItems.some((f) => f.siteId === id) ||
      contractWorkLogs.some((l) => l.siteId === id) ||
      contractItems.some((it) => it.siteId === id) ||
      discussionItems.some((d) => d.siteId === id);
    if (hasData) { showToast("此案場已有叫貨、領料、出工或發包紀錄，無法刪除"); return; }
    const next = sites.filter((s) => s.id !== id);
    setSites(next);
    persistSites(next);
    if (currentSiteId === id) setCurrentSiteId(next[0]?.id || "all");
  };

  const saveName = () => {
    setEditingName(false);
    persistMeta(projectName, categories);
  };

  // ---- total cost digit meter ----
  const meterDigits = String(Math.round(projectGrandTotal)).padStart(8, "0").split("");

  if (loading) {
    return (
      <div className="wel-app" style={{ minHeight: 480, display: "flex", alignItems: "center", justifyContent: "center" }}>
        <StyleBlock />
        <div style={{ color: "var(--text-muted)", display: "flex", gap: 10, alignItems: "center", fontFamily: "var(--font-body)" }}>
          <Loader2 size={20} className="wel-spin" /> 讀取資料中…
        </div>
      </div>
    );
  }

  return (
    <div className="wel-app">
      <StyleBlock />
      {toast && (
        <div className="wel-toast" onClick={() => setToast(null)}>
          <AlertTriangle size={15} />
          <span>{toast}</span>
        </div>
      )}
      {view === "sites" ? (
        <SiteLandingPage
          projectName={projectName}
          setProjectName={setProjectName}
          editingName={editingName}
          setEditingName={setEditingName}
          saveName={saveName}
          sites={sites}
          addSite={addSite}
          renameSite={renameSite}
          removeSite={removeSite}
          siteBreakdown={siteBreakdown}
          onEnterSite={enterSite}
        />
      ) : (
      <div className="wel-shell">
        {/* Sidebar */}
        <aside className="wel-sidebar">
          <div className="wel-brand">
            <Droplets size={18} color="var(--teal)" />
            <Zap size={16} color="var(--amber)" style={{ marginLeft: -6 }} />
            {editingName ? (
              <input
                autoFocus
                value={projectName}
                onChange={(e) => setProjectName(e.target.value)}
                onBlur={saveName}
                onKeyDown={(e) => e.key === "Enter" && saveName()}
                className="wel-brand-input"
              />
            ) : (
              <span onClick={() => setEditingName(true)} className="wel-brand-text" title="點擊編輯專案名稱">
                {projectName}
              </span>
            )}
          </div>
          <button type="button" className="wel-back-to-sites" onClick={returnToSites}>
            <ChevronRight size={14} style={{ transform: "rotate(180deg)" }} />
            <span className="wel-back-current">{isAllSites ? "全部案場（總覽）" : (siteById[currentSiteId]?.name || "未選擇案場")}</span>
            <span className="wel-back-label">返回案場列表</span>
          </button>
          <nav className="wel-nav">
            <NavBtn icon={<HardHat size={16} />} label="師傅出工" active={tab === "labor"} onClick={() => setTab("labor")} />
            <NavBtn icon={<ClipboardList size={16} />} label="叫貨紀錄" active={tab === "orders"} onClick={() => setTab("orders")} />
            <NavBtn icon={<Camera size={16} />} label="收貨單" active={tab === "receipts"} onClick={() => setTab("receipts")} />
            <NavBtn icon={<Calculator size={16} />} label="估價單" active={tab === "estimate"} onClick={() => setTab("estimate")} />
            <NavBtn icon={<Boxes size={16} />} label="庫存現況" active={tab === "inventory"} onClick={() => setTab("inventory")} />
            <NavBtn icon={<Tags size={16} />} label="品項類別" active={tab === "categories"} onClick={() => setTab("categories")} />
            <NavBtn icon={<Handshake size={16} />} label="發包與領款" active={tab === "contracts"} onClick={() => setTab("contracts")} />
            <NavBtn icon={<ListOrdered size={16} />} label="項目總表" active={tab === "summary"} onClick={() => setTab("summary")} />
            <NavBtn icon={<MessageSquare size={16} />} label="備註討論" active={tab === "discussion"} onClick={() => setTab("discussion")} />
            <NavBtn icon={<NotebookPen size={16} />} label="大項記事本" active={tab === "notebook"} onClick={() => setTab("notebook")} />
            <NavBtn icon={<Workflow size={16} />} label="施工流程" active={tab === "workflow"} onClick={() => setTab("workflow")} />
            <NavBtn icon={<Wallet size={16} />} label="甲方收款" active={tab === "client"} onClick={() => setTab("client")} />
            <NavBtn icon={<PiggyBank size={16} />} label="投入成本" active={tab === "capital"} onClick={() => setTab("capital")} />
            <NavBtn icon={<Hammer size={16} />} label="固定資產" active={tab === "assets"} onClick={() => setTab("assets")} />
            <NavBtn icon={<Receipt size={16} />} label="其他支出" active={tab === "otherexpenses"} onClick={() => setTab("otherexpenses")} />
            <NavBtn icon={<LayoutDashboard size={16} />} label="儀表板" active={tab === "dashboard"} onClick={() => setTab("dashboard")} />
          </nav>
          <div className="wel-sidebar-foot">
            {saving ? (
              <span className="wel-saving"><Loader2 size={12} className="wel-spin" /> 同步中</span>
            ) : (
              <span className="wel-saving">✓ 資料已同步（共用）</span>
            )}
            <button type="button" className="wel-btn-ghost" style={{ marginTop: 8, width: "100%" }} onClick={lockApp}>鎖定（需重新輸入密碼）</button>
          </div>
        </aside>

        {/* Main */}
        <main className="wel-main">
          {tab === "dashboard" && (
            <Dashboard
              totalCost={totalCost}
              thisMonthCost={thisMonthCost}
              meterDigits={meterDigits}
              categoryBreakdown={categoryBreakdown}
              monthlyTrend={monthlyTrend}
              months={months}
              monthFilter={monthFilter}
              setMonthFilter={setMonthFilter}
              inventory={inventory}
              catById={catById}
              orderCount={filteredOrders.length}
              laborTotalAllTime={laborTotalAllTime}
              grandTotal={projectGrandTotal}
              workerCount={workers.length}
              isAllSites={isAllSites}
              siteBreakdown={siteBreakdown}
              currentSiteName={siteById[currentSiteId]?.name}
              contractPaidTotal={contractPaidTotal}
              contractTotalPrice={contractTotalPrice}
              clientPaidTotal={clientPaidTotal}
              profitLoss={profitLoss}
              openDiscussionCount={openDiscussionCount}
              siteDiscussionItems={siteDiscussionItems}
              drillData={isAllSites ? {
                sites, siteById, catById, workerById, workLogRate, contracts, orders, workLogs, payments, otherExpenses, fixedAssets, clientPayments, capitalContributions,
                goTo: (siteKey, t) => { if (siteKey !== "all" && siteKey !== "__orphan") setCurrentSiteId(siteKey); setTab(t); },
                initial: pendingDrill, clearInitial: () => setPendingDrill(null),
              } : undefined}
            />
          )}
          {tab === "orders" && (
            <OrdersTab
              orders={siteOrders}
              categories={categories}
              catById={catById}
              siteById={siteById}
              orderForm={orderForm}
              setOrderForm={setOrderForm}
              submitOrder={submitOrder}
              deleteOrder={deleteOrder}
              updateOrder={updateOrder}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              materialItems={materialItems}
              addMaterialItem={addMaterialItem}
              updateMaterialItem={updateMaterialItem}
              removeMaterialItem={removeMaterialItem}
              suppliers={suppliers}
              addSupplier={addSupplier}
              updateSupplier={updateSupplier}
              removeSupplier={removeSupplier}
              showToast={showToast}
              confirmScannedOrders={confirmScannedOrders}
            />
          )}
          {tab === "estimate" && (
            <EstimateTab
              estimates={siteEstimates}
              orders={siteOrders}
              categories={categories}
              catById={catById}
              materialItems={materialItems}
              siteById={siteById}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              currentSiteId={currentSiteId}
              estimateFiles={estimateFilesBySite[currentSiteId] || []}
              addEstimateFile={addEstimateFile}
              deleteEstimateFile={deleteEstimateFile}
              openContractFile={openContractFile}
              addEstimateLine={addEstimateLine}
              updateEstimateLine={updateEstimateLine}
              deleteEstimateLine={deleteEstimateLine}
              confirmScannedEstimate={confirmScannedEstimate}
              showToast={showToast}
            />
          )}
          {tab === "receipts" && (
            <ReceiptsTab
              receipts={siteReceipts}
              siteById={siteById}
              suppliers={suppliers}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              addReceipt={addReceipt}
              updateReceipt={updateReceipt}
              deleteReceipt={deleteReceipt}
              showToast={showToast}
            />
          )}
          {tab === "inventory" && (
            <InventoryTab
              inventory={inventory}
              catById={catById}
              categories={categories}
              orders={siteOrders}
              usages={siteUsages}
              siteById={siteById}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              consumeFromInventory={consumeFromInventory}
              reassignUsages={reassignUsages}
              deleteUsagesByKey={deleteUsagesByKey}
              deleteUsage={deleteUsage}
              showToast={showToast}
            />
          )}
          {tab === "labor" && (
            <LaborTab
              workers={workers}
              workLogs={siteWorkLogsAll}
              workerById={workerById}
              siteById={siteById}
              newWorker={newWorker}
              setNewWorker={setNewWorker}
              addWorker={addWorker}
              updateWorkerRate={updateWorkerRate}
              removeWorker={removeWorker}
              workLogForm={workLogForm}
              setWorkLogForm={setWorkLogForm}
              submitWorkLog={submitWorkLog}
              deleteWorkLog={deleteWorkLog}
              updateWorkLog={updateWorkLog}
              laborRange={laborRange}
              setLaborRange={setLaborRange}
              laborSummaryByWorker={laborSummaryByWorker}
              laborRangeTotal={laborRangeTotal}
              laborRangeDays={laborRangeDays}
              laborTotalAllTime={laborTotalAllTime}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              contractWorkLogs={siteContractWorkLogs}
              contractRows={contractRows}
              addAttendance={addAttendance}
              deleteAttendance={deleteAttendance}
              updateAttendance={updateAttendance}
            />
          )}
          {tab === "categories" && (
            <CategoriesTab
              categories={categories}
              orders={orders}
              usages={usages}
              newCatName={newCatName}
              setNewCatName={setNewCatName}
              addCategory={addCategory}
              removeCategory={removeCategory}
              renameCategory={renameCategory}
              materialItems={materialItems}
              addMaterialItem={addMaterialItem}
              updateMaterialItem={updateMaterialItem}
              removeMaterialItem={removeMaterialItem}
              confirmScannedItems={confirmScannedItems}
              showToast={showToast}
            />
          )}
          {tab === "contracts" && (
            <ContractsTab
              updateContract={updateContract}
              renameVendor={renameVendor}
              addContractFile={addContractFile}
              deleteContractFile={deleteContractFile}
              openContractFile={openContractFile}
              contractRows={contractRows}
              payments={sitePayments}
              siteById={siteById}
              contractForm={contractForm}
              setContractForm={setContractForm}
              submitContract={submitContract}
              deleteContract={deleteContract}
              paymentForm={paymentForm}
              setPaymentForm={setPaymentForm}
              submitPayment={submitPayment}
              deletePayment={deletePayment}
              updatePayment={updatePayment}
              contractPaidTotal={contractPaidTotal}
              contractTotalPrice={contractTotalPrice}
              isAllSites={isAllSites}
              floorItems={siteFloorItems}
              addTemplateRow={addTemplateRow}
              updateTemplateRow={updateTemplateRow}
              removeTemplateRow={removeTemplateRow}
              applyTemplateToFloors={applyTemplateToFloors}
              addFloorItem={addFloorItem}
              updateFloorItem={updateFloorItem}
              deleteFloorItem={deleteFloorItem}
              deleteFloorGroup={deleteFloorGroup}
              addAttendance={addAttendance}
              deleteAttendance={deleteAttendance}
              updateAttendance={updateAttendance}
              contractWorkLogs={siteContractWorkLogs}
              contractItems={siteContractItems}
              addContractItem={addContractItem}
              addContractItemsBatch={addContractItemsBatch}
              addPaymentFromItem={addPaymentFromItem}
              updateContractItem={updateContractItem}
              deleteContractItem={deleteContractItem}
              deleteContractItemsByFloor={deleteContractItemsByFloor}
              deleteContractItemsByName={deleteContractItemsByName}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              showToast={showToast}
            />
          )}
          {tab === "client" && (
            <ClientPaymentsTab
              clientPayments={siteClientPayments}
              siteById={siteById}
              currentSite={siteById[currentSiteId]}
              clientPaymentForm={clientPaymentForm}
              setClientPaymentForm={setClientPaymentForm}
              submitClientPayment={submitClientPayment}
              deleteClientPayment={deleteClientPayment}
              updateClientPayment={updateClientPayment}
              clientPaidTotal={clientPaidTotal}
              projectGrandTotal={projectGrandTotal}
              profitLoss={profitLoss}
              isAllSites={isAllSites}
              currentSiteId={currentSiteId}
              templateRows={siteClientTemplateRows}
              floorItems={siteClientFloorItems}
              clientFloorTotal={clientFloorTotal}
              addClientTemplateRow={addClientTemplateRow}
              addClientFloorItem={addClientFloorItem}
              addClientFloorItemsBatch={addClientFloorItemsBatch}
              updateClientFloorItem={updateClientFloorItem}
              deleteClientFloorItem={deleteClientFloorItem}
              deleteClientFloorItemsByFloor={deleteClientFloorItemsByFloor}
              deleteClientFloorItemsByName={deleteClientFloorItemsByName}
              addClientPaymentFromItem={addClientPaymentFromItem}
              updateSiteClientPricing={updateSiteClientPricing}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              showToast={showToast}
            />
          )}
          {tab === "summary" && (
            <MasterSummaryTab
              currentSiteName={siteById[currentSiteId]?.name}
              isAllSites={isAllSites}
              clientPaidTotal={clientPaidTotal}
              clientPayments={siteClientPayments}
              clientFloorItems={siteClientFloorItems}
              totalCost={totalCost}
              laborTotalAllTime={laborTotalAllTime}
              categoryBreakdown={categoryBreakdown}
              laborSummaryByWorker={laborSummaryByWorker}
              workerById={workerById}
              contractRows={contractRows}
              floorItems={siteFloorItems}
              projectGrandTotal={projectGrandTotal}
              profitLoss={profitLoss}
              orders={siteOrders}
              usages={siteUsages}
              catById={catById}
              contractPayments={sitePayments}
              contractItems={siteContractItems}
              discussionItems={siteDiscussionItems}
              contractorsTotalPrice={contractTotalPrice}
            />
          )}
          {tab === "discussion" && (
            <DiscussionTab
              discussionItems={siteDiscussionItems}
              siteById={siteById}
              discussionForm={discussionForm}
              setDiscussionForm={setDiscussionForm}
              submitDiscussion={submitDiscussion}
              toggleDiscussionResolved={toggleDiscussionResolved}
              updateDiscussionResult={updateDiscussionResult}
              updateDiscussion={updateDiscussion}
              deleteDiscussion={deleteDiscussion}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              currentSiteId={currentSiteId}
              showToast={showToast}
            />
          )}
          {tab === "workflow" && <WorkflowTab workflow={workflow} saveWorkflow={saveWorkflow} showToast={showToast} />}
          {tab === "notebook" && (
            <NotebookTab
              notebookItems={siteNotebookItems}
              siteById={siteById}
              notebookForm={notebookForm}
              setNotebookForm={setNotebookForm}
              submitNotebookItem={submitNotebookItem}
              updateNotebookItem={updateNotebookItem}
              deleteNotebookItem={deleteNotebookItem}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              showToast={showToast}
            />
          )}
          {tab === "capital" && (
            <CapitalTab
              capitalContributions={siteCapitalContributions}
              siteById={siteById}
              capitalForm={capitalForm}
              setCapitalForm={setCapitalForm}
              submitCapitalContribution={submitCapitalContribution}
              deleteCapitalContribution={deleteCapitalContribution}
              updateCapitalContribution={updateCapitalContribution}
              capitalTotal={capitalTotal}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
            />
          )}
          {tab === "assets" && (
            <FixedAssetsTab
              fixedAssets={siteFixedAssets}
              siteById={siteById}
              fixedAssetForm={fixedAssetForm}
              setFixedAssetForm={setFixedAssetForm}
              submitFixedAsset={submitFixedAsset}
              deleteFixedAsset={deleteFixedAsset}
              updateFixedAsset={updateFixedAsset}
              fixedAssetsTotal={fixedAssetsTotal}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
            />
          )}
          {tab === "otherexpenses" && (
            <OtherExpensesTab
              otherExpenses={siteOtherExpenses}
              siteById={siteById}
              otherExpenseForm={otherExpenseForm}
              setOtherExpenseForm={setOtherExpenseForm}
              submitOtherExpense={submitOtherExpense}
              deleteOtherExpense={deleteOtherExpense}
              updateOtherExpense={updateOtherExpense}
              otherExpensesTotal={otherExpensesTotal}
              materialItems={materialItems}
              suppliers={suppliers}
              addSupplier={addSupplier}
              updateSupplier={updateSupplier}
              removeSupplier={removeSupplier}
              isAllSites={isAllSites}
              sites={sites}
              setCurrentSiteId={setCurrentSiteId}
              showToast={showToast}
            />
          )}
        </main>
      </div>
      )}
    </div>
  );
}

function AllSitesNotice({ sites, setCurrentSiteId, action }) {
  return (
    <div className="wel-notice wel-notice-action">
      <span>目前檢視「全部案場」，新增紀錄前請先切換到單一案場才能{action}。以下列表為所有案場的紀錄。</span>
      <select
        defaultValue=""
        onChange={(e) => { if (e.target.value) setCurrentSiteId(e.target.value); }}
      >
        <option value="" disabled>切換到案場…</option>
        {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
      </select>
    </div>
  );
}

function NavBtn({ icon, label, active, onClick }) {
  return (
    <button className={"wel-navbtn" + (active ? " active" : "")} onClick={onClick}>
      {icon}
      <span>{label}</span>
    </button>
  );
}

function ExportBar({ onExcel, onWord, label }) {
  return (
    <div className="wel-export-bar">
      {label && <span className="muted" style={{ fontSize: 11.5, fontFamily: "var(--font-mono)" }}>{label}</span>}
      <button type="button" className="wel-export-btn" onClick={onExcel}><FileSpreadsheet size={13} /> 匯出 Excel</button>
      <button type="button" className="wel-export-btn" onClick={onWord}><FileText size={13} /> 匯出 Word</button>
    </div>
  );
}

const DRILL_KINDS = {
  collected: { label: "收入（甲方收款）", tab: "client", tabName: "甲方收款" },
  material: { label: "材料（叫貨）", tab: "orders", tabName: "叫貨紀錄" },
  labor: { label: "人工（出工）", tab: "labor", tabName: "師傅出工" },
  contract: { label: "發包（領款）", tab: "contracts", tabName: "發包與領款" },
  other: { label: "其他支出", tab: "otherexpenses", tabName: "其他支出" },
  asset: { label: "固定資產", tab: "assets", tabName: "固定資產" },
  capital: { label: "投入成本（參考）", tab: "capital", tabName: "投入成本" },
};
const EXPENSE_KINDS = ["material", "labor", "contract", "other", "asset"];

// 全部案場總結算：收入、各項支出、總淨損益；投入成本只做參考（已在收入中回收，不計入損益）。
// 每個數字都能點開，看到是哪些紀錄加起來的。
function SettlementCard({ siteBreakdown, drillData }) {
  const sum = (k) => siteBreakdown.reduce((acc, b) => acc + (b[k] || 0), 0);
  const T = { id: "all", collected: sum("collected"), material: sum("material"), labor: sum("labor"), contract: sum("contract"), other: sum("other"), asset: sum("asset"), total: sum("total"), profit: sum("profit"), capital: sum("capital") };
  const signed = (n) => (n >= 0 ? "+" : "") + fmtMoney(n);
  const color = (n) => (n >= 0 ? "var(--green)" : "var(--red)");
  const [drill, setDrill] = useState(drillData && drillData.initial ? drillData.initial : null);
  useEffect(() => { if (drillData && drillData.initial) drillData.clearInitial(); }, []); // eslint-disable-line react-hooks/exhaustive-deps

  const toRow = (name, b) => ({ 案場: name, 收入: b.collected, 材料: b.material, 人工: b.labor, 發包: b.contract, 其他支出: b.other, 固定資產: b.asset, 總支出: b.total, 淨損益: b.profit, "投入成本（參考）": b.capital });
  const handleExportExcel = () => exportExcel("全部案場總結算", [{ name: "總結算", rows: [...siteBreakdown.map((b) => toRow(b.name, b)), toRow("全部案場合計", T)] }]);
  const handleExportWord = () => exportWord("全部案場總結算", "全部案場總結算", [{
    title: "各案場與合計",
    note: "淨損益 ＝ 收入 － 總支出。投入成本已在收入中回收，只列出供參考，不計入損益。",
    headers: ["案場", "收入", "材料", "人工", "發包", "其他支出", "固定資產", "總支出", "淨損益", "投入成本（參考）"],
    rows: [...siteBreakdown.map((b) => [b.name, ...Object.values(toRow(b.name, b)).slice(1).map(fmtMoney)]), ["全部案場合計", ...Object.values(toRow("", T)).slice(1).map(fmtMoney)]],
  }]);

  const cell = { padding: "6px 10px", textAlign: "right", whiteSpace: "nowrap" };
  const isActive = (siteKey, kind) => drill && drill.siteKey === siteKey && drill.kind === kind;
  const num = (siteKey, kind, text, tone) => (
    <button type="button" className={"wel-num-link" + (isActive(siteKey, kind) ? " active" : "")} style={tone ? { color: tone } : undefined} onClick={() => setDrill({ siteKey, kind })} title="點一下看明細">{text}</button>
  );

  // ---- drill-down panel ----
  const d = drillData;
  const siteIds = d ? new Set(d.sites.map((s) => s.id)) : new Set();
  const siteName = (key) => (key === "all" ? "全部案場合計" : key === "__orphan" ? "（已刪除案場的紀錄）" : (d && d.siteById[key]?.name) || "");
  const inScope = (x, key) => (key === "all" ? true : key === "__orphan" ? !siteIds.has(x.siteId) : x.siteId === key);
  const recordSite = (x) => (d && d.siteById[x.siteId]?.name) || "（已刪除案場）";

  const buildRecords = (kind, key) => {
    const contractById = {};
    d.contracts.forEach((c) => (contractById[c.id] = c));
    const defs = {
      material: { headers: ["類別", "品項", "數量 × 單價", "廠商"], list: d.orders, row: (o) => ({ cells: [d.catById[o.categoryId]?.name || "未分類", o.itemName, `${fmtNum(o.quantity)} ${o.unit || ""} × ${fmtMoney(o.unitPrice)}`, o.supplier || ""], amount: Number(o.amount) || 0 }) },
      labor: { headers: ["師傅", "天數 × 日薪", "備註"], list: d.workLogs, row: (l) => ({ cells: [d.workerById[l.workerId]?.name || "已刪除師傅", `${fmtNum(l.days)} 天 × ${fmtMoney(d.workLogRate(l))}`, l.note || ""], amount: (Number(l.days) || 0) * d.workLogRate(l) }) },
      contract: { headers: ["發包項目", "廠商", "備註"], list: d.payments, row: (p) => ({ cells: [contractById[p.contractId]?.name || "已刪除項目", contractById[p.contractId]?.contractor || "", p.note || ""], amount: Number(p.amount) || 0 }) },
      other: { headers: ["項目", "廠商", "備註"], list: d.otherExpenses, row: (e) => ({ cells: [e.itemName, e.supplier || "", e.note || ""], amount: Number(e.amount) || 0 }) },
      asset: { headers: ["品名", "數量 × 單價", "發票"], list: d.fixedAssets, row: (f) => ({ cells: [f.name, `${fmtNum(f.quantity)} × ${fmtMoney(f.amount)}`, f.hasInvoice ? "有發票" : "無發票"], amount: (Number(f.amount) || 0) * (Number(f.quantity) || 1) }) },
      collected: { headers: ["請款項目", "備註"], list: d.clientPayments, row: (p) => ({ cells: [p.item || "", p.note || ""], amount: Number(p.amount) || 0 }) },
      capital: { headers: ["備註"], list: d.capitalContributions, row: (c) => ({ cells: [c.note || ""], amount: Number(c.amount) || 0 }) },
    };
    const def = defs[kind];
    const rows = def.list.filter((x) => inScope(x, key)).map((x) => ({ id: x.id, date: x.date || "", site: recordSite(x), ...def.row(x) })).sort((a, b) => (a.date < b.date ? 1 : -1));
    return { headers: def.headers, rows, total: rows.reduce((s, r) => s + r.amount, 0) };
  };

  const panelBox = { marginTop: 14, padding: 14, border: "1px solid var(--teal)", borderRadius: 10, background: "rgba(74,144,164,0.06)" };
  const renderDrill = () => {
    if (!drill || !d) return null;
    const { siteKey, kind } = drill;
    const b = siteKey === "all" ? T : siteBreakdown.find((x) => x.id === siteKey);
    if (!b) return null;
    const title = (text) => (
      <div className="wel-card-title-row" style={{ marginBottom: 8 }}>
        <div className="wel-card-title"><ScanLine size={14} /> {siteName(siteKey)} · {text}</div>
        <button type="button" className="wel-btn-ghost" onClick={() => setDrill(null)}>關閉明細</button>
      </div>
    );
    if (kind === "total") {
      return (
        <div style={panelBox}>
          {title("總支出由哪些項目組成")}
          <table className="wel-table">
            <thead><tr><th>項目</th><th className="right">金額</th><th className="right">佔總支出</th></tr></thead>
            <tbody>
              {EXPENSE_KINDS.map((k) => (
                <tr key={k}>
                  <td>{DRILL_KINDS[k].label}</td>
                  <td className="mono" style={cell}>{num(siteKey, k, fmtMoney(b[k]))}</td>
                  <td className="mono muted" style={cell}>{b.total ? ((b[k] / b.total) * 100).toFixed(1) + "%" : "—"}</td>
                </tr>
              ))}
              <tr style={{ borderTop: "2px solid var(--border)", background: "var(--surface2)" }}>
                <td className="strong">總支出</td><td className="mono strong" style={cell}>{fmtMoney(b.total)}</td><td className="mono muted" style={cell}>100%</td>
              </tr>
            </tbody>
          </table>
          <div className="muted" style={{ fontSize: 11.5, marginTop: 6 }}>點上面任一金額，可以再往下看是哪幾筆紀錄。</div>
        </div>
      );
    }
    if (kind === "profit") {
      return (
        <div style={panelBox}>
          {title("淨損益怎麼算出來的")}
          <table className="wel-table">
            <tbody>
              <tr><td>收入（甲方收款）</td><td className="mono" style={cell}>{num(siteKey, "collected", fmtMoney(b.collected))}</td></tr>
              <tr><td>減：總支出</td><td className="mono" style={cell}>{num(siteKey, "total", fmtMoney(b.total))}</td></tr>
              <tr style={{ borderTop: "2px solid var(--border)", background: "var(--surface2)" }}>
                <td className="strong">淨損益（{b.profit >= 0 ? "毛利" : "虧損"}）</td><td className="mono strong" style={{ ...cell, color: color(b.profit) }}>{signed(b.profit)}</td>
              </tr>
            </tbody>
          </table>
          <div className="muted" style={{ fontSize: 11.5, marginTop: 6 }}>投入成本已在收入中回收，不計入損益。點上面的金額可以再往下追。</div>
        </div>
      );
    }
    const info = DRILL_KINDS[kind];
    const rec = buildRecords(kind, siteKey);
    const matches = Math.abs(rec.total - (b[kind] || 0)) < 1;
    return (
      <div style={panelBox}>
        {title(`${info.label}：${rec.rows.length} 筆，合計 ${fmtMoney(rec.total)}`)}
        <div style={{ maxHeight: 380, overflow: "auto", border: "1px solid var(--border)", borderRadius: 8 }}>
          <table className="wel-table">
            <thead><tr><th>日期</th><th>案場</th>{rec.headers.map((h) => <th key={h}>{h}</th>)}<th className="right">金額</th></tr></thead>
            <tbody>
              {rec.rows.length === 0 && <tr><td colSpan={rec.headers.length + 3}><Empty text="沒有這類紀錄" /></td></tr>}
              {rec.rows.map((r) => (
                <tr key={r.id}>
                  <td className="mono">{r.date}</td>
                  <td className="muted">{r.site}</td>
                  {r.cells.map((c, i) => <td key={i}>{c || "—"}</td>)}
                  <td className="mono strong" style={cell}>{fmtMoney(r.amount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginTop: 8 }}>
          <span className="muted" style={{ fontSize: 12 }}>{matches ? "✓ 這些紀錄加起來，和上面表格的數字一致" : `⚠ 與表格數字不同（表格 ${fmtMoney(b[kind] || 0)}），請回報給我檢查`}</span>
          <button type="button" className="wel-btn-ghost" onClick={() => d.goTo(siteKey, info.tab)}>到「{info.tabName}」查看／修改這些紀錄 →</button>
        </div>
      </div>
    );
  };

  const tile = (label, value, tone, siteKey, kind) => (
    <button type="button" className="wel-labor-stat wel-stat-btn" onClick={() => setDrill({ siteKey, kind })} title="點一下看明細">
      <span>{label}</span>
      <b style={{ color: tone }}>{value}</b>
    </button>
  );
  const rowCells = (key, b, strong) => (
    <>
      <td className="mono" style={cell}>{num(key, "collected", fmtMoney(b.collected))}</td>
      <td className="mono" style={cell}>{num(key, "material", fmtMoney(b.material))}</td>
      <td className="mono" style={cell}>{num(key, "labor", fmtMoney(b.labor))}</td>
      <td className="mono" style={cell}>{num(key, "contract", fmtMoney(b.contract))}</td>
      <td className="mono" style={cell}>{num(key, "other", fmtMoney(b.other))}</td>
      <td className="mono" style={cell}>{num(key, "asset", fmtMoney(b.asset))}</td>
      <td className="mono strong" style={cell}>{num(key, "total", fmtMoney(b.total))}</td>
      <td className="mono strong" style={cell}>{num(key, "profit", signed(b.profit), color(b.profit))}</td>
      <td className="mono" style={cell}>{num(key, "capital", fmtMoney(b.capital))}</td>
    </>
  );
  return (
    <div className="wel-card">
      <div className="wel-card-title-row">
        <div className="wel-card-title"><Layers size={14} /> 全部案場總結算</div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>
      <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>點任何一個數字，就能看到它是哪些紀錄加起來的，也可以直接跳到那一頁修改。</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: 12 }}>
        {tile("總收入（甲方收款）", fmtMoney(T.collected), "var(--teal)", "all", "collected")}
        {tile("總支出", fmtMoney(T.total), "var(--amber)", "all", "total")}
        {tile("總淨損益（收入 － 總支出）", signed(T.profit), color(T.profit), "all", "profit")}
        {tile("總投入成本（參考）", fmtMoney(T.capital), "var(--teal)", "all", "capital")}
      </div>
      <div style={{ overflowX: "auto", marginTop: 14 }}>
        <table className="wel-table">
          <thead>
            <tr><th>案場</th><th className="right">收入</th><th className="right">材料</th><th className="right">人工</th><th className="right">發包</th><th className="right">其他支出</th><th className="right">固定資產</th><th className="right">總支出</th><th className="right">淨損益</th><th className="right">投入成本（參考）</th></tr>
          </thead>
          <tbody>
            {siteBreakdown.map((b) => (
              <tr key={b.id}>
                <td style={{ whiteSpace: "nowrap" }}>{b.name}</td>
                {rowCells(b.id, b)}
              </tr>
            ))}
            <tr style={{ borderTop: "2px solid var(--border)", background: "var(--surface2)" }}>
              <td className="strong" style={{ whiteSpace: "nowrap" }}>全部案場合計</td>
              {rowCells("all", T)}
            </tr>
          </tbody>
        </table>
      </div>
      {renderDrill()}
      <div className="muted" style={{ fontSize: 11.5, marginTop: 10 }}>
        總支出 ＝ 材料 ＋ 人工 ＋ 發包 ＋ 其他支出 ＋ 固定資產；總淨損益 ＝ 收入 － 總支出。投入成本已在收入中回收，只列出供參考，不計入損益。
      </div>
    </div>
  );
}

function Dashboard({ totalCost, thisMonthCost, meterDigits, categoryBreakdown, monthlyTrend, months, monthFilter, setMonthFilter, inventory, catById, orderCount, laborTotalAllTime, grandTotal, workerCount, isAllSites, siteBreakdown, currentSiteName, contractPaidTotal, contractTotalPrice, clientPaidTotal, profitLoss, openDiscussionCount, siteDiscussionItems, drillData }) {
  const lowStock = inventory.filter((i) => i.remaining <= 0 && i.ordered > 0);
  const maxCat = Math.max(1, ...categoryBreakdown.map((d) => d.value));
  const maxSite = Math.max(1, ...siteBreakdown.map((d) => d.total));

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">總覽 · OVERVIEW{isAllSites ? " · 全部案場" : ""}</div>
          <h1 className="wel-h1">{isAllSites ? "全案場總支出儀表板" : `${currentSiteName || ""} 支出儀表板`}</h1>
        </div>
        <div className="wel-filter">
          <select value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
            <option value="all">全部月份</option>
            {months.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
        </div>
      </div>

      {/* Meter */}
      <div className="wel-meter-card">
        <div className="wel-meter-label">
          <Gauge size={14} /> 累計總支出（材料＋人工＋發包領款，如同水電錶累計度數）
        </div>
        <div className="wel-meter">
          {meterDigits.map((d, i) => (
            <span key={i} className="wel-meter-digit">{d}</span>
          ))}
          <span className="wel-meter-unit">元</span>
        </div>
        <div className="wel-meter-sub">
          <span>材料支出：<b>{fmtMoney(totalCost)}</b></span>
          <span>人工薪資：<b>{fmtMoney(laborTotalAllTime)}</b></span>
          <span>發包已付款：<b>{fmtMoney(contractPaidTotal)}</b>{contractTotalPrice > 0 && <span className="muted"> / 發包總價 {fmtMoney(contractTotalPrice)}</span>}</span>
          <span>本月材料支出：<b>{fmtMoney(thisMonthCost)}</b></span>
          <span>叫貨筆數：<b>{orderCount}</b> 筆</span>
          <span>師傅人數：<b>{workerCount}</b> 人</span>
        </div>
      </div>

      {/* Profit / loss vs. client collections */}
      <div className={"wel-profit-card" + (profitLoss >= 0 ? " positive" : " negative")}>
        <div className="wel-profit-col">
          <span className="wel-profit-label"><Wallet size={13} /> 甲方累計收款</span>
          <b className="mono">{fmtMoney(clientPaidTotal)}</b>
        </div>
        <div className="wel-profit-sign">－</div>
        <div className="wel-profit-col">
          <span className="wel-profit-label"><Gauge size={13} /> 專案總支出</span>
          <b className="mono">{fmtMoney(grandTotal)}</b>
        </div>
        <div className="wel-profit-sign">=</div>
        <div className="wel-profit-col">
          <span className="wel-profit-label">{profitLoss >= 0 ? <TrendingUp size={13} /> : <TrendingDown size={13} />} {profitLoss >= 0 ? "毛利" : "虧損"}</span>
          <b className="mono wel-profit-value">{profitLoss >= 0 ? "+" : ""}{fmtMoney(profitLoss)}</b>
        </div>
      </div>

      {isAllSites && siteBreakdown.length > 0 && <SettlementCard siteBreakdown={siteBreakdown} drillData={drillData} />}

      {isAllSites && siteBreakdown.length > 0 && (
        <div className="wel-card">
          <div className="wel-card-title"><Building2 size={14} /> 各案場支出比較</div>
          <div className="wel-pipes">
            {siteBreakdown.map((s, i) => (
              <div key={s.id} className="wel-pipe-row">
                <div className="wel-pipe-label">{s.name}</div>
                <div className="wel-pipe-track">
                  <div
                    className="wel-pipe-fill"
                    style={{ width: `${Math.max(6, (s.total / maxSite) * 100)}%`, background: i % 2 === 0 ? "var(--teal)" : "var(--amber)" }}
                  />
                </div>
                <div className="wel-pipe-value">{fmtMoney(s.total)}</div>
              </div>
            ))}
          </div>
          <div className="wel-site-profit-list">
            {siteBreakdown.map((s) => (
              <div key={s.id} className="wel-site-profit-row">
                <span className="muted">{s.name}</span>
                <span className="mono">{fmtMoney(s.collected)} 收款</span>
                <span className={"mono" + (s.profit >= 0 ? " wel-profit-pos" : " wel-profit-neg")}>
                  {s.profit >= 0 ? "+" : ""}{fmtMoney(s.profit)}
                </span>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="wel-grid2">
        {/* Category pipes */}
        <div className="wel-card">
          <div className="wel-card-title">各品項支出（管線粗細＝金額）</div>
          {categoryBreakdown.length === 0 && <Empty text="尚無叫貨紀錄" />}
          <div className="wel-pipes">
            {categoryBreakdown.map((d) => (
              <div key={d.id} className="wel-pipe-row">
                <div className="wel-pipe-label">{d.name}</div>
                <div className="wel-pipe-track">
                  <div
                    className="wel-pipe-fill"
                    style={{ width: `${Math.max(6, (d.value / maxCat) * 100)}%`, background: d.color }}
                  />
                </div>
                <div className="wel-pipe-value">{fmtMoney(d.value)}</div>
              </div>
            ))}
          </div>
        </div>

        {/* Monthly trend */}
        <div className="wel-card">
          <div className="wel-card-title"><TrendingUp size={14} /> 每月支出趨勢</div>
          {monthlyTrend.length === 0 ? (
            <Empty text="尚無月度資料" />
          ) : (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={monthlyTrend} margin={{ top: 8, right: 8, left: -18, bottom: 0 }}>
                <CartesianGrid stroke="var(--border)" strokeDasharray="3 3" vertical={false} />
                <XAxis dataKey="month" tick={{ fill: "var(--text-muted)", fontSize: 11, fontFamily: "var(--font-mono)" }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                <YAxis tick={{ fill: "var(--text-muted)", fontSize: 11, fontFamily: "var(--font-mono)" }} axisLine={false} tickLine={false} width={54} tickFormatter={(v) => (v >= 1000 ? `${Math.round(v / 1000)}k` : v)} />
                <Tooltip
                  formatter={(v) => fmtMoney(v)}
                  contentStyle={{ background: "var(--surface2)", border: "1px solid var(--border)", borderRadius: 6, fontFamily: "var(--font-body)", fontSize: 12 }}
                  labelStyle={{ color: "var(--text)" }}
                  cursor={{ fill: "rgba(255,255,255,0.04)" }}
                />
                <Bar dataKey="total" radius={[4, 4, 0, 0]}>
                  {monthlyTrend.map((_, i) => (
                    <Cell key={i} fill={i === monthlyTrend.length - 1 ? "var(--amber)" : "var(--teal)"} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>

      {lowStock.length > 0 && (
        <div className="wel-card wel-alert">
          <div className="wel-card-title" style={{ color: "var(--red)" }}>
            <AlertTriangle size={14} /> 庫存告警（剩餘 ≤ 0）
          </div>
          <div className="wel-alert-list">
            {lowStock.map((i) => (
              <span key={i.key} className="wel-chip">
                {catById[i.categoryId]?.name || "未分類"} · {i.itemName}：{fmtNum(i.remaining)} {i.unit}
              </span>
            ))}
          </div>
        </div>
      )}

      {openDiscussionCount > 0 && (
        <div className="wel-card wel-alert-info">
          <div className="wel-card-title" style={{ color: "var(--teal)" }}>
            <MessageSquare size={14} /> 尚待處理的討論項目（{openDiscussionCount}）
          </div>
          <div className="wel-alert-list">
            {siteDiscussionItems.filter((d) => !d.resolved).slice(0, 8).map((d) => (
              <span key={d.id} className="wel-chip wel-chip-info">{d.topic}</span>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function OrdersTab({ orders, categories, catById, siteById, orderForm, setOrderForm, submitOrder, deleteOrder, updateOrder, isAllSites, sites, setCurrentSiteId, materialItems, addMaterialItem, updateMaterialItem, removeMaterialItem, suppliers, addSupplier, updateSupplier, removeSupplier, showToast, confirmScannedOrders }) {
  const [scanOpen, setScanOpen] = useState(false);
  const [bulkOpen, setBulkOpen] = useState(false);
  const [expandedOrderId, setExpandedOrderId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const startEditOrder = (o) => {
    setExpandedOrderId(o.id);
    setEditDraft({
      date: o.date, categoryId: o.categoryId, itemName: o.itemName, supplier: o.supplier || "",
      quantity: o.quantity, unit: o.unit || "", unitPrice: o.unitPrice, note: o.note || "",
      hasInvoice: !!o.hasInvoice, applyTax: !!o.applyTax, taxRate: o.taxRate ?? 5,
    });
  };
  const cancelEditOrder = () => { setExpandedOrderId(null); setEditDraft(null); };
  const saveEditOrder = (id) => {
    if (!editDraft.itemName.trim() || !editDraft.quantity || !editDraft.unitPrice) {
      showToast("請填寫「品項名稱」「數量」「單價」後再儲存");
      return;
    }
    updateOrder(id, { ...editDraft, amount: Number(editDraft.quantity) * Number(editDraft.unitPrice) });
    setExpandedOrderId(null);
    setEditDraft(null);
  };
  const amount = (Number(orderForm.quantity) || 0) * (Number(orderForm.unitPrice) || 0);
  const [manageOpen, setManageOpen] = useState(false);
  const [supplierManageOpen, setSupplierManageOpen] = useState(false);
  const [newMaterial, setNewMaterial] = useState({ name: "", unit: "", defaultPrice: "" });
  const [newSupplier, setNewSupplier] = useState("");
  const [openOrderMonths, setOpenOrderMonths] = useState({});
  const [invoiceFilter, setInvoiceFilter] = useState("all"); // all | yes | no
  const filteredOrders = useMemo(
    () => orders.filter((o) => (invoiceFilter === "all" ? true : invoiceFilter === "yes" ? !!o.hasInvoice : !o.hasInvoice)),
    [orders, invoiceFilter]
  );
  const sortedMaterials = useMemo(() => materialItems.slice().sort((a, b) => a.name.localeCompare(b.name, "zh-Hant")), [materialItems]);
  const sortedSuppliers = useMemo(() => suppliers.slice().sort((a, b) => a.name.localeCompare(b.name, "zh-Hant")), [suppliers]);

  const categoryMaterials = sortedMaterials.filter((m) => m.categoryId === orderForm.categoryId);
  const uncategorizedMaterials = sortedMaterials.filter((m) => !m.categoryId || !catById[m.categoryId]);
  const pickableMaterials = [...categoryMaterials, ...uncategorizedMaterials];
  const renderMaterialOption = (m) => (
    <option key={m.id} value={m.name}>
      {m.name}{m.unit ? `（${m.unit}）` : ""}{m.defaultPrice !== "" && m.defaultPrice != null ? ` · ${fmtMoney(m.defaultPrice)}` : ""}
    </option>
  );
  const handlePickCategory = (e) => {
    const categoryId = e.target.value;
    const stillValid = materialItems.some((m) => m.name === orderForm.itemName && (m.categoryId === categoryId || !m.categoryId || !catById[m.categoryId]));
    setOrderForm({ ...orderForm, categoryId, ...(stillValid ? {} : { itemName: "" }) });
  };
  const handlePick = (e) => {
    const val = e.target.value;
    if (val === "__manage__") { setManageOpen(true); return; }
    const m = materialItems.find((x) => x.name === val);
    setOrderForm({
      ...orderForm,
      itemName: val,
      unit: m && m.unit ? m.unit : orderForm.unit,
      unitPrice: m && m.defaultPrice !== "" && m.defaultPrice != null ? m.defaultPrice : orderForm.unitPrice,
      categoryId: m && m.categoryId ? m.categoryId : orderForm.categoryId,
    });
  };
  const handlePickSupplier = (e) => {
    const val = e.target.value;
    if (val === "__manage__") { setSupplierManageOpen(true); return; }
    setOrderForm({ ...orderForm, supplier: val });
  };

  const taxOf = (o) => (o.hasInvoice && o.applyTax ? (Number(o.amount) || 0) * (Number(o.taxRate) || 0) / 100 : 0);
  const exportRows = () => filteredOrders.map((o) => ({
    日期: o.date, 類別: catById[o.categoryId]?.name || "", 品項: o.itemName, 廠商: o.supplier || "",
    數量: Number(o.quantity) || 0, 單位: o.unit || "", 單價: Number(o.unitPrice) || 0, 金額: Number(o.amount) || 0,
    發票: o.hasInvoice ? "有發票" : "無發票", 稅額: taxOf(o), 備註: o.note || "",
  }));
  const handleExportExcel = () => exportExcel("叫貨紀錄", [{ name: "叫貨紀錄", rows: exportRows() }]);
  const handleExportWord = () => exportWord("叫貨紀錄", "叫貨紀錄", [{
    title: "叫貨紀錄",
    headers: ["日期", "類別", "品項", "廠商", "數量", "單位", "單價", "金額", "發票", "稅額", "備註"],
    rows: filteredOrders.map((o) => [o.date, catById[o.categoryId]?.name || "", o.itemName, o.supplier || "", o.quantity, o.unit || "", fmtMoney(o.unitPrice), fmtMoney(o.amount), o.hasInvoice ? "有發票" : "無發票", taxOf(o) ? fmtMoney(taxOf(o)) : "—", o.note || ""]),
  }]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">採購 · PURCHASE ORDERS</div>
          <h1 className="wel-h1">叫貨紀錄</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增叫貨紀錄" />}

      {!isAllSites && (
        <div style={{ marginBottom: 12 }}>
          {scanOpen ? (
            <ReceiptScanner
              mode="order" categories={categories} materialItems={materialItems} suppliers={suppliers}
              defaultCategoryId={orderForm.categoryId} onConfirmOrders={confirmScannedOrders}
              onClose={() => setScanOpen(false)} showToast={showToast}
            />
          ) : bulkOpen ? (
            <BulkOrderPanel
              categories={categories} catById={catById} materialItems={materialItems} suppliers={suppliers}
              defaultCategoryId={orderForm.categoryId} onConfirm={confirmScannedOrders}
              onClose={() => setBulkOpen(false)} showToast={showToast}
            />
          ) : (
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button type="button" className="wel-btn-ghost" onClick={() => setBulkOpen(true)}><ListChecks size={14} /> 多項一次叫貨（勾選多個品項一起輸入）</button>
              <button type="button" className="wel-btn-ghost" onClick={() => setScanOpen(true)}><Camera size={14} /> 拍照辨識收貨單（先校對再存入叫貨）</button>
            </div>
          )}
        </div>
      )}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={orderForm.date} onChange={(e) => setOrderForm({ ...orderForm, date: e.target.value })} />
          </Field>
          <Field label="類別（先選類別，品項會依類別列出）">
            <select value={orderForm.categoryId} onChange={handlePickCategory}>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="品項名稱" wide>
            <select value={pickableMaterials.some((m) => m.name === orderForm.itemName) ? orderForm.itemName : ""} onChange={handlePick}>
              <option value="" disabled>{pickableMaterials.length ? "請選擇品項…" : "此類別尚無品項，請先新增"}</option>
              {categoryMaterials.map(renderMaterialOption)}
              {uncategorizedMaterials.length > 0 && (
                <optgroup label="未歸類品項">{uncategorizedMaterials.map(renderMaterialOption)}</optgroup>
              )}
              <option value="__manage__">＋ 管理品項清單…</option>
            </select>
            {orderForm.itemName && !manageOpen && (
              <div className="muted" style={{ fontSize: 11.5, marginTop: 4 }}>已選擇：{orderForm.itemName}（單價已帶入品項清單裡設定的預設價，這筆可以再調整，不會回寫品項清單）</div>
            )}
            {manageOpen && (
              <div className="wel-material-manager">
                <div className="wel-material-add-row">
                  <input placeholder="新品項名稱" value={newMaterial.name} onChange={(e) => setNewMaterial({ ...newMaterial, name: e.target.value })} />
                  <input placeholder="單位" value={newMaterial.unit} onChange={(e) => setNewMaterial({ ...newMaterial, unit: e.target.value })} style={{ width: 70 }} />
                  <input type="number" placeholder="預設單價" value={newMaterial.defaultPrice} onChange={(e) => setNewMaterial({ ...newMaterial, defaultPrice: e.target.value })} style={{ width: 90 }} />
                  <button
                    type="button" className="wel-btn-ghost"
                    onClick={() => { addMaterialItem(newMaterial.name, newMaterial.unit, newMaterial.defaultPrice, orderForm.categoryId); setNewMaterial({ name: "", unit: "", defaultPrice: "" }); }}
                  >
                    <Plus size={13} /> 新增到「{catById[orderForm.categoryId]?.name || "此類別"}」
                  </button>
                </div>
                <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>
                  目前只顯示「{catById[orderForm.categoryId]?.name || "此類別"}」的品項（{categoryMaterials.length}）；要管理其他類別，請先在上方切換類別。
                </div>
                <div className="wel-manager-list">
                  {categoryMaterials.length === 0 && <span className="muted" style={{ fontSize: 12 }}>這個類別還沒有品項，可在上方新增</span>}
                  {categoryMaterials.map((m) => (
                    <div key={m.id} className="wel-manager-row">
                      <input
                        className="wel-manager-name"
                        value={m.name}
                        onChange={(e) => updateMaterialItem(m.id, { name: e.target.value })}
                        onBlur={(e) => { if (!e.target.value.trim()) updateMaterialItem(m.id, { name: m.name }); }}
                      />
                      <input className="wel-manager-mini" placeholder="單位" value={m.unit || ""} onChange={(e) => updateMaterialItem(m.id, { unit: e.target.value })} />
                      <input className="wel-manager-mini" type="number" placeholder="單價" value={m.defaultPrice === "" || m.defaultPrice == null ? "" : m.defaultPrice} onChange={(e) => updateMaterialItem(m.id, { defaultPrice: e.target.value })} />
                      <select className="wel-manager-mini" style={{ width: 110 }} value={m.categoryId || ""} onChange={(e) => updateMaterialItem(m.id, { categoryId: e.target.value })}>
                        <option value="">未歸類</option>
                        {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                      </select>
                      <button type="button" className="wel-icon-btn" onClick={() => removeMaterialItem(m.id)}><Trash2 size={13} /></button>
                    </div>
                  ))}
                </div>
                {uncategorizedMaterials.length > 0 && (
                  <div className="muted" style={{ fontSize: 11 }}>另有 {uncategorizedMaterials.length} 個未歸類品項，請到「品項類別」頁指定類別後就會出現在對應類別裡。</div>
                )}
                <div className="muted" style={{ fontSize: 11 }}>提示：直接修改上面的名稱／單位／單價即可更新，改名稱會自動同步更新過去的叫貨紀錄；設定類別後，之後挑選這個品項會自動帶入該類別（不影響過去已新增的紀錄）。</div>
                <button type="button" className="wel-btn-ghost" onClick={() => setManageOpen(false)}>收合品項管理</button>
              </div>
            )}
          </Field>
          <Field label="廠商 / 供應商" wide>
            <select value={sortedSuppliers.some((s) => s.name === orderForm.supplier) ? orderForm.supplier : ""} onChange={handlePickSupplier}>
              <option value="">（未指定）</option>
              {sortedSuppliers.map((s) => <option key={s.id} value={s.name}>{s.name}</option>)}
              <option value="__manage__">＋ 管理廠商清單…</option>
            </select>
            {supplierManageOpen && (
              <div className="wel-material-manager">
                <div className="wel-material-add-row">
                  <input placeholder="新廠商名稱" value={newSupplier} onChange={(e) => setNewSupplier(e.target.value)} />
                  <button type="button" className="wel-btn-ghost" onClick={() => { addSupplier(newSupplier); setNewSupplier(""); }}>
                    <Plus size={13} /> 新增
                  </button>
                </div>
                <div className="wel-manager-list">
                  {sortedSuppliers.length === 0 && <span className="muted" style={{ fontSize: 12 }}>尚未建立任何廠商</span>}
                  {sortedSuppliers.map((s) => (
                    <div key={s.id} className="wel-manager-row">
                      <input className="wel-manager-name" value={s.name} onChange={(e) => updateSupplier(s.id, { name: e.target.value })} />
                      <button type="button" className="wel-icon-btn" onClick={() => removeSupplier(s.id)}><Trash2 size={13} /></button>
                    </div>
                  ))}
                </div>
                <div className="muted" style={{ fontSize: 11 }}>提示：改名稱會自動同步更新過去的叫貨紀錄。</div>
                <button type="button" className="wel-btn-ghost" onClick={() => setSupplierManageOpen(false)}>收合廠商管理</button>
              </div>
            )}
          </Field>
          <Field label="數量">
            <input type="number" step="any" min="0" value={orderForm.quantity} onChange={(e) => setOrderForm({ ...orderForm, quantity: e.target.value })} />
          </Field>
          <Field label="單位">
            <input placeholder="支 / 米 / 個" value={orderForm.unit} onChange={(e) => setOrderForm({ ...orderForm, unit: e.target.value })} />
          </Field>
          <Field label="單價">
            <input type="number" step="any" min="0" value={orderForm.unitPrice} onChange={(e) => setOrderForm({ ...orderForm, unitPrice: e.target.value })} />
          </Field>
          <Field label="小計">
            <div className="wel-amount-preview">{fmtMoney(amount)}</div>
          </Field>
          <Field label="有無發票">
            <select value={orderForm.hasInvoice ? "yes" : "no"} onChange={(e) => setOrderForm({ ...orderForm, hasInvoice: e.target.value === "yes" })}>
              <option value="yes">有發票</option>
              <option value="no">無發票</option>
            </select>
          </Field>
          {orderForm.hasInvoice && (
            <Field label="營業稅" wide>
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                  <input
                    type="checkbox"
                    checked={orderForm.applyTax}
                    onChange={(e) => setOrderForm({ ...orderForm, applyTax: e.target.checked })}
                    style={{ width: 16, height: 16 }}
                  />
                  計算稅額
                </label>
                {orderForm.applyTax && (
                  <>
                    <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                      稅率
                      <input
                        type="number" min="0" step="0.1" value={orderForm.taxRate}
                        onChange={(e) => setOrderForm({ ...orderForm, taxRate: e.target.value })}
                        style={{ width: 60, height: 32 }}
                      />
                      %
                    </span>
                    <span className="muted" style={{ fontFamily: "var(--font-mono)", fontSize: 13 }}>
                      稅額：{fmtMoney((amount * (Number(orderForm.taxRate) || 0)) / 100)}
                    </span>
                  </>
                )}
              </div>
            </Field>
          )}
          <Field label="備註" wide>
            <input placeholder="選填" value={orderForm.note} onChange={(e) => setOrderForm({ ...orderForm, note: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitOrder}><Plus size={15} /> 新增叫貨紀錄</button>
      </div>

      <div className="wel-card">
        <div className="wel-card-title-row">
          <div className="wel-card-title"><CalendarRange size={14} /> 叫貨紀錄（依年月點開查詢）</div>
          <div className="wel-filter">
            <select value={invoiceFilter} onChange={(e) => setInvoiceFilter(e.target.value)}>
              <option value="all">全部（{orders.length}）</option>
              <option value="yes">有發票（{orders.filter((o) => o.hasInvoice).length}）</option>
              <option value="no">無發票（{orders.filter((o) => !o.hasInvoice).length}）</option>
            </select>
          </div>
        </div>
        {filteredOrders.length === 0 && <Empty text="尚未新增任何叫貨紀錄" />}
        {attendanceByMonth(filteredOrders).map((mg) => {
          const subtotal = mg.items.reduce((s, o) => s + (Number(o.amount) || 0), 0);
          const mOpen = !!openOrderMonths[mg.month];
          return (
            <div key={mg.month} className="wel-item-floor-group">
              <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => setOpenOrderMonths((prev) => ({ ...prev, [mg.month]: !prev[mg.month] }))}>
                {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                {mg.month} · {mg.items.length} 筆 · {fmtMoney(subtotal)}
              </button>
              {mOpen && (
                <div style={{ overflowX: "auto" }}>
                  <table className="wel-table">
                    <thead>
                      <tr>
                        {isAllSites && <th>案場</th>}
                        <th>日期</th><th>類別</th><th>品項</th><th>廠商</th><th className="right">數量</th><th className="right">單價</th><th className="right">金額</th><th>發票</th><th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {mg.items.map((o) => {
                        const isOpen = expandedOrderId === o.id;
                        return (
                        <React.Fragment key={o.id}>
                        <tr className="wel-row-clickable" onClick={() => (isOpen ? cancelEditOrder() : startEditOrder(o))}>
                          {isAllSites && <td className="muted">{siteById[o.siteId]?.name || "—"}</td>}
                          <td className="mono">{o.date}</td>
                          <td><span className="wel-tag" style={{ borderColor: catById[o.categoryId]?.color }}>{catById[o.categoryId]?.name || "未分類"}</span></td>
                          <td>
                            <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                              {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                              {o.itemName}
                            </span>
                            {o.note && <div className="muted" style={{ fontSize: 11.5, marginTop: 2, marginLeft: 18 }}>{o.note}</div>}
                          </td>
                          <td className="muted">{o.supplier || "—"}</td>
                          <td className="right mono">{fmtNum(o.quantity)} {o.unit}</td>
                          <td className="right mono">{fmtMoney(o.unitPrice)}</td>
                          <td className="right mono strong">{fmtMoney(o.amount)}</td>
                          <td>
                            {o.hasInvoice
                              ? <span className="wel-chip-info">有發票</span>
                              : <span className="wel-chip">無發票</span>}
                            {o.hasInvoice && o.applyTax && (
                              <div className="muted mono" style={{ fontSize: 11, marginTop: 3 }}>
                                稅額 {fmtMoney((Number(o.amount) || 0) * (Number(o.taxRate) || 0) / 100)}（{fmtNum(o.taxRate)}%）
                              </div>
                            )}
                          </td>
                          <td><button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteOrder(o.id); }}><Trash2 size={14} /></button></td>
                        </tr>
                        {isOpen && editDraft && (
                          <tr onClick={(e) => e.stopPropagation()}>
                            <td colSpan={isAllSites ? 9 : 8} style={{ padding: 0 }}>
                              <div className="wel-timeline-wrap">
                                <div className="wel-form-grid">
                                  <Field label="日期">
                                    <input type="date" value={editDraft.date} onChange={(e) => setEditDraft({ ...editDraft, date: e.target.value })} />
                                  </Field>
                                  <Field label="類別">
                                    <select value={editDraft.categoryId} onChange={(e) => setEditDraft({ ...editDraft, categoryId: e.target.value })}>
                                      {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                                    </select>
                                  </Field>
                                  <Field label="品項名稱" wide>
                                    <input value={editDraft.itemName} onChange={(e) => setEditDraft({ ...editDraft, itemName: e.target.value })} />
                                  </Field>
                                  <Field label="廠商 / 供應商">
                                    <input value={editDraft.supplier} onChange={(e) => setEditDraft({ ...editDraft, supplier: e.target.value })} />
                                  </Field>
                                  <Field label="數量">
                                    <input type="number" step="any" min="0" value={editDraft.quantity} onChange={(e) => setEditDraft({ ...editDraft, quantity: e.target.value })} />
                                  </Field>
                                  <Field label="單位">
                                    <input value={editDraft.unit} onChange={(e) => setEditDraft({ ...editDraft, unit: e.target.value })} />
                                  </Field>
                                  <Field label="單價">
                                    <input type="number" step="any" min="0" value={editDraft.unitPrice} onChange={(e) => setEditDraft({ ...editDraft, unitPrice: e.target.value })} />
                                  </Field>
                                  <Field label="小計">
                                    <div className="wel-amount-preview">{fmtMoney((Number(editDraft.quantity) || 0) * (Number(editDraft.unitPrice) || 0))}</div>
                                  </Field>
                                  <Field label="有無發票">
                                    <select value={editDraft.hasInvoice ? "yes" : "no"} onChange={(e) => setEditDraft({ ...editDraft, hasInvoice: e.target.value === "yes" })}>
                                      <option value="yes">有發票</option>
                                      <option value="no">無發票</option>
                                    </select>
                                  </Field>
                                  {editDraft.hasInvoice && (
                                    <Field label="營業稅" wide>
                                      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                                          <input
                                            type="checkbox" checked={editDraft.applyTax}
                                            onChange={(e) => setEditDraft({ ...editDraft, applyTax: e.target.checked })}
                                            style={{ width: 16, height: 16 }}
                                          />
                                          計算稅額
                                        </label>
                                        {editDraft.applyTax && (
                                          <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                                            稅率
                                            <input
                                              type="number" min="0" step="0.1" value={editDraft.taxRate}
                                              onChange={(e) => setEditDraft({ ...editDraft, taxRate: e.target.value })}
                                              style={{ width: 60, height: 32 }}
                                            />
                                            %
                                          </span>
                                        )}
                                      </div>
                                    </Field>
                                  )}
                                  <Field label="備註" wide>
                                    <input value={editDraft.note} onChange={(e) => setEditDraft({ ...editDraft, note: e.target.value })} />
                                  </Field>
                                </div>
                                <div style={{ display: "flex", gap: 8 }}>
                                  <button type="button" className="wel-btn-primary" onClick={() => saveEditOrder(o.id)}>儲存修改</button>
                                  <button type="button" className="wel-btn-ghost" onClick={cancelEditOrder}>取消</button>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                        </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function InventoryTab({ inventory, catById, categories, orders, usages, siteById, isAllSites, sites, setCurrentSiteId, consumeFromInventory, reassignUsages, deleteUsagesByKey, deleteUsage, showToast }) {
  const [form, setForm] = useState({ date: todayStr(), categoryId: categories[0]?.id || "", itemKey: "", quantity: "", location: "", type: "out" });
  const [filterCat, setFilterCat] = useState("all");
  const [openMonths, setOpenMonths] = useState({});

  const itemsInFormCategory = useMemo(() => inventory.filter((i) => i.categoryId === form.categoryId), [inventory, form.categoryId]);
  const selectedItem = inventory.find((i) => i.key === form.itemKey);
  const visibleInventory = filterCat === "all" ? inventory : inventory.filter((i) => i.categoryId === filterCat);

  const pickRow = (i) => setForm((f) => ({ ...f, categoryId: i.categoryId, itemKey: i.key }));
  const submitConsume = () => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能消貨"); return; }
    if (!selectedItem) { showToast("請先選擇「種類」與「品項」"); return; }
    const isReturn = form.type === "return";
    const qty = Number(form.quantity);
    if (!qty || qty <= 0) { showToast(isReturn ? "請填寫要退回的「數量」" : "請填寫要消貨的「數量」"); return; }
    if (!isReturn && qty > selectedItem.remaining) { showToast(`庫存不足：「${selectedItem.itemName}」剩餘 ${fmtNum(selectedItem.remaining)} ${selectedItem.unit || ""}`); return; }
    consumeFromInventory({ date: form.date, categoryId: selectedItem.categoryId, itemName: selectedItem.itemName, quantity: qty, unit: selectedItem.unit || "", location: form.location, type: form.type });
    setForm((f) => ({ ...f, quantity: "", location: "" }));
    showToast(`${isReturn ? "已退回入庫" : "已消貨"}：${selectedItem.itemName} ${fmtNum(qty)} ${selectedItem.unit || ""}`);
  };

  // 負數且完全沒有對應叫貨的列：多半是消貨紀錄的類別/品項和叫貨對不上，提供「併入」或「刪除」來校正
  const fixTargets = (row) => inventory
    .filter((x) => x.key !== row.key && x.ordered > 0)
    .sort((a, b) => (b.itemName === row.itemName) - (a.itemName === row.itemName));
  const usageCountOf = (row) => usages.filter((u) => u.categoryId === row.categoryId && u.itemName === row.itemName).length;
  const doFix = (row, value) => {
    if (!value) return;
    const where = `${catById[row.categoryId]?.name || "未分類"}・${row.itemName}`;
    if (value === "delete") {
      if (window.confirm(`確定要刪除「${where}」的 ${usageCountOf(row)} 筆消貨紀錄嗎？刪除後無法復原。`)) deleteUsagesByKey(row.categoryId, row.itemName);
      return;
    }
    const target = inventory.find((x) => x.key === value);
    if (target && window.confirm(`把「${where}」的 ${usageCountOf(row)} 筆消貨紀錄，併入「${catById[target.categoryId]?.name || "未分類"}・${target.itemName}」嗎？`)) {
      reassignUsages(row.categoryId, row.itemName, target.categoryId, target.itemName);
    }
  };

  const movements = useMemo(() => {
    const list = [];
    orders.forEach((o) => list.push({ id: "o" + o.id, date: o.date, kind: "in", categoryId: o.categoryId, itemName: o.itemName, quantity: Number(o.quantity) || 0, unit: o.unit || "", note: o.supplier || "", siteId: o.siteId }));
    usages.forEach((u) => list.push({ id: "u" + u.id, usageId: u.id, date: u.date, kind: u.type === "return" ? "return" : "out", categoryId: u.categoryId, itemName: u.itemName, quantity: Number(u.quantity) || 0, unit: u.unit || "", note: u.location || "", siteId: u.siteId }));
    return list;
  }, [orders, usages]);
  const kindLabel = { in: "進貨", out: "消貨", return: "退回" };

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">庫存 · INVENTORY</div>
          <h1 className="wel-h1">剩餘庫存現況</h1>
        </div>
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="消貨" />}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Boxes size={14} /> 消貨 / 退回（依剩餘庫存領出使用）</div>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
          </Field>
          <Field label="種類">
            <select value={form.categoryId} onChange={(e) => setForm({ ...form, categoryId: e.target.value, itemKey: "" })}>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="品項（依種類列出）">
            <select value={form.itemKey} onChange={(e) => setForm({ ...form, itemKey: e.target.value })}>
              <option value="" disabled>{itemsInFormCategory.length ? "請選擇品項…" : "此種類目前沒有庫存品項"}</option>
              {itemsInFormCategory.map((i) => (
                <option key={i.key} value={i.key}>{i.itemName}（剩餘 {fmtNum(i.remaining)} {i.unit || ""}）</option>
              ))}
            </select>
          </Field>
          <Field label="類型">
            <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
              <option value="out">消貨（領出使用）</option>
              <option value="return">退回（沒用完，還回庫存）</option>
            </select>
          </Field>
          <Field label={(form.type === "return" ? "退回數量" : "消貨數量") + (selectedItem ? `（${selectedItem.unit || "單位"}）` : "")}>
            <input type="number" step="any" min="0" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} />
          </Field>
          <Field label="用途 / 備註" wide>
            <input placeholder="選填，例：3樓衛浴配管" value={form.location} onChange={(e) => setForm({ ...form, location: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitConsume}><Plus size={15} /> {form.type === "return" ? "退回入庫" : "消貨"}</button>
        <div className="muted" style={{ fontSize: 11, marginTop: 8 }}>消貨會扣減剩餘庫存並記在下方「庫存異動紀錄」；用不完的選「退回」加回庫存。記錯的可在異動紀錄按垃圾桶刪除。</div>
      </div>

      <div className="wel-card-title-row" style={{ marginTop: 14 }}>
        <div className="wel-card-title"><Boxes size={14} /> 庫存明細（點一列可帶入上方消貨表單）</div>
        <div className="wel-filter">
          <select value={filterCat} onChange={(e) => setFilterCat(e.target.value)}>
            <option value="all">全部種類</option>
            {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </div>
      </div>
      <div className="wel-card" style={{ padding: 0, overflow: "hidden" }}>
        <table className="wel-table">
          <thead>
            <tr><th>類別</th><th>品項</th><th className="right">已叫貨</th><th className="right">已使用</th><th className="right">剩餘庫存</th><th className="right">平均單價</th><th className="right">庫存價值</th><th>校正</th></tr>
          </thead>
          <tbody>
            {visibleInventory.length === 0 && (<tr><td colSpan={8}><Empty text="尚無庫存資料，請先新增叫貨紀錄" /></td></tr>)}
            {visibleInventory.map((i) => (
              <tr key={i.key} className="wel-row-clickable" onClick={() => pickRow(i)}>
                <td><span className="wel-tag" style={{ borderColor: catById[i.categoryId]?.color }}>{catById[i.categoryId]?.name || "未分類"}</span></td>
                <td>{i.itemName}</td>
                <td className="right mono">{fmtNum(i.ordered)} {i.unit}</td>
                <td className="right mono">{fmtNum(i.used)} {i.unit}</td>
                <td className={"right mono strong" + (i.remaining <= 0 ? " wel-low" : "")}>{fmtNum(i.remaining)} {i.unit}</td>
                <td className="right mono muted">{fmtMoney(i.avgPrice)}</td>
                <td className="right mono">{fmtMoney(i.value)}</td>
                <td onClick={(e) => e.stopPropagation()}>
                  {i.remaining < 0 && i.ordered === 0 && (
                    <select className="wel-scan-input" style={{ maxWidth: 170 }} value="" onChange={(e) => doFix(i, e.target.value)} title="這筆消貨沒有對應的叫貨，可併入正確品項或刪除">
                      <option value="">校正負數…</option>
                      {fixTargets(i).map((t) => <option key={t.key} value={t.key}>併入「{catById[t.categoryId]?.name || "未分類"}・{t.itemName}」</option>)}
                      <option value="delete">刪除這些消貨紀錄</option>
                    </select>
                  )}
                  {i.remaining < 0 && i.ordered > 0 && <span className="muted" style={{ fontSize: 11 }}>叫貨數量不足，請補登叫貨</span>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="wel-card" style={{ marginTop: 14 }}>
        <div className="wel-card-title"><CalendarRange size={14} /> 庫存異動紀錄（進貨／消貨／退回，依年月點開查詢）</div>
        {movements.length === 0 && <Empty text="尚無庫存異動紀錄" />}
        {attendanceByMonth(movements).map((mg) => {
          const mOpen = !!openMonths[mg.month];
          const inCount = mg.items.filter((m) => m.kind === "in").length;
          const outCount = mg.items.filter((m) => m.kind === "out").length;
          const retCount = mg.items.filter((m) => m.kind === "return").length;
          return (
            <div key={mg.month} className="wel-item-floor-group">
              <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => setOpenMonths((prev) => ({ ...prev, [mg.month]: !prev[mg.month] }))}>
                {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                {mg.month} · 進貨 {inCount} 筆 · 消貨 {outCount} 筆{retCount ? ` · 退回 ${retCount} 筆` : ""}
              </button>
              {mOpen && (
                <div style={{ overflowX: "auto" }}>
                  <table className="wel-table">
                    <thead>
                      <tr>{isAllSites && <th>案場</th>}<th>日期</th><th>類型</th><th>類別</th><th>品項</th><th className="right">數量</th><th>廠商 / 用途</th><th></th></tr>
                    </thead>
                    <tbody>
                      {mg.items.map((m) => (
                        <tr key={m.id}>
                          {isAllSites && <td className="muted">{siteById[m.siteId]?.name || "—"}</td>}
                          <td className="mono">{m.date}</td>
                          <td>{m.kind === "out" ? <span className="wel-chip">{kindLabel[m.kind]}</span> : <span className="wel-chip-info">{kindLabel[m.kind]}</span>}</td>
                          <td><span className="wel-tag" style={{ borderColor: catById[m.categoryId]?.color }}>{catById[m.categoryId]?.name || "未分類"}</span></td>
                          <td>{m.itemName}</td>
                          <td className="right mono" style={m.kind === "out" ? undefined : { color: "var(--green)" }}>{m.kind === "out" ? "−" : "+"}{fmtNum(m.quantity)} {m.unit}</td>
                          <td className="muted">{m.note || "—"}</td>
                          <td>{m.usageId && <button className="wel-icon-btn" onClick={() => deleteUsage(m.usageId)}><Trash2 size={13} /></button>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function LaborTab({
  workers, workLogs, workerById, siteById, newWorker, setNewWorker, addWorker, updateWorkerRate, removeWorker,
  workLogForm, setWorkLogForm, submitWorkLog, deleteWorkLog, updateWorkLog,
  laborRange, setLaborRange, laborSummaryByWorker, laborRangeTotal, laborRangeDays, laborTotalAllTime, isAllSites, sites, setCurrentSiteId,
  contractWorkLogs, contractRows, addAttendance, deleteAttendance, updateAttendance,
}) {
  const [expandedWorker, setExpandedWorker] = useState(null);
  const [expandedWageRow, setExpandedWageRow] = useState(null);
  const [expandedContract, setExpandedContract] = useState(null);
  const [quickAttend, setQuickAttend] = useState({ contractId: "", date: todayStr(), headcount: 1, note: "" });
  const [openMonths, setOpenMonths] = useState({});
  const rateOf = (l) => Number(l.dailyRate ?? workerById[l.workerId]?.dailyRate) || 0;
  const [editingLogId, setEditingLogId] = useState(null);
  const [logDraft, setLogDraft] = useState(null);
  const startEditLog = (l) => { setEditingLogId(l.id); setLogDraft({ date: l.date, days: l.days, note: l.note || "", dailyRate: l.dailyRate ?? workerById[l.workerId]?.dailyRate ?? "" }); };
  const cancelEditLog = () => { setEditingLogId(null); setLogDraft(null); };
  const saveEditLog = (id) => { updateWorkLog(id, { ...logDraft, days: Number(logDraft.days) || 0, dailyRate: Number(logDraft.dailyRate) || 0 }); setEditingLogId(null); setLogDraft(null); };
  const [editingAttendId, setEditingAttendId] = useState(null);
  const [attendDraft, setAttendDraft] = useState(null);
  const startEditAttend = (l) => { setEditingAttendId(l.id); setAttendDraft({ date: l.date, headcount: l.headcount, note: l.note || "" }); };
  const cancelEditAttend = () => { setEditingAttendId(null); setAttendDraft(null); };
  const saveEditAttend = (id) => { updateAttendance(id, { ...attendDraft, headcount: Number(attendDraft.headcount) || 1 }); setEditingAttendId(null); setAttendDraft(null); };
  const toggleMonth = (key) => setOpenMonths((prev) => ({ ...prev, [key]: !prev[key] }));
  const rangeLabel = laborRange.start || laborRange.end
    ? `${laborRange.start || "最早"} ～ ${laborRange.end || "最新"}`
    : "全部區間";

  const logsByWorker = useMemo(() => {
    const map = {};
    workLogs.forEach((l) => {
      if (!map[l.workerId]) map[l.workerId] = [];
      map[l.workerId].push(l);
    });
    return map;
  }, [workLogs]);

  const monthGroup = (logs) => {
    const map = {};
    logs.forEach((l) => {
      const mk = (l.date || "").slice(0, 7) || "未知月份";
      if (!map[mk]) map[mk] = [];
      map[mk].push(l);
    });
    return Object.entries(map)
      .sort((a, b) => (a[0] < b[0] ? 1 : -1))
      .map(([month, items]) => ({
        month,
        items: items.slice().sort((a, b) => (a.date < b.date ? 1 : -1)),
        days: items.reduce((s, l) => s + (Number(l.days ?? l.headcount) || 0), 0),
        wage: items.reduce((s, l) => s + (Number(l.days) || 0) * rateOf(l), 0),
      }));
  };

  const logsByContract = useMemo(() => {
    const map = {};
    contractWorkLogs.forEach((l) => {
      if (!map[l.contractId]) map[l.contractId] = [];
      map[l.contractId].push(l);
    });
    return map;
  }, [contractWorkLogs]);

  const contractorGroups = useMemo(() => {
    return contractRows.map((c) => ({
      contractId: c.id,
      vendor: c.contractor || "未指定廠商",
      contractName: c.name,
      logs: logsByContract[c.id] || [],
    }));
  }, [contractRows, logsByContract]);

  const quickRange = (days) => {
    const end = new Date();
    const start = new Date();
    start.setDate(end.getDate() - (days - 1));
    setLaborRange({ start: start.toISOString().slice(0, 10), end: end.toISOString().slice(0, 10) });
  };

  const handleExportExcel = () => exportExcel("師傅出工與薪資", [
    { name: "出工明細", rows: workLogs.map((l) => ({ 日期: l.date, 師傅: workerById[l.workerId]?.name || "已刪除師傅", 天數: Number(l.days) || 0, 日薪: rateOf(l), 薪資: (Number(l.days) || 0) * rateOf(l), 備註: l.note || "" })) },
    { name: "依師傅彙總（目前區間）", rows: laborSummaryByWorker.map((r) => ({ 師傅: workerById[r.workerId]?.name || "已刪除師傅", 出工天數: r.days, 應付薪資: r.wage })) },
  ]);
  const handleExportWord = () => exportWord("師傅出工與薪資", "師傅出工與薪資", [
    { title: "薪資彙總（" + rangeLabel + "）", headers: ["師傅", "出工天數", "應付薪資"], rows: laborSummaryByWorker.map((r) => [workerById[r.workerId]?.name || "已刪除師傅", fmtNum(r.days), fmtMoney(r.wage)]) },
    { title: "出工明細", headers: ["日期", "師傅", "天數", "日薪", "薪資", "備註"], rows: workLogs.map((l) => [l.date, workerById[l.workerId]?.name || "已刪除師傅", l.days, fmtMoney(rateOf(l)), fmtMoney((Number(l.days) || 0) * rateOf(l)), l.note || ""]) },
  ]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">人工 · LABOR &amp; WAGES</div>
          <h1 className="wel-h1">師傅出工與薪資</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {/* Worker roster */}
      <div className="wel-card wel-form">
        <div className="wel-card-title"><Users size={14} /> 師傅名單（可隨時增減）</div>
        <div className="wel-form-grid wel-form-grid-btn" style={{ gridTemplateColumns: "2fr 1fr auto" }}>
          <Field label="師傅姓名">
            <input placeholder="例：陳師傅" value={newWorker.name} onChange={(e) => setNewWorker({ ...newWorker, name: e.target.value })} />
          </Field>
          <Field label="日薪（元/天）">
            <select value={newWorker.dailyRate} onChange={(e) => setNewWorker({ ...newWorker, dailyRate: e.target.value })}>
              <option value="" disabled>請選擇日薪…</option>
              {WAGE_OPTIONS.map((v) => <option key={v} value={v}>{fmtMoney(v)}</option>)}
            </select>
          </Field>
          <button type="button" className="wel-btn-primary" style={{ alignSelf: "end" }} onClick={addWorker}><Plus size={15} /> 新增師傅</button>
        </div>
        <div className="wel-cat-list" style={{ marginTop: 4 }}>
          {workers.length === 0 && <Empty text="尚未新增師傅，請先在上方新增" />}
          {workers.map((w) => (
            <div key={w.id} className="wel-cat-item">
              <span className="wel-dot" style={{ background: "var(--teal)" }} />
              <span className="wel-cat-name">{w.name}</span>
              <span className="muted" style={{ fontSize: 12, fontFamily: "var(--font-mono)" }}>日薪</span>
              <select
                className="wel-inline-rate"
                value={w.dailyRate}
                onChange={(e) => updateWorkerRate(w.id, e.target.value)}
              >
                {!WAGE_OPTIONS.includes(Number(w.dailyRate)) && w.dailyRate !== "" && w.dailyRate != null && (
                  <option value={w.dailyRate}>{fmtMoney(w.dailyRate)}（原設定）</option>
                )}
                {WAGE_OPTIONS.map((v) => <option key={v} value={v}>{fmtMoney(v)}</option>)}
              </select>
              <button className="wel-icon-btn" onClick={() => removeWorker(w.id)}><X size={14} /></button>
            </div>
          ))}
        </div>
      </div>

      {/* Attendance entry */}
      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="登記出工" />}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><HardHat size={14} /> 登記出工</div>
        <div className="wel-form-grid">
          <Field label="師傅">
            <select
              value={workLogForm.workerId}
              onChange={(e) => {
                const workerId = e.target.value;
                const w = workers.find((x) => x.id === workerId);
                setWorkLogForm({ ...workLogForm, workerId, dailyRate: w?.dailyRate ?? workLogForm.dailyRate });
              }}
            >
              <option value="" disabled>請選擇</option>
              {workers.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
            </select>
          </Field>
          <Field label="日期">
            <input type="date" value={workLogForm.date} onChange={(e) => setWorkLogForm({ ...workLogForm, date: e.target.value })} />
          </Field>
          <Field label="出工天數">
            <input type="number" min="0" step="0.5" value={workLogForm.days} onChange={(e) => setWorkLogForm({ ...workLogForm, days: e.target.value })} />
          </Field>
          <Field label="當天日薪（元/天）">
            <select value={workLogForm.dailyRate} onChange={(e) => setWorkLogForm({ ...workLogForm, dailyRate: e.target.value })}>
              <option value="" disabled>請選擇日薪…</option>
              {!WAGE_OPTIONS.includes(Number(workLogForm.dailyRate)) && workLogForm.dailyRate !== "" && workLogForm.dailyRate != null && (
                <option value={workLogForm.dailyRate}>{fmtMoney(workLogForm.dailyRate)}（原設定）</option>
              )}
              {WAGE_OPTIONS.map((v) => <option key={v} value={v}>{fmtMoney(v)}</option>)}
            </select>
          </Field>
          <Field label="備註" wide>
            <input placeholder="選填，例：3樓配管" value={workLogForm.note} onChange={(e) => setWorkLogForm({ ...workLogForm, note: e.target.value })} />
          </Field>
        </div>
        <div className="muted" style={{ fontSize: 11, marginTop: -6, marginBottom: 8 }}>提示：日薪預設帶入該師傅目前設定的日薪，這筆可以再調整，只會套用在這一天，不會影響其他已登記的出工紀錄或師傅的預設日薪。</div>
        <button type="button" className="wel-btn-primary" onClick={submitWorkLog} disabled={workers.length === 0}><Plus size={15} /> 新增出工紀錄</button>
      </div>

      {/* Adjustable range summary */}
      <div className="wel-card">
        <div className="wel-card-title"><CalendarRange size={14} /> 薪資試算（可隨時調整區間）</div>
        <div className="wel-range-bar">
          <label>起：<input type="date" value={laborRange.start} onChange={(e) => setLaborRange({ ...laborRange, start: e.target.value })} /></label>
          <label>迄：<input type="date" value={laborRange.end} onChange={(e) => setLaborRange({ ...laborRange, end: e.target.value })} /></label>
          <button type="button" className="wel-btn-ghost" onClick={() => quickRange(7)}>近 7 天</button>
          <button type="button" className="wel-btn-ghost" onClick={() => quickRange(30)}>近 30 天</button>
          <button type="button" className="wel-btn-ghost" onClick={() => setLaborRange({ start: "", end: "" })}>清除區間</button>
          <span className="wel-range-label muted">{rangeLabel}</span>
        </div>

        <div className="wel-labor-summary">
          <div className="wel-labor-stat">
            <span>區間出工天數</span>
            <b>{fmtNum(laborRangeDays)} 天</b>
          </div>
          <div className="wel-labor-stat">
            <span>區間應付薪資</span>
            <b style={{ color: "var(--amber)" }}>{fmtMoney(laborRangeTotal)}</b>
          </div>
          <div className="wel-labor-stat">
            <span>累計總薪資（不受區間限制）</span>
            <b>{fmtMoney(laborTotalAllTime)}</b>
          </div>
        </div>

        <table className="wel-table" style={{ marginTop: 12 }}>
          <thead>
            <tr><th>師傅</th><th className="right">目前日薪</th><th className="right">出工天數</th><th className="right">應付薪資</th></tr>
          </thead>
          <tbody>
            {laborSummaryByWorker.length === 0 && (<tr><td colSpan={4}><Empty text="此區間尚無出工紀錄" /></td></tr>)}
            {laborSummaryByWorker.map((r) => {
              const isOpen = expandedWageRow === r.workerId;
              const detailLogs = workLogs
                .filter((l) => l.workerId === r.workerId)
                .filter((l) => !laborRange.start || l.date >= laborRange.start)
                .filter((l) => !laborRange.end || l.date <= laborRange.end)
                .sort((a, b) => (a.date < b.date ? 1 : -1));
              return (
                <React.Fragment key={r.workerId}>
                  <tr
                    className="wel-row-clickable"
                    onClick={() => setExpandedWageRow(isOpen ? null : r.workerId)}
                  >
                    <td>
                      <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                        {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                        {workerById[r.workerId]?.name || "已刪除師傅"}
                      </span>
                    </td>
                    <td className="right mono muted">{fmtMoney(workerById[r.workerId]?.dailyRate)}</td>
                    <td className="right mono">{fmtNum(r.days)} 天</td>
                    <td className="right mono strong">{fmtMoney(r.wage)}</td>
                  </tr>
                  {isOpen && (
                    <tr>
                      <td colSpan={4} style={{ padding: 0 }}>
                        <div className="wel-timeline-wrap">
                          {detailLogs.length === 0 ? (
                            <div className="muted" style={{ fontSize: 12 }}>此區間沒有出工明細</div>
                          ) : (
                            <table className="wel-table">
                              <thead>
                                <tr><th>日期</th><th>上下班時間</th><th className="right">天數</th><th className="right">日薪</th><th>備註</th></tr>
                              </thead>
                              <tbody>
                                {detailLogs.map((l) => (
                                  <tr key={l.id}>
                                    <td className="mono">{l.date}</td>
                                    <td className="mono">
                                      {l.startTime && l.endTime ? `${l.startTime} - ${l.endTime}` : "—"}
                                      {l.overtimeHours > 0 ? <span className="muted" style={{ marginLeft: 6 }}>（加班 {l.overtimeHours} 小時）</span> : null}
                                    </td>
                                    <td className="right mono">{fmtNum(l.days)}</td>
                                    <td className="right mono">{fmtMoney(rateOf(l))}</td>
                                    <td className="muted">{l.note || "—"}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Per-worker attendance, expandable, grouped by month */}
      <div className="wel-card">
        <div className="wel-card-title"><HardHat size={14} /> 各師傅出工紀錄（點開查詢，依月份分組）</div>
        {workers.length === 0 && <Empty text="尚未新增師傅" />}
        {workers.map((w) => {
          const logs = logsByWorker[w.id] || [];
          const totalDays = logs.reduce((s, l) => s + (Number(l.days) || 0), 0);
          const totalWage = logs.reduce((s, l) => s + (Number(l.days) || 0) * rateOf(l), 0);
          const isOpen = expandedWorker === w.id;
          const months = isOpen ? monthGroup(logs) : [];
          return (
            <div key={w.id} className="wel-vendor-group">
              <button type="button" className="wel-vendor-head wel-vendor-head-btn" onClick={() => setExpandedWorker(isOpen ? null : w.id)}>
                {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                <span className="wel-vendor-name">{w.name}</span>
                <span className="muted mono" style={{ fontSize: 11 }}>{fmtNum(totalDays)} 天</span>
                <span className="wel-contract-amt">{fmtMoney(totalWage)}</span>
              </button>
              {isOpen && (
                <div style={{ padding: "4px 4px 8px" }}>
                  {months.length === 0 && <Empty text="這位師傅尚無出工紀錄" />}
                  {months.map((mg) => {
                    const mKey = "w|" + w.id + "|" + mg.month;
                    const mOpen = !!openMonths[mKey];
                    return (
                      <div key={mg.month} className="wel-item-floor-group">
                        <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => toggleMonth(mKey)}>
                          {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                          {mg.month} · {fmtNum(mg.days)} 天 · {fmtMoney(mg.wage)}
                        </button>
                        {mOpen && (
                          <div className="wel-template-rows">
                            {mg.items.map((l) => (
                              editingLogId === l.id ? (
                                <div key={l.id} className="wel-attendance-item">
                                  <input type="date" value={logDraft.date} onChange={(e) => setLogDraft({ ...logDraft, date: e.target.value })} style={{ width: 130 }} />
                                  <input type="number" min="0" step="0.5" value={logDraft.days} onChange={(e) => setLogDraft({ ...logDraft, days: e.target.value })} style={{ width: 60 }} />
                                  <select
                                    className="wel-inline-rate"
                                    value={logDraft.dailyRate}
                                    onChange={(e) => setLogDraft({ ...logDraft, dailyRate: e.target.value })}
                                  >
                                    {!WAGE_OPTIONS.includes(Number(logDraft.dailyRate)) && logDraft.dailyRate !== "" && logDraft.dailyRate != null && (
                                      <option value={logDraft.dailyRate}>{fmtMoney(logDraft.dailyRate)}（原設定）</option>
                                    )}
                                    {WAGE_OPTIONS.map((v) => <option key={v} value={v}>{fmtMoney(v)}</option>)}
                                  </select>
                                  <input placeholder="備註" value={logDraft.note} onChange={(e) => setLogDraft({ ...logDraft, note: e.target.value })} style={{ flex: 1 }} />
                                  <button className="wel-icon-btn" onClick={() => saveEditLog(l.id)}><CheckCircle2 size={13} /></button>
                                  <button className="wel-icon-btn" onClick={cancelEditLog}><X size={13} /></button>
                                </div>
                              ) : (
                                <div key={l.id} className="wel-attendance-item wel-row-clickable" onClick={() => startEditLog(l)}>
                                  <span className="mono">{l.date}</span>
                                  <span className="mono muted">{fmtNum(l.days)} 天</span>
                                  <span className="mono muted">日薪 {fmtMoney(rateOf(l))}</span>
                                  <span className="muted" style={{ flex: 1 }}>{l.note || "—"}</span>
                                  <button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteWorkLog(l.id); }}><X size={13} /></button>
                                </div>
                              )
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Linked portal: 發包廠商出工 (subcontractor attendance), same expand/click pattern */}
      <div className="wel-card">
        <div className="wel-card-title"><Handshake size={14} /> 發包廠商出工（與「發包與領款」連動，方便直接點選）</div>
        {contractorGroups.length === 0 && <Empty text="尚未建立發包項目" />}
        {contractorGroups.map((g) => {
          const totalDays = g.logs.length;
          const totalManDays = g.logs.reduce((s, l) => s + (Number(l.headcount) || 0), 0);
          const isOpen = expandedContract === g.contractId;
          const months = isOpen ? monthGroup(g.logs) : [];
          return (
            <div key={g.contractId} className="wel-vendor-group">
              <button type="button" className="wel-vendor-head wel-vendor-head-btn" onClick={() => setExpandedContract(isOpen ? null : g.contractId)}>
                {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                <span className="wel-vendor-name">{g.vendor}</span>
                <span className="muted mono" style={{ fontSize: 11 }}>{g.contractName}</span>
                <span className="wel-contract-amt">{totalDays} 天 · 工天人次 {totalManDays}</span>
              </button>
              {isOpen && (
                <div style={{ padding: "4px 4px 8px" }}>
                  <div className="wel-attendance-row" style={{ marginBottom: 8 }}>
                    <input type="date" value={quickAttend.contractId === g.contractId ? quickAttend.date : todayStr()} onChange={(e) => setQuickAttend({ contractId: g.contractId, date: e.target.value, headcount: quickAttend.headcount, note: quickAttend.note })} />
                    <label>人數<input type="number" min="1" step="1" style={{ width: 56 }} value={quickAttend.contractId === g.contractId ? quickAttend.headcount : 1} onChange={(e) => setQuickAttend({ contractId: g.contractId, date: quickAttend.date, headcount: e.target.value, note: quickAttend.note })} /></label>
                    <input placeholder="備註" value={quickAttend.contractId === g.contractId ? quickAttend.note : ""} onChange={(e) => setQuickAttend({ contractId: g.contractId, date: quickAttend.date, headcount: quickAttend.headcount, note: e.target.value })} />
                    <button
                      type="button" className="wel-btn-ghost"
                      onClick={() => {
                        const d = quickAttend.contractId === g.contractId ? quickAttend : { date: todayStr(), headcount: 1, note: "" };
                        addAttendance(g.contractId, d.date, d.headcount, d.note);
                        setQuickAttend({ contractId: "", date: todayStr(), headcount: 1, note: "" });
                      }}
                    >
                      <Plus size={13} /> 登記出工
                    </button>
                  </div>
                  {months.length === 0 && <Empty text="這個發包項目尚無出工紀錄" />}
                  {months.map((mg) => {
                    const mKey = "c|" + g.contractId + "|" + mg.month;
                    const mOpen = !!openMonths[mKey];
                    return (
                      <div key={mg.month} className="wel-item-floor-group">
                        <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => toggleMonth(mKey)}>
                          {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                          {mg.month} · {mg.items.length} 天
                        </button>
                        {mOpen && (
                          <div className="wel-template-rows">
                            {mg.items.map((l) => (
                              editingAttendId === l.id ? (
                                <div key={l.id} className="wel-attendance-item">
                                  <input type="date" value={attendDraft.date} onChange={(e) => setAttendDraft({ ...attendDraft, date: e.target.value })} style={{ width: 130 }} />
                                  <input type="number" min="1" step="1" value={attendDraft.headcount} onChange={(e) => setAttendDraft({ ...attendDraft, headcount: e.target.value })} style={{ width: 60 }} />
                                  <input placeholder="備註" value={attendDraft.note} onChange={(e) => setAttendDraft({ ...attendDraft, note: e.target.value })} style={{ flex: 1 }} />
                                  <button className="wel-icon-btn" onClick={() => saveEditAttend(l.id)}><CheckCircle2 size={13} /></button>
                                  <button className="wel-icon-btn" onClick={cancelEditAttend}><X size={13} /></button>
                                </div>
                              ) : (
                                <div key={l.id} className="wel-attendance-item wel-row-clickable" onClick={() => startEditAttend(l)}>
                                  <span className="mono">{l.date}</span>
                                  <span className="mono muted">{l.headcount} 人</span>
                                  <span className="muted" style={{ flex: 1 }}>{l.note || "—"}</span>
                                  <button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteAttendance(l.id); }}><X size={13} /></button>
                                </div>
                              )
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ContractsTab({
  updateContract, renameVendor,
  addContractFile, deleteContractFile, openContractFile,
  contractRows, payments, siteById, contractForm, setContractForm, submitContract, deleteContract,
  paymentForm, setPaymentForm, submitPayment, deletePayment, updatePayment, contractPaidTotal, contractTotalPrice, isAllSites,
  floorItems, addTemplateRow, updateTemplateRow, removeTemplateRow, applyTemplateToFloors,
  addFloorItem, updateFloorItem, deleteFloorItem, deleteFloorGroup,
  addAttendance, deleteAttendance, updateAttendance, contractWorkLogs, sites, setCurrentSiteId, showToast,
  contractItems, addContractItem, addContractItemsBatch, addPaymentFromItem, updateContractItem, deleteContractItem, deleteContractItemsByFloor, deleteContractItemsByName,
}) {
  const [expandedId, setExpandedId] = useState(null);
  const [editingContractId, setEditingContractId] = useState(null);
  const [contractDraft, setContractDraft] = useState(null);
  const [editingVendor, setEditingVendor] = useState(null);
  const [vendorDraft, setVendorDraft] = useState("");
  const startEditContract = (c) => {
    setEditingContractId(c.id);
    setContractDraft({ contractor: c.contractor || "", name: c.name || "", date: c.date || "", totalPrice: c.totalPrice ?? "", note: c.note || "" });
  };
  const cancelEditContract = () => { setEditingContractId(null); setContractDraft(null); };
  const saveEditContract = (id) => {
    if (!contractDraft.name.trim()) { showToast("請填寫「工程總稱」後再儲存"); return; }
    updateContract(id, {
      contractor: contractDraft.contractor.trim(),
      name: contractDraft.name.trim(),
      date: contractDraft.date,
      totalPrice: contractDraft.totalPrice === "" ? "" : Number(contractDraft.totalPrice) || 0,
      note: contractDraft.note.trim(),
    });
    cancelEditContract();
  };
  const [expandedPaymentId, setExpandedPaymentId] = useState(null);
  const [paymentDraft, setPaymentDraft] = useState(null);
  const startEditPayment = (p) => { setExpandedPaymentId(p.id); setPaymentDraft({ date: p.date, contractId: p.contractId, amount: p.amount, note: p.note || "" }); };
  const cancelEditPayment = () => { setExpandedPaymentId(null); setPaymentDraft(null); };
  const saveEditPayment = (id) => { updatePayment(id, { ...paymentDraft, amount: Number(paymentDraft.amount) || 0 }); setExpandedPaymentId(null); setPaymentDraft(null); };
  const sortedPayments = useMemo(() => payments.slice().sort((a, b) => (a.date < b.date ? 1 : -1)), [payments]);

  const vendorGroups = useMemo(() => {
    const map = {};
    contractRows.forEach((c) => {
      const key = (c.contractor || "").trim() || "未指定廠商";
      if (!map[key]) map[key] = { vendor: key, contracts: [], paid: 0, totalPrice: 0, itemsTotal: 0, floorTotal: 0 };
      map[key].contracts.push(c);
      map[key].paid += c.paid;
      map[key].totalPrice += Number(c.totalPrice) || 0;
      map[key].itemsTotal += c.itemsTotal;
      map[key].floorTotal += c.floorTotal;
    });
    return Object.values(map).sort((a, b) => b.paid - a.paid);
  }, [contractRows]);

  const contractNameById = useMemo(() => {
    const m = {};
    contractRows.forEach((c) => (m[c.id] = { name: c.name, contractor: c.contractor || "未指定廠商" }));
    return m;
  }, [contractRows]);

  const handleExportExcel = () => exportExcel("發包與領款", [
    { name: "發包項目", rows: contractRows.map((c) => ({ 廠商: c.contractor || "未指定廠商", 工程總稱: c.name, 發包日期: c.date, 發包總價: Number(c.totalPrice) || 0, 已領款: c.paid, 尚未付款: c.remaining, 項目明細合計: c.itemsTotal, 備註: c.note || "" })) },
    { name: "請款項目明細", rows: contractItems.map((it) => ({ 廠商: contractNameById[it.contractId]?.contractor || "", 工程總稱: contractNameById[it.contractId]?.name || "", 日期: it.date, 項目: it.name, 樓層: it.floor || "", 百分比: it.percent === undefined ? "" : it.percent, 金額: Number(it.amount) || 0, 已匯入領款: it.paidImported ? "是" : "否", 備註: it.note || "" })) },
    { name: "領款紀錄", rows: payments.map((p) => ({ 廠商: contractNameById[p.contractId]?.contractor || "", 工程總稱: contractNameById[p.contractId]?.name || "", 日期: p.date, 金額: Number(p.amount) || 0, 備註: p.note || "" })) },
    { name: "出工紀錄", rows: contractWorkLogs.map((l) => ({ 廠商: contractNameById[l.contractId]?.contractor || "", 工程總稱: contractNameById[l.contractId]?.name || "", 日期: l.date, 人數: l.headcount, 備註: l.note || "" })) },
  ]);
  const handleExportWord = () => exportWord("發包與領款", "發包項目與領款紀錄", [
    { title: "發包項目總覽", headers: ["廠商", "工程總稱", "發包總價", "已領款", "尚未付款"], rows: contractRows.map((c) => [c.contractor || "未指定廠商", c.name, fmtMoney(c.totalPrice), fmtMoney(c.paid), fmtMoney(c.remaining)]) },
    { title: "請款項目明細", headers: ["廠商", "工程總稱", "日期", "項目", "樓層", "百分比", "金額"], rows: contractItems.map((it) => [contractNameById[it.contractId]?.contractor || "", contractNameById[it.contractId]?.name || "", it.date, it.name, it.floor || "—", it.percent === undefined ? "—" : `${it.percent}%`, fmtMoney(it.amount)]) },
    { title: "領款紀錄", headers: ["廠商", "工程總稱", "日期", "金額", "備註"], rows: payments.map((p) => [contractNameById[p.contractId]?.contractor || "", contractNameById[p.contractId]?.name || "", p.date, fmtMoney(p.amount), p.note || ""]) },
    { title: "出工紀錄", headers: ["廠商", "工程總稱", "日期", "人數", "備註"], rows: contractWorkLogs.map((l) => [contractNameById[l.contractId]?.contractor || "", contractNameById[l.contractId]?.name || "", l.date, l.headcount, l.note || ""]) },
  ]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">發包 · SUBCONTRACT &amp; PAYMENTS</div>
          <h1 className="wel-h1">發包項目與領款紀錄</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增發包項目、登記領款或編輯樓層工項" />}

      {/* Contract summary */}
      <div className="wel-labor-summary">
        <div className="wel-labor-stat">
          <span>發包總價（已建立項目）</span>
          <b>{fmtMoney(contractTotalPrice)}</b>
        </div>
        <div className="wel-labor-stat">
          <span>已領款（列入成本）</span>
          <b style={{ color: "var(--amber)" }}>{fmtMoney(contractPaidTotal)}</b>
        </div>
        <div className="wel-labor-stat">
          <span>尚未付款</span>
          <b>{fmtMoney(Math.max(0, contractTotalPrice - contractPaidTotal))}</b>
        </div>
      </div>

      {/* New contract */}
      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Handshake size={14} /> 新增廠商 / 發包項目</div>
        <div className="wel-form-grid">
          <Field label="廠商 / 工班">
            <input placeholder="例：全興水電工程行" value={contractForm.contractor} onChange={(e) => setContractForm({ ...contractForm, contractor: e.target.value })} />
          </Field>
          <Field label="工程總稱">
            <input placeholder="例：3F-5F 水電工程" value={contractForm.name} onChange={(e) => setContractForm({ ...contractForm, name: e.target.value })} />
          </Field>
          <Field label="發包日期">
            <input type="date" value={contractForm.date} onChange={(e) => setContractForm({ ...contractForm, date: e.target.value })} />
          </Field>
          <Field label="發包總價（選填，可留空改用項目明細加總）">
            <input type="number" min="0" step="any" value={contractForm.totalPrice} onChange={(e) => setContractForm({ ...contractForm, totalPrice: e.target.value })} />
          </Field>
          <Field label="備註" wide>
            <input placeholder="選填" value={contractForm.note} onChange={(e) => setContractForm({ ...contractForm, note: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitContract}><Plus size={15} /> 新增發包項目</button>
      </div>

      {/* Vendor-grouped contract list with progress + floor/item drill-down */}
      <div className="wel-card">
        <div className="wel-card-title"><Layers size={14} /> 依廠商列出的發包項目、付款進度與明細</div>
        {contractRows.length === 0 && <Empty text="尚未建立發包項目" />}
        {vendorGroups.map((vg) => (
          <div key={vg.vendor} className="wel-vendor-group">
            <div className="wel-vendor-head">
              <Building2 size={14} color="var(--teal)" />
              {editingVendor === vg.vendor ? (
                <>
                  <input
                    autoFocus className="wel-scan-input" style={{ flex: 1 }} value={vendorDraft} placeholder="廠商 / 工班名稱"
                    onChange={(e) => setVendorDraft(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") { renameVendor(vg.vendor, vendorDraft); setEditingVendor(null); } if (e.key === "Escape") setEditingVendor(null); }}
                  />
                  <button type="button" className="wel-btn-primary" style={{ padding: "5px 12px", fontSize: 12 }} onClick={() => { renameVendor(vg.vendor, vendorDraft); setEditingVendor(null); }}>儲存</button>
                  <button type="button" className="wel-btn-ghost" onClick={() => setEditingVendor(null)}>取消</button>
                </>
              ) : (
                <>
                  <span className="wel-vendor-name">{vg.vendor}</span>
                  <button className="wel-icon-btn" title="修改廠商名稱（底下所有發包項目一併更新）" onClick={() => { setEditingVendor(vg.vendor); setVendorDraft(vg.vendor === "未指定廠商" ? "" : vg.vendor); }}><Pencil size={13} /></button>
                  <span className="muted mono" style={{ fontSize: 11 }}>{vg.contracts.length} 個項目</span>
                  <span className="wel-contract-amt">{fmtMoney(vg.paid)} / {fmtMoney(vg.totalPrice)}</span>
                </>
              )}
            </div>
            <div className="wel-pipes">
              {vg.contracts.map((c) => {
            const pct = c.totalPrice > 0 ? Math.min(100, (c.paid / c.totalPrice) * 100) : 0;
            const isOpen = expandedId === c.id;
            const mismatch = c.floorTotal > 0 && c.totalPrice > 0 && Math.round(c.floorTotal) !== Math.round(c.totalPrice);
            return (
              <div key={c.id} className="wel-contract-row">
                <div className="wel-contract-head">
                  <button type="button" className="wel-expand-btn" onClick={() => setExpandedId(isOpen ? null : c.id)}>
                    {isOpen ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                  </button>
                  <span className="wel-contract-name">{c.name}</span>
                  {c.contractor && <span className="muted mono" style={{ fontSize: 11.5 }}>{c.contractor}</span>}
                  <span className="wel-contract-amt">{fmtMoney(c.paid)} / {fmtMoney(c.totalPrice)}</span>
                  <button className="wel-icon-btn" title="修改廠商／發包資料" onClick={() => (editingContractId === c.id ? cancelEditContract() : startEditContract(c))}><Pencil size={14} /></button>
                  <button className="wel-icon-btn" onClick={() => deleteContract(c.id)}><Trash2 size={14} /></button>
                </div>
                {editingContractId === c.id && contractDraft && (
                  <div className="wel-timeline-wrap">
                    <div className="wel-form-grid">
                      <Field label="廠商 / 工班"><input value={contractDraft.contractor} onChange={(e) => setContractDraft({ ...contractDraft, contractor: e.target.value })} /></Field>
                      <Field label="工程總稱"><input value={contractDraft.name} onChange={(e) => setContractDraft({ ...contractDraft, name: e.target.value })} /></Field>
                      <Field label="發包日期"><input type="date" value={contractDraft.date} onChange={(e) => setContractDraft({ ...contractDraft, date: e.target.value })} /></Field>
                      <Field label="發包總價"><input type="number" min="0" step="any" value={contractDraft.totalPrice} onChange={(e) => setContractDraft({ ...contractDraft, totalPrice: e.target.value })} /></Field>
                      <Field label="備註" wide><textarea rows={3} value={contractDraft.note} onChange={(e) => setContractDraft({ ...contractDraft, note: e.target.value })} /></Field>
                    </div>
                    <div style={{ display: "flex", gap: 8 }}>
                      <button type="button" className="wel-btn-primary" onClick={() => saveEditContract(c.id)}>儲存修改</button>
                      <button type="button" className="wel-btn-ghost" onClick={cancelEditContract}>取消</button>
                    </div>
                  </div>
                )}
                {c.note && editingContractId !== c.id && (
                  <div className="wel-floor-summary-line" style={{ whiteSpace: "pre-wrap" }}><span className="muted">備註：</span>{c.note}</div>
                )}
                <div className="wel-pipe-track">
                  <div className="wel-pipe-fill" style={{ width: `${Math.max(pct, c.paid > 0 ? 4 : 0)}%`, background: pct >= 100 ? "var(--green)" : "var(--amber)" }} />
                </div>
                {c.floorTotal > 0 && (
                  <div className="wel-floor-summary-line muted">
                    樓層工項合計：<span className="mono">{fmtMoney(c.floorTotal)}</span>
                    {mismatch && <span className="wel-mismatch"> · 與發包總價不符，請確認</span>}
                  </div>
                )}
                {c.itemsTotal > 0 && (
                  <div className="wel-floor-summary-line muted">
                    項目明細合計：<span className="mono">{fmtMoney(c.itemsTotal)}</span>
                  </div>
                )}
                {c.attendanceDays > 0 && (
                  <div className="wel-floor-summary-line muted">
                    出工紀錄：<span className="mono">{c.attendanceDays}</span> 天 · 工天人次 <span className="mono">{c.attendanceManDays}</span>
                  </div>
                )}
                <ContractFiles contract={c} addContractFile={addContractFile} deleteContractFile={deleteContractFile} openContractFile={openContractFile} />
                {isOpen && (
                  <FloorItemsPanel
                    contract={c}
                    floorItems={floorItems.filter((f) => f.contractId === c.id)}
                    isAllSites={isAllSites}
                    addTemplateRow={addTemplateRow}
                    updateTemplateRow={updateTemplateRow}
                    removeTemplateRow={removeTemplateRow}
                    applyTemplateToFloors={applyTemplateToFloors}
                    addFloorItem={addFloorItem}
                    updateFloorItem={updateFloorItem}
                    deleteFloorItem={deleteFloorItem}
                    deleteFloorGroup={deleteFloorGroup}
                    addAttendance={addAttendance}
                    deleteAttendance={deleteAttendance}
                    updateAttendance={updateAttendance}
                    contractWorkLogs={contractWorkLogs.filter((l) => l.contractId === c.id)}
                    contractItems={contractItems.filter((it) => it.contractId === c.id)}
                    addContractItem={addContractItem}
                    addContractItemsBatch={addContractItemsBatch}
                    addPaymentFromItem={addPaymentFromItem}
                    updateContractItem={updateContractItem}
                    deleteContractItem={deleteContractItem}
                    deleteContractItemsByFloor={deleteContractItemsByFloor}
                    deleteContractItemsByName={deleteContractItemsByName}
                    showToast={showToast}
                  />
                )}
              </div>
            );
          })}
            </div>
          </div>
        ))}
      </div>

      {/* Payment entry */}
      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Banknote size={14} /> 登記領款</div>
        <div className="wel-form-grid">
          <Field label="發包項目">
            <select value={paymentForm.contractId} onChange={(e) => setPaymentForm({ ...paymentForm, contractId: e.target.value })}>
              <option value="" disabled>請選擇</option>
              {contractRows.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="領款日期">
            <input type="date" value={paymentForm.date} onChange={(e) => setPaymentForm({ ...paymentForm, date: e.target.value })} />
          </Field>
          <Field label="金額">
            <input type="number" min="0" step="any" value={paymentForm.amount} onChange={(e) => setPaymentForm({ ...paymentForm, amount: e.target.value })} />
          </Field>
          <Field label="備註" wide>
            <input placeholder="選填，例：第二期款" value={paymentForm.note} onChange={(e) => setPaymentForm({ ...paymentForm, note: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitPayment} disabled={contractRows.length === 0}><Plus size={15} /> 新增領款紀錄</button>
      </div>

      {/* Payment log */}
      <div className="wel-card" style={{ padding: 0, overflow: "hidden" }}>
        <table className="wel-table">
          <thead>
            <tr>{isAllSites && <th>案場</th>}<th>日期</th><th>發包項目</th><th className="right">金額</th><th>備註</th><th></th></tr>
          </thead>
          <tbody>
            {sortedPayments.length === 0 && (<tr><td colSpan={isAllSites ? 6 : 5}><Empty text="尚無領款紀錄" /></td></tr>)}
            {sortedPayments.map((p) => {
              const c = contractRows.find((cr) => cr.id === p.contractId);
              const isOpen = expandedPaymentId === p.id;
              return (
                <React.Fragment key={p.id}>
                <tr className="wel-row-clickable" onClick={() => (isOpen ? cancelEditPayment() : startEditPayment(p))}>
                  {isAllSites && <td className="muted">{siteById[p.siteId]?.name || "—"}</td>}
                  <td className="mono">{p.date}</td>
                  <td>
                    <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                      {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                      {c?.name || "已刪除項目"}
                    </span>
                  </td>
                  <td className="right mono strong">{fmtMoney(p.amount)}</td>
                  <td className="muted">{p.note || "—"}</td>
                  <td><button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deletePayment(p.id); }}><Trash2 size={14} /></button></td>
                </tr>
                {isOpen && paymentDraft && (
                  <tr onClick={(e) => e.stopPropagation()}>
                    <td colSpan={isAllSites ? 6 : 5} style={{ padding: 0 }}>
                      <div className="wel-timeline-wrap">
                        <div className="wel-form-grid">
                          <Field label="發包項目">
                            <select value={paymentDraft.contractId} onChange={(e) => setPaymentDraft({ ...paymentDraft, contractId: e.target.value })}>
                              {contractRows.map((cr) => <option key={cr.id} value={cr.id}>{cr.name}</option>)}
                            </select>
                          </Field>
                          <Field label="領款日期">
                            <input type="date" value={paymentDraft.date} onChange={(e) => setPaymentDraft({ ...paymentDraft, date: e.target.value })} />
                          </Field>
                          <Field label="金額">
                            <input type="number" min="0" step="any" value={paymentDraft.amount} onChange={(e) => setPaymentDraft({ ...paymentDraft, amount: e.target.value })} />
                          </Field>
                          <Field label="備註" wide>
                            <input value={paymentDraft.note} onChange={(e) => setPaymentDraft({ ...paymentDraft, note: e.target.value })} />
                          </Field>
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                          <button type="button" className="wel-btn-primary" onClick={() => saveEditPayment(p.id)}>儲存修改</button>
                          <button type="button" className="wel-btn-ghost" onClick={cancelEditPayment}>取消</button>
                        </div>
                      </div>
                    </td>
                  </tr>
                )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ReceiptPhoto({ file }) {
  const [url, setUrl] = useState("");
  const [failed, setFailed] = useState(false);
  useEffect(() => {
    let alive = true;
    let objectUrl = "";
    (async () => {
      try {
        const blob = base64ToBlob(await fileLoad(file.id, file.chunks), file.type);
        objectUrl = URL.createObjectURL(blob);
        if (alive) setUrl(objectUrl);
      } catch (e) {
        if (alive) setFailed(true);
      }
    })();
    return () => { alive = false; if (objectUrl) URL.revokeObjectURL(objectUrl); };
  }, [file.id, file.chunks, file.type]);
  if (failed) return <div className="muted" style={{ fontSize: 12 }}>照片讀取失敗，請稍後再試</div>;
  if (!url) return <div className="muted" style={{ fontSize: 12 }}>照片載入中…</div>;
  return <img src={url} alt={file.name} onClick={() => window.open(url, "_blank")} style={{ maxWidth: "100%", width: 380, borderRadius: 8, border: "1px solid var(--border)", cursor: "zoom-in" }} />;
}

function ReceiptsTab({ receipts, siteById, suppliers, isAllSites, sites, setCurrentSiteId, addReceipt, updateReceipt, deleteReceipt, showToast }) {
  const [form, setForm] = useState({ date: todayStr(), title: "", note: "" });
  const [photos, setPhotos] = useState([]);
  const [busy, setBusy] = useState(false);
  const [query, setQuery] = useState("");
  const [openMonths, setOpenMonths] = useState({});
  const [openId, setOpenId] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState(null);
  const camRef = useRef(null);
  const galRef = useRef(null);

  const onPick = (e) => {
    const list = Array.from(e.target.files || []);
    e.target.value = "";
    if (list.length) setPhotos((prev) => [...prev, ...list]);
  };
  const submit = async () => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增收貨單"); return; }
    if (!photos.length) { showToast("請先拍照或選擇收貨單照片"); return; }
    setBusy(true);
    const ok = await addReceipt({ ...form, files: photos });
    setBusy(false);
    if (ok) { setPhotos([]); setForm({ date: todayStr(), title: "", note: "" }); }
  };
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return receipts;
    return receipts.filter((r) => (r.title || "").toLowerCase().includes(q) || (r.note || "").toLowerCase().includes(q) || (r.date || "").includes(q));
  }, [receipts, query]);
  const startEdit = (r) => { setEditingId(r.id); setDraft({ date: r.date, title: r.title || "", note: r.note || "" }); };
  const saveEdit = (id) => { updateReceipt(id, { date: draft.date, title: draft.title.trim(), note: draft.note.trim() }); setEditingId(null); setDraft(null); };

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">收貨單 · DELIVERY SLIPS</div>
          <h1 className="wel-h1">收貨單相簿</h1>
        </div>
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增收貨單" />}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Camera size={14} /> 新增收貨單照片</div>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} />
          </Field>
          <Field label="廠商 / 標題" wide>
            <input list="wel-receipt-suppliers" placeholder="例：賀倫水電材料行" value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} />
            <datalist id="wel-receipt-suppliers">{suppliers.map((s) => <option key={s.id} value={s.name} />)}</datalist>
          </Field>
          <Field label="備註" wide>
            <input placeholder="選填，例：10/20 水管進貨" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} />
          </Field>
        </div>
        <input ref={camRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={onPick} />
        <input ref={galRef} type="file" accept="image/*" multiple style={{ display: "none" }} onChange={onPick} />
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
          <button type="button" className="wel-btn-ghost" onClick={() => camRef.current && camRef.current.click()}><Camera size={14} /> 拍照</button>
          <button type="button" className="wel-btn-ghost" onClick={() => galRef.current && galRef.current.click()}><ImagePlus size={14} /> 選擇照片（可多張）</button>
        </div>
        {photos.length > 0 && (
          <div className="muted" style={{ fontSize: 12, marginBottom: 8, display: "flex", flexDirection: "column", gap: 4 }}>
            {photos.map((f, i) => (
              <div key={i} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <span style={{ wordBreak: "break-all" }}>{f.name}</span><span className="mono">{fmtSize(f.size)}</span>
                <button type="button" className="wel-icon-btn" onClick={() => setPhotos((prev) => prev.filter((_, k) => k !== i))}><X size={12} /></button>
              </div>
            ))}
          </div>
        )}
        <button type="button" className="wel-btn-primary" disabled={busy} onClick={submit}>{busy ? "上傳中…（照片較大時需要一點時間）" : <><Plus size={15} /> 儲存收貨單</>}</button>
      </div>

      <div className="wel-card">
        <div className="wel-card-title-row">
          <div className="wel-card-title"><CalendarRange size={14} /> 收貨單紀錄（依年月點開查詢）</div>
          <div className="wel-filter">
            <input className="wel-scan-input" placeholder="搜尋廠商／備註／日期" value={query} onChange={(e) => setQuery(e.target.value)} style={{ width: 190 }} />
          </div>
        </div>
        {filtered.length === 0 && <Empty text={receipts.length ? "沒有符合的收貨單" : "尚未新增任何收貨單"} />}
        {attendanceByMonth(filtered).map((mg) => {
          const mOpen = !!openMonths[mg.month] || !!query.trim();
          return (
            <div key={mg.month} className="wel-item-floor-group">
              <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => setOpenMonths((prev) => ({ ...prev, [mg.month]: !prev[mg.month] }))}>
                {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                {mg.month} · {mg.items.length} 張
              </button>
              {mOpen && mg.items.map((r) => {
                const isOpen = openId === r.id;
                return (
                  <div key={r.id} className="wel-discussion-card">
                    {editingId === r.id ? (
                      <div className="wel-timeline-wrap" style={{ padding: 0 }}>
                        <div className="wel-form-grid">
                          <Field label="日期"><input type="date" value={draft.date} onChange={(e) => setDraft({ ...draft, date: e.target.value })} /></Field>
                          <Field label="廠商 / 標題" wide><input value={draft.title} onChange={(e) => setDraft({ ...draft, title: e.target.value })} /></Field>
                          <Field label="備註" wide><input value={draft.note} onChange={(e) => setDraft({ ...draft, note: e.target.value })} /></Field>
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                          <button type="button" className="wel-btn-primary" onClick={() => saveEdit(r.id)}>儲存修改</button>
                          <button type="button" className="wel-btn-ghost" onClick={() => { setEditingId(null); setDraft(null); }}>取消</button>
                        </div>
                      </div>
                    ) : (
                      <>
                        <div className="wel-discussion-head wel-row-clickable" onClick={() => setOpenId(isOpen ? null : r.id)}>
                          {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                          <span className="mono">{r.date}</span>
                          <span className="wel-discussion-topic">{r.title || "（未填標題）"}</span>
                          {isAllSites && <span className="wel-tag">{siteById[r.siteId]?.name || "—"}</span>}
                          <span className="muted mono" style={{ fontSize: 11.5 }}>{r.files.length} 張</span>
                          <button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); startEdit(r); }}><Pencil size={14} /></button>
                          <button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteReceipt(r.id); }}><Trash2 size={14} /></button>
                        </div>
                        {r.note && <div className="wel-discussion-note muted">{r.note}</div>}
                        {isOpen && (
                          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 8 }}>
                            {r.files.map((f) => <ReceiptPhoto key={f.id} file={f} />)}
                          </div>
                        )}
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// Photo -> OCR -> editable table. Nothing is saved until the user presses the confirm button.
function ReceiptScanner({ mode, categories, materialItems, suppliers, defaultCategoryId, onConfirmOrders, onConfirmItems, onConfirmEstimate, onClose, showToast }) {
  const isEstimate = mode === "estimate";
  const isOrder = mode === "order" || isEstimate;
  const [phase, setPhase] = useState("pick");
  const [file, setFile] = useState(null);
  const [previewUrl, setPreviewUrl] = useState("");
  const [progress, setProgress] = useState({ label: "", pct: 0 });
  const [rawText, setRawText] = useState("");
  const [rows, setRows] = useState([]);
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);
  const [meta, setMeta] = useState({ date: todayStr(), supplier: "", hasInvoice: true, applyTax: false, taxRate: 5, saveToAlbum: true, addNewItems: true, categoryId: defaultCategoryId || categories[0]?.id || "" });
  const camRef = useRef(null);
  const galRef = useRef(null);
  const ocrRef = useRef(null);
  const seq = useRef(0);

  useEffect(() => () => { if (ocrRef.current) ocrRef.current.terminateOcr(); }, []);
  useEffect(() => () => { if (previewUrl) URL.revokeObjectURL(previewUrl); }, [previewUrl]);

  const scan = async (f) => {
    setFile(f);
    setPreviewUrl(URL.createObjectURL(f));
    setPhase("scanning");
    setError("");
    setProgress({ label: "載入辨識引擎…（第一次使用需下載約 7 MB，之後會快很多）", pct: 0 });
    try {
      const ocr = ocrRef.current || (ocrRef.current = await import("./ocr.js"));
      const text = await ocr.recognizeText(f, (m) => {
        const st = String(m.status || "");
        const label = st.includes("recognizing") ? "辨識文字中…" : st.includes("language") ? "載入中文字庫…" : "載入辨識引擎…";
        setProgress({ label, pct: Math.round((m.progress || 0) * 100) });
      });
      setRawText(text);
      const parsed = isOrder ? parseReceiptText(text, materialItems) : parseItemListText(text, materialItems);
      setRows(parsed.map((r) => (isOrder
        ? { ...r, include: true, categoryId: r.categoryId || meta.categoryId }
        : { ...r, include: !r.exists && (r.unit !== "" || r.defaultPrice !== ""), categoryId: meta.categoryId })));
    } catch (e) {
      console.error("ocr failed", e);
      setRows([]);
      setError("辨識失敗了。請重新拍一張（光線充足、拍正、對焦清楚），或直接在下方手動新增。");
    }
    setPhase("review");
  };
  const onPick = (e) => {
    const f = (e.target.files || [])[0];
    e.target.value = "";
    if (f) scan(f);
  };

  const patchRow = (id, patch) => setRows((prev) => prev.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const removeRow = (id) => setRows((prev) => prev.filter((r) => r.id !== id));
  const addRow = () => setRows((prev) => [...prev, isOrder
    ? { id: "m" + ++seq.current, raw: "", name: "", matchedId: "", categoryId: meta.categoryId, quantity: "", unit: "", unitPrice: "", include: true }
    : { id: "m" + ++seq.current, raw: "", name: "", unit: "", defaultPrice: "", exists: false, include: true, categoryId: meta.categoryId }]);
  const setName = (id, name) => {
    const item = materialItems.find((m) => m.name === name.trim());
    setRows((prev) => prev.map((r) => {
      if (r.id !== id) return r;
      if (!item) return { ...r, name, matchedId: "" };
      return isOrder
        ? { ...r, name, matchedId: item.id, categoryId: item.categoryId || r.categoryId, unit: r.unit || item.unit || "", unitPrice: r.unitPrice === "" ? (item.defaultPrice ?? "") : r.unitPrice }
        : { ...r, name, exists: true, include: false };
    }));
  };

  const confirm = async () => {
    const chosen = rows.filter((r) => r.include);
    if (!chosen.length) { showToast("請至少勾選一列要加入的項目"); return; }
    for (const r of chosen) {
      const n = rows.indexOf(r) + 1;
      if (!r.name.trim()) { showToast(`第 ${n} 列沒有品項名稱`); return; }
      if (isOrder) {
        if (!(Number(r.quantity) > 0)) { showToast(`第 ${n} 列「${r.name}」的數量要大於 0`); return; }
        if (r.unitPrice === "" || !(Number(r.unitPrice) >= 0)) { showToast(`第 ${n} 列「${r.name}」缺少單價`); return; }
      }
    }
    setSaving(true);
    try {
      if (isEstimate) await onConfirmEstimate(chosen);
      else if (isOrder) await onConfirmOrders(chosen, meta, meta.saveToAlbum ? file : null);
      else onConfirmItems(chosen);
      onClose();
    } finally {
      setSaving(false);
    }
  };

  const total = rows.filter((r) => r.include).reduce((sum, r) => sum + (isOrder ? (Number(r.quantity) || 0) * (Number(r.unitPrice) || 0) : 0), 0);
  const taxOn = isOrder && meta.hasInvoice && meta.applyTax;
  const taxTotal = taxOn ? total * (Number(meta.taxRate) || 0) / 100 : 0;
  const catSelect = (value, onChange) => (
    <select className="wel-scan-input" value={value} onChange={onChange}>{categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}</select>
  );

  return (
    <div className="wel-card wel-form" style={{ borderColor: "var(--teal)" }}>
      <div className="wel-card-title-row">
        <div className="wel-card-title"><ScanLine size={14} /> {isEstimate ? "拍照辨識估價單 → 校對 → 加入估價項目" : isOrder ? "拍照辨識收貨單 → 校對 → 存入叫貨" : "拍照掃描品項 → 校對 → 加入品項清單"}</div>
        <button type="button" className="wel-btn-ghost" onClick={onClose}>關閉</button>
      </div>
      <input ref={camRef} type="file" accept="image/*" capture="environment" style={{ display: "none" }} onChange={onPick} />
      <input ref={galRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onPick} />

      {phase === "pick" && (
        <>
          <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>
            {isEstimate ? "拍下簽約估價單，網頁會在你的手機或電腦上辨識品名、數量與單價（照片不會傳給第三方）。辨識結果只是「建議」，一定會先讓你校對修改，按確認後才會加入。" : isOrder ? "拍下收貨單／送貨單，網頁會在你的手機或電腦上直接辨識品名與數量（照片不會傳給第三方）。辨識結果只是「建議」，一定會先讓你校對修改，按確認後才會存入。" : "拍下品項清單、價目表或貨品標示，網頁會辨識出品名、單位與單價，讓你校對修改後再加入品項清單。"}
            <br />小提醒：光線充足、整張拍正、字要清楚；手寫字辨識較不準，請多留意校對。
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button type="button" className="wel-btn-primary" onClick={() => camRef.current && camRef.current.click()}><Camera size={15} /> 拍照</button>
            <button type="button" className="wel-btn-ghost" onClick={() => galRef.current && galRef.current.click()}><ImagePlus size={15} /> 選擇照片</button>
          </div>
        </>
      )}

      {phase === "scanning" && (
        <div style={{ padding: "12px 0" }}>
          <div style={{ fontSize: 13, marginBottom: 6 }}>{progress.label} {progress.pct ? progress.pct + "%" : ""}</div>
          <div className="wel-pipe-track"><div className="wel-pipe-fill" style={{ width: Math.max(progress.pct, 4) + "%", background: "var(--teal)" }} /></div>
        </div>
      )}

      {phase === "review" && (
        <>
          {error && <div className="wel-mismatch" style={{ marginBottom: 8 }}>{error}</div>}
          {isEstimate ? null : isOrder ? (
            <div className="wel-form-grid">
              <Field label="日期"><input type="date" value={meta.date} onChange={(e) => setMeta({ ...meta, date: e.target.value })} /></Field>
              <Field label="廠商 / 供應商">
                <select value={meta.supplier} onChange={(e) => setMeta({ ...meta, supplier: e.target.value })}>
                  <option value="">（未指定）</option>
                  {suppliers.map((s) => <option key={s.id} value={s.name}>{s.name}</option>)}
                </select>
              </Field>
              <Field label="有無發票">
                <select value={meta.hasInvoice ? "yes" : "no"} onChange={(e) => setMeta({ ...meta, hasInvoice: e.target.value === "yes" })}>
                  <option value="yes">有發票</option><option value="no">無發票</option>
                </select>
              </Field>
              {meta.hasInvoice && (
                <Field label="營業稅" wide>
                  <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                      <input type="checkbox" checked={meta.applyTax} onChange={(e) => setMeta({ ...meta, applyTax: e.target.checked })} style={{ width: 16, height: 16 }} />
                      計算稅額
                    </label>
                    {meta.applyTax && (
                      <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                        稅率
                        <input type="number" min="0" step="0.1" value={meta.taxRate} onChange={(e) => setMeta({ ...meta, taxRate: e.target.value })} style={{ width: 60, height: 32 }} />
                        %
                      </span>
                    )}
                  </div>
                </Field>
              )}
            </div>
          ) : (
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginBottom: 8 }}>
              <span className="muted" style={{ fontSize: 12 }}>全部歸入類別：</span>
              {catSelect(meta.categoryId, (e) => setMeta({ ...meta, categoryId: e.target.value }))}
              <button type="button" className="wel-btn-ghost" onClick={() => setRows((prev) => prev.map((r) => ({ ...r, categoryId: meta.categoryId })))}>套用到全部列</button>
            </div>
          )}
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap", alignItems: "flex-start" }}>
            {previewUrl && (
              <div style={{ flex: "0 0 240px", maxWidth: "100%" }}>
                <img src={previewUrl} alt="收貨單照片" onClick={() => window.open(previewUrl, "_blank")} style={{ width: "100%", borderRadius: 8, border: "1px solid var(--border)", cursor: "zoom-in" }} />
                <div className="muted" style={{ fontSize: 11, marginTop: 4 }}>點照片可放大對照</div>
                <button type="button" className="wel-btn-ghost" style={{ marginTop: 6 }} onClick={() => { setPhase("pick"); setRows([]); setRawText(""); }}>重新拍一張</button>
              </div>
            )}
            <div style={{ flex: "1 1 420px", minWidth: 0, overflowX: "auto" }}>
              <datalist id="wel-scan-items">{materialItems.map((m) => <option key={m.id} value={m.name} />)}</datalist>
              <table className="wel-table">
                <thead>
                  {isOrder
                    ? <tr><th></th><th>品項</th><th>類別</th><th className="right">數量</th><th>單位</th><th className="right">單價</th><th className="right">小計</th>{taxOn && <th className="right">稅額</th>}<th></th></tr>
                    : <tr><th></th><th>品項名稱</th><th>類別</th><th>單位</th><th className="right">預設單價</th><th></th></tr>}
                </thead>
                <tbody>
                  {rows.length === 0 && <tr><td colSpan={isOrder ? (taxOn ? 9 : 8) : 6}><Empty text="沒有辨識到項目，請按下方「新增一列」手動輸入" /></td></tr>}
                  {rows.map((r, idx) => (
                    <tr key={r.id} style={isOrder && !r.matchedId ? { background: "rgba(232,163,61,0.08)" } : undefined}>
                      <td><input type="checkbox" checked={r.include} onChange={(e) => patchRow(r.id, { include: e.target.checked })} style={{ width: 16, height: 16 }} /></td>
                      <td style={{ minWidth: 190 }}>
                        <input className="wel-scan-input wel-scan-name" list="wel-scan-items" value={r.name} onChange={(e) => setName(r.id, e.target.value)} placeholder="品項名稱" />
                        <div className="muted" style={{ fontSize: 11, marginTop: 2 }}>
                          {isOrder ? (r.matchedId ? "✓ 已對應品項清單" : "未在品項清單內（可改選清單品項，或當作新品項）") : (r.exists ? "已在品項清單內（預設不加入）" : "新品項")}
                          {r.raw ? ` · 原文：${r.raw}` : ""}
                        </div>
                      </td>
                      <td>{catSelect(r.categoryId, (e) => patchRow(r.id, { categoryId: e.target.value }))}</td>
                      {isOrder ? (
                        <>
                          <td className="right"><input className="wel-scan-input" type="number" step="any" min="0" style={{ width: 74 }} value={r.quantity} onChange={(e) => patchRow(r.id, { quantity: e.target.value })} /></td>
                          <td><input className="wel-scan-input" style={{ width: 56 }} value={r.unit} onChange={(e) => patchRow(r.id, { unit: e.target.value })} /></td>
                          <td className="right"><input className="wel-scan-input" type="number" step="any" min="0" style={{ width: 86 }} value={r.unitPrice} onChange={(e) => patchRow(r.id, { unitPrice: e.target.value })} /></td>
                          <td className="right mono">{fmtMoney((Number(r.quantity) || 0) * (Number(r.unitPrice) || 0))}</td>
                          {taxOn && <td className="right mono muted">{fmtMoney((Number(r.quantity) || 0) * (Number(r.unitPrice) || 0) * (Number(meta.taxRate) || 0) / 100)}</td>}
                        </>
                      ) : (
                        <>
                          <td><input className="wel-scan-input" style={{ width: 64 }} value={r.unit} onChange={(e) => patchRow(r.id, { unit: e.target.value })} /></td>
                          <td className="right"><input className="wel-scan-input" type="number" step="any" min="0" style={{ width: 90 }} value={r.defaultPrice} onChange={(e) => patchRow(r.id, { defaultPrice: e.target.value })} /></td>
                        </>
                      )}
                      <td><button type="button" className="wel-icon-btn" title={`刪除第 ${idx + 1} 列`} onClick={() => removeRow(r.id)}><Trash2 size={13} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
                <button type="button" className="wel-btn-ghost" onClick={addRow}><Plus size={13} /> 新增一列</button>
                {isOrder && (
                  <span className="muted mono" style={{ fontSize: 12 }}>
                    已勾選合計：{fmtMoney(total)}{taxOn ? `　稅額（${Number(meta.taxRate) || 0}%）：${fmtMoney(taxTotal)}　含稅：${fmtMoney(total + taxTotal)}` : ""}
                  </span>
                )}
              </div>
              {rawText && (
                <details style={{ marginTop: 8 }}>
                  <summary className="muted" style={{ fontSize: 12, cursor: "pointer" }}>查看辨識到的原始文字</summary>
                  <pre style={{ whiteSpace: "pre-wrap", fontSize: 11.5, margin: "6px 0 0" }}>{rawText}</pre>
                </details>
              )}
            </div>
          </div>
          {isOrder && !isEstimate && (
            <div style={{ display: "flex", gap: 16, flexWrap: "wrap", margin: "10px 0", fontSize: 13 }}>
              <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
                <input type="checkbox" checked={meta.addNewItems} onChange={(e) => setMeta({ ...meta, addNewItems: e.target.checked })} style={{ width: 16, height: 16 }} /> 新品項同時加入品項清單
              </label>
              <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
                <input type="checkbox" checked={meta.saveToAlbum} onChange={(e) => setMeta({ ...meta, saveToAlbum: e.target.checked })} style={{ width: 16, height: 16 }} /> 同時把照片存到「收貨單」相簿
              </label>
            </div>
          )}
          <button type="button" className="wel-btn-primary" disabled={saving} onClick={confirm} style={{ marginTop: 8 }}>
            {saving ? "儲存中…" : isEstimate ? `確認無誤，加入估價單（${rows.filter((r) => r.include).length} 項）` : isOrder ? `確認無誤，加入叫貨（${rows.filter((r) => r.include).length} 筆）` : `確認無誤，加入品項清單（${rows.filter((r) => r.include).length} 項）`}
          </button>
        </>
      )}
    </div>
  );
}

// Free-text field that saves when you leave it (not on every keystroke) and picks up outside changes.
function NoteInput({ value, onCommit, placeholder = "備註" }) {
  const [draft, setDraft] = useState(value || "");
  useEffect(() => { setDraft(value || ""); }, [value]);
  const commit = () => { if (draft !== (value || "")) onCommit(draft.trim()); };
  return (
    <input
      value={draft} placeholder={placeholder} title={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
    />
  );
}

// 多項一次叫貨：依類別勾選多個品項、各自填數量，一次寫入叫貨紀錄。
function BulkOrderPanel({ categories, catById, materialItems, suppliers, defaultCategoryId, onConfirm, onClose, showToast }) {
  const [categoryId, setCategoryId] = useState(defaultCategoryId || categories[0]?.id || "");
  const [meta, setMeta] = useState({ date: todayStr(), supplier: "", hasInvoice: true, applyTax: false, taxRate: 5 });
  const [sel, setSel] = useState({});
  const [saving, setSaving] = useState(false);

  const items = useMemo(
    () => materialItems.filter((m) => m.categoryId === categoryId).sort((a, b) => a.name.localeCompare(b.name, "zh-Hant")),
    [materialItems, categoryId]
  );
  const uncategorizedCount = materialItems.filter((m) => !m.categoryId || !catById[m.categoryId]).length;
  const base = (m) => ({ include: false, quantity: "", unit: m.unit || "", unitPrice: m.defaultPrice === "" || m.defaultPrice == null ? "" : m.defaultPrice });
  const rowOf = (m) => sel[m.id] || base(m);
  const patch = (m, p) => setSel((prev) => ({ ...prev, [m.id]: { ...(prev[m.id] || base(m)), ...p } }));
  const chosen = materialItems.filter((m) => sel[m.id] && sel[m.id].include);
  const lineTotal = (m) => (Number(sel[m.id].quantity) || 0) * (Number(sel[m.id].unitPrice) || 0);
  const total = chosen.reduce((sum, m) => sum + lineTotal(m), 0);
  const taxOn = meta.hasInvoice && meta.applyTax;
  const taxTotal = taxOn ? total * (Number(meta.taxRate) || 0) / 100 : 0;

  const confirm = async () => {
    if (!chosen.length) { showToast("請先勾選要叫貨的品項並填寫數量"); return; }
    for (const m of chosen) {
      const r = sel[m.id];
      if (!(Number(r.quantity) > 0)) { showToast(`「${m.name}」的數量要大於 0`); return; }
      if (r.unitPrice === "" || !(Number(r.unitPrice) >= 0)) { showToast(`「${m.name}」缺少單價`); return; }
    }
    setSaving(true);
    try {
      await onConfirm(
        chosen.map((m) => ({ name: m.name, categoryId: m.categoryId, quantity: sel[m.id].quantity, unit: sel[m.id].unit, unitPrice: sel[m.id].unitPrice })),
        { ...meta, note: "", addNewItems: false, saveToAlbum: false },
        null
      );
      onClose();
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="wel-card wel-form" style={{ borderColor: "var(--teal)" }}>
      <div className="wel-card-title-row">
        <div className="wel-card-title"><ListChecks size={14} /> 多項一次叫貨（勾選品項、填數量，一次存入）</div>
        <button type="button" className="wel-btn-ghost" onClick={onClose}>關閉</button>
      </div>
      <div className="wel-form-grid">
        <Field label="日期"><input type="date" value={meta.date} onChange={(e) => setMeta({ ...meta, date: e.target.value })} /></Field>
        <Field label="廠商 / 供應商">
          <select value={meta.supplier} onChange={(e) => setMeta({ ...meta, supplier: e.target.value })}>
            <option value="">（未指定）</option>
            {suppliers.map((s) => <option key={s.id} value={s.name}>{s.name}</option>)}
          </select>
        </Field>
        <Field label="有無發票">
          <select value={meta.hasInvoice ? "yes" : "no"} onChange={(e) => setMeta({ ...meta, hasInvoice: e.target.value === "yes" })}>
            <option value="yes">有發票</option><option value="no">無發票</option>
          </select>
        </Field>
        {meta.hasInvoice && (
          <Field label="營業稅" wide>
            <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                <input type="checkbox" checked={meta.applyTax} onChange={(e) => setMeta({ ...meta, applyTax: e.target.checked })} style={{ width: 16, height: 16 }} />
                計算稅額
              </label>
              {meta.applyTax && (
                <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                  稅率
                  <input type="number" min="0" step="0.1" value={meta.taxRate} onChange={(e) => setMeta({ ...meta, taxRate: e.target.value })} style={{ width: 60, height: 32 }} />
                  %
                </span>
              )}
            </div>
          </Field>
        )}
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", margin: "4px 0 8px" }}>
        <span className="muted" style={{ fontSize: 12 }}>類別：</span>
        <select className="wel-scan-input" value={categoryId} onChange={(e) => setCategoryId(e.target.value)}>
          {categories.map((c) => <option key={c.id} value={c.id}>{c.name}{materialItems.filter((m) => m.categoryId === c.id).length ? "" : "（無品項）"}</option>)}
        </select>
        <span className="muted" style={{ fontSize: 12 }}>切換類別時，已勾選的品項都會保留，可以跨類別一起叫貨。</span>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table className="wel-table">
          <thead><tr><th></th><th>品項</th><th className="right">數量</th><th>單位</th><th className="right">單價</th><th className="right">小計</th></tr></thead>
          <tbody>
            {items.length === 0 && <tr><td colSpan={6}><Empty text="這個類別還沒有品項，可到「品項類別」或叫貨頁的「管理品項清單」新增" /></td></tr>}
            {items.map((m) => {
              const r = rowOf(m);
              return (
                <tr key={m.id} style={r.include ? { background: "rgba(74,144,164,0.10)" } : undefined}>
                  <td><input type="checkbox" checked={r.include} onChange={(e) => patch(m, { include: e.target.checked })} style={{ width: 16, height: 16 }} /></td>
                  <td style={{ minWidth: 200 }}>{m.name}</td>
                  <td className="right"><input className="wel-scan-input" type="number" step="any" min="0" style={{ width: 80 }} value={r.quantity} onChange={(e) => patch(m, { quantity: e.target.value, include: Number(e.target.value) > 0 })} /></td>
                  <td><input className="wel-scan-input" style={{ width: 56 }} value={r.unit} onChange={(e) => patch(m, { unit: e.target.value })} /></td>
                  <td className="right"><input className="wel-scan-input" type="number" step="any" min="0" style={{ width: 90 }} value={r.unitPrice} onChange={(e) => patch(m, { unitPrice: e.target.value })} /></td>
                  <td className="right mono">{fmtMoney((Number(r.quantity) || 0) * (Number(r.unitPrice) || 0))}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {uncategorizedCount > 0 && <div className="muted" style={{ fontSize: 11, marginTop: 6 }}>另有 {uncategorizedCount} 個未歸類品項沒有列在這裡，請到「品項類別」頁指定類別。</div>}

      <div style={{ marginTop: 10, fontSize: 13 }}>
        <div className="muted" style={{ fontSize: 12, marginBottom: 4 }}>
          已選 {chosen.length} 項{chosen.length ? "：" : ""}
          {chosen.map((m) => `${m.name}×${sel[m.id].quantity || 0}`).join("、")}
        </div>
        <div className="mono" style={{ fontSize: 13 }}>
          合計：{fmtMoney(total)}{taxOn ? `　稅額（${Number(meta.taxRate) || 0}%）：${fmtMoney(taxTotal)}　含稅：${fmtMoney(total + taxTotal)}` : ""}
        </div>
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <button type="button" className="wel-btn-primary" disabled={saving} onClick={confirm}>{saving ? "儲存中…" : `確認叫貨（${chosen.length} 筆）`}</button>
        <button type="button" className="wel-btn-ghost" onClick={() => setSel({})}>清除全部勾選</button>
      </div>
    </div>
  );
}

function ContractFiles({ contract, addContractFile, deleteContractFile, openContractFile, label = "合約檔案", uploadText = "上傳合約" }) {
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);
  const files = contract.files || [];
  const onPick = async (e) => {
    const list = Array.from(e.target.files || []);
    e.target.value = "";
    if (!list.length) return;
    setBusy(true);
    for (const f of list) await addContractFile(contract.id, f);
    setBusy(false);
  };
  return (
    <div className="wel-floor-summary-line" style={{ display: "flex", flexDirection: "column", gap: 4 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
        <span className="muted" style={{ display: "inline-flex", alignItems: "center", gap: 4 }}><Paperclip size={12} /> {label}（{files.length}）</span>
        <input ref={inputRef} type="file" accept="application/pdf,image/*,.doc,.docx" multiple style={{ display: "none" }} onChange={onPick} />
        <button type="button" className="wel-btn-ghost" disabled={busy} onClick={() => inputRef.current && inputRef.current.click()}>
          {busy ? "上傳中…" : <><Upload size={12} /> {uploadText}</>}
        </button>
      </div>
      {files.map((f) => (
        <div key={f.id} className="wel-attendance-item">
          <button type="button" className="wel-btn-ghost" style={{ padding: "2px 8px" }} onClick={() => openContractFile(f)}><Download size={12} /> 開啟</button>
          <span style={{ flex: 1, wordBreak: "break-all" }}>{f.name}</span>
          <span className="mono muted">{fmtSize(f.size)}</span>
          <span className="mono muted">{f.uploadedAt}</span>
          <button className="wel-icon-btn" onClick={() => deleteContractFile(contract.id, f)}><Trash2 size={13} /></button>
        </div>
      ))}
    </div>
  );
}

function naturalFloorSort(a, b) {
  const na = parseInt((a.match(/\d+/) || [])[0], 10);
  const nb = parseInt((b.match(/\d+/) || [])[0], 10);
  if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
  return a.localeCompare(b);
}

function attendanceByMonth(logs) {
  const map = {};
  logs.forEach((l) => {
    const mk = (l.date || "").slice(0, 7) || "未知月份";
    if (!map[mk]) map[mk] = [];
    map[mk].push(l);
  });
  return Object.entries(map)
    .sort((a, b) => (a[0] < b[0] ? 1 : -1))
    .map(([month, items]) => ({ month, items: items.slice().sort((a, b) => (a.date < b.date ? 1 : -1)) }));
}

function FloorItemsPanel({
  contract, floorItems, isAllSites,
  addTemplateRow, updateTemplateRow, removeTemplateRow, applyTemplateToFloors,
  addFloorItem, updateFloorItem, deleteFloorItem, deleteFloorGroup,
  addAttendance, deleteAttendance, updateAttendance, contractWorkLogs, sites, setCurrentSiteId, showToast,
  contractItems, addContractItem, addContractItemsBatch, addPaymentFromItem, updateContractItem, deleteContractItem, deleteContractItemsByFloor, deleteContractItemsByName,
}) {
  const [range, setRange] = useState({ fromN: "", toN: "", format: "{n}F", extra: "", overwrite: false });
  const [quickAdd, setQuickAdd] = useState({ floor: "", itemName: "", amount: "" });
  const [attend, setAttend] = useState({ date: todayStr(), headcount: 1, note: "" });
  const [openAttendMonths, setOpenAttendMonths] = useState({});
  const [editingAttendId, setEditingAttendId] = useState(null);
  const [attendDraft, setAttendDraft] = useState(null);
  const startEditAttend = (l) => { setEditingAttendId(l.id); setAttendDraft({ date: l.date, headcount: l.headcount, note: l.note || "" }); };
  const cancelEditAttend = () => { setEditingAttendId(null); setAttendDraft(null); };
  const saveEditAttend = (id) => { updateAttendance(id, { ...attendDraft, headcount: Number(attendDraft.headcount) || 1 }); setEditingAttendId(null); setAttendDraft(null); };
  const [newItem, setNewItem] = useState({ name: "", amount: "", percent: "", date: todayStr(), note: "", floor: "" });
  const [itemFloorRange, setItemFloorRange] = useState({ fromN: "", toN: "", format: "{n}F", extra: "" });
  const [newTemplateRow, setNewTemplateRow] = useState({ name: "", amount: "" });
  const [addingNewType, setAddingNewType] = useState(false);
  const [newTypeName, setNewTypeName] = useState("");
  const [addingNewFloor, setAddingNewFloor] = useState(false);
  const [newFloorName, setNewFloorName] = useState("");
  const [openFloors, setOpenFloors] = useState({}); // { [name+"|"+floor]: true }
  const template = contract.templateItems || [];
  const itemTypeNames = useMemo(() => {
    const seen = new Set();
    const list = [];
    template.forEach((t) => { if (t.name.trim() && !seen.has(t.name)) { seen.add(t.name); list.push(t.name); } });
    contractItems.forEach((it) => { if (it.name.trim() && !seen.has(it.name)) { seen.add(it.name); list.push(it.name); } });
    return list.sort((a, b) => a.localeCompare(b, "zh-Hant"));
  }, [template, contractItems]);
  const knownFloors = useMemo(() => {
    const seen = new Set();
    const list = [];
    contractItems.forEach((it) => { if (it.floor && !seen.has(it.floor)) { seen.add(it.floor); list.push(it.floor); } });
    return list.sort(naturalFloorSort);
  }, [contractItems]);
  const itemsTotal = contractItems.reduce((s, it) => s + (Number(it.amount) || 0), 0);
  const contractTotalPriceNum = Number(contract.totalPrice) || 0;

  const onItemPercentChange = (id, val) => {
    const patch = { percent: val };
    if (contractTotalPriceNum > 0 && val !== "") patch.amount = Math.round(contractTotalPriceNum * Number(val) / 100);
    updateContractItem(id, patch);
  };
  const onItemAmountChange = (id, val) => {
    const patch = { amount: val };
    if (contractTotalPriceNum > 0 && val !== "") patch.percent = Number((Number(val) / contractTotalPriceNum * 100).toFixed(1));
    updateContractItem(id, patch);
  };
  const onNewPercentChange = (val) => {
    setNewItem((f) => ({ ...f, percent: val, amount: contractTotalPriceNum > 0 && val !== "" ? Math.round(contractTotalPriceNum * Number(val) / 100) : f.amount }));
  };
  const onNewAmountChange = (val) => {
    setNewItem((f) => ({ ...f, amount: val, percent: contractTotalPriceNum > 0 && val !== "" ? Number((Number(val) / contractTotalPriceNum * 100).toFixed(1)) : f.percent }));
  };
  const handleTypeSelect = (val) => {
    if (val === "__new__") { setAddingNewType(true); return; }
    setNewItem((f) => ({ ...f, name: val }));
  };
  const confirmNewType = () => {
    const nm = newTypeName.trim();
    if (!nm) { showToast("請輸入新項目類型名稱"); return; }
    addTemplateRow(contract.id, nm, "");
    setNewItem((f) => ({ ...f, name: nm }));
    setNewTypeName("");
    setAddingNewType(false);
  };
  const handleFloorSelect = (val) => {
    if (val === "__new__") { setAddingNewFloor(true); return; }
    setNewItem((f) => ({ ...f, floor: val }));
  };
  const confirmNewFloor = () => {
    const nm = newFloorName.trim();
    if (!nm) { showToast("請輸入新樓層名稱"); return; }
    setNewItem((f) => ({ ...f, floor: nm }));
    setNewFloorName("");
    setAddingNewFloor(false);
  };
  const toggleFloorOpen = (name, floor) => {
    const key = name + "|" + floor;
    setOpenFloors((prev) => ({ ...prev, [key]: !prev[key] }));
  };
  const buildItemFloorList = () => {
    const list = [];
    const from = parseInt(itemFloorRange.fromN, 10);
    const to = parseInt(itemFloorRange.toN, 10);
    if (!isNaN(from) && !isNaN(to) && itemFloorRange.format.includes("{n}")) {
      const step = to >= from ? 1 : -1;
      for (let n = from; step > 0 ? n <= to : n >= to; n += step) list.push(itemFloorRange.format.replace("{n}", n));
    }
    itemFloorRange.extra.split(",").map((s) => s.trim()).filter(Boolean).forEach((f) => { if (!list.includes(f)) list.push(f); });
    return list;
  };
  const handleAddNewItem = () => {
    if (!newItem.name.trim()) { showToast("請先從下拉選單選擇項目，或新增一個項目類型"); return; }
    const rangeFloors = buildItemFloorList();
    if (rangeFloors.length > 0) {
      addContractItemsBatch(contract.id, newItem.name, newItem.amount, newItem.date, newItem.note, newItem.percent, rangeFloors);
    } else {
      addContractItem(contract.id, newItem.name, newItem.amount, newItem.date, newItem.note, newItem.percent, newItem.floor || "");
      if (newItem.floor) setOpenFloors((prev) => ({ ...prev, [newItem.name.trim() + "|" + newItem.floor]: true }));
    }
    setNewItem({ name: "", amount: "", percent: "", date: todayStr(), note: "", floor: "" });
    setItemFloorRange({ fromN: "", toN: "", format: "{n}F", extra: "" });
  };
  const importAsPayment = (it) => {
    addPaymentFromItem(contract.id, it.date, it.amount, `${it.name}${it.floor ? `（${it.floor}）` : ""}`, it.id);
  };

  const itemGroups = useMemo(() => {
    const byName = {};
    contractItems.forEach((it) => {
      const key = it.name || "（未命名項目）";
      if (!byName[key]) byName[key] = [];
      byName[key].push(it);
    });
    const sortByDate = (arr) => arr.slice().sort((a, b) => (a.date < b.date ? 1 : -1));
    return Object.entries(byName)
      .map(([name, items]) => {
        const general = items.filter((it) => !it.floor);
        const byFloorMap = {};
        items.filter((it) => it.floor).forEach((it) => {
          if (!byFloorMap[it.floor]) byFloorMap[it.floor] = [];
          byFloorMap[it.floor].push(it);
        });
        const floorGroups = [];
        if (general.length) floorGroups.push({ floor: "", items: sortByDate(general) });
        Object.entries(byFloorMap)
          .sort((a, b) => naturalFloorSort(b[0], a[0]))
          .forEach(([floor, floorItemsArr]) => floorGroups.push({ floor, items: sortByDate(floorItemsArr) }));
        return { name, subtotal: items.reduce((s, it) => s + (Number(it.amount) || 0), 0), floorGroups };
      })
      .sort((a, b) => b.subtotal - a.subtotal);
  }, [contractItems]);

  const renameItemGroup = (oldName, newName) => {
    const nm = newName.trim();
    if (!nm || nm === oldName) return;
    contractItems.filter((it) => it.name === oldName).forEach((it) => updateContractItem(it.id, { name: nm }));
  };

  const buildFloorList = () => {
    const list = [];
    const from = parseInt(range.fromN, 10);
    const to = parseInt(range.toN, 10);
    if (!isNaN(from) && !isNaN(to) && range.format.includes("{n}")) {
      const step = to >= from ? 1 : -1;
      for (let n = from; step > 0 ? n <= to : n >= to; n += step) {
        list.push(range.format.replace("{n}", n));
      }
    }
    range.extra.split(",").map((s) => s.trim()).filter(Boolean).forEach((f) => {
      if (!list.includes(f)) list.push(f);
    });
    return list;
  };

  const handleApply = () => {
    const floors = buildFloorList();
    if (!floors.length) { showToast("請輸入樓層範圍或額外樓層"); return; }
    if (!template.some((t) => t.name.trim() && t.amount !== "")) { showToast("請先在標準層範本中新增至少一個項目"); return; }
    applyTemplateToFloors(contract.id, floors, range.overwrite);
  };

  const disabled = isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined;

  return (
    <div className="wel-floor-panel">
      {/* Multi-item breakdown — one contractor, many separately priced items, each pickable by wheel and applicable across floors */}
      <div className="wel-floor-block">
        <div className="wel-floor-block-title"><ListOrdered size={13} /> 請款項目明細（依項目分組，樓層點開查看；可直接匯入領款）</div>
        {contractTotalPriceNum === 0 && <div className="muted" style={{ fontSize: 11.5, marginBottom: 4 }}>提示：先在上方填寫「發包總價／總工程價格」，才能用百分比自動換算金額。</div>}
        {contractTotalPriceNum > 0 && (
          <div style={{ marginBottom: 10 }}>
            <div className="wel-pipe-track" style={{ height: 14 }}>
              <div className="wel-pipe-fill" style={{ width: `${Math.max(Math.min(100, (itemsTotal / contractTotalPriceNum) * 100), itemsTotal > 0 ? 3 : 0)}%`, background: itemsTotal / contractTotalPriceNum >= 1 ? "var(--green)" : "var(--amber)" }} />
            </div>
            <div className="wel-floor-summary-line muted" style={{ marginTop: 6 }}>
              項目明細合計：<span className="mono">{fmtMoney(itemsTotal)}</span>（約 {(itemsTotal / contractTotalPriceNum * 100).toFixed(1)}% 總工程價格）
            </div>
          </div>
        )}

        {/* Existing items, grouped by item type, then by floor (click a floor to open it) */}
        {contractItems.length === 0 && <div className="muted" style={{ fontSize: 12 }}>尚未新增項目明細</div>}
        {itemGroups.map((g) => (
          <div key={g.name} className="wel-item-name-group">
            <div className="wel-item-name-head">
              <ListOrdered size={13} color="var(--amber)" />
              <ItemNameHeader name={g.name} onRename={(newName) => renameItemGroup(g.name, newName)} />
              <span className="mono muted" style={{ fontSize: 12 }}>{fmtMoney(g.subtotal)}</span>
              <button className="wel-icon-btn" title="刪除這個項目類型的所有樓層紀錄" onClick={() => deleteContractItemsByName(contract.id, g.name)}><Trash2 size={13} /></button>
            </div>
            {g.floorGroups.map((fg) => {
              const key = g.name + "|" + fg.floor;
              const isOpen = !fg.floor || !!openFloors[key]; // general (no-floor) group always shown open
              const subtotal = fg.items.reduce((s, it) => s + (Number(it.amount) || 0), 0);
              return (
                <div key={fg.floor || "__general__"} className="wel-item-floor-group">
                  {fg.floor && (
                    <div className="wel-item-group-toggle-row">
                      <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => toggleFloorOpen(g.name, fg.floor)}>
                        {isOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                        {fg.floor}
                        <span className="muted" style={{ marginLeft: 4 }}>{fmtMoney(subtotal)}</span>
                      </button>
                      <button className="wel-icon-btn" title="刪除這個樓層的所有紀錄" onClick={() => deleteContractItemsByFloor(contract.id, g.name, fg.floor)}><Trash2 size={12} /></button>
                    </div>
                  )}
                  {isOpen && (
                    <div className="wel-template-rows">
                      {fg.items.map((it) => (
                        <div key={it.id} className="wel-template-row wel-template-row-item" style={disabled}>
                          <input type="date" value={it.date || ""} onChange={(e) => updateContractItem(it.id, { date: e.target.value })} />
                          <input placeholder="樓層（留空＝一般項目）" value={it.floor || ""} onChange={(e) => updateContractItem(it.id, { floor: e.target.value })} />
                          <input type="number" min="0" step="any" placeholder="%" title="佔總工程價格百分比" value={it.percent === undefined ? "" : it.percent} onChange={(e) => onItemPercentChange(it.id, e.target.value)} />
                          <input type="number" min="0" step="any" placeholder="金額" value={it.amount} onChange={(e) => onItemAmountChange(it.id, e.target.value)} />
                          <NoteInput value={it.note} onCommit={(v) => updateContractItem(it.id, { note: v })} />
                          <button
                            type="button"
                            className={"wel-pay-import-btn" + (it.paidImported ? " imported" : "")}
                            title={it.paidImported ? "已匯入領款，仍可再次匯入" : "點擊直接匯入到已領款項目"}
                            onClick={() => importAsPayment(it)}
                          >
                            {it.paidImported ? <CheckCircle2 size={13} /> : <Banknote size={13} />}
                          </button>
                          <button className="wel-icon-btn" onClick={() => deleteContractItem(it.id)}><X size={14} /></button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ))}

        {/* New item — wheel-style pick for 樓層／項目, plain custom % input */}
        <div className="wel-new-item-block" style={disabled}>
          <div className="wel-new-item-row">
            <label className="wel-mini-label">日期
              <input type="date" value={newItem.date} onChange={(e) => setNewItem({ ...newItem, date: e.target.value })} />
            </label>
            <label className="wel-mini-label" style={{ flex: 1, minWidth: 140 }}>項目
              <select value={itemTypeNames.includes(newItem.name) ? newItem.name : ""} onChange={(e) => handleTypeSelect(e.target.value)}>
                <option value="" disabled>{itemTypeNames.length ? "請選擇項目…" : "尚無項目類型"}</option>
                {itemTypeNames.map((n) => <option key={n} value={n}>{n}</option>)}
                <option value="__new__">＋ 新增項目類型…</option>
              </select>
            </label>
            <label className="wel-mini-label">樓層
              <select value={knownFloors.includes(newItem.floor) ? newItem.floor : (newItem.floor === "" ? "" : newItem.floor)} onChange={(e) => handleFloorSelect(e.target.value)}>
                <option value="">（一般項目）</option>
                {knownFloors.map((f) => <option key={f} value={f}>{f}</option>)}
                <option value="__new__">＋ 新增樓層…</option>
              </select>
            </label>
            <label className="wel-mini-label">百分比
              <input type="number" min="0" step="any" placeholder="自訂 %" value={newItem.percent} onChange={(e) => onNewPercentChange(e.target.value)} style={{ width: 72 }} />
            </label>
            <label className="wel-mini-label">金額
              <input type="number" min="0" step="any" value={newItem.amount} onChange={(e) => onNewAmountChange(e.target.value)} style={{ width: 90 }} />
            </label>
          </div>

          {addingNewType && (
            <div className="wel-new-type-row">
              <input placeholder="新項目類型名稱，例：配電箱更換" value={newTypeName} onChange={(e) => setNewTypeName(e.target.value)} autoFocus />
              <button type="button" className="wel-btn-ghost" onClick={confirmNewType}>確定新增類型</button>
              <button type="button" className="wel-btn-ghost" onClick={() => { setAddingNewType(false); setNewTypeName(""); }}>取消</button>
            </div>
          )}
          {addingNewFloor && (
            <div className="wel-new-type-row">
              <input placeholder="新樓層名稱，例：3F 或 RF" value={newFloorName} onChange={(e) => setNewFloorName(e.target.value)} autoFocus />
              <button type="button" className="wel-btn-ghost" onClick={confirmNewFloor}>確定新增樓層</button>
              <button type="button" className="wel-btn-ghost" onClick={() => { setAddingNewFloor(false); setNewFloorName(""); }}>取消</button>
            </div>
          )}

          <details className="wel-floor-batch-advanced">
            <summary>進階：一次套用到多個樓層（可選）</summary>
            <div className="wel-item-floor-apply">
              <label>從第<input type="number" value={itemFloorRange.fromN} onChange={(e) => setItemFloorRange({ ...itemFloorRange, fromN: e.target.value })} placeholder="3" /></label>
              <label>到第<input type="number" value={itemFloorRange.toN} onChange={(e) => setItemFloorRange({ ...itemFloorRange, toN: e.target.value })} placeholder="15" /></label>
              <label>格式<input value={itemFloorRange.format} onChange={(e) => setItemFloorRange({ ...itemFloorRange, format: e.target.value })} style={{ width: 56 }} /></label>
              <label>額外<input value={itemFloorRange.extra} onChange={(e) => setItemFloorRange({ ...itemFloorRange, extra: e.target.value })} placeholder="RF, B1" style={{ width: 80 }} /></label>
            </div>
          </details>

          <input className="wel-scan-input" placeholder="備註（選填）" value={newItem.note} onChange={(e) => setNewItem({ ...newItem, note: e.target.value })} style={{ width: "100%" }} />

          <button type="button" className="wel-btn-primary" style={{ alignSelf: "flex-start" }} onClick={handleAddNewItem}>
            <Plus size={14} /> 新增項目
          </button>
        </div>
      </div>

      {/* Subcontractor attendance */}
      <div className="wel-floor-block">
        <div className="wel-floor-block-title"><UserCheck size={13} /> 發包廠商出工紀錄</div>
        <div className="wel-attendance-row" style={disabled}>
          <input type="date" value={attend.date} onChange={(e) => setAttend({ ...attend, date: e.target.value })} />
          <label>人數<input type="number" min="1" step="1" style={{ width: 56 }} value={attend.headcount} onChange={(e) => setAttend({ ...attend, headcount: e.target.value })} /></label>
          <input placeholder="備註，例：3樓配管進場" value={attend.note} onChange={(e) => setAttend({ ...attend, note: e.target.value })} />
          <button
            type="button" className="wel-btn-ghost"
            onClick={() => { addAttendance(contract.id, attend.date, attend.headcount, attend.note); setAttend({ date: todayStr(), headcount: 1, note: "" }); }}
          >
            <Plus size={13} /> 登記出工
          </button>
        </div>
        {contractWorkLogs.length === 0 ? (
          <div className="muted" style={{ fontSize: 12 }}>尚無出工紀錄</div>
        ) : (
          attendanceByMonth(contractWorkLogs).map((mg) => {
            const mOpen = !!openAttendMonths[mg.month];
            return (
              <div key={mg.month} className="wel-item-floor-group">
                <button
                  type="button" className="wel-item-group-label wel-item-group-toggle"
                  onClick={() => setOpenAttendMonths((prev) => ({ ...prev, [mg.month]: !prev[mg.month] }))}
                >
                  {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                  {mg.month} · {mg.items.length} 天
                </button>
                {mOpen && (
                  <div className="wel-attendance-list">
                    {mg.items.map((l) => (
                      editingAttendId === l.id ? (
                        <div key={l.id} className="wel-attendance-item">
                          <input type="date" value={attendDraft.date} onChange={(e) => setAttendDraft({ ...attendDraft, date: e.target.value })} style={{ width: 130 }} />
                          <input type="number" min="1" step="1" value={attendDraft.headcount} onChange={(e) => setAttendDraft({ ...attendDraft, headcount: e.target.value })} style={{ width: 60 }} />
                          <input placeholder="備註" value={attendDraft.note} onChange={(e) => setAttendDraft({ ...attendDraft, note: e.target.value })} style={{ flex: 1 }} />
                          <button className="wel-icon-btn" onClick={() => saveEditAttend(l.id)}><CheckCircle2 size={13} /></button>
                          <button className="wel-icon-btn" onClick={cancelEditAttend}><X size={13} /></button>
                        </div>
                      ) : (
                        <div key={l.id} className="wel-attendance-item wel-row-clickable" onClick={() => startEditAttend(l)}>
                          <span className="mono">{l.date}</span>
                          <span className="mono muted">{l.headcount} 人</span>
                          <span className="muted" style={{ flex: 1 }}>{l.note || "—"}</span>
                          <button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteAttendance(l.id); }}><X size={13} /></button>
                        </div>
                      )
                    ))}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

function ItemNameHeader({ name, onRename }) {
  const [draft, setDraft] = useState(name);
  useEffect(() => setDraft(name), [name]);
  const commit = () => {
    if (draft.trim() && draft.trim() !== name) onRename(draft.trim());
    else setDraft(name);
  };
  return (
    <input
      className="wel-item-name-input"
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => { if (e.key === "Enter") e.target.blur(); }}
      title="修改名稱並移出焦點，會同步更新這個項目在所有樓層的紀錄"
    />
  );
}

function ClientPaymentsTab({
  clientPayments, siteById, currentSite, clientPaymentForm, setClientPaymentForm, submitClientPayment, deleteClientPayment, updateClientPayment,
  clientPaidTotal, projectGrandTotal, profitLoss, isAllSites, currentSiteId,
  templateRows, floorItems, clientFloorTotal,
  addClientTemplateRow, addClientFloorItem, addClientFloorItemsBatch, updateClientFloorItem, deleteClientFloorItem,
  deleteClientFloorItemsByFloor, deleteClientFloorItemsByName,
  addClientPaymentFromItem, updateSiteClientPricing,
  sites, setCurrentSiteId, showToast,
}) {
  const pct = projectGrandTotal > 0 ? Math.min(100, (clientPaidTotal / projectGrandTotal) * 100) : 0;
  const [expandedPaymentId, setExpandedPaymentId] = useState(null);
  const [paymentDraft, setPaymentDraft] = useState(null);
  const startEditPayment = (p) => { setExpandedPaymentId(p.id); setPaymentDraft({ date: p.date, item: p.item || "", amount: p.amount, note: p.note || "" }); };
  const cancelEditPayment = () => { setExpandedPaymentId(null); setPaymentDraft(null); };
  const saveEditPayment = (id) => { updateClientPayment(id, { ...paymentDraft, amount: Number(paymentDraft.amount) || 0 }); setExpandedPaymentId(null); setPaymentDraft(null); };
  const sortedClientPayments = useMemo(() => clientPayments.slice().sort((a, b) => (a.date < b.date ? 1 : -1)), [clientPayments]);
  const handleExportExcel = () => exportExcel("甲方收款", [
    { name: "收款紀錄", rows: clientPayments.map((p) => ({ 日期: p.date, 請款項目: p.item || "", 金額: Number(p.amount) || 0, 備註: p.note || "" })) },
    { name: "樓層請款項目", rows: floorItems.map((f) => ({ 日期: f.date, 項目: f.itemName, 樓層: f.floor || "", 百分比: f.percent === undefined ? "" : f.percent, 金額: Number(f.amount) || 0 })) },
  ]);
  const handleExportWord = () => exportWord("甲方收款", "甲方清款紀錄", [
    { title: "收款摘要", headers: ["項目", "金額"], rows: [["甲方累計收款", fmtMoney(clientPaidTotal)], ["專案總支出", fmtMoney(projectGrandTotal)], [profitLoss >= 0 ? "毛利" : "虧損", fmtMoney(profitLoss)]] },
    { title: "收款紀錄", headers: ["日期", "請款項目", "金額", "備註"], rows: clientPayments.map((p) => [p.date, p.item || "", fmtMoney(p.amount), p.note || ""]) },
    { title: "樓層請款項目", headers: ["日期", "項目", "樓層", "百分比", "金額"], rows: floorItems.map((f) => [f.date, f.itemName, f.floor || "—", f.percent === undefined ? "—" : `${f.percent}%`, fmtMoney(f.amount)]) },
  ]);
  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">甲方 · CLIENT COLLECTIONS</div>
          <h1 className="wel-h1">甲方清款紀錄</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增甲方收款紀錄或編輯樓層請款項目" />}

      {/* Collections vs total spend */}
      <div className="wel-meter-card">
        <div className="wel-meter-label">
          <Wallet size={14} /> 甲方累計收款 vs 專案總支出
        </div>
        <div className="wel-meter-sub" style={{ marginTop: 0, marginBottom: 12 }}>
          <span>甲方累計收款：<b style={{ color: "var(--teal)" }}>{fmtMoney(clientPaidTotal)}</b></span>
          <span>專案總支出：<b>{fmtMoney(projectGrandTotal)}</b></span>
          <span>{profitLoss >= 0 ? "毛利" : "虧損"}：<b style={{ color: profitLoss >= 0 ? "var(--green)" : "var(--red)" }}>{profitLoss >= 0 ? "+" : ""}{fmtMoney(profitLoss)}</b></span>
        </div>
        <div className="wel-pipe-track" style={{ height: 16 }}>
          <div className="wel-pipe-fill" style={{ width: `${Math.max(pct, clientPaidTotal > 0 ? 3 : 0)}%`, background: profitLoss >= 0 ? "var(--green)" : "var(--red)" }} />
        </div>
        <div className="wel-floor-summary-line muted" style={{ marginTop: 6 }}>
          收款進度：{pct.toFixed(1)}%（相對於總支出）
        </div>
      </div>

      {/* New collection entry */}
      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Wallet size={14} /> 新增甲方收款</div>
        <div className="wel-form-grid">
          <Field label="收款日期">
            <input type="date" value={clientPaymentForm.date} onChange={(e) => setClientPaymentForm({ ...clientPaymentForm, date: e.target.value })} />
          </Field>
          <Field label="請款項目 / 期別" wide>
            <input placeholder="例：第一期進度款，或從下方項目點選" value={clientPaymentForm.item} onChange={(e) => setClientPaymentForm({ ...clientPaymentForm, item: e.target.value })} />
            {templateRows.filter((t) => t.name.trim()).length > 0 && (
              <div className="wel-item-chips">
                <span className="muted" style={{ fontSize: 11, fontFamily: "var(--font-mono)" }}>快速選擇：</span>
                {templateRows.filter((t) => t.name.trim()).map((t) => (
                  <button
                    key={t.id} type="button" className="wel-chip-btn"
                    onClick={() => setClientPaymentForm({ ...clientPaymentForm, item: t.name, amount: clientPaymentForm.amount || t.amount })}
                  >
                    {t.name}{t.amount !== "" && <span className="muted"> · {fmtMoney(t.amount)}</span>}
                  </button>
                ))}
              </div>
            )}
          </Field>
          <Field label="金額">
            <input type="number" min="0" step="any" value={clientPaymentForm.amount} onChange={(e) => setClientPaymentForm({ ...clientPaymentForm, amount: e.target.value })} />
          </Field>
          <Field label="備註" wide>
            <input placeholder="選填" value={clientPaymentForm.note} onChange={(e) => setClientPaymentForm({ ...clientPaymentForm, note: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitClientPayment}><Plus size={15} /> 新增收款紀錄</button>
      </div>

      {/* Collections log */}
      <div className="wel-card" style={{ padding: 0, overflow: "hidden" }}>
        <table className="wel-table">
          <thead>
            <tr>{isAllSites && <th>案場</th>}<th>日期</th><th>請款項目</th><th className="right">金額</th><th>備註</th><th></th></tr>
          </thead>
          <tbody>
            {sortedClientPayments.length === 0 && (<tr><td colSpan={isAllSites ? 6 : 5}><Empty text="尚無甲方收款紀錄" /></td></tr>)}
            {sortedClientPayments.map((p) => {
              const isOpen = expandedPaymentId === p.id;
              return (
              <React.Fragment key={p.id}>
              <tr className="wel-row-clickable" onClick={() => (isOpen ? cancelEditPayment() : startEditPayment(p))}>
                {isAllSites && <td className="muted">{siteById[p.siteId]?.name || "—"}</td>}
                <td className="mono">{p.date}</td>
                <td>
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                    {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                    {p.item || "—"}
                  </span>
                </td>
                <td className="right mono strong" style={{ color: "var(--teal)" }}>{fmtMoney(p.amount)}</td>
                <td className="muted">{p.note || "—"}</td>
                <td><button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteClientPayment(p.id); }}><Trash2 size={14} /></button></td>
              </tr>
              {isOpen && paymentDraft && (
                <tr onClick={(e) => e.stopPropagation()}>
                  <td colSpan={isAllSites ? 6 : 5} style={{ padding: 0 }}>
                    <div className="wel-timeline-wrap">
                      <div className="wel-form-grid">
                        <Field label="收款日期">
                          <input type="date" value={paymentDraft.date} onChange={(e) => setPaymentDraft({ ...paymentDraft, date: e.target.value })} />
                        </Field>
                        <Field label="請款項目 / 期別" wide>
                          <input value={paymentDraft.item} onChange={(e) => setPaymentDraft({ ...paymentDraft, item: e.target.value })} />
                        </Field>
                        <Field label="金額">
                          <input type="number" min="0" step="any" value={paymentDraft.amount} onChange={(e) => setPaymentDraft({ ...paymentDraft, amount: e.target.value })} />
                        </Field>
                        <Field label="備註" wide>
                          <input value={paymentDraft.note} onChange={(e) => setPaymentDraft({ ...paymentDraft, note: e.target.value })} />
                        </Field>
                      </div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button type="button" className="wel-btn-primary" onClick={() => saveEditPayment(p.id)}>儲存修改</button>
                        <button type="button" className="wel-btn-ghost" onClick={cancelEditPayment}>取消</button>
                      </div>
                    </div>
                  </td>
                </tr>
              )}
              </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Client-side floor billing items — same design as 發包廠商: item -> floor -> percent */}
      <div className="wel-card">
        <div className="wel-card-title"><Layers size={14} /> 甲方樓層請款項目</div>
        <ClientFloorItemsPanel
          siteId={currentSiteId}
          site={currentSite}
          isAllSites={isAllSites}
          templateRows={templateRows}
          floorItems={floorItems}
          clientFloorTotal={clientFloorTotal}
          addTemplateRow={addClientTemplateRow}
          addFloorItem={addClientFloorItem}
          addFloorItemsBatch={addClientFloorItemsBatch}
          updateFloorItem={updateClientFloorItem}
          deleteFloorItem={deleteClientFloorItem}
          deleteFloorItemsByFloor={deleteClientFloorItemsByFloor}
          deleteFloorItemsByName={deleteClientFloorItemsByName}
          addPaymentFromItem={addClientPaymentFromItem}
          updateSiteClientPricing={updateSiteClientPricing}
          showToast={showToast}
        />
      </div>
    </div>
  );
}

function ClientFloorItemsPanel({
  siteId, site, isAllSites, templateRows, floorItems, clientFloorTotal,
  addTemplateRow, addFloorItem, addFloorItemsBatch, updateFloorItem, deleteFloorItem,
  deleteFloorItemsByFloor, deleteFloorItemsByName,
  addPaymentFromItem, updateSiteClientPricing, showToast,
}) {
  const [newItem, setNewItem] = useState({ name: "", amount: "", percent: "", date: todayStr(), note: "", floor: "" });
  const [itemFloorRange, setItemFloorRange] = useState({ fromN: "", toN: "", format: "{n}F", extra: "" });
  const [addingNewType, setAddingNewType] = useState(false);
  const [newTypeName, setNewTypeName] = useState("");
  const [addingNewFloor, setAddingNewFloor] = useState(false);
  const [newFloorName, setNewFloorName] = useState("");
  const [openFloors, setOpenFloors] = useState({});
  const totalPriceNum = Number(site?.clientTotalPrice) || 0;

  const itemTypeNames = useMemo(() => {
    const seen = new Set();
    const list = [];
    templateRows.forEach((t) => { if (t.name.trim() && !seen.has(t.name)) { seen.add(t.name); list.push(t.name); } });
    floorItems.forEach((f) => { if (f.itemName.trim() && !seen.has(f.itemName)) { seen.add(f.itemName); list.push(f.itemName); } });
    return list.sort((a, b) => a.localeCompare(b, "zh-Hant"));
  }, [templateRows, floorItems]);
  const knownFloors = useMemo(() => {
    const seen = new Set();
    const list = [];
    floorItems.forEach((f) => { if (f.floor && !seen.has(f.floor)) { seen.add(f.floor); list.push(f.floor); } });
    return list.sort(naturalFloorSort);
  }, [floorItems]);

  const onItemPercentChange = (id, val) => {
    const patch = { percent: val };
    if (totalPriceNum > 0 && val !== "") patch.amount = Math.round(totalPriceNum * Number(val) / 100);
    updateFloorItem(id, patch);
  };
  const onItemAmountChange = (id, val) => {
    const patch = { amount: val };
    if (totalPriceNum > 0 && val !== "") patch.percent = Number((Number(val) / totalPriceNum * 100).toFixed(1));
    updateFloorItem(id, patch);
  };
  const onNewPercentChange = (val) => {
    setNewItem((f) => ({ ...f, percent: val, amount: totalPriceNum > 0 && val !== "" ? Math.round(totalPriceNum * Number(val) / 100) : f.amount }));
  };
  const onNewAmountChange = (val) => {
    setNewItem((f) => ({ ...f, amount: val, percent: totalPriceNum > 0 && val !== "" ? Number((Number(val) / totalPriceNum * 100).toFixed(1)) : f.percent }));
  };
  const handleTypeSelect = (val) => {
    if (val === "__new__") { setAddingNewType(true); return; }
    setNewItem((f) => ({ ...f, name: val }));
  };
  const confirmNewType = () => {
    const nm = newTypeName.trim();
    if (!nm) { showToast("請輸入新項目類型名稱"); return; }
    addTemplateRow(siteId, nm, "");
    setNewItem((f) => ({ ...f, name: nm }));
    setNewTypeName("");
    setAddingNewType(false);
  };
  const handleFloorSelect = (val) => {
    if (val === "__new__") { setAddingNewFloor(true); return; }
    setNewItem((f) => ({ ...f, floor: val }));
  };
  const confirmNewFloor = () => {
    const nm = newFloorName.trim();
    if (!nm) { showToast("請輸入新樓層名稱"); return; }
    setNewItem((f) => ({ ...f, floor: nm }));
    setNewFloorName("");
    setAddingNewFloor(false);
  };
  const toggleFloorOpen = (name, floor) => {
    const key = name + "|" + floor;
    setOpenFloors((prev) => ({ ...prev, [key]: !prev[key] }));
  };
  const buildItemFloorList = () => {
    const list = [];
    const from = parseInt(itemFloorRange.fromN, 10);
    const to = parseInt(itemFloorRange.toN, 10);
    if (!isNaN(from) && !isNaN(to) && itemFloorRange.format.includes("{n}")) {
      const step = to >= from ? 1 : -1;
      for (let n = from; step > 0 ? n <= to : n >= to; n += step) list.push(itemFloorRange.format.replace("{n}", n));
    }
    itemFloorRange.extra.split(",").map((s) => s.trim()).filter(Boolean).forEach((f) => { if (!list.includes(f)) list.push(f); });
    return list;
  };
  const handleAddNewItem = () => {
    if (!newItem.name.trim()) { showToast("請先從下拉選單選擇項目，或新增一個項目類型"); return; }
    const rangeFloors = buildItemFloorList();
    if (rangeFloors.length > 0) {
      addFloorItemsBatch(siteId, newItem.name, newItem.amount, newItem.date, newItem.note, newItem.percent, rangeFloors);
    } else {
      addFloorItem(siteId, newItem.name, newItem.amount, newItem.date, newItem.note, newItem.percent, newItem.floor || "");
      if (newItem.floor) setOpenFloors((prev) => ({ ...prev, [newItem.name.trim() + "|" + newItem.floor]: true }));
    }
    setNewItem({ name: "", amount: "", percent: "", date: todayStr(), note: "", floor: "" });
    setItemFloorRange({ fromN: "", toN: "", format: "{n}F", extra: "" });
  };
  const importAsPayment = (it) => {
    addPaymentFromItem(siteId, it.date, it.amount, `${it.itemName}${it.floor ? `（${it.floor}）` : ""}`, it.id);
  };

  const renameItemGroup = (oldName, newName) => {
    const nm = newName.trim();
    if (!nm || nm === oldName) return;
    floorItems.filter((f) => f.itemName === oldName).forEach((f) => updateFloorItem(f.id, { itemName: nm }));
  };

  const itemGroups = useMemo(() => {
    const byName = {};
    floorItems.forEach((f) => {
      const key = f.itemName || "（未命名項目）";
      if (!byName[key]) byName[key] = [];
      byName[key].push(f);
    });
    const sortByDate = (arr) => arr.slice().sort((a, b) => (a.date < b.date ? 1 : -1));
    return Object.entries(byName)
      .map(([name, items]) => {
        const general = items.filter((it) => !it.floor);
        const byFloorMap = {};
        items.filter((it) => it.floor).forEach((it) => {
          if (!byFloorMap[it.floor]) byFloorMap[it.floor] = [];
          byFloorMap[it.floor].push(it);
        });
        const floorGroups = [];
        if (general.length) floorGroups.push({ floor: "", items: sortByDate(general) });
        Object.entries(byFloorMap)
          .sort((a, b) => naturalFloorSort(b[0], a[0]))
          .forEach(([floor, arr]) => floorGroups.push({ floor, items: sortByDate(arr) }));
        return { name, subtotal: items.reduce((s, it) => s + (Number(it.amount) || 0), 0), floorGroups };
      })
      .sort((a, b) => b.subtotal - a.subtotal);
  }, [floorItems]);

  const disabled = isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined;
  const usedPct = totalPriceNum > 0 ? Math.min(100, (clientFloorTotal / totalPriceNum) * 100) : 0;

  return (
    <div className="wel-floor-panel" style={{ margin: 0 }}>
      {/* Total price basis + progress, linked to every item added below */}
      <div className="wel-floor-block">
        <div className="wel-floor-block-title"><Gauge size={13} /> 甲方總工程價格（百分比計算基準）</div>
        <div className="wel-new-item-row">
          <label className="wel-mini-label">每戶價格
            <input
              type="number" min="0" step="any" placeholder="選填" style={{ width: 110 }}
              className="wel-no-spinner"
              value={site?.clientUnitPrice ?? ""}
              onChange={(e) => updateSiteClientPricing(siteId, { clientUnitPrice: e.target.value })}
              disabled={isAllSites}
            />
          </label>
          <span className="muted" style={{ alignSelf: "center", marginTop: 14 }}>×</span>
          <label className="wel-mini-label">戶數
            <input
              type="number" min="0" step="1" placeholder="選填" style={{ width: 80 }}
              className="wel-no-spinner"
              value={site?.clientUnitCount ?? ""}
              onChange={(e) => updateSiteClientPricing(siteId, { clientUnitCount: e.target.value })}
              disabled={isAllSites}
            />
          </label>
          <span className="muted" style={{ alignSelf: "center", marginTop: 14 }}>＝</span>
          <label className="wel-mini-label">總工程價格
            <input
              type="number" min="0" step="any" placeholder="或直接輸入總價" style={{ width: 140 }}
              value={site?.clientTotalPrice ?? ""}
              onChange={(e) => updateSiteClientPricing(siteId, { clientTotalPrice: e.target.value })}
              disabled={isAllSites}
            />
          </label>
        </div>
        <div className="muted" style={{ fontSize: 11, marginTop: 6 }}>可以直接輸入「總工程價格」，或用「每戶價格 × 戶數」自動算出來（填了會覆蓋總工程價格）。</div>
        {totalPriceNum > 0 && (
          <>
            <div className="wel-pipe-track" style={{ height: 14, marginTop: 8 }}>
              <div className="wel-pipe-fill" style={{ width: `${Math.max(usedPct, clientFloorTotal > 0 ? 3 : 0)}%`, background: usedPct >= 100 ? "var(--green)" : "var(--amber)" }} />
            </div>
            <div className="wel-floor-summary-line muted" style={{ marginTop: 6 }}>
              已請款項目合計：<span className="mono">{fmtMoney(clientFloorTotal)}</span>（約 {usedPct.toFixed(1)}% 總工程價格）
            </div>
          </>
        )}
      </div>

      {/* Existing items, grouped by item type, then by floor (click a floor to open it) */}
      <div className="wel-floor-block">
        <div className="wel-floor-block-title"><ListOrdered size={13} /> 樓層請款項目明細（依項目分組，樓層點開查看；可直接匯入甲方收款）</div>
        {totalPriceNum === 0 && <div className="muted" style={{ fontSize: 11.5, marginBottom: 4 }}>提示：先在上方填寫「甲方總工程價格」，才能用百分比自動換算金額。</div>}
        {floorItems.length === 0 && <div className="muted" style={{ fontSize: 12 }}>尚未新增項目明細</div>}
        {itemGroups.map((g) => (
          <div key={g.name} className="wel-item-name-group">
            <div className="wel-item-name-head">
              <ListOrdered size={13} color="var(--amber)" />
              <ItemNameHeader name={g.name} onRename={(newName) => renameItemGroup(g.name, newName)} />
              <span className="mono muted" style={{ fontSize: 12 }}>{fmtMoney(g.subtotal)}</span>
              <button className="wel-icon-btn" title="刪除這個項目類型的所有樓層紀錄" onClick={() => deleteFloorItemsByName(siteId, g.name)}><Trash2 size={13} /></button>
            </div>
            {g.floorGroups.map((fg) => {
              const key = g.name + "|" + fg.floor;
              const isOpen = !fg.floor || !!openFloors[key];
              const subtotal = fg.items.reduce((s, it) => s + (Number(it.amount) || 0), 0);
              return (
                <div key={fg.floor || "__general__"} className="wel-item-floor-group">
                  {fg.floor && (
                    <div className="wel-item-group-toggle-row">
                      <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => toggleFloorOpen(g.name, fg.floor)}>
                        {isOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                        {fg.floor}
                        <span className="muted" style={{ marginLeft: 4 }}>{fmtMoney(subtotal)}</span>
                      </button>
                      <button className="wel-icon-btn" title="刪除這個樓層的所有紀錄" onClick={() => deleteFloorItemsByFloor(siteId, g.name, fg.floor)}><Trash2 size={12} /></button>
                    </div>
                  )}
                  {isOpen && (
                    <div className="wel-template-rows">
                      {fg.items.map((it) => (
                        <div key={it.id} className="wel-template-row wel-template-row-item" style={disabled}>
                          <input type="date" value={it.date || ""} onChange={(e) => updateFloorItem(it.id, { date: e.target.value })} />
                          <input placeholder="樓層（留空＝一般項目）" value={it.floor || ""} onChange={(e) => updateFloorItem(it.id, { floor: e.target.value })} />
                          <input type="number" min="0" step="any" placeholder="%" title="佔總工程價格百分比" value={it.percent === undefined ? "" : it.percent} onChange={(e) => onItemPercentChange(it.id, e.target.value)} />
                          <input type="number" min="0" step="any" placeholder="金額" value={it.amount} onChange={(e) => onItemAmountChange(it.id, e.target.value)} />
                          <NoteInput value={it.note} onCommit={(v) => updateFloorItem(it.id, { note: v })} />
                          <button
                            type="button"
                            className={"wel-pay-import-btn" + (it.paidImported ? " imported" : "")}
                            title={it.paidImported ? "已匯入收款，仍可再次匯入" : "點擊直接匯入到甲方收款"}
                            onClick={() => importAsPayment(it)}
                          >
                            {it.paidImported ? <CheckCircle2 size={13} /> : <Banknote size={13} />}
                          </button>
                          <button className="wel-icon-btn" onClick={() => deleteFloorItem(it.id)}><X size={14} /></button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ))}

        {/* New item — wheel-style pick for 樓層／項目, plain custom % input */}
        <div className="wel-new-item-block" style={disabled}>
          <div className="wel-new-item-row">
            <label className="wel-mini-label">日期
              <input type="date" value={newItem.date} onChange={(e) => setNewItem({ ...newItem, date: e.target.value })} />
            </label>
            <label className="wel-mini-label" style={{ flex: 1, minWidth: 140 }}>項目
              <select value={itemTypeNames.includes(newItem.name) ? newItem.name : ""} onChange={(e) => handleTypeSelect(e.target.value)}>
                <option value="" disabled>{itemTypeNames.length ? "請選擇項目…" : "尚無項目類型"}</option>
                {itemTypeNames.map((n) => <option key={n} value={n}>{n}</option>)}
                <option value="__new__">＋ 新增項目類型…</option>
              </select>
            </label>
            <label className="wel-mini-label">樓層
              <select value={newItem.floor} onChange={(e) => handleFloorSelect(e.target.value)}>
                <option value="">（一般項目）</option>
                {knownFloors.map((f) => <option key={f} value={f}>{f}</option>)}
                <option value="__new__">＋ 新增樓層…</option>
              </select>
            </label>
            <label className="wel-mini-label">百分比
              <input type="number" min="0" step="any" placeholder="自訂 %" value={newItem.percent} onChange={(e) => onNewPercentChange(e.target.value)} style={{ width: 72 }} />
            </label>
            <label className="wel-mini-label">金額
              <input type="number" min="0" step="any" value={newItem.amount} onChange={(e) => onNewAmountChange(e.target.value)} style={{ width: 90 }} />
            </label>
          </div>

          {addingNewType && (
            <div className="wel-new-type-row">
              <input placeholder="新項目類型名稱，例：水電材料進場款" value={newTypeName} onChange={(e) => setNewTypeName(e.target.value)} autoFocus />
              <button type="button" className="wel-btn-ghost" onClick={confirmNewType}>確定新增類型</button>
              <button type="button" className="wel-btn-ghost" onClick={() => { setAddingNewType(false); setNewTypeName(""); }}>取消</button>
            </div>
          )}
          {addingNewFloor && (
            <div className="wel-new-type-row">
              <input placeholder="新樓層名稱，例：3F 或 RF" value={newFloorName} onChange={(e) => setNewFloorName(e.target.value)} autoFocus />
              <button type="button" className="wel-btn-ghost" onClick={confirmNewFloor}>確定新增樓層</button>
              <button type="button" className="wel-btn-ghost" onClick={() => { setAddingNewFloor(false); setNewFloorName(""); }}>取消</button>
            </div>
          )}

          <details className="wel-floor-batch-advanced">
            <summary>進階：一次套用到多個樓層（可選）</summary>
            <div className="wel-item-floor-apply">
              <label>從第<input type="number" value={itemFloorRange.fromN} onChange={(e) => setItemFloorRange({ ...itemFloorRange, fromN: e.target.value })} placeholder="3" /></label>
              <label>到第<input type="number" value={itemFloorRange.toN} onChange={(e) => setItemFloorRange({ ...itemFloorRange, toN: e.target.value })} placeholder="15" /></label>
              <label>格式<input value={itemFloorRange.format} onChange={(e) => setItemFloorRange({ ...itemFloorRange, format: e.target.value })} style={{ width: 56 }} /></label>
              <label>額外<input value={itemFloorRange.extra} onChange={(e) => setItemFloorRange({ ...itemFloorRange, extra: e.target.value })} placeholder="RF, B1" style={{ width: 80 }} /></label>
            </div>
          </details>

          <input className="wel-scan-input" placeholder="備註（選填）" value={newItem.note} onChange={(e) => setNewItem({ ...newItem, note: e.target.value })} style={{ width: "100%" }} />

          <button type="button" className="wel-btn-primary" style={{ alignSelf: "flex-start" }} onClick={handleAddNewItem}>
            <Plus size={14} /> 新增項目
          </button>
        </div>
      </div>
    </div>
  );
}

function MasterSummaryTab({
  currentSiteName, isAllSites, clientPaidTotal, clientPayments, clientFloorItems,
  totalCost, laborTotalAllTime, categoryBreakdown, laborSummaryByWorker, workerById,
  contractRows, floorItems, projectGrandTotal, profitLoss,
  orders, usages, catById, contractPayments, contractItems, discussionItems, contractorsTotalPrice,
}) {
  const clientFloorByFloor = useMemo(() => {
    const map = {};
    clientFloorItems.forEach((f) => {
      if (!map[f.floor]) map[f.floor] = [];
      map[f.floor].push(f);
    });
    return Object.entries(map)
      .map(([floor, items]) => ({ floor, items, subtotal: items.reduce((s, i) => s + (Number(i.amount) || 0), 0) }))
      .sort((a, b) => naturalFloorSort(b.floor, a.floor));
  }, [clientFloorItems]);
  const clientFloorTotal = clientFloorByFloor.reduce((s, g) => s + g.subtotal, 0);

  const meCostTotal = totalCost + laborTotalAllTime;
  const contractorsTotal = contractRows.reduce((s, c) => s + c.paid, 0);

  const contractFloorByContract = useMemo(() => {
    const map = {};
    floorItems.forEach((f) => {
      if (!map[f.contractId]) map[f.contractId] = [];
      map[f.contractId].push(f);
    });
    return map;
  }, [floorItems]);

  const handleExportExcel = () => exportExcel(`項目總表${currentSiteName ? "_" + currentSiteName : ""}`, [
    { name: "甲方收款", rows: clientPayments.map((p) => ({ 日期: p.date, 請款項目: p.item || "", 金額: Number(p.amount) || 0 })) },
    { name: "甲方樓層請款", rows: clientFloorItems.map((f) => ({ 樓層: f.floor, 項目: f.itemName, 金額: Number(f.amount) || 0 })) },
    { name: "材料叫貨", rows: orders.map((o) => ({ 日期: o.date, 類別: catById[o.categoryId]?.name || "", 品項: o.itemName, 數量: o.quantity, 單位: o.unit || "", 單價: Number(o.unitPrice) || 0, 金額: Number(o.amount) || 0 })) },
    { name: "領用紀錄", rows: usages.map((u) => ({ 日期: u.date, 品項: u.itemName, 數量: u.quantity, 單位: u.unit || "" })) },
    { name: "人工薪資", rows: laborSummaryByWorker.filter((r) => r.wage > 0).map((r) => ({ 師傅: workerById[r.workerId]?.name || "", 出工天數: r.days, 應付薪資: r.wage })) },
    { name: "發包項目", rows: contractRows.map((c) => ({ 廠商: c.contractor || "", 工程總稱: c.name, 發包總價: Number(c.totalPrice) || 0, 已領款: c.paid })) },
    { name: "發包請款明細", rows: contractItems.map((it) => ({ 日期: it.date, 項目: it.name, 樓層: it.floor || "", 百分比: it.percent === undefined ? "" : it.percent, 金額: Number(it.amount) || 0 })) },
    { name: "發包領款", rows: contractPayments.map((p) => ({ 日期: p.date, 金額: Number(p.amount) || 0, 備註: p.note || "" })) },
    { name: "備註討論", rows: discussionItems.map((d) => ({ 日期: d.date, 項目: d.topic, 是否有結果: d.resolved ? "已有結果" : "尚無結果", 結果: d.result || "" })) },
  ]);
  const handleExportWord = () => exportWord(`項目總表${currentSiteName ? "_" + currentSiteName : ""}`, `項目總表${currentSiteName ? " · " + currentSiteName : ""}`, [
    { title: "1. 甲方（業主收款與請款項目）", note: `合計：${fmtMoney(clientPaidTotal)}`, headers: ["日期", "請款項目 / 樓層", "金額"], rows: [
      ...clientPayments.map((p) => [p.date, p.item || "", fmtMoney(p.amount)]),
      ...clientFloorItems.map((f) => ["—", `${f.floor} · ${f.itemName}`, fmtMoney(f.amount)]),
    ] },
    { title: "2. 我（材料採購）", note: `合計：${fmtMoney(totalCost)}`, headers: ["類別", "金額"], rows: categoryBreakdown.map((c) => [c.name, fmtMoney(c.value)]) },
    { title: "2. 我（人工薪資）", note: `合計：${fmtMoney(laborTotalAllTime)}`, headers: ["師傅", "出工天數", "薪資"], rows: laborSummaryByWorker.filter((r) => r.wage > 0).map((r) => [workerById[r.workerId]?.name || "", fmtNum(r.days), fmtMoney(r.wage)]) },
    { title: "3. 發包廠商", note: `已領款合計：${fmtMoney(contractorsTotal)}`, headers: ["廠商", "工程總稱", "已領款", "發包總價"], rows: contractRows.map((c) => [c.contractor || "", c.name, fmtMoney(c.paid), fmtMoney(c.totalPrice)]) },
    { title: "備註討論項目", headers: ["日期", "項目", "是否有結果"], rows: discussionItems.map((d) => [d.date, d.topic, d.resolved ? "已有結果" : "尚無結果"]) },
    { title: "總結", headers: ["項目", "金額"], rows: [["甲方收款", fmtMoney(clientPaidTotal)], ["專案總支出", fmtMoney(projectGrandTotal)], [profitLoss >= 0 ? "毛利" : "虧損", fmtMoney(profitLoss)]] },
  ]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">總表 · MASTER ITEM LIST{isAllSites ? " · 全部案場" : ""}</div>
          <h1 className="wel-h1">項目總表{!isAllSites && currentSiteName ? ` · ${currentSiteName}` : ""}</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} label="匯出全部：" />
      </div>
      <div className="wel-summary-note muted">依序排列：甲方 → 我 → 發包廠商</div>

      {/* 1. 甲方 */}
      <div className="wel-summary-section">
        <div className="wel-summary-heading">
          <span className="wel-summary-index">1</span>
          <Wallet size={15} color="var(--teal)" />
          <span>甲方（業主收款與請款項目）</span>
          <span className="wel-summary-total mono">{fmtMoney(clientPaidTotal)}</span>
        </div>
        <div className="wel-card">
          <div className="wel-card-title">收款紀錄</div>
          {clientPayments.length === 0 ? <Empty text="尚無甲方收款紀錄" /> : (
            <div className="wel-summary-rows">
              {clientPayments.map((p) => (
                <div key={p.id} className="wel-summary-row">
                  <span className="mono muted">{p.date}</span>
                  <span style={{ flex: 1 }}>{p.item || "—"}</span>
                  <span className="mono" style={{ color: "var(--teal)" }}>{fmtMoney(p.amount)}</span>
                </div>
              ))}
            </div>
          )}
          {clientFloorByFloor.length > 0 && (
            <>
              <div className="wel-card-title" style={{ marginTop: 16 }}>樓層請款項目（合計 {fmtMoney(clientFloorTotal)}）</div>
              <div className="wel-summary-rows">
                {clientFloorByFloor.map((g) => (
                  <div key={g.floor} className="wel-summary-row">
                    <span className="mono muted">{g.floor}</span>
                    <span style={{ flex: 1 }}>{g.items.map((i) => i.itemName).join("、")}</span>
                    <span className="mono">{fmtMoney(g.subtotal)}</span>
                  </div>
                ))}
              </div>
            </>
          )}
        </div>
      </div>

      {/* 2. 我 */}
      <div className="wel-summary-section">
        <div className="wel-summary-heading">
          <span className="wel-summary-index">2</span>
          <Gauge size={15} color="var(--amber)" />
          <span>我（材料採購與人工薪資）</span>
          <span className="wel-summary-total mono">{fmtMoney(meCostTotal)}</span>
        </div>
        <div className="wel-card">
          <div className="wel-card-title">材料支出（依品項類別，合計 {fmtMoney(totalCost)}）</div>
          {categoryBreakdown.length === 0 ? <Empty text="尚無叫貨紀錄" /> : (
            <div className="wel-summary-rows">
              {categoryBreakdown.map((c) => (
                <div key={c.id} className="wel-summary-row">
                  <span className="wel-dot" style={{ background: c.color }} />
                  <span style={{ flex: 1 }}>{c.name}</span>
                  <span className="mono">{fmtMoney(c.value)}</span>
                </div>
              ))}
            </div>
          )}
          <div className="wel-card-title" style={{ marginTop: 16 }}>人工薪資（依師傅，合計 {fmtMoney(laborTotalAllTime)}）</div>
          {laborSummaryByWorker.filter((r) => r.wage > 0).length === 0 ? <Empty text="尚無出工紀錄" /> : (
            <div className="wel-summary-rows">
              {laborSummaryByWorker.filter((r) => r.wage > 0).map((r) => (
                <div key={r.workerId} className="wel-summary-row">
                  <span style={{ flex: 1 }}>{workerById[r.workerId]?.name || "已刪除師傅"}</span>
                  <span className="mono muted">{fmtNum(r.days)} 天</span>
                  <span className="mono">{fmtMoney(r.wage)}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {/* 3. 發包廠商 */}
      <div className="wel-summary-section">
        <div className="wel-summary-heading">
          <span className="wel-summary-index">3</span>
          <Handshake size={15} color="var(--red)" />
          <span>發包廠商（各發包項目已付款）</span>
          <span className="wel-summary-total mono">{fmtMoney(contractorsTotal)}</span>
        </div>
        {contractRows.length === 0 ? (
          <div className="wel-card"><Empty text="尚未建立發包項目" /></div>
        ) : (
          contractRows.map((c) => {
            const cFloors = contractFloorByContract[c.id] || [];
            return (
              <div key={c.id} className="wel-card" style={{ marginBottom: 10 }}>
                <div className="wel-card-title">
                  {c.name}{c.contractor ? ` · ${c.contractor}` : ""}
                  <span style={{ marginLeft: "auto", color: "var(--amber)" }}>{fmtMoney(c.paid)} / {fmtMoney(c.totalPrice)}</span>
                </div>
                {c.attendanceDays > 0 && (
                  <div className="wel-floor-summary-line muted">出工 {c.attendanceDays} 天 · 工天人次 {c.attendanceManDays}</div>
                )}
                {cFloors.length > 0 && (
                  <div className="wel-summary-rows" style={{ marginTop: 8 }}>
                    {cFloors.map((f) => (
                      <div key={f.id} className="wel-summary-row">
                        <span className="mono muted">{f.floor}</span>
                        <span style={{ flex: 1 }}>{f.itemName}</span>
                        <span className="mono">{fmtMoney(f.amount)}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      {/* Grand total */}
      <div className={"wel-profit-card" + (profitLoss >= 0 ? " positive" : " negative")}>
        <div className="wel-profit-col">
          <span className="wel-profit-label">甲方收款</span>
          <b className="mono">{fmtMoney(clientPaidTotal)}</b>
        </div>
        <div className="wel-profit-sign">－</div>
        <div className="wel-profit-col">
          <span className="wel-profit-label">總支出（我＋發包）</span>
          <b className="mono">{fmtMoney(projectGrandTotal)}</b>
        </div>
        <div className="wel-profit-sign">=</div>
        <div className="wel-profit-col">
          <span className="wel-profit-label">{profitLoss >= 0 ? "毛利" : "虧損"}</span>
          <b className="mono wel-profit-value">{profitLoss >= 0 ? "+" : ""}{fmtMoney(profitLoss)}</b>
        </div>
      </div>
    </div>
  );
}

function CategoriesTab({ categories, orders, usages, newCatName, setNewCatName, addCategory, removeCategory, renameCategory, materialItems, addMaterialItem, updateMaterialItem, removeMaterialItem, confirmScannedItems, showToast }) {
  const [scanOpen, setScanOpen] = useState(false);
  const countFor = (id) => orders.filter((o) => o.categoryId === id).length + usages.filter((u) => u.categoryId === id).length;
  const [newMaterial, setNewMaterial] = useState({ name: "", unit: "", defaultPrice: "" });
  const sortedMaterials = useMemo(() => materialItems.slice().sort((a, b) => a.name.localeCompare(b.name, "zh-Hant")), [materialItems]);
  const orderCountFor = (name) => orders.filter((o) => o.itemName === name).length;
  const catIds = useMemo(() => new Set(categories.map((c) => c.id)), [categories]);
  const itemsOf = (id) => sortedMaterials.filter((m) => m.categoryId === id);
  const uncategorizedItems = sortedMaterials.filter((m) => !m.categoryId || !catIds.has(m.categoryId));
  const itemChips = (list) => (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
      {list.map((m) => (
        <span key={m.id} className="wel-chip">
          {m.name}{m.unit ? `（${m.unit}）` : ""}{m.defaultPrice !== "" && m.defaultPrice != null ? ` · ${fmtMoney(m.defaultPrice)}` : ""}
        </span>
      ))}
    </div>
  );
  const [expandedCat, setExpandedCat] = useState(null);
  const [expandedMaterial, setExpandedMaterial] = useState(null);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">設定 · CATEGORIES</div>
          <h1 className="wel-h1">品項類別管理</h1>
        </div>
      </div>
      <div className="wel-card wel-form">
        <div className="wel-form-grid wel-form-grid-btn" style={{ gridTemplateColumns: "1fr auto" }}>
          <Field label="新增類別名稱">
            <input placeholder="例：消防管路" value={newCatName} onChange={(e) => setNewCatName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addCategory())} />
          </Field>
          <button type="button" className="wel-btn-primary" style={{ alignSelf: "end" }} onClick={addCategory}><Plus size={15} /> 新增</button>
        </div>
      </div>
      <div className="wel-cat-list">
        {categories.map((c) => {
          const isOpen = expandedCat === c.id;
          return (
            <div key={c.id}>
              <div className="wel-cat-item">
                <button type="button" className="wel-expand-btn" onClick={() => setExpandedCat(isOpen ? null : c.id)} title="查詢時間軸紀錄">
                  {isOpen ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                </button>
                <span className="wel-dot" style={{ background: c.color }} />
                <ItemNameHeader name={c.name} onRename={(newName) => renameCategory(c.id, newName)} />
                <span className="muted mono" style={{ fontSize: 12 }}>{itemsOf(c.id).length} 項品項 · {countFor(c.id)} 筆紀錄</span>
                <button className="wel-icon-btn" onClick={() => removeCategory(c.id)}><X size={14} /></button>
              </div>
              {isOpen && (
                <div className="wel-timeline-wrap">
                  <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>這個類別底下的小項品項（{itemsOf(c.id).length}）</div>
                  {itemsOf(c.id).length === 0
                    ? <div className="muted" style={{ fontSize: 12, marginBottom: 10 }}>尚無品項，可在下方「品項清單」把品項的類別設成「{c.name}」，或用拍照掃描新增</div>
                    : itemChips(itemsOf(c.id))}
                  <OrderTimeline items={orders.filter((o) => o.categoryId === c.id)} />
                </div>
              )}
            </div>
          );
        })}
        {uncategorizedItems.length > 0 && (
          <div>
            <div className="wel-cat-item">
              <button type="button" className="wel-expand-btn" onClick={() => setExpandedCat(expandedCat === "__none" ? null : "__none")}>
                {expandedCat === "__none" ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
              </button>
              <span className="wel-dot" style={{ background: "var(--text-muted)" }} />
              <span className="wel-cat-name">未歸類</span>
              <span className="muted mono" style={{ fontSize: 12 }}>{uncategorizedItems.length} 項品項</span>
            </div>
            {expandedCat === "__none" && (
              <div className="wel-timeline-wrap">
                <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>還沒指定類別的品項，到下方「品項清單」的類別欄位歸類後，叫貨時就會依類別列出</div>
                {itemChips(uncategorizedItems)}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Material item catalog — same list used by the 叫貨 wheel-picker, fully manageable here */}
      {scanOpen && (
        <div style={{ marginTop: 22 }}>
          <ReceiptScanner
            mode="items" categories={categories} materialItems={materialItems} suppliers={[]}
            defaultCategoryId={categories[0]?.id} onConfirmItems={confirmScannedItems}
            onClose={() => setScanOpen(false)} showToast={showToast}
          />
        </div>
      )}
      <div className="wel-card wel-form" style={{ marginTop: 22 }}>
        <div className="wel-card-title"><Package size={14} /> 品項清單（連動叫貨的品項選單，可重新命名、刪除、查詢時間軸）</div>
        {!scanOpen && (
          <div style={{ marginBottom: 10 }}>
            <button type="button" className="wel-btn-ghost" onClick={() => setScanOpen(true)}><Camera size={14} /> 拍照掃描品項（先校對再加入）</button>
          </div>
        )}
        <div className="wel-material-add-row">
          <input placeholder="新品項名稱" value={newMaterial.name} onChange={(e) => setNewMaterial({ ...newMaterial, name: e.target.value })} />
          <input placeholder="單位" value={newMaterial.unit} onChange={(e) => setNewMaterial({ ...newMaterial, unit: e.target.value })} style={{ width: 90 }} />
          <input type="number" placeholder="預設單價" value={newMaterial.defaultPrice} onChange={(e) => setNewMaterial({ ...newMaterial, defaultPrice: e.target.value })} style={{ width: 100 }} />
          <button
            type="button" className="wel-btn-primary"
            onClick={() => { addMaterialItem(newMaterial.name, newMaterial.unit, newMaterial.defaultPrice); setNewMaterial({ name: "", unit: "", defaultPrice: "" }); }}
          >
            <Plus size={15} /> 新增品項
          </button>
        </div>
        <div className="wel-manager-list" style={{ marginTop: 12, maxHeight: "none" }}>
          {sortedMaterials.length === 0 && <div className="muted" style={{ fontSize: 12 }}>尚未建立任何品項</div>}
          {sortedMaterials.map((m) => {
            const isOpen = expandedMaterial === m.id;
            return (
              <div key={m.id}>
                <div className="wel-manager-row">
                  <button type="button" className="wel-expand-btn" onClick={() => setExpandedMaterial(isOpen ? null : m.id)} title="查詢時間軸紀錄">
                    {isOpen ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
                  </button>
                  <input
                    className="wel-manager-name"
                    value={m.name}
                    onChange={(e) => updateMaterialItem(m.id, { name: e.target.value })}
                    onBlur={(e) => { if (!e.target.value.trim()) updateMaterialItem(m.id, { name: m.name }); }}
                  />
                  <input className="wel-manager-mini" placeholder="單位" value={m.unit || ""} onChange={(e) => updateMaterialItem(m.id, { unit: e.target.value })} />
                  <input className="wel-manager-mini" type="number" placeholder="單價" value={m.defaultPrice === "" || m.defaultPrice == null ? "" : m.defaultPrice} onChange={(e) => updateMaterialItem(m.id, { defaultPrice: e.target.value })} />
                  <select className="wel-manager-mini" style={{ width: 110 }} value={m.categoryId || ""} onChange={(e) => updateMaterialItem(m.id, { categoryId: e.target.value })}>
                    <option value="">未歸類</option>
                    {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
                  <span className="muted mono" style={{ fontSize: 11.5, width: 64, textAlign: "right" }}>{orderCountFor(m.name)} 筆紀錄</span>
                  <button type="button" className="wel-icon-btn" onClick={() => removeMaterialItem(m.id)}><Trash2 size={13} /></button>
                </div>
                {isOpen && (
                  <div className="wel-timeline-wrap">
                    <OrderTimeline items={orders.filter((o) => o.itemName === m.name)} />
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <div className="muted" style={{ fontSize: 11, marginTop: 8 }}>提示：改名稱會自動同步更新過去的叫貨/領用紀錄；設定類別後，之後在叫貨紀錄挑選這個品項會自動帶入該類別（不會更動過去已新增的紀錄）；刪除只會移除選單中的品項，不會刪除過去已經新增的叫貨紀錄。</div>
      </div>
    </div>
  );
}

function OrderTimeline({ items }) {
  const [openMonths, setOpenMonths] = useState({});
  const months = attendanceByMonth(items);
  const total = items.reduce((s, o) => s + (Number(o.amount) || 0), 0);

  if (items.length === 0) {
    return <div className="wel-timeline"><div className="muted" style={{ fontSize: 12, padding: "8px 0" }}>尚無叫貨紀錄</div></div>;
  }

  return (
    <div className="wel-timeline">
      {months.map((mg) => {
        const subtotal = mg.items.reduce((s, o) => s + (Number(o.amount) || 0), 0);
        const mOpen = !!openMonths[mg.month];
        return (
          <div key={mg.month} className="wel-item-floor-group">
            <button
              type="button" className="wel-item-group-label wel-item-group-toggle"
              onClick={() => setOpenMonths((prev) => ({ ...prev, [mg.month]: !prev[mg.month] }))}
            >
              {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
              {mg.month} · {mg.items.length} 筆 · {fmtMoney(subtotal)}
            </button>
            {mOpen && (
              <table className="wel-table">
                <thead>
                  <tr><th>日期</th><th>品項</th><th className="right">數量</th><th className="right">單價</th><th className="right">金額</th><th>廠商</th></tr>
                </thead>
                <tbody>
                  {mg.items.map((o) => (
                    <tr key={o.id}>
                      <td className="mono">{o.date}</td>
                      <td>{o.itemName}</td>
                      <td className="right mono">{fmtNum(o.quantity)} {o.unit}</td>
                      <td className="right mono">{fmtMoney(o.unitPrice)}</td>
                      <td className="right mono strong">{fmtMoney(o.amount)}</td>
                      <td className="muted">{o.supplier || "—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        );
      })}
      <div className="wel-floor-summary-line muted" style={{ marginTop: 6 }}>總合計：<span className="mono">{fmtMoney(total)}</span>（{items.length} 筆）</div>
    </div>
  );
}

function DiscussionTab({
  discussionItems, siteById, discussionForm, setDiscussionForm, submitDiscussion,
  toggleDiscussionResolved, updateDiscussionResult, updateDiscussion, deleteDiscussion, isAllSites, sites, setCurrentSiteId,
  currentSiteId, showToast,
}) {
  const copyShareLink = async () => {
    const url = `${window.location.origin}/tasks?site=${currentSiteId}`;
    try {
      await navigator.clipboard.writeText(url);
      showToast("連結已複製，可以直接分享出去");
    } catch (e) {
      showToast(url);
    }
  };
  const [filter, setFilter] = useState("all"); // all | open | resolved
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const startEdit = (d) => { setEditingId(d.id); setEditDraft({ date: d.date, topic: d.topic, note: d.note || "" }); };
  const cancelEdit = () => { setEditingId(null); setEditDraft(null); };
  const saveEdit = (id) => {
    if (!editDraft.topic.trim()) { showToast("請填寫「討論項目」後再儲存"); return; }
    updateDiscussion(id, editDraft);
    setEditingId(null);
    setEditDraft(null);
  };
  const filtered = discussionItems
    .filter((d) => {
      if (filter === "open") return !d.resolved;
      if (filter === "resolved") return !!d.resolved;
      return true;
    })
    .slice()
    .sort((a, b) => (a.date < b.date ? 1 : -1));
  const openCount = discussionItems.filter((d) => !d.resolved).length;

  const handleExportExcel = () => exportExcel("備註討論項目", [{
    name: "討論項目",
    rows: discussionItems.map((d) => ({ 日期: d.date, 討論項目: d.topic, 備註: d.note || "", 是否有結果: d.resolved ? "已有結果" : "尚無結果", 結果說明: d.result || "" })),
  }]);
  const handleExportWord = () => exportWord("備註討論項目", "備註討論項目", [{
    title: "討論項目一覽",
    headers: ["日期", "討論項目", "是否有結果", "結果說明"],
    rows: discussionItems.map((d) => [d.date, d.topic, d.resolved ? "已有結果" : "尚無結果", d.result || "—"]),
  }]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">備註 · DISCUSSION &amp; NOTES</div>
          <h1 className="wel-h1">備註討論項目</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
        <div className="wel-export-bar">
          <button type="button" className="wel-export-btn" onClick={copyShareLink} disabled={isAllSites}>
            <Copy size={13} /> 分享連結
          </button>
          <div className="wel-filter">
            <select value={filter} onChange={(e) => setFilter(e.target.value)}>
              <option value="all">全部（{discussionItems.length}）</option>
              <option value="open">待處理（{openCount}）</option>
              <option value="resolved">已有結果（{discussionItems.length - openCount}）</option>
            </select>
          </div>
        </div>
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增討論項目" />}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><MessageSquare size={14} /> 新增討論項目</div>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={discussionForm.date} onChange={(e) => setDiscussionForm({ ...discussionForm, date: e.target.value })} />
          </Field>
          <Field label="討論項目" wide>
            <input placeholder="例：3樓管線與空調衝突，需與空調師傅確認" value={discussionForm.topic} onChange={(e) => setDiscussionForm({ ...discussionForm, topic: e.target.value })} />
          </Field>
          <Field label="備註 / 細節" wide>
            <input placeholder="選填" value={discussionForm.note} onChange={(e) => setDiscussionForm({ ...discussionForm, note: e.target.value })} />
          </Field>
          <Field label="是否已有結果">
            <select value={discussionForm.resolved ? "yes" : "no"} onChange={(e) => setDiscussionForm({ ...discussionForm, resolved: e.target.value === "yes" })}>
              <option value="no">尚無結果</option>
              <option value="yes">已有結果</option>
            </select>
          </Field>
          {discussionForm.resolved && (
            <Field label="結果說明" wide>
              <input placeholder="例：改走天花板側邊，已與空調師傅確認" value={discussionForm.result} onChange={(e) => setDiscussionForm({ ...discussionForm, result: e.target.value })} />
            </Field>
          )}
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitDiscussion}><Plus size={15} /> 新增討論項目</button>
      </div>

      <div className="wel-discussion-list">
        {filtered.length === 0 && <div className="wel-card"><Empty text="尚無討論項目" /></div>}
        {filtered.map((d) => (
          <div key={d.id} className={"wel-discussion-card" + (d.resolved ? " resolved" : "")}>
            {editingId === d.id ? (
              <div className="wel-timeline-wrap" style={{ padding: 0 }}>
                <div className="wel-form-grid">
                  <Field label="日期">
                    <input type="date" value={editDraft.date} onChange={(e) => setEditDraft({ ...editDraft, date: e.target.value })} />
                  </Field>
                  <Field label="討論項目" wide>
                    <input value={editDraft.topic} onChange={(e) => setEditDraft({ ...editDraft, topic: e.target.value })} />
                  </Field>
                  <Field label="備註 / 細節" wide>
                    <input value={editDraft.note} onChange={(e) => setEditDraft({ ...editDraft, note: e.target.value })} />
                  </Field>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button type="button" className="wel-btn-primary" onClick={() => saveEdit(d.id)}>儲存修改</button>
                  <button type="button" className="wel-btn-ghost" onClick={cancelEdit}>取消</button>
                </div>
              </div>
            ) : (
              <>
                <div className="wel-discussion-head">
                  <button type="button" className="wel-status-btn" onClick={() => toggleDiscussionResolved(d.id)} title="點擊切換是否有結果">
                    {d.resolved ? <CheckCircle2 size={17} color="var(--green)" /> : <Circle size={17} color="var(--text-muted)" />}
                  </button>
                  <span className="wel-discussion-topic">{d.topic}</span>
                  {isAllSites && <span className="wel-tag">{siteById[d.siteId]?.name || "—"}</span>}
                  <span className="mono muted" style={{ fontSize: 11.5 }}>{d.date}</span>
                  <button className="wel-icon-btn" onClick={() => startEdit(d)}><Pencil size={14} /></button>
                  <button className="wel-icon-btn" onClick={() => deleteDiscussion(d.id)}><Trash2 size={14} /></button>
                </div>
                {d.note && <div className="wel-discussion-note muted">{d.note}</div>}
                <div className="wel-discussion-result">
                  <span className="muted" style={{ fontSize: 11.5, fontFamily: "var(--font-mono)" }}>結果：</span>
                  <input
                    placeholder={d.resolved ? "輸入結果說明…" : "尚無結果（點左側圓圈標記為已有結果）"}
                    value={d.result || ""}
                    disabled={!d.resolved}
                    onChange={(e) => updateDiscussionResult(d.id, e.target.value)}
                  />
                </div>
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

const WORKFLOW_TEMPLATE = ["接案與估價", "簽約與發包", "備料與叫貨", "進場施工", "驗收與請款", "結案與保固"];

// 施工流程：階段 → 步驟，寫給之後接手的人看的經驗手冊（全案場共用）。
function WorkflowTab({ workflow, saveWorkflow, showToast }) {
  const [openPhases, setOpenPhases] = useState({});
  const [editing, setEditing] = useState(null); // { type: "phase" | "step", phaseId, stepId, isNew }
  const [draft, setDraft] = useState(null);
  const isOpen = (id) => openPhases[id] !== false;
  const togglePhase = (id) => setOpenPhases((prev) => ({ ...prev, [id]: !isOpen(id) }));
  const stop = () => { setEditing(null); setDraft(null); };

  const move = (arr, i, d) => {
    const j = i + d;
    if (j < 0 || j >= arr.length) return arr;
    const copy = arr.slice();
    [copy[i], copy[j]] = [copy[j], copy[i]];
    return copy;
  };
  const startNewPhase = () => { setEditing({ type: "phase", phaseId: null, stepId: null, isNew: true }); setDraft({ title: "", summary: "" }); };
  const startNewStep = (phaseId) => {
    setOpenPhases((prev) => ({ ...prev, [phaseId]: true }));
    setEditing({ type: "step", phaseId, stepId: null, isNew: true });
    setDraft({ title: "", details: "", tips: "" });
  };
  const startEditPhase = (p) => { setEditing({ type: "phase", phaseId: p.id, stepId: null, isNew: false }); setDraft({ title: p.title, summary: p.summary || "" }); };
  const startEditStep = (p, st) => { setEditing({ type: "step", phaseId: p.id, stepId: st.id, isNew: false }); setDraft({ title: st.title, details: st.details || "", tips: st.tips || "" }); };

  const save = () => {
    if (!draft.title.trim()) { showToast("請先填寫標題"); return; }
    if (editing.type === "phase") {
      const body = { title: draft.title.trim(), summary: draft.summary.trim() };
      saveWorkflow(editing.isNew
        ? [...workflow, { id: uid(), ...body, steps: [] }]
        : workflow.map((p) => (p.id === editing.phaseId ? { ...p, ...body } : p)));
    } else {
      const body = { title: draft.title.trim(), details: draft.details.trim(), tips: draft.tips.trim() };
      saveWorkflow(workflow.map((p) => {
        if (p.id !== editing.phaseId) return p;
        return { ...p, steps: editing.isNew ? [...p.steps, { id: uid(), ...body }] : p.steps.map((st) => (st.id === editing.stepId ? { ...st, ...body } : st)) };
      }));
    }
    stop();
  };
  const deletePhase = (p) => {
    if (!window.confirm(`確定要刪除階段「${p.title}」與底下 ${p.steps.length} 個步驟嗎？刪除後無法復原。`)) return;
    saveWorkflow(workflow.filter((x) => x.id !== p.id));
  };
  const deleteStep = (p, st) => {
    if (!window.confirm(`確定要刪除步驟「${st.title}」嗎？刪除後無法復原。`)) return;
    saveWorkflow(workflow.map((x) => (x.id === p.id ? { ...x, steps: x.steps.filter((y) => y.id !== st.id) } : x)));
  };
  const applyTemplate = () => saveWorkflow(WORKFLOW_TEMPLATE.map((title) => ({ id: uid(), title, summary: "", steps: [] })));

  const handleExportExcel = () => exportExcel("施工流程", [{
    name: "施工流程",
    rows: workflow.flatMap((p, i) => (p.steps.length ? p.steps : [{ title: "", details: "", tips: "" }]).map((st, j) => ({
      階段: `${i + 1}. ${p.title}`, 階段說明: p.summary || "", 步驟: st.title ? `${i + 1}-${j + 1} ${st.title}` : "", 做法細節: st.details || "", "經驗／注意事項": st.tips || "",
    }))),
  }]);
  const handleExportWord = () => exportWord("施工流程", "施工流程（經驗傳承）", workflow.map((p, i) => ({
    title: `${i + 1}. ${p.title}`,
    note: p.summary || undefined,
    headers: ["步驟", "做法細節", "經驗／注意事項"],
    rows: p.steps.map((st, j) => [`${i + 1}-${j + 1} ${st.title}`, st.details || "", st.tips || ""]),
  })));

  const form = (
    <div className="wel-card wel-form" style={{ borderColor: "var(--teal)", marginTop: 10 }}>
      <div className="wel-form-grid">
        <Field label={editing && editing.type === "phase" ? "階段名稱" : "步驟名稱"} wide>
          <input autoFocus placeholder={editing && editing.type === "phase" ? "例：進場施工" : "例：進場前先跟業主確認樓層動線"} value={draft ? draft.title : ""} onChange={(e) => setDraft({ ...draft, title: e.target.value })} />
        </Field>
        {editing && editing.type === "phase" ? (
          <Field label="這個階段在做什麼（選填）" wide>
            <textarea rows={3} value={draft.summary} onChange={(e) => setDraft({ ...draft, summary: e.target.value })} />
          </Field>
        ) : (
          <>
            <Field label="做法細節（一步一步寫清楚，可以換行）" wide>
              <textarea rows={7} placeholder={"例：\n1. 先到現場量測、拍照存證\n2. 跟業主確認配管位置…"} value={draft ? draft.details : ""} onChange={(e) => setDraft({ ...draft, details: e.target.value })} />
            </Field>
            <Field label="經驗／注意事項（踩過的坑、小技巧）" wide>
              <textarea rows={4} value={draft ? draft.tips : ""} onChange={(e) => setDraft({ ...draft, tips: e.target.value })} />
            </Field>
          </>
        )}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button type="button" className="wel-btn-primary" onClick={save}>{editing && editing.isNew ? "新增" : "儲存修改"}</button>
        <button type="button" className="wel-btn-ghost" onClick={stop}>取消</button>
      </div>
    </div>
  );
  const editingPhaseForm = (p) => editing && editing.type === "phase" && !editing.isNew && editing.phaseId === p.id;
  const editingStepForm = (p, st) => editing && editing.type === "step" && !editing.isNew && editing.phaseId === p.id && editing.stepId === st.id;

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">流程 · WORKFLOW</div>
          <h1 className="wel-h1">施工流程（經驗傳承）</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>
      <div className="muted" style={{ fontSize: 12.5, marginBottom: 12 }}>
        把從開始到結束的做法、順序與經驗寫在這裡，給之後接手的人照著做。所有案場共用同一份，隨時可以修改、調整順序或刪除。
      </div>

      {workflow.length === 0 && !editing && (
        <div className="wel-card">
          <Empty text="還沒有寫任何流程" />
          <div style={{ display: "flex", gap: 8, justifyContent: "center", flexWrap: "wrap", marginTop: 8 }}>
            <button type="button" className="wel-btn-primary" onClick={applyTemplate}>套用建議的階段（之後都可改）</button>
            <button type="button" className="wel-btn-ghost" onClick={startNewPhase}><Plus size={14} /> 自己新增階段</button>
          </div>
        </div>
      )}

      {workflow.map((p, i) => (
        <div key={p.id} className="wel-card" style={{ marginBottom: 12 }}>
          {editingPhaseForm(p) ? form : (
            <>
              <div className="wel-vendor-head" style={{ marginBottom: 6 }}>
                <button type="button" className="wel-expand-btn" onClick={() => togglePhase(p.id)}>{isOpen(p.id) ? <ChevronDown size={15} /> : <ChevronRight size={15} />}</button>
                <span className="wel-vendor-name" style={{ cursor: "pointer" }} onClick={() => togglePhase(p.id)}>{i + 1}. {p.title}</span>
                <span className="muted mono" style={{ fontSize: 11 }}>{p.steps.length} 個步驟</span>
                <button className="wel-icon-btn" title="上移" disabled={i === 0} onClick={() => saveWorkflow(move(workflow, i, -1))}><ChevronUp size={14} /></button>
                <button className="wel-icon-btn" title="下移" disabled={i === workflow.length - 1} onClick={() => saveWorkflow(move(workflow, i, 1))}><ChevronDown size={14} /></button>
                <button className="wel-icon-btn" title="修改階段" onClick={() => startEditPhase(p)}><Pencil size={14} /></button>
                <button className="wel-icon-btn" title="刪除階段" onClick={() => deletePhase(p)}><Trash2 size={14} /></button>
              </div>
              {p.summary && <div className="muted" style={{ fontSize: 13, whiteSpace: "pre-wrap", marginBottom: 8 }}>{p.summary}</div>}
            </>
          )}
          {isOpen(p.id) && !editingPhaseForm(p) && (
            <>
              {p.steps.map((st, j) => (
                <div key={st.id} className="wel-discussion-card" style={{ marginTop: 8 }}>
                  {editingStepForm(p, st) ? form : (
                    <>
                      <div className="wel-discussion-head">
                        <span className="mono" style={{ color: "var(--amber)" }}>{i + 1}-{j + 1}</span>
                        <span className="wel-discussion-topic">{st.title}</span>
                        <button className="wel-icon-btn" title="上移" disabled={j === 0} onClick={() => saveWorkflow(workflow.map((x) => (x.id === p.id ? { ...x, steps: move(x.steps, j, -1) } : x)))}><ChevronUp size={14} /></button>
                        <button className="wel-icon-btn" title="下移" disabled={j === p.steps.length - 1} onClick={() => saveWorkflow(workflow.map((x) => (x.id === p.id ? { ...x, steps: move(x.steps, j, 1) } : x)))}><ChevronDown size={14} /></button>
                        <button className="wel-icon-btn" title="修改步驟" onClick={() => startEditStep(p, st)}><Pencil size={14} /></button>
                        <button className="wel-icon-btn" title="刪除步驟" onClick={() => deleteStep(p, st)}><Trash2 size={14} /></button>
                      </div>
                      {st.details && <div style={{ whiteSpace: "pre-wrap", fontSize: 13.5, lineHeight: 1.7, marginTop: 6 }}>{st.details}</div>}
                      {st.tips && (
                        <div style={{ whiteSpace: "pre-wrap", fontSize: 13, lineHeight: 1.65, marginTop: 8, padding: "6px 10px", borderLeft: "3px solid var(--amber)", background: "rgba(232,163,61,0.08)", borderRadius: 4 }}>
                          <span className="muted" style={{ fontSize: 11.5 }}>經驗／注意事項</span>{"\n"}{st.tips}
                        </div>
                      )}
                    </>
                  )}
                </div>
              ))}
              {editing && editing.isNew && editing.type === "step" && editing.phaseId === p.id && form}
              {!(editing && editing.type === "step" && editing.isNew && editing.phaseId === p.id) && (
                <button type="button" className="wel-btn-ghost" style={{ marginTop: 10 }} onClick={() => startNewStep(p.id)}><Plus size={13} /> 新增步驟</button>
              )}
            </>
          )}
        </div>
      ))}

      {editing && editing.isNew && editing.type === "phase" && form}
      {workflow.length > 0 && !(editing && editing.isNew && editing.type === "phase") && (
        <button type="button" className="wel-btn-primary" onClick={startNewPhase}><Plus size={15} /> 新增階段</button>
      )}
    </div>
  );
}

// 估價單：簽約估價單的項目，與實際叫貨對照（估價 − 實際；負數＝超支）。僅供對照，不計入任何總結算。
function EstimateTab({
  estimates, orders, categories, catById, materialItems, siteById, isAllSites, sites, setCurrentSiteId, currentSiteId,
  estimateFiles, addEstimateFile, deleteEstimateFile, openContractFile,
  addEstimateLine, updateEstimateLine, deleteEstimateLine, confirmScannedEstimate, showToast,
}) {
  const emptyForm = { categoryId: categories[0]?.id || "", name: "", quantity: "", unit: "", unitPrice: "", note: "" };
  const [form, setForm] = useState(emptyForm);
  const [scanOpen, setScanOpen] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState(null);

  const lineAmount = (l) => (Number(l.quantity) || 0) * (Number(l.unitPrice) || 0);
  const catName = (id) => catById[id]?.name || "未分類";
  const signed = (n) => (n >= 0 ? "+" : "") + fmtMoney(n);
  const tone = (n) => (n >= 0 ? "var(--green)" : "var(--red)");

  const estTotal = estimates.reduce((s, l) => s + lineAmount(l), 0);
  const actualTotal = orders.reduce((s, o) => s + (Number(o.amount) || 0), 0);
  const diffTotal = estTotal - actualTotal;
  const usedPct = estTotal > 0 ? (actualTotal / estTotal) * 100 : 0;

  const byCategory = useMemo(() => {
    const m = {};
    const row = (id) => m[id] || (m[id] = { id, est: 0, actual: 0 });
    estimates.forEach((l) => { row(l.categoryId || "").est += (Number(l.quantity) || 0) * (Number(l.unitPrice) || 0); });
    orders.forEach((o) => { row(o.categoryId || "").actual += Number(o.amount) || 0; });
    return Object.values(m).filter((r) => r.est || r.actual).sort((a, b) => (b.est + b.actual) - (a.est + a.actual));
  }, [estimates, orders]);

  const actualByName = useMemo(() => {
    const m = {};
    orders.forEach((o) => {
      const k = (o.itemName || "").trim();
      const r = m[k] || (m[k] = { qty: 0, amount: 0, unit: o.unit || "" });
      r.qty += Number(o.quantity) || 0;
      r.amount += Number(o.amount) || 0;
    });
    return m;
  }, [orders]);

  // 每個估價項目可對應一個或多個叫貨品項（名稱相同自動對應，不同的可手動指定）；同一個叫貨品項只算給最前面的那一列
  const itemRows = useMemo(() => {
    const claimed = new Set();
    return estimates.map((l) => {
      const own = (l.name || "").trim();
      const names = [...new Set([own, ...(l.matchNames || [])])].filter((n) => n && n in actualByName && !claimed.has(n));
      names.forEach((n) => claimed.add(n));
      const actual = names.length ? names.reduce((acc, n) => ({ qty: acc.qty + actualByName[n].qty, amount: acc.amount + actualByName[n].amount }), { qty: 0, amount: 0 }) : null;
      return { line: l, names, matched: names.length > 0, actual, actualUnit: names.length ? actualByName[names[0]].unit : "", ownTaken: !names.length && own in actualByName };
    });
  }, [estimates, actualByName]);
  const allOrderNames = useMemo(() => Object.keys(actualByName).filter(Boolean).sort((a, b) => a.localeCompare(b, "zh-Hant")), [actualByName]);
  const claimedNames = useMemo(() => new Set(itemRows.flatMap((r) => r.names)), [itemRows]);
  const freeOrderNames = allOrderNames.filter((n) => !claimedNames.has(n));
  const addMatch = (line, name) => updateEstimateLine(line.id, { matchNames: [...new Set([...(line.matchNames || []), name])] });
  const removeMatch = (line, name) => updateEstimateLine(line.id, { matchNames: (line.matchNames || []).filter((n) => n !== name) });

  const catItems = materialItems.filter((m) => m.categoryId === form.categoryId);
  const setName = (name) => {
    const item = materialItems.find((m) => m.name === name.trim());
    setForm((f) => ({ ...f, name, ...(item ? { categoryId: item.categoryId || f.categoryId, unit: f.unit || item.unit || "", unitPrice: f.unitPrice === "" ? (item.defaultPrice ?? "") : f.unitPrice } : {}) }));
  };
  const submit = () => {
    if (isAllSites) { showToast("請先在左側選擇一個案場，才能新增估價項目"); return; }
    if (!form.name.trim()) { showToast("請填寫「項目名稱」"); return; }
    if (!(Number(form.quantity) > 0) || form.unitPrice === "" || !(Number(form.unitPrice) >= 0)) { showToast("請填寫「數量」與「單價」"); return; }
    addEstimateLine(form);
    setForm({ ...emptyForm, categoryId: form.categoryId });
  };
  const startEdit = (l) => { setEditingId(l.id); setDraft({ categoryId: l.categoryId, name: l.name, quantity: l.quantity, unit: l.unit || "", unitPrice: l.unitPrice, note: l.note || "", matchNames: l.matchNames || [] }); };
  const saveEdit = (id) => {
    if (!draft.name.trim()) { showToast("請填寫「項目名稱」"); return; }
    updateEstimateLine(id, { ...draft, name: draft.name.trim(), quantity: Number(draft.quantity) || 0, unitPrice: Number(draft.unitPrice) || 0, matchNames: draft.matchNames });
    setEditingId(null);
    setDraft(null);
  };

  const handleExportExcel = () => exportExcel("估價單對照", [
    { name: "依類別對照", rows: [...byCategory.map((r) => ({ 類別: catName(r.id), 估價金額: r.est, 實際叫貨: r.actual, "差異（估價－實際）": r.est - r.actual })), { 類別: "合計", 估價金額: estTotal, 實際叫貨: actualTotal, "差異（估價－實際）": diffTotal }] },
    { name: "逐項對照", rows: itemRows.map(({ line: l, matched, actual, names }) => ({ 類別: catName(l.categoryId), 項目: l.name, 對應叫貨品項: names.join("、"), 估價數量: Number(l.quantity) || 0, 單位: l.unit || "", 估價單價: Number(l.unitPrice) || 0, 估價金額: lineAmount(l), 實際數量: actual ? actual.qty : "", 實際金額: actual ? actual.amount : "", "差異金額（估價－實際）": actual ? lineAmount(l) - actual.amount : "", 備註: l.note || (matched ? "" : "未對應叫貨品項") })) },
  ]);
  const handleExportWord = () => exportWord("估價單對照", "估價單與實際叫貨對照", [
    { title: "總覽", note: "差異 ＝ 估價 － 實際叫貨；負數代表超出估價。僅供對照，不計入總結算。", headers: ["項目", "金額"], rows: [["估價合計", fmtMoney(estTotal)], ["實際叫貨合計", fmtMoney(actualTotal)], ["差異（估價淨損益）", signed(diffTotal)]] },
    { title: "依類別對照", headers: ["類別", "估價金額", "實際叫貨", "差異"], rows: byCategory.map((r) => [catName(r.id), fmtMoney(r.est), fmtMoney(r.actual), signed(r.est - r.actual)]) },
    { title: "逐項對照", headers: ["類別", "項目", "對應叫貨品項", "估價數量", "估價金額", "實際數量", "實際金額", "差異金額"], rows: itemRows.map(({ line: l, actual, names, actualUnit }) => [catName(l.categoryId), l.name, names.join("、") || "—", `${fmtNum(l.quantity)} ${l.unit || ""}`, fmtMoney(lineAmount(l)), actual ? `${fmtNum(actual.qty)} ${actualUnit || ""}` : "—", actual ? fmtMoney(actual.amount) : "—", actual ? signed(lineAmount(l) - actual.amount) : "—"]) },
  ]);

  const cell = { padding: "6px 10px", textAlign: "right", whiteSpace: "nowrap" };
  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">估價 · ESTIMATE</div>
          <h1 className="wel-h1">估價單與叫貨對照</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>
      <div className="muted" style={{ fontSize: 12.5, marginBottom: 12 }}>
        把簽約估價單的項目放進來，系統會拿你之後叫的貨去對照。差異 ＝ 估價 － 實際叫貨，<b>負數代表超出估價</b>。這裡的數字只用來統計估價與淨損益，<b>不會計入「全部案場總結算」</b>。
      </div>
      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增估價項目與上傳估價單" />}

      <div className="wel-labor-summary">
        <div className="wel-labor-stat"><span>估價合計</span><b style={{ color: "var(--teal)" }}>{fmtMoney(estTotal)}</b></div>
        <div className="wel-labor-stat"><span>實際叫貨合計</span><b style={{ color: "var(--amber)" }}>{fmtMoney(actualTotal)}</b></div>
        <div className="wel-labor-stat"><span>估價淨損益（估價 － 實際）</span><b style={{ color: tone(diffTotal) }}>{signed(diffTotal)}</b></div>
        <div className="wel-labor-stat"><span>已用估價</span><b>{estTotal > 0 ? usedPct.toFixed(1) + "%" : "—"}</b></div>
      </div>

      {!isAllSites && (
        <div className="wel-card" style={{ marginTop: 12 }}>
          <ContractFiles
            contract={{ id: currentSiteId, files: estimateFiles }} label="估價單檔案（簽約估價單原檔，PDF／照片）" uploadText="上傳估價單"
            addContractFile={(id, f) => addEstimateFile(f)} deleteContractFile={(id, f) => deleteEstimateFile(f)} openContractFile={openContractFile}
          />
        </div>
      )}

      {!isAllSites && (
        <div style={{ margin: "12px 0" }}>
          {scanOpen ? (
            <ReceiptScanner
              mode="estimate" categories={categories} materialItems={materialItems} suppliers={[]}
              defaultCategoryId={form.categoryId} onConfirmEstimate={confirmScannedEstimate}
              onClose={() => setScanOpen(false)} showToast={showToast}
            />
          ) : (
            <button type="button" className="wel-btn-ghost" onClick={() => setScanOpen(true)}><Camera size={14} /> 拍照辨識估價單（先校對再加入）</button>
          )}
        </div>
      )}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Calculator size={14} /> 新增估價項目</div>
        <div className="wel-form-grid">
          <Field label="類別">
            <select value={form.categoryId} onChange={(e) => setForm({ ...form, categoryId: e.target.value })}>
              {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </Field>
          <Field label="項目名稱（和叫貨品項同名才會逐項對照）" wide>
            <input list="wel-estimate-items" placeholder="可從品項清單選，也可自己輸入" value={form.name} onChange={(e) => setName(e.target.value)} />
            <datalist id="wel-estimate-items">{catItems.map((m) => <option key={m.id} value={m.name} />)}</datalist>
          </Field>
          <Field label="數量"><input type="number" step="any" min="0" value={form.quantity} onChange={(e) => setForm({ ...form, quantity: e.target.value })} /></Field>
          <Field label="單位"><input placeholder="支 / 米 / 式" value={form.unit} onChange={(e) => setForm({ ...form, unit: e.target.value })} /></Field>
          <Field label="估價單價"><input type="number" step="any" min="0" value={form.unitPrice} onChange={(e) => setForm({ ...form, unitPrice: e.target.value })} /></Field>
          <Field label="小計"><div className="wel-amount-preview">{fmtMoney((Number(form.quantity) || 0) * (Number(form.unitPrice) || 0))}</div></Field>
          <Field label="備註" wide><input placeholder="選填" value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} /></Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submit}><Plus size={15} /> 新增估價項目</button>
      </div>

      <div className="wel-card" style={{ marginTop: 12, padding: 0, overflow: "hidden" }}>
        <div className="wel-card-title" style={{ padding: "12px 14px 0" }}><Scale size={14} /> 依類別對照</div>
        <div style={{ overflowX: "auto" }}>
          <table className="wel-table">
            <thead><tr><th>類別</th><th className="right">估價金額</th><th className="right">實際叫貨</th><th className="right">差異（估價－實際）</th><th className="right">已用</th></tr></thead>
            <tbody>
              {byCategory.length === 0 && <tr><td colSpan={5}><Empty text="尚無估價項目，也沒有叫貨紀錄" /></td></tr>}
              {byCategory.map((r) => (
                <tr key={r.id}>
                  <td><span className="wel-tag" style={{ borderColor: catById[r.id]?.color }}>{catName(r.id)}</span>{r.est === 0 && <span className="muted" style={{ fontSize: 11, marginLeft: 6 }}>未列入估價</span>}</td>
                  <td className="mono" style={cell}>{fmtMoney(r.est)}</td>
                  <td className="mono" style={cell}>{fmtMoney(r.actual)}</td>
                  <td className="mono strong" style={{ ...cell, color: tone(r.est - r.actual) }}>{signed(r.est - r.actual)}</td>
                  <td className="mono muted" style={cell}>{r.est > 0 ? ((r.actual / r.est) * 100).toFixed(0) + "%" : "—"}</td>
                </tr>
              ))}
              {byCategory.length > 0 && (
                <tr style={{ borderTop: "2px solid var(--border)", background: "var(--surface2)" }}>
                  <td className="strong">合計</td>
                  <td className="mono strong" style={cell}>{fmtMoney(estTotal)}</td>
                  <td className="mono strong" style={cell}>{fmtMoney(actualTotal)}</td>
                  <td className="mono strong" style={{ ...cell, color: tone(diffTotal) }}>{signed(diffTotal)}</td>
                  <td className="mono strong" style={cell}>{estTotal > 0 ? usedPct.toFixed(0) + "%" : "—"}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="wel-card" style={{ marginTop: 12, padding: 0, overflow: "hidden" }}>
        <div className="wel-card-title" style={{ padding: "12px 14px 0" }}><ListOrdered size={14} /> 逐項對照（點一列可修改）</div>
        <div className="muted" style={{ fontSize: 11.5, padding: "4px 14px 8px" }}>名稱和叫貨品項相同會自動對應；名稱不同時，列下方會出現「建議對應」可一鍵套用，或點該列自己指定要對應哪些叫貨品項（只要指定一次，之後都會記住）。同一個叫貨品項只會算給最前面的那一列。</div>
        <div style={{ overflowX: "auto" }}>
          <table className="wel-table">
            <thead>
              <tr>{isAllSites && <th>案場</th>}<th>類別</th><th>項目</th><th className="right">估價數量</th><th className="right">估價金額</th><th className="right">實際數量</th><th className="right">實際金額</th><th className="right">差異金額</th><th></th></tr>
            </thead>
            <tbody>
              {itemRows.length === 0 && <tr><td colSpan={isAllSites ? 9 : 8}><Empty text="尚未新增估價項目" /></td></tr>}
              {itemRows.map(({ line: l, matched, actual, actualUnit, names, ownTaken }) => {
                const sugg = !matched && !ownTaken ? suggestNames(l.name, freeOrderNames, 3).filter((n) => !(l.matchNames || []).includes(n)) : [];
                const isOpen = editingId === l.id;
                return (
                  <React.Fragment key={l.id}>
                    <tr className="wel-row-clickable" onClick={() => (isOpen ? (setEditingId(null), setDraft(null)) : startEdit(l))}>
                      {isAllSites && <td className="muted">{siteById[l.siteId]?.name || "—"}</td>}
                      <td><span className="wel-tag" style={{ borderColor: catById[l.categoryId]?.color }}>{catName(l.categoryId)}</span></td>
                      <td>
                        <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>{isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}{l.name}</span>
                        {l.note && <div className="muted" style={{ fontSize: 11.5, marginLeft: 18 }}>{l.note}</div>}
                        {matched && (l.matchNames || []).length > 0 && <div className="muted" style={{ fontSize: 11.5, marginLeft: 18 }}>對應叫貨：{names.join("、")}</div>}
                        {sugg.length > 0 && (
                          <div style={{ fontSize: 11.5, marginLeft: 18, display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                            <span className="muted">建議對應：</span>
                            {sugg.map((n) => <button key={n} type="button" className="wel-chip-btn" onClick={(e) => { e.stopPropagation(); addMatch(l, n); }} title="點一下，把這個叫貨品項算給這一列">{n}</button>)}
                          </div>
                        )}
                      </td>
                      <td className="mono" style={cell}>{fmtNum(l.quantity)} {l.unit}</td>
                      <td className="mono" style={cell}>{fmtMoney(lineAmount(l))}</td>
                      <td className="mono" style={cell}>{actual ? `${fmtNum(actual.qty)} ${actualUnit || ""}` : "—"}</td>
                      <td className="mono" style={cell}>{actual ? fmtMoney(actual.amount) : (ownTaken ? <span className="muted" style={{ fontSize: 11 }}>同名併入上方</span> : "—")}</td>
                      <td className="mono strong" style={{ ...cell, color: actual ? tone(lineAmount(l) - actual.amount) : undefined }}>{actual ? signed(lineAmount(l) - actual.amount) : "—"}</td>
                      <td><button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteEstimateLine(l.id); }}><Trash2 size={14} /></button></td>
                    </tr>
                    {isOpen && draft && (
                      <tr onClick={(e) => e.stopPropagation()}>
                        <td colSpan={isAllSites ? 9 : 8} style={{ padding: 0 }}>
                          <div className="wel-timeline-wrap">
                            <div className="wel-form-grid">
                              <Field label="類別">
                                <select value={draft.categoryId} onChange={(e) => setDraft({ ...draft, categoryId: e.target.value })}>
                                  {categories.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                                </select>
                              </Field>
                              <Field label="項目名稱" wide><input list="wel-estimate-items" value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} /></Field>
                              <Field label="數量"><input type="number" step="any" min="0" value={draft.quantity} onChange={(e) => setDraft({ ...draft, quantity: e.target.value })} /></Field>
                              <Field label="單位"><input value={draft.unit} onChange={(e) => setDraft({ ...draft, unit: e.target.value })} /></Field>
                              <Field label="估價單價"><input type="number" step="any" min="0" value={draft.unitPrice} onChange={(e) => setDraft({ ...draft, unitPrice: e.target.value })} /></Field>
                              <Field label="備註" wide><input value={draft.note} onChange={(e) => setDraft({ ...draft, note: e.target.value })} /></Field>
                              <Field label="對應叫貨品項（估價名稱和叫貨品項名稱不同時，在這裡指定，可選多個）" wide>
                                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center", minHeight: 34 }}>
                                  {draft.matchNames.map((n) => (
                                    <span key={n} className="wel-chip" style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                                      {n}
                                      <button type="button" className="wel-icon-btn" style={{ padding: 0 }} onClick={() => setDraft({ ...draft, matchNames: draft.matchNames.filter((x) => x !== n) })}><X size={11} /></button>
                                    </span>
                                  ))}
                                  <select className="wel-scan-input" style={{ maxWidth: 260 }} value="" onChange={(e) => { if (e.target.value) setDraft({ ...draft, matchNames: [...draft.matchNames, e.target.value] }); }}>
                                    <option value="">＋ 加入叫貨品項…</option>
                                    {allOrderNames.filter((n) => !draft.matchNames.includes(n)).map((n) => <option key={n} value={n}>{n}{claimedNames.has(n) ? "（已對應其他列）" : ""}</option>)}
                                  </select>
                                </div>
                              </Field>
                            </div>
                            <div style={{ display: "flex", gap: 8 }}>
                              <button type="button" className="wel-btn-primary" onClick={() => saveEdit(l.id)}>儲存修改</button>
                              <button type="button" className="wel-btn-ghost" onClick={() => { setEditingId(null); setDraft(null); }}>取消</button>
                            </div>
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function NotebookTab({
  notebookItems, siteById, notebookForm, setNotebookForm, submitNotebookItem, updateNotebookItem, deleteNotebookItem,
  isAllSites, sites, setCurrentSiteId, showToast,
}) {
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const startEdit = (n) => { setEditingId(n.id); setEditDraft({ date: n.date, title: n.title, content: n.content || "" }); };
  const cancelEdit = () => { setEditingId(null); setEditDraft(null); };
  const saveEdit = (id) => {
    if (!editDraft.title.trim()) { showToast("請填寫「標題」後再儲存"); return; }
    updateNotebookItem(id, editDraft);
    setEditingId(null);
    setEditDraft(null);
  };
  const sortedItems = useMemo(() => notebookItems.slice().sort((a, b) => (a.date < b.date ? 1 : -1)), [notebookItems]);

  const handleExportExcel = () => exportExcel("大項記事本", [{
    name: "記事本",
    rows: sortedItems.map((n) => ({ 日期: n.date, 標題: n.title, 內容: n.content || "" })),
  }]);
  const handleExportWord = () => exportWord("大項記事本", "大項記事本", [{
    title: "大項記事本",
    headers: ["日期", "標題", "內容"],
    rows: sortedItems.map((n) => [n.date, n.title, n.content || ""]),
  }]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">記事 · NOTEBOOK</div>
          <h1 className="wel-h1">大項記事本</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增記事本項目" />}

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><NotebookPen size={14} /> 新增記事本項目</div>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={notebookForm.date} onChange={(e) => setNotebookForm({ ...notebookForm, date: e.target.value })} />
          </Field>
          <Field label="標題" wide>
            <input placeholder="例：3樓浴室防水工程" value={notebookForm.title} onChange={(e) => setNotebookForm({ ...notebookForm, title: e.target.value })} />
          </Field>
          <Field label="內容" wide>
            <input placeholder="選填" value={notebookForm.content} onChange={(e) => setNotebookForm({ ...notebookForm, content: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitNotebookItem}><Plus size={15} /> 新增記事本項目</button>
      </div>

      <div className="wel-discussion-list">
        {sortedItems.length === 0 && <div className="wel-card"><Empty text="尚無記事本項目" /></div>}
        {sortedItems.map((n) => (
          <div key={n.id} className="wel-discussion-card">
            {editingId === n.id ? (
              <div className="wel-timeline-wrap" style={{ padding: 0 }}>
                <div className="wel-form-grid">
                  <Field label="日期">
                    <input type="date" value={editDraft.date} onChange={(e) => setEditDraft({ ...editDraft, date: e.target.value })} />
                  </Field>
                  <Field label="標題" wide>
                    <input value={editDraft.title} onChange={(e) => setEditDraft({ ...editDraft, title: e.target.value })} />
                  </Field>
                  <Field label="內容" wide>
                    <input value={editDraft.content} onChange={(e) => setEditDraft({ ...editDraft, content: e.target.value })} />
                  </Field>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button type="button" className="wel-btn-primary" onClick={() => saveEdit(n.id)}>儲存修改</button>
                  <button type="button" className="wel-btn-ghost" onClick={cancelEdit}>取消</button>
                </div>
              </div>
            ) : (
              <>
                <div className="wel-discussion-head">
                  <span className="wel-discussion-topic">{n.title}</span>
                  {isAllSites && <span className="wel-tag">{siteById[n.siteId]?.name || "—"}</span>}
                  <span className="mono muted" style={{ fontSize: 11.5 }}>{n.date}</span>
                  <button className="wel-icon-btn" onClick={() => startEdit(n)}><Pencil size={14} /></button>
                  <button className="wel-icon-btn" onClick={() => deleteNotebookItem(n.id)}><Trash2 size={14} /></button>
                </div>
                {n.content && <div className="wel-discussion-note muted">{n.content}</div>}
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

function CapitalTab({
  capitalContributions, siteById, capitalForm, setCapitalForm, submitCapitalContribution, deleteCapitalContribution, updateCapitalContribution,
  capitalTotal, isAllSites, sites, setCurrentSiteId,
}) {
  const [expandedId, setExpandedId] = useState(null);
  const [draft, setDraft] = useState(null);
  const startEdit = (c) => { setExpandedId(c.id); setDraft({ date: c.date, amount: c.amount, note: c.note || "" }); };
  const cancelEdit = () => { setExpandedId(null); setDraft(null); };
  const saveEdit = (id) => { updateCapitalContribution(id, { ...draft, amount: Number(draft.amount) || 0 }); setExpandedId(null); setDraft(null); };
  const sortedContributions = useMemo(() => capitalContributions.slice().sort((a, b) => (a.date < b.date ? 1 : -1)), [capitalContributions]);
  const handleExportExcel = () => exportExcel("投入成本", [
    { name: "投入成本", rows: sortedContributions.map((c) => ({ 日期: c.date, 金額: Number(c.amount) || 0, 備註: c.note || "" })) },
  ]);
  const handleExportWord = () => exportWord("投入成本", "投入成本紀錄", [
    { title: "投入成本紀錄", headers: ["日期", "金額", "備註"], rows: sortedContributions.map((c) => [c.date, fmtMoney(c.amount), c.note || ""]) },
  ]);
  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">資金 · CAPITAL CONTRIBUTIONS</div>
          <h1 className="wel-h1">投入成本</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增投入成本紀錄" />}

      <div className="wel-meter-card">
        <div className="wel-meter-label"><PiggyBank size={14} /> 累計投入成本</div>
        <div className="wel-meter-sub" style={{ marginTop: 0 }}>
          <span>你投入的資金，已在收入中回收，所以只記錄、供參考，不計入總支出與損益</span>
        </div>
        <div style={{ marginTop: 10, fontFamily: "var(--font-mono)", fontSize: 26, fontWeight: 600, color: "var(--teal)" }}>{fmtMoney(capitalTotal)}</div>
      </div>

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><PiggyBank size={14} /> 新增投入成本</div>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={capitalForm.date} onChange={(e) => setCapitalForm({ ...capitalForm, date: e.target.value })} />
          </Field>
          <Field label="金額">
            <input type="number" min="0" step="any" value={capitalForm.amount} onChange={(e) => setCapitalForm({ ...capitalForm, amount: e.target.value })} />
          </Field>
          <Field label="備註" wide>
            <input placeholder="選填，例：開工前先墊付材料款" value={capitalForm.note} onChange={(e) => setCapitalForm({ ...capitalForm, note: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitCapitalContribution}><Plus size={15} /> 新增投入成本</button>
      </div>

      <div className="wel-card" style={{ padding: 0, overflow: "hidden" }}>
        <table className="wel-table">
          <thead>
            <tr>{isAllSites && <th>案場</th>}<th>日期</th><th className="right">金額</th><th>備註</th><th></th></tr>
          </thead>
          <tbody>
            {sortedContributions.length === 0 && (<tr><td colSpan={isAllSites ? 4 : 3}><Empty text="尚無投入成本紀錄" /></td></tr>)}
            {sortedContributions.map((c) => {
              const isOpen = expandedId === c.id;
              return (
              <React.Fragment key={c.id}>
              <tr className="wel-row-clickable" onClick={() => (isOpen ? cancelEdit() : startEdit(c))}>
                {isAllSites && <td className="muted">{siteById[c.siteId]?.name || "—"}</td>}
                <td className="mono">
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                    {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                    {c.date}
                  </span>
                </td>
                <td className="right mono strong" style={{ color: "var(--teal)" }}>{fmtMoney(c.amount)}</td>
                <td className="muted">{c.note || "—"}</td>
                <td><button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteCapitalContribution(c.id); }}><Trash2 size={14} /></button></td>
              </tr>
              {isOpen && draft && (
                <tr onClick={(e) => e.stopPropagation()}>
                  <td colSpan={isAllSites ? 4 : 3} style={{ padding: 0 }}>
                    <div className="wel-timeline-wrap">
                      <div className="wel-form-grid">
                        <Field label="日期">
                          <input type="date" value={draft.date} onChange={(e) => setDraft({ ...draft, date: e.target.value })} />
                        </Field>
                        <Field label="金額">
                          <input type="number" min="0" step="any" value={draft.amount} onChange={(e) => setDraft({ ...draft, amount: e.target.value })} />
                        </Field>
                        <Field label="備註" wide>
                          <input value={draft.note} onChange={(e) => setDraft({ ...draft, note: e.target.value })} />
                        </Field>
                      </div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button type="button" className="wel-btn-primary" onClick={() => saveEdit(c.id)}>儲存修改</button>
                        <button type="button" className="wel-btn-ghost" onClick={cancelEdit}>取消</button>
                      </div>
                    </div>
                  </td>
                </tr>
              )}
              </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FixedAssetsTab({
  fixedAssets, siteById, fixedAssetForm, setFixedAssetForm, submitFixedAsset, deleteFixedAsset, updateFixedAsset,
  fixedAssetsTotal, isAllSites, sites, setCurrentSiteId,
}) {
  const [expandedId, setExpandedId] = useState(null);
  const [draft, setDraft] = useState(null);
  const startEdit = (f) => { setExpandedId(f.id); setDraft({ name: f.name, quantity: f.quantity, amount: f.amount, date: f.date, hasInvoice: !!f.hasInvoice, applyTax: !!f.applyTax, taxRate: f.taxRate ?? 5 }); };
  const cancelEdit = () => { setExpandedId(null); setDraft(null); };
  const saveEdit = (id) => { updateFixedAsset(id, { ...draft, quantity: Number(draft.quantity) || 1, amount: Number(draft.amount) || 0 }); setExpandedId(null); setDraft(null); };
  const sortedAssets = useMemo(() => fixedAssets.slice().sort((a, b) => (a.date < b.date ? 1 : -1)), [fixedAssets]);
  const taxOf = (f) => (f.hasInvoice && f.applyTax ? (Number(f.amount) || 0) * (Number(f.quantity) || 0) * (Number(f.taxRate) || 0) / 100 : 0);
  const handleExportExcel = () => exportExcel("固定資產", [
    { name: "固定資產清單", rows: sortedAssets.map((f) => ({ 工具名稱: f.name, 數量: Number(f.quantity) || 0, 購買金額: Number(f.amount) || 0, 小計: (Number(f.quantity) || 0) * (Number(f.amount) || 0), 發票: f.hasInvoice ? "有發票" : "無發票", 稅額: taxOf(f), 購買日期: f.date })) },
  ]);
  const handleExportWord = () => exportWord("固定資產", "固定資產清點清單", [
    { title: "固定資產清單", headers: ["工具名稱", "數量", "購買金額", "小計", "發票", "稅額", "購買日期"], rows: sortedAssets.map((f) => [f.name, f.quantity, fmtMoney(f.amount), fmtMoney((Number(f.quantity) || 0) * (Number(f.amount) || 0)), f.hasInvoice ? "有發票" : "無發票", taxOf(f) ? fmtMoney(taxOf(f)) : "—", f.date]) },
  ]);
  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">資產 · FIXED ASSETS</div>
          <h1 className="wel-h1">固定資產清點</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增固定資產" />}

      <div className="wel-meter-card">
        <div className="wel-meter-label"><Hammer size={14} /> 固定資產總值</div>
        <div className="wel-meter-sub" style={{ marginTop: 0 }}>
          <span>生財工具、設備清點，會計入這個案場的總支出／毛利計算</span>
        </div>
        <div style={{ marginTop: 10, fontFamily: "var(--font-mono)", fontSize: 26, fontWeight: 600, color: "var(--amber)" }}>{fmtMoney(fixedAssetsTotal)}</div>
      </div>

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Hammer size={14} /> 新增固定資產</div>
        <div className="wel-form-grid">
          <Field label="工具名稱" wide>
            <input placeholder="例：電鑽、鷹架、廂型車" value={fixedAssetForm.name} onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, name: e.target.value })} />
          </Field>
          <Field label="數量">
            <input type="number" min="1" step="1" value={fixedAssetForm.quantity} onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, quantity: e.target.value })} />
          </Field>
          <Field label="購買金額">
            <input type="number" min="0" step="any" value={fixedAssetForm.amount} onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, amount: e.target.value })} />
          </Field>
          <Field label="購買日期">
            <input type="date" value={fixedAssetForm.date} onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, date: e.target.value })} />
          </Field>
          <Field label="有無發票">
            <select value={fixedAssetForm.hasInvoice ? "yes" : "no"} onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, hasInvoice: e.target.value === "yes" })}>
              <option value="yes">有發票</option>
              <option value="no">無發票</option>
            </select>
          </Field>
          {fixedAssetForm.hasInvoice && (
            <Field label="營業稅" wide>
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                  <input
                    type="checkbox"
                    checked={fixedAssetForm.applyTax}
                    onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, applyTax: e.target.checked })}
                    style={{ width: 16, height: 16 }}
                  />
                  計算稅額
                </label>
                {fixedAssetForm.applyTax && (
                  <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                    稅率
                    <input
                      type="number" min="0" step="0.1" value={fixedAssetForm.taxRate}
                      onChange={(e) => setFixedAssetForm({ ...fixedAssetForm, taxRate: e.target.value })}
                      style={{ width: 60, height: 32 }}
                    />
                    %
                  </span>
                )}
              </div>
            </Field>
          )}
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitFixedAsset}><Plus size={15} /> 新增固定資產</button>
      </div>

      <div className="wel-card" style={{ padding: 0, overflow: "hidden" }}>
        <table className="wel-table">
          <thead>
            <tr>{isAllSites && <th>案場</th>}<th>工具名稱</th><th className="right">數量</th><th className="right">購買金額</th><th className="right">小計</th><th>購買日期</th><th>發票</th><th></th></tr>
          </thead>
          <tbody>
            {sortedAssets.length === 0 && (<tr><td colSpan={isAllSites ? 8 : 7}><Empty text="尚無固定資產紀錄" /></td></tr>)}
            {sortedAssets.map((f) => {
              const isOpen = expandedId === f.id;
              return (
              <React.Fragment key={f.id}>
              <tr className="wel-row-clickable" onClick={() => (isOpen ? cancelEdit() : startEdit(f))}>
                {isAllSites && <td className="muted">{siteById[f.siteId]?.name || "—"}</td>}
                <td>
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                    {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                    {f.name}
                  </span>
                </td>
                <td className="right mono">{fmtNum(f.quantity)}</td>
                <td className="right mono">{fmtMoney(f.amount)}</td>
                <td className="right mono strong" style={{ color: "var(--amber)" }}>{fmtMoney((Number(f.quantity) || 0) * (Number(f.amount) || 0))}</td>
                <td className="mono muted">{f.date}</td>
                <td>
                  {f.hasInvoice ? <span className="wel-chip-info">有發票</span> : <span className="wel-chip">無發票</span>}
                  {f.hasInvoice && f.applyTax && (
                    <div className="muted mono" style={{ fontSize: 11, marginTop: 3 }}>
                      稅額 {fmtMoney(taxOf(f))}（{fmtNum(f.taxRate)}%）
                    </div>
                  )}
                </td>
                <td><button className="wel-icon-btn" onClick={(e) => { e.stopPropagation(); deleteFixedAsset(f.id); }}><Trash2 size={14} /></button></td>
              </tr>
              {isOpen && draft && (
                <tr onClick={(e) => e.stopPropagation()}>
                  <td colSpan={isAllSites ? 8 : 7} style={{ padding: 0 }}>
                    <div className="wel-timeline-wrap">
                      <div className="wel-form-grid">
                        <Field label="工具名稱" wide>
                          <input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} />
                        </Field>
                        <Field label="數量">
                          <input type="number" min="1" step="1" value={draft.quantity} onChange={(e) => setDraft({ ...draft, quantity: e.target.value })} />
                        </Field>
                        <Field label="購買金額">
                          <input type="number" min="0" step="any" value={draft.amount} onChange={(e) => setDraft({ ...draft, amount: e.target.value })} />
                        </Field>
                        <Field label="購買日期">
                          <input type="date" value={draft.date} onChange={(e) => setDraft({ ...draft, date: e.target.value })} />
                        </Field>
                        <Field label="有無發票">
                          <select value={draft.hasInvoice ? "yes" : "no"} onChange={(e) => setDraft({ ...draft, hasInvoice: e.target.value === "yes" })}>
                            <option value="yes">有發票</option>
                            <option value="no">無發票</option>
                          </select>
                        </Field>
                        {draft.hasInvoice && (
                          <Field label="營業稅" wide>
                            <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                                <input
                                  type="checkbox" checked={draft.applyTax}
                                  onChange={(e) => setDraft({ ...draft, applyTax: e.target.checked })}
                                  style={{ width: 16, height: 16 }}
                                />
                                計算稅額
                              </label>
                              {draft.applyTax && (
                                <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                                  稅率
                                  <input
                                    type="number" min="0" step="0.1" value={draft.taxRate}
                                    onChange={(e) => setDraft({ ...draft, taxRate: e.target.value })}
                                    style={{ width: 60, height: 32 }}
                                  />
                                  %
                                </span>
                              )}
                            </div>
                          </Field>
                        )}
                      </div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <button type="button" className="wel-btn-primary" onClick={() => saveEdit(f.id)}>儲存修改</button>
                        <button type="button" className="wel-btn-ghost" onClick={cancelEdit}>取消</button>
                      </div>
                    </div>
                  </td>
                </tr>
              )}
              </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function OtherExpensesTab({
  otherExpenses, siteById, otherExpenseForm, setOtherExpenseForm, submitOtherExpense, deleteOtherExpense, updateOtherExpense,
  otherExpensesTotal, materialItems, suppliers, addSupplier, updateSupplier, removeSupplier, isAllSites, sites, setCurrentSiteId, showToast,
}) {
  const [expandedExpenseId, setExpandedExpenseId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const [openExpenseMonths, setOpenExpenseMonths] = useState({});
  const [invoiceFilter, setInvoiceFilter] = useState("all"); // all | yes | no
  const [supplierManageOpen, setSupplierManageOpen] = useState(false);
  const [newSupplier, setNewSupplier] = useState("");

  const startEditExpense = (e) => {
    setExpandedExpenseId(e.id);
    setEditDraft({
      date: e.date, itemName: e.itemName, supplier: e.supplier || "", amount: e.amount, note: e.note || "",
      hasInvoice: !!e.hasInvoice, applyTax: !!e.applyTax, taxRate: e.taxRate ?? 5,
    });
  };
  const cancelEditExpense = () => { setExpandedExpenseId(null); setEditDraft(null); };
  const saveEditExpense = (id) => {
    if (!editDraft.itemName.trim() || !editDraft.amount) {
      showToast("請填寫「項目名稱」與「金額」後再儲存");
      return;
    }
    updateOtherExpense(id, { ...editDraft, amount: Number(editDraft.amount) });
    setExpandedExpenseId(null);
    setEditDraft(null);
  };

  const sortedSuppliers = useMemo(() => suppliers.slice().sort((a, b) => a.name.localeCompare(b.name, "zh-Hant")), [suppliers]);
  const handlePickSupplier = (e) => {
    const val = e.target.value;
    if (val === "__manage__") { setSupplierManageOpen(true); return; }
    setOtherExpenseForm({ ...otherExpenseForm, supplier: val });
  };

  const filteredExpenses = useMemo(
    () => otherExpenses.filter((e) => (invoiceFilter === "all" ? true : invoiceFilter === "yes" ? !!e.hasInvoice : !e.hasInvoice)),
    [otherExpenses, invoiceFilter]
  );
  const taxOf = (e) => (e.hasInvoice && e.applyTax ? (Number(e.amount) || 0) * (Number(e.taxRate) || 0) / 100 : 0);

  const handleExportExcel = () => exportExcel("其他支出", [
    { name: "其他支出", rows: filteredExpenses.map((e) => ({ 日期: e.date, 項目: e.itemName, 廠商: e.supplier || "", 金額: Number(e.amount) || 0, 發票: e.hasInvoice ? "有發票" : "無發票", 稅額: taxOf(e), 備註: e.note || "" })) },
  ]);
  const handleExportWord = () => exportWord("其他支出", "其他支出紀錄", [
    { title: "其他支出紀錄", headers: ["日期", "項目", "廠商", "金額", "發票", "稅額", "備註"], rows: filteredExpenses.map((e) => [e.date, e.itemName, e.supplier || "", fmtMoney(e.amount), e.hasInvoice ? "有發票" : "無發票", taxOf(e) ? fmtMoney(taxOf(e)) : "—", e.note || ""]) },
  ]);

  return (
    <div className="wel-page">
      <div className="wel-page-head">
        <div>
          <div className="wel-eyebrow">支出 · OTHER EXPENSES</div>
          <h1 className="wel-h1">其他支出</h1>
        </div>
        <ExportBar onExcel={handleExportExcel} onWord={handleExportWord} />
      </div>

      {isAllSites && <AllSitesNotice sites={sites} setCurrentSiteId={setCurrentSiteId} action="新增其他支出" />}

      <div className="wel-meter-card">
        <div className="wel-meter-label"><Receipt size={14} /> 累計其他支出</div>
        <div className="wel-meter-sub" style={{ marginTop: 0 }}>
          <span>計入專案總支出／毛利計算，但不會影響領料使用或庫存現況的數量</span>
        </div>
        <div style={{ marginTop: 10, fontFamily: "var(--font-mono)", fontSize: 26, fontWeight: 600, color: "var(--amber)" }}>{fmtMoney(otherExpensesTotal)}</div>
      </div>

      <div className="wel-card wel-form" style={isAllSites ? { opacity: 0.5, pointerEvents: "none" } : undefined}>
        <div className="wel-card-title"><Receipt size={14} /> 新增其他支出</div>
        <div className="wel-form-grid">
          <Field label="日期">
            <input type="date" value={otherExpenseForm.date} onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, date: e.target.value })} />
          </Field>
          <Field label="項目名稱" wide>
            <input
              list="wel-other-expense-items"
              placeholder="例：油錢、餐費，也可從品項清單挑選"
              value={otherExpenseForm.itemName}
              onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, itemName: e.target.value })}
            />
            <datalist id="wel-other-expense-items">
              {materialItems.map((m) => <option key={m.id} value={m.name} />)}
            </datalist>
          </Field>
          <Field label="廠商 / 供應商" wide>
            <select value={sortedSuppliers.some((s) => s.name === otherExpenseForm.supplier) ? otherExpenseForm.supplier : ""} onChange={handlePickSupplier}>
              <option value="">（未指定）</option>
              {sortedSuppliers.map((s) => <option key={s.id} value={s.name}>{s.name}</option>)}
              <option value="__manage__">＋ 管理廠商清單…</option>
            </select>
            {supplierManageOpen && (
              <div className="wel-material-manager">
                <div className="wel-material-add-row">
                  <input placeholder="新廠商名稱" value={newSupplier} onChange={(e) => setNewSupplier(e.target.value)} />
                  <button type="button" className="wel-btn-ghost" onClick={() => { addSupplier(newSupplier); setNewSupplier(""); }}>
                    <Plus size={13} /> 新增
                  </button>
                </div>
                <div className="wel-manager-list">
                  {sortedSuppliers.length === 0 && <span className="muted" style={{ fontSize: 12 }}>尚未建立任何廠商</span>}
                  {sortedSuppliers.map((s) => (
                    <div key={s.id} className="wel-manager-row">
                      <input className="wel-manager-name" value={s.name} onChange={(e) => updateSupplier(s.id, { name: e.target.value })} />
                      <button type="button" className="wel-icon-btn" onClick={() => removeSupplier(s.id)}><Trash2 size={13} /></button>
                    </div>
                  ))}
                </div>
                <div className="muted" style={{ fontSize: 11 }}>提示：改名稱會自動同步更新過去的叫貨／其他支出紀錄。</div>
                <button type="button" className="wel-btn-ghost" onClick={() => setSupplierManageOpen(false)}>收合廠商管理</button>
              </div>
            )}
          </Field>
          <Field label="金額">
            <input type="number" min="0" step="any" value={otherExpenseForm.amount} onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, amount: e.target.value })} />
          </Field>
          <Field label="有無發票">
            <select value={otherExpenseForm.hasInvoice ? "yes" : "no"} onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, hasInvoice: e.target.value === "yes" })}>
              <option value="yes">有發票</option>
              <option value="no">無發票</option>
            </select>
          </Field>
          {otherExpenseForm.hasInvoice && (
            <Field label="營業稅" wide>
              <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                  <input
                    type="checkbox"
                    checked={otherExpenseForm.applyTax}
                    onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, applyTax: e.target.checked })}
                    style={{ width: 16, height: 16 }}
                  />
                  計算稅額
                </label>
                {otherExpenseForm.applyTax && (
                  <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                    稅率
                    <input
                      type="number" min="0" step="0.1" value={otherExpenseForm.taxRate}
                      onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, taxRate: e.target.value })}
                      style={{ width: 60, height: 32 }}
                    />
                    %
                  </span>
                )}
              </div>
            </Field>
          )}
          <Field label="備註" wide>
            <input placeholder="選填" value={otherExpenseForm.note} onChange={(e) => setOtherExpenseForm({ ...otherExpenseForm, note: e.target.value })} />
          </Field>
        </div>
        <button type="button" className="wel-btn-primary" onClick={submitOtherExpense}><Plus size={15} /> 新增其他支出</button>
      </div>

      <div className="wel-card">
        <div className="wel-card-title-row">
          <div className="wel-card-title"><CalendarRange size={14} /> 其他支出紀錄（依年月點開查詢）</div>
          <div className="wel-filter">
            <select value={invoiceFilter} onChange={(e) => setInvoiceFilter(e.target.value)}>
              <option value="all">全部（{otherExpenses.length}）</option>
              <option value="yes">有發票（{otherExpenses.filter((e) => e.hasInvoice).length}）</option>
              <option value="no">無發票（{otherExpenses.filter((e) => !e.hasInvoice).length}）</option>
            </select>
          </div>
        </div>
        {filteredExpenses.length === 0 && <Empty text="尚無其他支出紀錄" />}
        {attendanceByMonth(filteredExpenses).map((mg) => {
          const subtotal = mg.items.reduce((s, e) => s + (Number(e.amount) || 0), 0);
          const mOpen = !!openExpenseMonths[mg.month];
          return (
            <div key={mg.month} className="wel-item-floor-group">
              <button type="button" className="wel-item-group-label wel-item-group-toggle" onClick={() => setOpenExpenseMonths((prev) => ({ ...prev, [mg.month]: !prev[mg.month] }))}>
                {mOpen ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
                {mg.month} · {mg.items.length} 筆 · {fmtMoney(subtotal)}
              </button>
              {mOpen && (
                <div style={{ overflowX: "auto" }}>
                  <table className="wel-table">
                    <thead>
                      <tr>
                        {isAllSites && <th>案場</th>}
                        <th>日期</th><th>項目</th><th>廠商</th><th className="right">金額</th><th>發票</th><th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {mg.items.map((e) => {
                        const isOpen = expandedExpenseId === e.id;
                        return (
                        <React.Fragment key={e.id}>
                        <tr className="wel-row-clickable" onClick={() => (isOpen ? cancelEditExpense() : startEditExpense(e))}>
                          {isAllSites && <td className="muted">{siteById[e.siteId]?.name || "—"}</td>}
                          <td className="mono">{e.date}</td>
                          <td>
                            <span style={{ display: "inline-flex", alignItems: "center", gap: 5 }}>
                              {isOpen ? <ChevronDown size={13} /> : <ChevronRight size={13} />}
                              {e.itemName}
                            </span>
                            {e.note && <div className="muted" style={{ fontSize: 11.5, marginTop: 2, marginLeft: 18 }}>{e.note}</div>}
                          </td>
                          <td className="muted">{e.supplier || "—"}</td>
                          <td className="right mono strong">{fmtMoney(e.amount)}</td>
                          <td>
                            {e.hasInvoice
                              ? <span className="wel-chip-info">有發票</span>
                              : <span className="wel-chip">無發票</span>}
                            {e.hasInvoice && e.applyTax && (
                              <div className="muted mono" style={{ fontSize: 11, marginTop: 3 }}>
                                稅額 {fmtMoney((Number(e.amount) || 0) * (Number(e.taxRate) || 0) / 100)}（{fmtNum(e.taxRate)}%）
                              </div>
                            )}
                          </td>
                          <td><button className="wel-icon-btn" onClick={(ev) => { ev.stopPropagation(); deleteOtherExpense(e.id); }}><Trash2 size={14} /></button></td>
                        </tr>
                        {isOpen && editDraft && (
                          <tr onClick={(ev) => ev.stopPropagation()}>
                            <td colSpan={isAllSites ? 7 : 6} style={{ padding: 0 }}>
                              <div className="wel-timeline-wrap">
                                <div className="wel-form-grid">
                                  <Field label="日期">
                                    <input type="date" value={editDraft.date} onChange={(ev) => setEditDraft({ ...editDraft, date: ev.target.value })} />
                                  </Field>
                                  <Field label="項目名稱" wide>
                                    <input
                                      list="wel-other-expense-items"
                                      value={editDraft.itemName}
                                      onChange={(ev) => setEditDraft({ ...editDraft, itemName: ev.target.value })}
                                    />
                                  </Field>
                                  <Field label="廠商 / 供應商">
                                    <input value={editDraft.supplier} onChange={(ev) => setEditDraft({ ...editDraft, supplier: ev.target.value })} />
                                  </Field>
                                  <Field label="金額">
                                    <input type="number" step="any" min="0" value={editDraft.amount} onChange={(ev) => setEditDraft({ ...editDraft, amount: ev.target.value })} />
                                  </Field>
                                  <Field label="有無發票">
                                    <select value={editDraft.hasInvoice ? "yes" : "no"} onChange={(ev) => setEditDraft({ ...editDraft, hasInvoice: ev.target.value === "yes" })}>
                                      <option value="yes">有發票</option>
                                      <option value="no">無發票</option>
                                    </select>
                                  </Field>
                                  {editDraft.hasInvoice && (
                                    <Field label="營業稅" wide>
                                      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", height: 40 }}>
                                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, color: "var(--text)", cursor: "pointer" }}>
                                          <input
                                            type="checkbox" checked={editDraft.applyTax}
                                            onChange={(ev) => setEditDraft({ ...editDraft, applyTax: ev.target.checked })}
                                            style={{ width: 16, height: 16 }}
                                          />
                                          計算稅額
                                        </label>
                                        {editDraft.applyTax && (
                                          <span style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12.5 }}>
                                            稅率
                                            <input
                                              type="number" min="0" step="0.1" value={editDraft.taxRate}
                                              onChange={(ev) => setEditDraft({ ...editDraft, taxRate: ev.target.value })}
                                              style={{ width: 60, height: 32 }}
                                            />
                                            %
                                          </span>
                                        )}
                                      </div>
                                    </Field>
                                  )}
                                  <Field label="備註" wide>
                                    <input value={editDraft.note} onChange={(ev) => setEditDraft({ ...editDraft, note: ev.target.value })} />
                                  </Field>
                                </div>
                                <div style={{ display: "flex", gap: 8 }}>
                                  <button type="button" className="wel-btn-primary" onClick={() => saveEditExpense(e.id)}>儲存修改</button>
                                  <button type="button" className="wel-btn-ghost" onClick={cancelEditExpense}>取消</button>
                                </div>
                              </div>
                            </td>
                          </tr>
                        )}
                        </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function SiteLandingPage({
  projectName, setProjectName, editingName, setEditingName, saveName,
  sites, addSite, renameSite, removeSite, siteBreakdown, onEnterSite,
}) {
  const [newSite, setNewSite] = useState({ name: "", address: "", copyFrom: "" });
  const [showAdd, setShowAdd] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState({ name: "", address: "" });
  const statsById = useMemo(() => {
    const m = {};
    siteBreakdown.forEach((s) => (m[s.id] = s));
    return m;
  }, [siteBreakdown]);
  const allSitesTotal = useMemo(
    () => siteBreakdown.reduce((s, b) => s + b.total, 0),
    [siteBreakdown]
  );
  const allSitesIncome = useMemo(
    () => siteBreakdown.reduce((s, b) => s + b.collected, 0),
    [siteBreakdown]
  );
  const allSitesProfit = allSitesIncome - allSitesTotal;
  const allSitesCapital = useMemo(
    () => siteBreakdown.reduce((s, b) => s + b.capital, 0),
    [siteBreakdown]
  );

  const startEdit = (e, s) => { e.stopPropagation(); setEditingId(s.id); setEditDraft({ name: s.name, address: s.address || "" }); };
  const saveEdit = (e) => { e.stopPropagation(); renameSite(editingId, editDraft.name.trim() || "未命名案場", editDraft.address.trim()); setEditingId(null); };
  const cancelEdit = (e) => { e.stopPropagation(); setEditingId(null); };
  const handleRemove = (e, id) => { e.stopPropagation(); removeSite(id); };

  return (
    <div className="wel-landing">
      <div className="wel-landing-header">
        <div className="wel-landing-brand">
          <Droplets size={22} color="var(--teal)" />
          <Zap size={19} color="var(--amber)" style={{ marginLeft: -8 }} />
          {editingName ? (
            <input
              autoFocus
              value={projectName}
              onChange={(e) => setProjectName(e.target.value)}
              onBlur={saveName}
              onKeyDown={(e) => e.key === "Enter" && saveName()}
              className="wel-landing-title-input"
            />
          ) : (
            <h1 className="wel-landing-title" onClick={() => setEditingName(true)} title="點擊編輯專案名稱">{projectName}</h1>
          )}
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div className="wel-eyebrow">案場列表 · SELECT A SITE</div>
          <button type="button" className="wel-btn-ghost" onClick={lockApp}>鎖定</button>
        </div>
      </div>

      <div className="wel-site-grid wel-site-grid-landing">
        {/* All-sites overview card */}
        {sites.length > 1 && (
          <div className="wel-site-card wel-site-card-all" onClick={() => onEnterSite("all")}>
            <div className="wel-site-card-head">
              <Layers size={18} color="var(--amber)" />
              <span className="wel-site-name">全部案場總覽</span>
            </div>
            <div className="wel-site-address muted">查看所有案場的加總與比較</div>
            <div className="wel-site-stats" style={{ display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: "12px 16px" }}>
              <div><span>案場數</span><b>{sites.length}</b></div>
              <div><span>合計收入</span><button type="button" className="wel-num-link" style={{ color: "var(--teal)" }} title="點一下看明細" onClick={(e) => { e.stopPropagation(); onEnterSite("all", { siteKey: "all", kind: "collected" }); }}>{fmtMoney(allSitesIncome)}</button></div>
              <div><span>合計總支出</span><button type="button" className="wel-num-link" style={{ color: "var(--amber)" }} title="點一下看明細" onClick={(e) => { e.stopPropagation(); onEnterSite("all", { siteKey: "all", kind: "total" }); }}>{fmtMoney(allSitesTotal)}</button></div>
              <div><span>總淨損益</span><button type="button" className="wel-num-link" style={{ color: allSitesProfit >= 0 ? "var(--green)" : "var(--red)" }} title="點一下看明細" onClick={(e) => { e.stopPropagation(); onEnterSite("all", { siteKey: "all", kind: "profit" }); }}>{allSitesProfit >= 0 ? "+" : ""}{fmtMoney(allSitesProfit)}</button></div>
              <div><span>合計投入成本（參考）</span><button type="button" className="wel-num-link" style={{ color: "var(--teal)" }} title="點一下看明細" onClick={(e) => { e.stopPropagation(); onEnterSite("all", { siteKey: "all", kind: "capital" }); }}>{fmtMoney(allSitesCapital)}</button></div>
            </div>
            <div className="muted" style={{ fontSize: 11, marginTop: 8 }}>點數字可直接看這個數字是由哪些紀錄組成</div>
          </div>
        )}

        {sites.map((s) => {
          const stat = statsById[s.id] || { material: 0, labor: 0, contract: 0, total: 0, collected: 0, profit: 0, capital: 0, other: 0, asset: 0 };
          const isEditing = editingId === s.id;
          return (
            <div
              key={s.id}
              className="wel-site-card"
              onClick={() => !isEditing && onEnterSite(s.id)}
              role="button"
              tabIndex={0}
            >
              {isEditing ? (
                <div className="wel-site-edit" onClick={(e) => e.stopPropagation()}>
                  <input value={editDraft.name} onChange={(e) => setEditDraft({ ...editDraft, name: e.target.value })} placeholder="案場名稱" autoFocus />
                  <input value={editDraft.address} onChange={(e) => setEditDraft({ ...editDraft, address: e.target.value })} placeholder="地址 / 備註" />
                  <div className="wel-site-edit-actions">
                    <button type="button" className="wel-btn-primary" style={{ padding: "5px 12px", fontSize: 12 }} onClick={saveEdit}>儲存</button>
                    <button type="button" className="wel-btn-ghost" onClick={cancelEdit}>取消</button>
                  </div>
                </div>
              ) : (
                <>
                  <div className="wel-site-card-head">
                    <Building2 size={16} color="var(--teal)" />
                    <span className="wel-site-name">{s.name}</span>
                  </div>
                  {s.address && <div className="wel-site-address muted">{s.address}</div>}
                  <div className="wel-site-stats">
                    <div><span>材料</span><b>{fmtMoney(stat.material)}</b></div>
                    <div><span>人工</span><b>{fmtMoney(stat.labor)}</b></div>
                    <div><span>發包</span><b>{fmtMoney(stat.contract)}</b></div>
                    <div><span>總計</span><b style={{ color: "var(--amber)" }}>{fmtMoney(stat.total)}</b></div>
                  </div>
                  <div className="wel-site-stats">
                    <div><span>甲方收款</span><b>{fmtMoney(stat.collected)}</b></div>
                    <div><span>{stat.profit >= 0 ? "毛利" : "虧損"}</span><b style={{ color: stat.profit >= 0 ? "var(--green)" : "var(--red)" }}>{stat.profit >= 0 ? "+" : ""}{fmtMoney(stat.profit)}</b></div>
                  </div>
                  <div className="wel-site-stats">
                    <div><span>投入成本</span><b style={{ color: "var(--teal)" }}>{fmtMoney(stat.capital)}</b></div>
                    <div><span>其他支出</span><b>{fmtMoney(stat.other)}</b></div>
                    <div><span>固定資產</span><b>{fmtMoney(stat.asset)}</b></div>
                  </div>
                  <div className="wel-site-actions">
                    <span className="wel-site-enter-hint">點擊進入案場 →</span>
                    <button type="button" className="wel-btn-ghost" onClick={(e) => startEdit(e, s)}>編輯</button>
                    <button className="wel-icon-btn" onClick={(e) => handleRemove(e, s.id)}><Trash2 size={14} /></button>
                  </div>
                </>
              )}
            </div>
          );
        })}

        {/* Add new site card */}
        <div className="wel-site-card wel-site-card-add" onClick={() => setShowAdd(true)}>
          {showAdd ? (
            <div className="wel-site-edit" onClick={(e) => e.stopPropagation()}>
              <input placeholder="案場名稱，例：台南安平新建案" value={newSite.name} onChange={(e) => setNewSite({ ...newSite, name: e.target.value })} autoFocus />
              <input placeholder="地址 / 備註（選填）" value={newSite.address} onChange={(e) => setNewSite({ ...newSite, address: e.target.value })} />
              {sites.length > 0 && (
                <select value={newSite.copyFrom} onChange={(e) => setNewSite({ ...newSite, copyFrom: e.target.value })}>
                  <option value="">不複製，從空白開始</option>
                  {sites.map((s) => <option key={s.id} value={s.id}>複製「{s.name}」的發包項目與請款明細</option>)}
                </select>
              )}
              <div className="wel-site-edit-actions">
                <button
                  type="button" className="wel-btn-primary" style={{ padding: "5px 12px", fontSize: 12 }}
                  onClick={() => { addSite(newSite.name, newSite.address, newSite.copyFrom || null); setNewSite({ name: "", address: "", copyFrom: "" }); setShowAdd(false); }}
                >
                  <Plus size={13} /> 新增
                </button>
                <button type="button" className="wel-btn-ghost" onClick={() => setShowAdd(false)}>取消</button>
              </div>
            </div>
          ) : (
            <div className="wel-site-add-inner">
              <Plus size={26} color="var(--text-muted)" />
              <span>新增案場</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function Field({ label, children, wide }) {
  return (
    <label className="wel-field" style={wide ? { gridColumn: "1 / -1" } : undefined}>
      <span>{label}</span>
      {children}
    </label>
  );
}

function Empty({ text }) {
  return (
    <div className="wel-empty">
      <Package size={22} style={{ opacity: 0.4 }} />
      <span>{text}</span>
    </div>
  );
}

function StyleBlock() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

      .wel-app {
        --bg: #1C2530;
        --surface: #212B37;
        --surface2: #2A3644;
        --border: #37424F;
        --amber: #E8A33D;
        --teal: #4A90A4;
        --green: #6B9B6E;
        --red: #C1543C;
        --text: #E8E4DA;
        --text-muted: #8B95A1;
        --font-display: 'Oswald', sans-serif;
        --font-body: 'IBM Plex Sans', sans-serif;
        --font-mono: 'IBM Plex Mono', monospace;
        background: var(--bg);
        color: var(--text);
        font-family: var(--font-body);
        border-radius: 10px;
        background-image:
          linear-gradient(rgba(255,255,255,0.025) 1px, transparent 1px),
          linear-gradient(90deg, rgba(255,255,255,0.025) 1px, transparent 1px);
        background-size: 28px 28px;
      }
      .wel-app * { box-sizing: border-box; }
      .wel-toast {
        position: sticky; top: 10px; z-index: 50; margin: 0 auto 0; width: fit-content; max-width: 90%;
        display: flex; align-items: center; gap: 8px; background: #2A1E16; border: 1px solid var(--amber);
        color: #F2D9AE; padding: 10px 16px; border-radius: 8px; font-size: 13px; font-family: var(--font-body);
        box-shadow: 0 4px 16px rgba(0,0,0,0.4); cursor: pointer;
      }
      .wel-spin { animation: wel-spin 1s linear infinite; }
      @keyframes wel-spin { to { transform: rotate(360deg); } }

      .wel-shell { display: flex; min-height: 640px; }

      .wel-sidebar {
        width: 208px; flex-shrink: 0; background: var(--surface);
        border-right: 1px solid var(--border); display: flex; flex-direction: column;
        padding: 18px 12px;
      }
      .wel-brand { display: flex; align-items: center; gap: 4px; padding: 0 6px 18px; border-bottom: 1px solid var(--border); margin-bottom: 14px; }
      .wel-brand-text { font-family: var(--font-display); font-weight: 600; font-size: 15px; letter-spacing: 0.02em; margin-left: 6px; cursor: text; }
      .wel-brand-input { font-family: var(--font-display); font-weight: 600; font-size: 15px; margin-left: 6px; background: var(--surface2); border: 1px solid var(--teal); border-radius: 4px; color: var(--text); padding: 2px 6px; width: 100%; }
      .wel-back-to-sites {
        display: flex; flex-direction: column; align-items: flex-start; gap: 2px; width: 100%;
        background: var(--surface2); border: 1px solid var(--border); border-radius: 8px; padding: 8px 10px 8px 8px;
        margin-bottom: 10px; cursor: pointer; text-align: left;
      }
      .wel-back-to-sites:hover { border-color: var(--teal); }
      .wel-back-to-sites svg { margin-bottom: 2px; color: var(--text-muted); }
      .wel-back-current { font-family: var(--font-body); font-size: 13px; font-weight: 600; color: var(--amber); }
      .wel-back-label { font-family: var(--font-mono); font-size: 10px; color: var(--text-muted); letter-spacing: 0.04em; }

      .wel-landing { padding: 40px 44px 56px; max-width: 1180px; margin: 0 auto; }
      .wel-landing-header { margin-bottom: 26px; }
      .wel-landing-brand { display: flex; align-items: center; gap: 4px; margin-bottom: 6px; }
      .wel-landing-title { font-family: var(--font-display); font-size: 28px; font-weight: 600; margin: 0 0 0 6px; cursor: text; letter-spacing: 0.01em; }
      .wel-landing-title-input { font-family: var(--font-display); font-size: 28px; font-weight: 600; margin-left: 6px; background: var(--surface2); border: 1px solid var(--teal); border-radius: 6px; color: var(--text); padding: 2px 10px; }
      .wel-site-grid-landing { margin-top: 6px; }
      .wel-site-card-all { border-color: rgba(232,163,61,0.4); background: linear-gradient(135deg, rgba(232,163,61,0.08), var(--surface)); }
      .wel-site-card-add {
        display: flex; align-items: center; justify-content: center; min-height: 148px;
        border: 1px dashed var(--border); background: transparent;
      }
      .wel-site-add-inner { display: flex; flex-direction: column; align-items: center; gap: 8px; color: var(--text-muted); font-size: 13px; }
      .wel-site-enter-hint { font-size: 11px; color: var(--teal); font-family: var(--font-mono); flex: 1; }
      .wel-notice { background: rgba(74,144,164,0.1); border: 1px solid rgba(74,144,164,0.4); color: #B9DCE6; font-size: 12.5px; padding: 10px 14px; border-radius: 8px; font-family: var(--font-body); }
      .wel-notice-action { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; justify-content: space-between; }
      .wel-notice-action select { background: var(--surface2); border: 1px solid var(--teal); color: var(--amber); border-radius: 6px; padding: 6px 9px; font-size: 12.5px; font-family: var(--font-body); font-weight: 500; flex-shrink: 0; }
      .wel-nav { display: flex; flex-direction: column; gap: 3px; flex: 1; }
      .wel-navbtn {
        display: flex; align-items: center; gap: 10px; padding: 9px 10px; border-radius: 6px;
        background: transparent; border: 1px solid transparent; color: var(--text-muted);
        font-family: var(--font-body); font-size: 13.5px; cursor: pointer; text-align: left; width: 100%;
      }
      .wel-navbtn:hover { background: var(--surface2); color: var(--text); }
      .wel-navbtn.active { background: var(--surface2); color: var(--amber); border-color: var(--border); font-weight: 500; }
      .wel-sidebar-foot { padding-top: 12px; border-top: 1px solid var(--border); }
      .wel-saving { font-size: 11px; color: var(--text-muted); display: flex; align-items: center; gap: 5px; font-family: var(--font-mono); }

      .wel-main { flex: 1; min-width: 0; padding: 26px 30px 40px; overflow-y: auto; max-height: 900px; }
      .wel-page { display: flex; flex-direction: column; gap: 18px; }
      .wel-page-head { display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 12px; }
      .wel-export-bar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
      .wel-export-btn {
        display: inline-flex; align-items: center; gap: 5px; background: var(--surface2); border: 1px solid var(--border);
        color: var(--text-muted); border-radius: 6px; padding: 6px 11px; font-size: 12px; font-family: var(--font-body); cursor: pointer;
      }
      .wel-export-btn:hover { color: var(--teal); border-color: var(--teal); background: rgba(74,144,164,0.08); }
      .wel-eyebrow { font-family: var(--font-mono); font-size: 11px; letter-spacing: 0.12em; color: var(--teal); margin-bottom: 4px; }
      .wel-h1 { font-family: var(--font-display); font-size: 26px; font-weight: 600; margin: 0; letter-spacing: 0.01em; }
      .wel-filter select { background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 7px 10px; font-family: var(--font-mono); font-size: 12.5px; }

      .wel-meter-card { background: #12181F; border: 1px solid var(--border); border-radius: 10px; padding: 20px 22px; box-shadow: inset 0 0 0 1px rgba(255,255,255,0.02); }
      .wel-meter-label { display: flex; align-items: center; gap: 6px; font-family: var(--font-mono); font-size: 11.5px; color: var(--text-muted); letter-spacing: 0.04em; margin-bottom: 10px; }
      .wel-meter { display: flex; align-items: baseline; gap: 3px; }
      .wel-meter-digit {
        font-family: var(--font-mono); font-weight: 600; font-size: 34px; color: var(--amber);
        background: #0B0F14; border: 1px solid #3A3226; border-radius: 4px; padding: 4px 6px 2px;
        min-width: 26px; text-align: center; text-shadow: 0 0 10px rgba(232,163,61,0.35);
      }
      .wel-meter-unit { font-family: var(--font-body); font-size: 14px; color: var(--text-muted); margin-left: 8px; }
      .wel-meter-sub { display: flex; gap: 22px; margin-top: 12px; font-size: 12.5px; color: var(--text-muted); font-family: var(--font-body); }
      .wel-meter-sub b { color: var(--text); font-family: var(--font-mono); font-weight: 500; }

      .wel-profit-card {
        display: flex; align-items: center; gap: 16px; flex-wrap: wrap;
        background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 16px 20px;
      }
      .wel-profit-card.positive { border-color: rgba(107,155,110,0.4); }
      .wel-profit-card.negative { border-color: rgba(193,84,60,0.4); }
      .wel-profit-col { display: flex; flex-direction: column; gap: 5px; }
      .wel-profit-label { display: flex; align-items: center; gap: 5px; font-size: 11px; color: var(--text-muted); font-family: var(--font-mono); text-transform: uppercase; letter-spacing: 0.03em; }
      .wel-profit-col b { font-size: 19px; font-weight: 600; }
      .wel-profit-sign { font-size: 18px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-profit-value { color: var(--green); }
      .wel-profit-card.negative .wel-profit-value { color: var(--red); }
      .wel-site-profit-list { display: flex; flex-direction: column; gap: 6px; margin-top: 12px; padding-top: 12px; border-top: 1px solid var(--border); }
      .wel-site-profit-row { display: flex; align-items: center; gap: 12px; font-size: 12px; }
      .wel-site-profit-row > *:first-child { flex: 1; }
      .wel-profit-pos { color: var(--green); }
      .wel-profit-neg { color: var(--red); }

      .wel-grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
      @media (max-width: 860px) { .wel-grid2 { grid-template-columns: 1fr; } .wel-shell { flex-direction: column; min-height: 0; } .wel-sidebar { width: 100%; flex-direction: row; align-items: center; padding: 10px 14px; flex-wrap: wrap; } .wel-nav { flex-direction: row; flex-wrap: wrap; } .wel-navbtn { width: auto; } .wel-sidebar-foot { display: none; } .wel-brand { border-bottom: none; margin-bottom: 0; padding-bottom: 0; } .wel-back-to-sites { width: auto; margin-bottom: 0; flex-direction: row; align-items: center; gap: 6px; } .wel-back-label { display: none; } .wel-landing { padding: 24px 18px 40px; } .wel-landing-title { font-size: 22px; } .wel-main { max-height: none; overflow-y: visible; padding: 18px 14px 32px; } .wel-template-rows { overflow-x: auto; padding-bottom: 4px; } .wel-template-row, .wel-template-row-3, .wel-template-row-4, .wel-template-row-5, .wel-template-row-item { min-width: 480px; } .wel-template-row-item { min-width: 640px; } }
      html, body { overflow-x: hidden; max-width: 100%; }
      .wel-app { overflow-x: hidden; max-width: 100vw; }

      .wel-card { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 18px 18px 16px; }
      .wel-card-title { display: flex; align-items: center; gap: 6px; font-family: var(--font-display); font-size: 14px; font-weight: 500; letter-spacing: 0.02em; color: var(--text); margin-bottom: 14px; text-transform: uppercase; }
      .wel-card-title-row { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 14px; }
      .wel-card-title-row .wel-card-title { margin-bottom: 0; }

      .wel-pipes { display: flex; flex-direction: column; gap: 11px; }
      .wel-pipe-row { display: grid; grid-template-columns: 110px 1fr 88px; align-items: center; gap: 10px; }
      .wel-pipe-label { font-size: 12.5px; color: var(--text-muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .wel-pipe-track { background: #171F27; border-radius: 20px; height: 14px; overflow: hidden; border: 1px solid var(--border); }
      .wel-pipe-fill { height: 100%; border-radius: 20px; transition: width 0.4s ease; }
      .wel-pipe-value { font-family: var(--font-mono); font-size: 12px; text-align: right; color: var(--text); }

      .wel-alert { border-color: rgba(193,84,60,0.5); background: rgba(193,84,60,0.06); }
      .wel-alert-info { border-color: rgba(74,144,164,0.5); background: rgba(74,144,164,0.06); }
      .wel-alert-list { display: flex; flex-wrap: wrap; gap: 8px; }
      .wel-chip-info { background: rgba(74,144,164,0.15); border-color: rgba(74,144,164,0.4); color: #B9DCE6; }

      .wel-discussion-list { display: flex; flex-direction: column; gap: 10px; }
      .wel-discussion-card { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 12px 16px; display: flex; flex-direction: column; gap: 8px; }
      .wel-discussion-card.resolved { border-color: rgba(107,155,110,0.4); }
      .wel-discussion-head { display: flex; align-items: center; gap: 10px; }
      .wel-status-btn { background: transparent; border: none; cursor: pointer; padding: 0; display: inline-flex; flex-shrink: 0; }
      .wel-discussion-topic { flex: 1; font-size: 13.5px; font-weight: 500; }
      .wel-discussion-note { font-size: 12.5px; padding-left: 27px; }
      .wel-discussion-result { display: flex; align-items: center; gap: 8px; padding-left: 27px; }
      .wel-discussion-result input { flex: 1; min-width: 0; background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 9px; font-size: 12.5px; font-family: var(--font-body); }
      .wel-discussion-result input:disabled { opacity: 0.5; cursor: not-allowed; }
      .wel-chip { background: rgba(193,84,60,0.15); border: 1px solid rgba(193,84,60,0.4); color: #E8B4A8; font-size: 12px; padding: 4px 9px; border-radius: 20px; font-family: var(--font-mono); }
      .wel-item-chips { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin-top: 8px; }
      .wel-chip-btn {
        background: rgba(74,144,164,0.12); border: 1px solid rgba(74,144,164,0.4); color: #B9DCE6;
        font-size: 11.5px; padding: 4px 10px; border-radius: 20px; font-family: var(--font-body); cursor: pointer;
      }
      .wel-chip-btn:hover { background: rgba(74,144,164,0.22); border-color: var(--teal); }
      .wel-chip-btn .muted { color: #7FA3AC; }

      .wel-material-manager { margin-top: 10px; padding: 10px 12px; background: var(--surface2); border: 1px solid var(--border); border-radius: 8px; display: flex; flex-direction: column; gap: 10px; }
      .wel-material-add-row { display: flex; gap: 8px; }
      .wel-material-add-row { flex-wrap: wrap; }
      .wel-material-add-row input { background: var(--surface); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 9px; font-size: 12.5px; font-family: var(--font-body); flex: 1; min-width: 90px; }
      .wel-material-chips { display: flex; flex-wrap: wrap; gap: 6px; }
      .wel-chip-material { display: inline-flex; align-items: center; gap: 5px; background: rgba(232,163,61,0.12); border: 1px solid rgba(232,163,61,0.35); color: #E8D2A8; }
      .wel-chip-material button { background: transparent; border: none; color: inherit; cursor: pointer; display: inline-flex; opacity: 0.7; padding: 0; }
      .wel-chip-material button:hover { opacity: 1; }
      .wel-manager-list { display: flex; flex-direction: column; gap: 6px; max-height: 260px; overflow-y: auto; }
      .wel-manager-row { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
      .wel-manager-name { flex: 1; min-width: 70px; background: var(--surface); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 8px; font-size: 12.5px; font-family: var(--font-body); }
      .wel-num-link { background: none; border: none; border-bottom: 1px dashed var(--text-muted); color: inherit; font: inherit; cursor: pointer; padding: 0; text-align: inherit; }
      .wel-num-link:hover, .wel-num-link.active { color: var(--teal) !important; border-bottom-color: var(--teal); }
      .wel-site-stats .wel-num-link { font-size: 13px; font-family: var(--font-mono); font-weight: 600; text-align: left; align-self: flex-start; }
      .wel-stat-btn { text-align: left; cursor: pointer; color: inherit; font: inherit; }
      .wel-stat-btn:hover { border-color: var(--teal); }
      .wel-scan-input { background: var(--surface); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 8px; font-size: 13px; min-width: 0; box-sizing: border-box; }
      .wel-scan-name { width: 100%; min-width: 210px; }
      .wel-manager-mini { width: 74px; flex-shrink: 0; background: var(--surface); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 8px; font-size: 12.5px; font-family: var(--font-mono); }

      .wel-form-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px 14px; margin-bottom: 14px; }
      @media (max-width: 700px) { .wel-form-grid { grid-template-columns: 1fr 1fr; } .wel-form-grid-btn { grid-template-columns: 1fr !important; } }
      @media (max-width: 480px) { .wel-form-grid { grid-template-columns: 1fr !important; } }
      .wel-field { display: flex; flex-direction: column; gap: 5px; font-size: 11.5px; color: var(--text-muted); font-family: var(--font-mono); min-width: 0; }
      .wel-field input, .wel-field select { background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 0 9px; font-family: var(--font-body); font-size: 13.5px; width: 100%; min-width: 0; height: 40px; box-sizing: border-box; }
      .wel-field textarea { background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 9px; font-family: var(--font-body); font-size: 13.5px; width: 100%; min-width: 0; box-sizing: border-box; line-height: 1.55; resize: vertical; }
      .wel-field input:focus, .wel-field select:focus, .wel-field textarea:focus { outline: none; border-color: var(--teal); }
      .wel-amount-preview { font-family: var(--font-mono); font-weight: 600; color: var(--amber); padding: 8px 0; font-size: 14px; }

      .wel-btn-primary {
        display: inline-flex; align-items: center; gap: 6px; background: var(--amber); color: #1C2530;
        border: none; border-radius: 6px; padding: 9px 16px; font-weight: 600; font-size: 13px;
        font-family: var(--font-body); cursor: pointer;
      }
      .wel-btn-primary:hover { filter: brightness(1.08); }

      .wel-table { width: 100%; border-collapse: collapse; font-size: 13px; }
      .wel-table th { text-align: left; font-family: var(--font-mono); font-size: 11px; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.04em; padding: 10px 14px; border-bottom: 1px solid var(--border); background: var(--surface2); }
      .wel-table td { padding: 9px 14px; border-bottom: 1px solid var(--border); }
      .wel-table tr:last-child td { border-bottom: none; }
      .wel-table .right { text-align: right; }
      .wel-table .mono { font-family: var(--font-mono); }
      .wel-table .muted { color: var(--text-muted); }
      .wel-row-clickable { cursor: pointer; }
      .wel-row-clickable:hover { background: var(--surface2); }
      .wel-table .strong { font-weight: 600; color: var(--amber); }
      .wel-low { color: var(--red) !important; }
      .wel-tag { border: 1px solid var(--border); border-radius: 20px; padding: 2px 9px; font-size: 11.5px; white-space: nowrap; }
      .wel-icon-btn { background: transparent; border: none; color: var(--text-muted); cursor: pointer; padding: 5px; border-radius: 4px; display: inline-flex; }
      .wel-icon-btn:hover { color: var(--red); background: rgba(193,84,60,0.12); }

      .wel-empty { display: flex; flex-direction: column; align-items: center; gap: 8px; padding: 34px 0; color: var(--text-muted); font-size: 13px; }

      .wel-inline-rate {
        width: 104px; flex-shrink: 0; background: var(--surface2); border: 1px solid var(--border); color: var(--amber);
        font-family: var(--font-mono); font-weight: 600; border-radius: 5px; padding: 5px 7px; font-size: 12.5px; text-align: right;
      }
      .wel-inline-rate:focus { outline: none; border-color: var(--teal); }

      .wel-range-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin-bottom: 16px; padding-bottom: 14px; border-bottom: 1px solid var(--border); }
      .wel-range-bar label { display: flex; align-items: center; gap: 6px; font-family: var(--font-mono); font-size: 11.5px; color: var(--text-muted); }
      .wel-range-bar input[type="date"] { background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 8px; font-family: var(--font-body); font-size: 12.5px; }
      .wel-btn-ghost { background: transparent; border: 1px solid var(--border); color: var(--text-muted); border-radius: 6px; padding: 6px 11px; font-size: 12px; font-family: var(--font-body); cursor: pointer; }
      .wel-btn-ghost:hover { color: var(--text); border-color: var(--teal); }
      .wel-range-label { font-size: 11.5px; margin-left: auto; font-family: var(--font-mono); }

      .wel-labor-summary { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }
      @media (max-width: 700px) { .wel-labor-summary { grid-template-columns: 1fr; } }
      .wel-labor-stat { background: var(--surface2); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; display: flex; flex-direction: column; gap: 6px; }
      .wel-labor-stat span { font-size: 11.5px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-labor-stat b { font-size: 18px; font-family: var(--font-mono); font-weight: 600; }

      .wel-cat-list { display: flex; flex-direction: column; gap: 8px; }
      .wel-cat-item { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 10px 14px; }
      .wel-timeline-wrap { background: #171F27; border: 1px solid var(--border); border-top: none; border-radius: 0 0 8px 8px; padding: 10px 14px; margin: -6px 0 6px; max-height: 340px; overflow-y: auto; }
      .wel-timeline .wel-table { font-size: 12px; }
      .wel-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
      .wel-cat-name { flex: 1; min-width: 60px; font-size: 13.5px; }

      .wel-site-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 14px; }
      .wel-site-card { background: var(--surface); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; display: flex; flex-direction: column; gap: 10px; cursor: pointer; transition: border-color 0.15s, transform 0.1s; }
      .wel-site-card:hover { border-color: var(--teal); }
      .wel-site-card:active { transform: scale(0.99); }
      .wel-site-card.active { border-color: var(--amber); box-shadow: 0 0 0 1px rgba(232,163,61,0.25); }
      .wel-site-card-head { display: flex; align-items: center; gap: 7px; }
      .wel-site-name { font-family: var(--font-display); font-size: 14.5px; font-weight: 500; flex: 1; }
      .wel-site-badge { font-size: 10px; font-family: var(--font-mono); background: rgba(232,163,61,0.15); color: var(--amber); border: 1px solid rgba(232,163,61,0.4); padding: 2px 7px; border-radius: 20px; white-space: nowrap; }
      .wel-site-address { font-size: 12px; margin-top: -4px; }
      .wel-site-stats { display: flex; flex-wrap: wrap; gap: 8px 14px; padding: 8px 0; border-top: 1px solid var(--border); border-bottom: 1px solid var(--border); }
      .wel-site-stats div { display: flex; flex-direction: column; gap: 3px; }
      .wel-site-stats span { font-size: 10.5px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-site-stats b, .wel-site-stats .wel-num-link { white-space: nowrap; }
      .wel-site-stats b { font-size: 13px; font-family: var(--font-mono); font-weight: 600; }
      .wel-site-actions { display: flex; align-items: center; gap: 6px; }
      .wel-site-actions .wel-btn-ghost { padding: 5px 9px; font-size: 11.5px; }
      .wel-site-actions .wel-btn-ghost:disabled { opacity: 0.4; cursor: default; }
      .wel-site-edit { display: flex; flex-direction: column; gap: 8px; }
      .wel-site-edit input, .wel-site-edit select { background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 7px 9px; font-size: 13px; font-family: var(--font-body); }
      .wel-site-edit-actions { display: flex; gap: 8px; }

      .wel-vendor-group { margin-bottom: 18px; padding-bottom: 6px; border-bottom: 1px solid var(--border); }
      .wel-vendor-group:last-child { border-bottom: none; margin-bottom: 0; }
      .wel-vendor-head { display: flex; align-items: center; gap: 8px; padding: 8px 10px; background: var(--surface2); border-radius: 7px; margin-bottom: 10px; }
      .wel-vendor-head-btn { width: 100%; border: none; cursor: pointer; text-align: left; color: var(--text); font-family: var(--font-body); }
      .wel-vendor-head-btn:hover { background: var(--surface); }
      .wel-vendor-name { font-family: var(--font-display); font-size: 14px; font-weight: 600; flex: 1; }
      .wel-contract-row { display: flex; flex-direction: column; gap: 6px; }
      .wel-contract-head { display: flex; align-items: center; gap: 10px; }
      .wel-contract-name { font-size: 13.5px; font-weight: 500; flex: 1; }
      .wel-contract-amt { font-family: var(--font-mono); font-size: 12.5px; color: var(--amber); white-space: nowrap; }
      .wel-expand-btn { background: transparent; border: none; color: var(--text-muted); cursor: pointer; padding: 2px; display: inline-flex; }
      .wel-expand-btn:hover { color: var(--teal); }
      .wel-floor-summary-line { font-size: 11.5px; padding-left: 2px; }
      .wel-mismatch { color: var(--red); }

      .wel-floor-panel { display: flex; flex-direction: column; gap: 14px; background: #171F27; border: 1px solid var(--border); border-radius: 8px; padding: 14px 16px; margin: 4px 0 8px; }
      .wel-floor-block { display: flex; flex-direction: column; gap: 8px; }
      .wel-floor-block-title { display: flex; align-items: center; gap: 6px; font-family: var(--font-mono); font-size: 11px; letter-spacing: 0.04em; color: var(--teal); text-transform: uppercase; }
      .wel-template-rows { display: flex; flex-direction: column; gap: 6px; }
      .wel-template-row { display: grid; grid-template-columns: 1fr 110px auto; gap: 8px; align-items: center; }
      .wel-template-row-3 { grid-template-columns: 1fr 100px 1fr auto; }
      .wel-template-row-4 { grid-template-columns: 130px 1fr 100px 1fr auto; }
      .wel-template-row-5 { grid-template-columns: 120px 1fr 64px 90px 1fr auto; }
      .wel-template-row-item { grid-template-columns: 118px 1fr 60px 90px 1.2fr auto auto; }

      .wel-item-name-group { margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid var(--border); }
      .wel-item-name-group:last-of-type { border-bottom: none; margin-bottom: 6px; }
      .wel-item-name-head { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }
      .wel-item-name-input {
        font-family: var(--font-display); font-size: 13.5px; font-weight: 600; color: var(--text); flex: 1;
        background: transparent; border: 1px solid transparent; border-radius: 6px; padding: 3px 6px;
      }
      .wel-item-name-input:hover, .wel-item-name-input:focus { background: var(--surface2); border-color: var(--border); outline: none; }
      .wel-item-floor-group { margin-bottom: 8px; padding-left: 4px; }
      .wel-item-floor-group:last-child { margin-bottom: 0; }
      .wel-item-group-label {
        display: inline-flex; align-items: center; gap: 3px; font-family: var(--font-mono); font-size: 11px; font-weight: 600;
        color: var(--teal); background: rgba(74,144,164,0.1); border: 1px solid rgba(74,144,164,0.3);
        padding: 2px 9px; border-radius: 20px; margin-bottom: 6px;
      }
      .wel-item-group-toggle { cursor: pointer; }
      .wel-item-group-toggle-row { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; }
      .wel-item-group-toggle-row .wel-item-group-label { margin-bottom: 0; }
      .wel-item-group-toggle:hover { background: rgba(74,144,164,0.2); }

      .wel-pay-import-btn {
        background: rgba(232,163,61,0.1); border: 1px solid rgba(232,163,61,0.35); color: var(--amber);
        border-radius: 6px; padding: 5px 7px; cursor: pointer; display: inline-flex; align-items: center;
      }
      .wel-pay-import-btn:hover { background: rgba(232,163,61,0.22); }
      .wel-pay-import-btn.imported { background: rgba(107,155,110,0.15); border-color: rgba(107,155,110,0.4); color: var(--green); }

      .wel-floor-batch-advanced { font-size: 11.5px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-floor-batch-advanced summary { cursor: pointer; padding: 2px 0; }
      .wel-floor-batch-advanced summary:hover { color: var(--teal); }
      .wel-floor-batch-advanced[open] summary { margin-bottom: 6px; }

      .wel-new-item-block { display: flex; flex-direction: column; gap: 10px; padding: 12px; margin-top: 8px; background: var(--surface2); border: 1px dashed var(--border); border-radius: 8px; }
      .wel-new-item-row { display: flex; flex-wrap: wrap; gap: 10px; align-items: flex-end; }
      .wel-mini-label { display: flex; flex-direction: column; gap: 4px; font-size: 10.5px; color: var(--text-muted); font-family: var(--font-mono); text-transform: uppercase; letter-spacing: 0.03em; }
      .wel-mini-label input, .wel-mini-label select {
        background: var(--surface); border: 1px solid var(--border); color: var(--text); border-radius: 6px;
        padding: 6px 8px; font-size: 12.5px; font-family: var(--font-body); font-weight: 500;
      }
      .wel-no-spinner::-webkit-outer-spin-button, .wel-no-spinner::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
      .wel-no-spinner { -moz-appearance: textfield; appearance: textfield; }
      .wel-new-type-row { display: flex; gap: 8px; align-items: center; }
      .wel-new-type-row input { flex: 1; min-width: 0; background: var(--surface); border: 1px solid var(--teal); color: var(--text); border-radius: 6px; padding: 6px 9px; font-size: 12.5px; }
      .wel-item-floor-apply { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; padding-top: 8px; border-top: 1px dashed var(--border); }
      .wel-item-floor-apply label { display: flex; align-items: center; gap: 5px; font-size: 11.5px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-item-floor-apply input { width: 52px; background: var(--surface); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 5px 7px; font-size: 12px; }
      .wel-icon-btn-add { color: var(--teal); }
      .wel-icon-btn-add:hover { background: rgba(74,144,164,0.15); color: var(--teal); }
      .wel-template-row input, .wel-floor-item-row input, .wel-quickadd-row input, .wel-apply-row input {
        background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 8px; font-size: 12.5px; font-family: var(--font-body);
      }
      .wel-apply-row { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; }
      .wel-apply-row label { display: flex; align-items: center; gap: 5px; font-size: 12px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-apply-row input { width: 56px; }
      .wel-checkbox-label { flex-direction: row-reverse; }
      .wel-checkbox-label input { width: auto; }
      .wel-floor-groups { display: flex; flex-direction: column; gap: 10px; }
      .wel-floor-group { background: var(--surface); border: 1px solid var(--border); border-radius: 7px; padding: 9px 11px; display: flex; flex-direction: column; gap: 6px; }
      .wel-floor-group-head { display: flex; align-items: center; gap: 8px; }
      .wel-floor-label { font-family: var(--font-display); font-weight: 500; font-size: 13px; flex: 1; }
      .wel-floor-item-row { display: grid; grid-template-columns: 1fr 100px auto; gap: 8px; align-items: center; }
      .wel-quickadd-row { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; padding-top: 4px; border-top: 1px dashed var(--border); }

      .wel-attendance-row { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; }
      .wel-attendance-row input { background: var(--surface2); border: 1px solid var(--border); color: var(--text); border-radius: 6px; padding: 6px 8px; font-size: 12.5px; font-family: var(--font-body); }
      .wel-attendance-row label { display: flex; align-items: center; gap: 5px; font-size: 12px; color: var(--text-muted); font-family: var(--font-mono); }
      .wel-attendance-list { display: flex; flex-direction: column; gap: 5px; margin-top: 8px; }
      .wel-attendance-item { display: flex; align-items: center; gap: 10px; font-size: 12px; padding: 5px 2px; border-bottom: 1px solid var(--border); }

      .wel-summary-note { font-family: var(--font-mono); font-size: 12px; margin-top: -6px; }
      .wel-summary-section { display: flex; flex-direction: column; gap: 10px; }
      .wel-summary-heading { display: flex; align-items: center; gap: 8px; font-family: var(--font-display); font-size: 15px; font-weight: 500; }
      .wel-summary-index {
        width: 22px; height: 22px; border-radius: 50%; background: var(--surface2); border: 1px solid var(--border);
        display: inline-flex; align-items: center; justify-content: center; font-family: var(--font-mono); font-size: 12px; color: var(--text-muted);
      }
      .wel-summary-total { margin-left: auto; font-size: 15px; color: var(--amber); }
      .wel-summary-rows { display: flex; flex-direction: column; gap: 6px; }
      .wel-summary-row { display: flex; align-items: center; gap: 10px; font-size: 12.5px; padding: 5px 2px; border-bottom: 1px solid var(--border); }
    `}</style>
  );
}
