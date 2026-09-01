import React, { useState, useEffect, useMemo } from "react";
import { HardHat, CheckCircle2, Clock } from "lucide-react";
import { storageGet, storageSet, storageGetOrThrow } from "./storage";

const DEFAULT_SITE = { id: "site-default", name: "預設案場", address: "" };
const uid = () => Math.random().toString(36).slice(2, 10);
const todayStr = () => new Date().toISOString().slice(0, 10);

function toMinutes(hhmm) {
  if (!hhmm) return null;
  const [h, m] = hhmm.split(":").map(Number);
  return h * 60 + m;
}

// 標準工時 08:00-17:00（含 12:00-13:00 休息 1 小時）＝ 8 小時 = 1 天
// 時數 >= 8 小時：1 天，超過的部分算加班
// 時數 >= 4 小時：0.5 天
// 時數 < 4 小時：0 天（不計入出工天數，但仍記錄時間供查詢）
function computeAttendance(startStr, endStr) {
  const start = toMinutes(startStr);
  const end = toMinutes(endStr);
  if (start == null || end == null || end <= start) return null;
  let minutes = end - start;
  const lunchStart = 12 * 60, lunchEnd = 13 * 60;
  const overlapStart = Math.max(start, lunchStart);
  const overlapEnd = Math.min(end, lunchEnd);
  if (overlapEnd > overlapStart) minutes -= (overlapEnd - overlapStart);
  const hours = minutes / 60;
  const days = hours >= 8 ? 1 : hours >= 4 ? 0.5 : 0;
  const overtimeHours = Math.max(0, Math.round((hours - 8) * 10) / 10);
  return { hours: Math.round(hours * 10) / 10, days, overtimeHours };
}

export default function CheckinPage() {
  const [loading, setLoading] = useState(true);
  const [workers, setWorkers] = useState([]);
  const [site, setSite] = useState(DEFAULT_SITE);
  const [saving, setSaving] = useState(false);
  const [done, setDone] = useState(null);
  const [form, setForm] = useState({
    workerId: "",
    date: todayStr(),
    start: "08:00",
    end: "17:00",
  });

  useEffect(() => {
    (async () => {
      const [wk, sites] = await Promise.all([
        storageGet("workers", []),
        storageGet("sites", [DEFAULT_SITE]),
      ]);
      setWorkers(wk);
      const params = new URLSearchParams(window.location.search);
      const siteId = params.get("site");
      const matched = (sites && sites.length ? sites : [DEFAULT_SITE]).find((s) => s.id === siteId);
      setSite(matched || (sites && sites[0]) || DEFAULT_SITE);
      setLoading(false);
    })();
  }, []);

  const result = useMemo(() => computeAttendance(form.start, form.end), [form.start, form.end]);
  const selectedWorker = workers.find((w) => w.id === form.workerId);

  const [error, setError] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    if (!form.workerId) return;
    if (!result || result.days <= 0) return;
    setSaving(true);
    setError("");
    try {
      // Read the current list fresh right before writing (this page has no
      // pre-loaded state to fall back on). If this read fails, we must NOT
      // fall back to an empty list — that would overwrite everyone else's
      // existing worklogs with just this one new record.
      const list = (await storageGetOrThrow("worklogs")) || [];
      const rec = {
        id: uid(),
        workerId: form.workerId,
        date: form.date,
        days: result.days,
        note: `打卡 ${form.start}-${form.end}${result.overtimeHours > 0 ? `，加班 ${result.overtimeHours} 小時` : ""}`,
        siteId: site.id,
        startTime: form.start,
        endTime: form.end,
        overtimeHours: result.overtimeHours,
      };
      const ok = await storageSet("worklogs", [rec, ...list]);
      if (!ok) throw new Error("write failed");
      setDone({ workerName: selectedWorker?.name, ...result, date: form.date });
      setForm({ workerId: "", date: todayStr(), start: "08:00", end: "17:00" });
    } catch (err) {
      console.error("checkin submit failed", err);
      setError("送出失敗，請檢查網路後再試一次（資料還沒有記錄）");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="ckn-app">
      <style>{`
        .ckn-app {
          min-height: 100vh; background: #1C2530; color: #E8E4DA;
          font-family: 'IBM Plex Sans', sans-serif; display: flex; justify-content: center;
          padding: 32px 18px 60px;
        }
        .ckn-app * { box-sizing: border-box; }
        .ckn-card { width: 100%; max-width: 420px; }
        .ckn-title { display: flex; align-items: center; gap: 8px; font-size: 20px; font-weight: 600; margin-bottom: 4px; }
        .ckn-site { color: #8B95A1; font-size: 13px; margin-bottom: 24px; }
        .ckn-field { display: flex; flex-direction: column; gap: 6px; margin-bottom: 16px; }
        .ckn-field label { font-size: 13px; color: #8B95A1; }
        .ckn-field input, .ckn-field select {
          background: #232E3A; border: 1px solid #3A4453; color: #E8E4DA; border-radius: 8px;
          padding: 0 12px; height: 48px; font-size: 15px; width: 100%;
        }
        .ckn-row { display: flex; gap: 12px; }
        .ckn-row > .ckn-field { flex: 1; min-width: 0; }
        .ckn-preview {
          background: #232E3A; border: 1px solid #3A4453; border-radius: 10px; padding: 14px 16px;
          margin-bottom: 20px; font-size: 13.5px; display: flex; flex-direction: column; gap: 6px;
        }
        .ckn-preview b { color: #E8A33D; font-family: monospace; }
        .ckn-warn { color: #C1543C; }
        .ckn-btn {
          width: 100%; background: #E8A33D; color: #1C2530; border: none; border-radius: 8px;
          padding: 14px; font-size: 16px; font-weight: 600; cursor: pointer;
        }
        .ckn-btn:disabled { opacity: 0.5; }
        .ckn-done {
          background: rgba(107,155,110,0.12); border: 1px solid rgba(107,155,110,0.4); border-radius: 10px;
          padding: 16px; margin-bottom: 20px; font-size: 14px; line-height: 1.7;
        }
        .ckn-loading { padding: 60px 0; text-align: center; color: #8B95A1; }
      `}</style>
      <div className="ckn-card">
        <div className="ckn-title"><HardHat size={22} /> 出工打卡</div>
        <div className="ckn-site">{loading ? "讀取中…" : `案場：${site.name}`}</div>

        {loading ? (
          <div className="ckn-loading">讀取中…</div>
        ) : (
          <>
            {done && (
              <div className="ckn-done">
                <div><CheckCircle2 size={16} style={{ verticalAlign: "-3px", marginRight: 6, color: "#6B9B6E" }} />已記錄！</div>
                <div>{done.workerName} · {done.date}</div>
                <div>工作時數 <b>{done.hours}</b> 小時，出工 <b>{done.days}</b> 天{done.overtimeHours > 0 ? <>，加班 <b>{done.overtimeHours}</b> 小時</> : null}</div>
              </div>
            )}
            <form onSubmit={submit}>
              <div className="ckn-field">
                <label>師傅</label>
                <select value={form.workerId} onChange={(e) => setForm({ ...form, workerId: e.target.value })} required>
                  <option value="" disabled>請選擇師傅</option>
                  {workers.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                </select>
              </div>
              <div className="ckn-field">
                <label>日期</label>
                <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} required />
              </div>
              <div className="ckn-row">
                <div className="ckn-field">
                  <label>上班時間</label>
                  <input type="time" value={form.start} onChange={(e) => setForm({ ...form, start: e.target.value })} required />
                </div>
                <div className="ckn-field">
                  <label>下班時間</label>
                  <input type="time" value={form.end} onChange={(e) => setForm({ ...form, end: e.target.value })} required />
                </div>
              </div>

              <div className="ckn-preview">
                <Clock size={14} style={{ verticalAlign: "-2px", marginRight: 4 }} />
                {result ? (
                  <>
                    <div>工作時數：<b>{result.hours}</b> 小時（已扣中午休息 1 小時）</div>
                    <div>出工天數：<b>{result.days}</b> 天{result.overtimeHours > 0 ? <> ｜ 加班：<b>{result.overtimeHours}</b> 小時</> : null}</div>
                    {result.days <= 0 && <div className="ckn-warn">未滿 4 小時，不計入出工天數</div>}
                  </>
                ) : (
                  <div className="ckn-warn">請確認下班時間晚於上班時間</div>
                )}
              </div>

              {workers.length === 0 ? (
                <div className="ckn-warn" style={{ marginBottom: 12 }}>目前還沒有師傅名單，請先請管理者在後台新增。</div>
              ) : null}

              <button type="submit" className="ckn-btn" disabled={saving || !form.workerId || !result || result.days <= 0}>
                {saving ? "送出中…" : "送出打卡"}
              </button>
              {error && <div className="ckn-warn" style={{ marginTop: 10 }}>{error}</div>}
            </form>
          </>
        )}
      </div>
    </div>
  );
}
