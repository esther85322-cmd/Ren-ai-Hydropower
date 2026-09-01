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

const WEEKDAY_LABELS = ["日", "一", "二", "三", "四", "五", "六"];

// 某位師傅在某個月份每一天的出班狀態：紫=有加班，紅=有出班（滿天），黃=半天，綠=休息（沒有紀錄）。
function buildMonthCalendar(workerId, monthDateStr, worklogs) {
  const [y, m] = monthDateStr.split("-").map(Number);
  const daysInMonth = new Date(y, m, 0).getDate();
  const startWeekday = new Date(y, m - 1, 1).getDay();
  const byDate = {};
  for (const l of worklogs) {
    if (l.workerId !== workerId) continue;
    if (!l.date.startsWith(monthDateStr.slice(0, 7))) continue;
    byDate[l.date] = l;
  }
  const days = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    const log = byDate[dateStr];
    const status = !log
      ? "rest"
      : Number(log.overtimeHours) > 0
      ? "overtime"
      : Number(log.days) >= 1
      ? "work"
      : Number(log.days) > 0
      ? "half"
      : "rest";
    days.push({ d, dateStr, status });
  }
  return { y, m, startWeekday, days };
}

export default function CheckinPage() {
  const [loading, setLoading] = useState(true);
  const [workers, setWorkers] = useState([]);
  const [sites, setSites] = useState([DEFAULT_SITE]);
  const [worklogs, setWorklogs] = useState([]);
  const [saving, setSaving] = useState(false);
  const [done, setDone] = useState(null);
  const [form, setForm] = useState({
    workerId: "",
    siteId: "",
    date: todayStr(),
    start: "08:00",
    end: "17:00",
  });

  useEffect(() => {
    (async () => {
      const [wk, st, wl] = await Promise.all([
        storageGet("workers", []),
        storageGet("sites", [DEFAULT_SITE]),
        storageGet("worklogs", []),
      ]);
      setWorkers(wk);
      setWorklogs(wl);
      const siteList = st && st.length ? st : [DEFAULT_SITE];
      setSites(siteList);
      // URL 可以帶 ?site=<id> 預先選好案場（方便針對單一案場分享固定連結），
      // 沒有帶的話預設選第一個，之後仍可在頁面上手動改選。
      const params = new URLSearchParams(window.location.search);
      const siteId = params.get("site");
      const matched = siteList.find((s) => s.id === siteId);
      setForm((f) => ({ ...f, siteId: (matched || siteList[0]).id }));
      setLoading(false);
    })();
  }, []);

  const result = useMemo(() => computeAttendance(form.start, form.end), [form.start, form.end]);
  const selectedWorker = workers.find((w) => w.id === form.workerId);
  const selectedSite = sites.find((s) => s.id === form.siteId);

  const calendar = form.workerId
    ? buildMonthCalendar(form.workerId, form.date, worklogs)
    : null;

  // 同一位師傅、同一天已經打過卡了 → 不新增一筆，而是修改原本那筆的時間。
  const existingLog = form.workerId
    ? worklogs.find((l) => l.workerId === form.workerId && l.date === form.date)
    : null;

  // 選了「已經打過卡」的師傅／日期時，把時間欄位帶出原本記錄的值，方便直接修改；
  // 換成別的師傅或別的日期、且沒有重複時，則回到預設的 08:00-17:00。
  useEffect(() => {
    if (loading) return;
    setForm((f) => ({
      ...f,
      start: existingLog?.startTime || "08:00",
      end: existingLog?.endTime || "17:00",
    }));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.workerId, form.date]);

  const [error, setError] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    if (!form.workerId || !form.siteId) return;
    if (!result || result.days <= 0) return;
    setSaving(true);
    setError("");
    try {
      // Read the current list fresh right before writing (this page has no
      // pre-loaded state to fall back on). If this read fails, we must NOT
      // fall back to an empty list — that would overwrite everyone else's
      // existing worklogs with just this one new record.
      const list = (await storageGetOrThrow("worklogs")) || [];
      const note = `打卡 ${form.start}-${form.end}${result.overtimeHours > 0 ? `，加班 ${result.overtimeHours} 小時` : ""}`;
      const matchIdx = list.findIndex((l) => l.workerId === form.workerId && l.date === form.date);
      let next;
      if (matchIdx >= 0) {
        next = list.map((l, i) =>
          i === matchIdx
            ? { ...l, days: result.days, note, siteId: form.siteId, startTime: form.start, endTime: form.end, overtimeHours: result.overtimeHours }
            : l
        );
      } else {
        const rec = {
          id: uid(),
          workerId: form.workerId,
          date: form.date,
          days: result.days,
          note,
          siteId: form.siteId,
          startTime: form.start,
          endTime: form.end,
          overtimeHours: result.overtimeHours,
        };
        next = [rec, ...list];
      }
      const ok = await storageSet("worklogs", next);
      if (!ok) throw new Error("write failed");
      setWorklogs(next);
      setDone({
        workerName: selectedWorker?.name, siteName: selectedSite?.name, ...result, date: form.date,
        updated: matchIdx >= 0,
      });
      setForm((f) => ({ ...f, workerId: "", date: todayStr(), start: "08:00", end: "17:00" }));
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
        @media (max-width: 420px) { .ckn-row { flex-direction: column; gap: 16px; } }
        .ckn-preview {
          background: #232E3A; border: 1px solid #3A4453; border-radius: 10px; padding: 14px 16px;
          margin-bottom: 20px; font-size: 13.5px; display: flex; flex-direction: column; gap: 6px;
        }
        .ckn-preview b { color: #E8A33D; font-family: monospace; }
        .ckn-warn { color: #C1543C; }
        .ckn-note {
          background: rgba(232,163,61,0.1); border: 1px solid rgba(232,163,61,0.35); color: #E8D2A8;
          border-radius: 8px; padding: 10px 12px; font-size: 12.5px; margin-bottom: 16px;
        }
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
        .ckn-cal { margin-bottom: 20px; }
        .ckn-cal-title { font-size: 12.5px; color: #8B95A1; margin-bottom: 8px; }
        .ckn-cal-grid { display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; }
        .ckn-cal-wd { text-align: center; font-size: 10.5px; color: #6B7684; padding-bottom: 2px; }
        .ckn-cal-day {
          aspect-ratio: 1; display: flex; align-items: center; justify-content: center;
          border-radius: 6px; font-size: 12px; font-family: monospace; color: #1C2530; font-weight: 600;
          border: none; cursor: pointer; padding: 0;
        }
        .ckn-cal-day.rest { background: rgba(107,155,110,0.35); color: #E8E4DA; }
        .ckn-cal-day.half { background: #E8D24A; }
        .ckn-cal-day.work { background: #D9503C; color: #fff; }
        .ckn-cal-day.overtime { background: #8B5CD9; color: #fff; }
        .ckn-cal-day.selected { box-shadow: 0 0 0 2px #E8A33D inset; }
        .ckn-cal-legend { display: flex; gap: 14px; margin-top: 8px; font-size: 11px; color: #8B95A1; }
        .ckn-cal-legend span { display: inline-flex; align-items: center; gap: 4px; }
        .ckn-cal-dot { width: 9px; height: 9px; border-radius: 3px; display: inline-block; }
      `}</style>
      <div className="ckn-card">
        <div className="ckn-title"><HardHat size={22} /> 出工打卡</div>
        <div className="ckn-site">{loading ? "讀取中…" : `${sites.length} 個案場可選`}</div>

        {loading ? (
          <div className="ckn-loading">讀取中…</div>
        ) : (
          <>
            {done && (
              <div className="ckn-done">
                <div><CheckCircle2 size={16} style={{ verticalAlign: "-3px", marginRight: 6, color: "#6B9B6E" }} />{done.updated ? "已更新！" : "已記錄！"}</div>
                <div>{done.siteName} · {done.workerName} · {done.date}</div>
                <div>工作時數 <b>{done.hours}</b> 小時，出工 <b>{done.days}</b> 天{done.overtimeHours > 0 ? <>，加班 <b>{done.overtimeHours}</b> 小時</> : null}</div>
              </div>
            )}
            <form onSubmit={submit}>
              <div className="ckn-field">
                <label>案場</label>
                <select value={form.siteId} onChange={(e) => setForm({ ...form, siteId: e.target.value })} required>
                  {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
              </div>
              <div className="ckn-field">
                <label>師傅</label>
                <select value={form.workerId} onChange={(e) => setForm({ ...form, workerId: e.target.value })} required>
                  <option value="" disabled>請選擇師傅</option>
                  {workers.map((w) => <option key={w.id} value={w.id}>{w.name}</option>)}
                </select>
              </div>
              {calendar && (
                <div className="ckn-cal">
                  <div className="ckn-cal-title">{selectedWorker?.name} · {calendar.y}年{calendar.m}月出班紀錄</div>
                  <div className="ckn-cal-grid">
                    {WEEKDAY_LABELS.map((w) => <div key={w} className="ckn-cal-wd">{w}</div>)}
                    {Array.from({ length: calendar.startWeekday }).map((_, i) => <div key={"pad" + i} />)}
                    {calendar.days.map(({ d, dateStr, status }) => (
                      <button
                        key={d} type="button"
                        className={`ckn-cal-day ${status}${dateStr === form.date ? " selected" : ""}`}
                        onClick={() => setForm((f) => ({ ...f, date: dateStr }))}
                      >
                        {d}
                      </button>
                    ))}
                  </div>
                  <div className="ckn-cal-legend">
                    <span><i className="ckn-cal-dot" style={{ background: "#D9503C" }} />有出班</span>
                    <span><i className="ckn-cal-dot" style={{ background: "#8B5CD9" }} />有加班</span>
                    <span><i className="ckn-cal-dot" style={{ background: "#E8D24A" }} />半天</span>
                    <span><i className="ckn-cal-dot" style={{ background: "rgba(107,155,110,0.6)" }} />休息</span>
                  </div>
                </div>
              )}

              <div className="ckn-field">
                <label>日期</label>
                <input type="date" value={form.date} onChange={(e) => setForm({ ...form, date: e.target.value })} required />
              </div>

              {existingLog && (
                <div className="ckn-note">
                  {selectedWorker?.name} 在 {form.date} 已經打過卡了（{existingLog.startTime || "—"}-{existingLog.endTime || "—"}）。
                  下面時間已經帶出原本的紀錄，直接改完送出就會更新這筆，不會變成兩筆。
                </div>
              )}

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

              <button type="submit" className="ckn-btn" disabled={saving || !form.workerId || !form.siteId || !result || result.days <= 0}>
                {saving ? "送出中…" : existingLog ? "更新工時" : "送出打卡"}
              </button>
              {error && <div className="ckn-warn" style={{ marginTop: 10 }}>{error}</div>}
            </form>
          </>
        )}
      </div>
    </div>
  );
}
