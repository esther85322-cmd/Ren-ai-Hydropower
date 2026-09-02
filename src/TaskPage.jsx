import React, { useState, useEffect } from "react";
import { MessageSquare, Circle, Plus } from "lucide-react";
import { storageGet, storageSet, storageGetOrThrow } from "./storage";

const DEFAULT_SITE = { id: "site-default", name: "預設案場", address: "" };
const uid = () => Math.random().toString(36).slice(2, 10);
const todayStr = () => new Date().toISOString().slice(0, 10);

export default function TaskPage() {
  const [loading, setLoading] = useState(true);
  const [sites, setSites] = useState([DEFAULT_SITE]);
  const [items, setItems] = useState([]);
  const [siteId, setSiteId] = useState("");
  const [topic, setTopic] = useState("");
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [busyId, setBusyId] = useState(null);

  useEffect(() => {
    (async () => {
      const [st, di] = await Promise.all([
        storageGet("sites", [DEFAULT_SITE]),
        storageGet("discussionitems", []),
      ]);
      const siteList = st && st.length ? st : [DEFAULT_SITE];
      setSites(siteList);
      setItems(di);
      const params = new URLSearchParams(window.location.search);
      const wantedSite = params.get("site");
      const matched = siteList.find((s) => s.id === wantedSite);
      setSiteId((matched || siteList[0]).id);
      setLoading(false);
    })();
  }, []);

  const openItems = items
    .filter((d) => d.siteId === siteId && !d.resolved)
    .sort((a, b) => (a.date < b.date ? 1 : -1));

  const addItem = async (e) => {
    e.preventDefault();
    if (!topic.trim() || !siteId) return;
    setSaving(true);
    setError("");
    try {
      // 送出前重新讀一次最新清單再寫回去，避免讀取失敗時用空清單蓋掉別人剛新增的項目。
      const list = (await storageGetOrThrow("discussionitems")) || [];
      const rec = { id: uid(), date: todayStr(), topic: topic.trim(), note: note.trim(), resolved: false, result: "", siteId };
      const next = [rec, ...list];
      const ok = await storageSet("discussionitems", next);
      if (!ok) throw new Error("write failed");
      setItems(next);
      setTopic("");
      setNote("");
    } catch (err) {
      console.error("add task failed", err);
      setError("送出失敗，請檢查網路後再試一次");
    } finally {
      setSaving(false);
    }
  };

  const completeItem = async (id) => {
    setBusyId(id);
    setError("");
    try {
      const list = (await storageGetOrThrow("discussionitems")) || [];
      const next = list.map((d) => (d.id === id ? { ...d, resolved: true } : d));
      const ok = await storageSet("discussionitems", next);
      if (!ok) throw new Error("write failed");
      setItems(next);
    } catch (err) {
      console.error("complete task failed", err);
      setError("更新失敗，請檢查網路後再試一次");
    } finally {
      setBusyId(null);
    }
  };

  const selectedSite = sites.find((s) => s.id === siteId);

  return (
    <div className="tsk-app">
      <style>{`
        .tsk-app {
          min-height: 100vh; background: #1C2530; color: #E8E4DA;
          font-family: 'IBM Plex Sans', sans-serif; display: flex; justify-content: center;
          padding: 32px 18px 60px;
        }
        .tsk-app * { box-sizing: border-box; }
        .tsk-card { width: 100%; max-width: 420px; }
        .tsk-title { display: flex; align-items: center; gap: 8px; font-size: 20px; font-weight: 600; margin-bottom: 4px; }
        .tsk-sub { color: #8B95A1; font-size: 13px; margin-bottom: 24px; }
        .tsk-field { display: flex; flex-direction: column; gap: 6px; margin-bottom: 14px; }
        .tsk-field label { font-size: 13px; color: #8B95A1; }
        .tsk-field input, .tsk-field select {
          background: #232E3A; border: 1px solid #3A4453; color: #E8E4DA; border-radius: 8px;
          padding: 0 12px; height: 46px; font-size: 15px; width: 100%;
        }
        .tsk-btn {
          width: 100%; background: #E8A33D; color: #1C2530; border: none; border-radius: 8px;
          padding: 13px; font-size: 15px; font-weight: 600; cursor: pointer; display: flex;
          align-items: center; justify-content: center; gap: 6px;
        }
        .tsk-btn:disabled { opacity: 0.5; }
        .tsk-warn { color: #C1543C; font-size: 12.5px; margin-top: 8px; }
        .tsk-loading { padding: 60px 0; text-align: center; color: #8B95A1; }
        .tsk-section-title { font-size: 12.5px; color: #8B95A1; margin: 28px 0 10px; }
        .tsk-item {
          display: flex; align-items: flex-start; gap: 10px; background: #232E3A; border: 1px solid #3A4453;
          border-radius: 10px; padding: 12px 14px; margin-bottom: 8px;
        }
        .tsk-item button { background: transparent; border: none; padding: 0; cursor: pointer; flex-shrink: 0; margin-top: 1px; }
        .tsk-item button:disabled { opacity: 0.4; cursor: default; }
        .tsk-item-body { flex: 1; min-width: 0; }
        .tsk-item-topic { font-size: 14px; line-height: 1.5; }
        .tsk-item-note { font-size: 12px; color: #8B95A1; margin-top: 4px; }
        .tsk-item-date { font-size: 10.5px; color: #6B7684; font-family: monospace; margin-top: 4px; }
        .tsk-empty { text-align: center; color: #6B7684; font-size: 13px; padding: 24px 0; }
      `}</style>
      <div className="tsk-card">
        <div className="tsk-title"><MessageSquare size={22} /> 待確認事項</div>
        <div className="tsk-sub">{loading ? "讀取中…" : `${sites.length} 個案場可選`}</div>

        {loading ? (
          <div className="tsk-loading">讀取中…</div>
        ) : (
          <>
            <div className="tsk-field">
              <label>案場</label>
              <select value={siteId} onChange={(e) => setSiteId(e.target.value)}>
                {sites.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
              </select>
            </div>

            <form onSubmit={addItem}>
              <div className="tsk-field">
                <label>今日待確認</label>
                <input placeholder="例：3樓管線與空調衝突，需與空調師傅確認" value={topic} onChange={(e) => setTopic(e.target.value)} required />
              </div>
              <div className="tsk-field">
                <label>備註（選填）</label>
                <input placeholder="選填" value={note} onChange={(e) => setNote(e.target.value)} />
              </div>
              <button type="submit" className="tsk-btn" disabled={saving || !topic.trim()}>
                <Plus size={16} /> {saving ? "送出中…" : "新增事項"}
              </button>
              {error && <div className="tsk-warn">{error}</div>}
            </form>

            <div className="tsk-section-title">{selectedSite?.name} · 待處理（{openItems.length}）</div>
            {openItems.length === 0 ? (
              <div className="tsk-empty">目前沒有待確認事項</div>
            ) : (
              openItems.map((d) => (
                <div key={d.id} className="tsk-item">
                  <button type="button" onClick={() => completeItem(d.id)} disabled={busyId === d.id} title="標記完成">
                    <Circle size={19} color="#8B95A1" />
                  </button>
                  <div className="tsk-item-body">
                    <div className="tsk-item-topic">{d.topic}</div>
                    {d.note && <div className="tsk-item-note">{d.note}</div>}
                    <div className="tsk-item-date">{d.date}</div>
                  </div>
                </div>
              ))
            )}
          </>
        )}
      </div>
    </div>
  );
}
