import React, { useState } from "react";

// 簡易密碼畫面：只擋住「完整後台」。密碼不以明文存放，輸入後比對雜湊值。
// 密碼驗證通過前，後台程式不會被載入，也不會讀取任何資料。
// 注意：這是網頁端的門檻，能擋一般人誤入，但不是銀行等級的防護。
const PASSWORD_HASH = "ea74f739a7c69e650fca997b19d0840dc0d86c6722e7373a691f53a1be358abc";
const SALT = "ren-ai-hydropower:";
const STORAGE_KEY = "wel-unlocked-at";
const REMEMBER_MS = 30 * 24 * 60 * 60 * 1000;

async function sha256Hex(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, "0")).join("");
}

function isRemembered() {
  try {
    const at = Number(localStorage.getItem(STORAGE_KEY));
    return at > 0 && Date.now() - at < REMEMBER_MS;
  } catch (e) {
    return false;
  }
}

export function lockApp() {
  try { localStorage.removeItem(STORAGE_KEY); } catch (e) { /* ignore */ }
  window.location.reload();
}

export default function PasswordGate({ children }) {
  const [unlocked, setUnlocked] = useState(isRemembered);
  const [value, setValue] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [fails, setFails] = useState(0);

  if (unlocked) return children;

  const submit = async (e) => {
    e.preventDefault();
    if (busy) return;
    setBusy(true);
    setError("");
    try {
      const hash = await sha256Hex(SALT + value.trim());
      if (hash === PASSWORD_HASH) {
        try { localStorage.setItem(STORAGE_KEY, String(Date.now())); } catch (err) { /* private mode: just unlock for this visit */ }
        setUnlocked(true);
        return;
      }
      const n = fails + 1;
      setFails(n);
      setError("密碼不對，請再試一次");
      setValue("");
      if (n >= 3) await new Promise((r) => setTimeout(r, Math.min(n, 8) * 1000));
    } catch (err) {
      setError("這個瀏覽器無法驗證密碼，請改用較新的瀏覽器");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#1C2530", color: "#E6EDF3", fontFamily: "system-ui, -apple-system, 'Noto Sans TC', sans-serif", padding: 20 }}>
      <form onSubmit={submit} style={{ width: "100%", maxWidth: 340, background: "#252F3D", border: "1px solid #354253", borderRadius: 14, padding: "28px 24px", textAlign: "center" }}>
        <div style={{ fontSize: 22, fontWeight: 700, marginBottom: 6 }}>仁愛水電工程紀錄帳本</div>
        <div style={{ fontSize: 13, color: "#8B98A8", marginBottom: 20 }}>請輸入密碼</div>
        <input
          type="password" inputMode="numeric" autoComplete="current-password" autoFocus
          value={value} onChange={(e) => setValue(e.target.value)} placeholder="密碼"
          style={{ width: "100%", height: 46, borderRadius: 8, border: "1px solid #3C4B5E", background: "#1C2530", color: "#E6EDF3", fontSize: 18, textAlign: "center", letterSpacing: 6, outline: "none", boxSizing: "border-box" }}
        />
        <div style={{ minHeight: 22, fontSize: 13, color: "#E5736A", marginTop: 10 }}>{error}</div>
        <button
          type="submit" disabled={busy || !value}
          style={{ width: "100%", height: 44, marginTop: 4, borderRadius: 8, border: "none", background: busy || !value ? "#6B5A33" : "#E8A33D", color: "#1C2530", fontSize: 15, fontWeight: 700, cursor: busy || !value ? "default" : "pointer" }}
        >
          {busy ? "驗證中…" : "進入"}
        </button>
      </form>
    </div>
  );
}
