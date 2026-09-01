// ---- Firestore-backed persistence layer (REST API) ----
// Every piece of app data (orders, contracts, sites, ...) is stored as ONE document
// in the "wel-data" collection, keyed by name — e.g. wel-data/orders.
// Reads/writes go through Firestore's plain HTTPS REST API rather than the
// firebase/firestore SDK, which routes even one-time reads through its
// real-time "Watch" channel — that channel can take a long time to establish
// on some networks. REST is a single ordinary HTTPS request per call.
//
// Shared between the main admin app (App.jsx) and the simplified check-in
// page (CheckinPage.jsx) so both read/write the exact same data.
import { auth, ensureSignedIn, FIRESTORE_BASE } from "./firebase";

const COLLECTION = "wel-data";

async function authHeader() {
  await ensureSignedIn();
  const token = await auth.currentUser.getIdToken();
  return { Authorization: `Bearer ${token}` };
}

// Fetches every document in the collection in a single round trip, instead of
// one request per key — used on initial load where all keys are needed at once.
//
// Returns { data, ok }. `ok` is false when the fetch itself failed (network,
// auth, transient error) — callers MUST check this before treating an empty
// `data` as "this is a brand-new database, safe to seed with defaults", or a
// transient read failure can look identical to "no data yet" and trigger a
// write that overwrites real data with empty defaults. (This is exactly what
// happened once during development: a flaky read was mistaken for an empty
// database and wiped the "workers"/"worklogs"/"sites" documents.)
export async function storageGetAll() {
  const byKey = {};
  try {
    const headers = await authHeader();
    const res = await fetch(`${FIRESTORE_BASE}/${COLLECTION}?pageSize=300`, { headers });
    if (!res.ok) throw new Error(`list failed: ${res.status}`);
    const data = await res.json();
    for (const d of data.documents || []) {
      const key = d.name.split("/").pop();
      const raw = d.fields?.value?.stringValue;
      if (raw === undefined) continue;
      try {
        byKey[key] = JSON.parse(raw);
      } catch (e) {
        console.error("parse failed", key, e);
      }
    }
  } catch (e) {
    console.error("storage get all failed", e);
    return { data: byKey, ok: false };
  }
  return { data: byKey, ok: true };
}

export async function storageGet(key, fallback) {
  try {
    const headers = await authHeader();
    const res = await fetch(`${FIRESTORE_BASE}/${COLLECTION}/${key}`, { headers });
    if (res.status === 404) return fallback;
    if (!res.ok) throw new Error(`get failed: ${res.status}`);
    const data = await res.json();
    const raw = data.fields?.value?.stringValue;
    return raw === undefined ? fallback : JSON.parse(raw);
  } catch (e) {
    console.error("storage get failed", key, e);
    return fallback;
  }
}

// Like storageGet, but for read-modify-write flows where silently falling
// back to a default on a failed read would mean writing that default back
// and clobbering whatever was already there (see storageGetAll's comment).
// Throws instead of swallowing the error, so the caller can abort the write.
export async function storageGetOrThrow(key) {
  const headers = await authHeader();
  const res = await fetch(`${FIRESTORE_BASE}/${COLLECTION}/${key}`, { headers });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`get failed: ${res.status}`);
  const data = await res.json();
  const raw = data.fields?.value?.stringValue;
  return raw === undefined ? null : JSON.parse(raw);
}

export async function storageSet(key, value) {
  try {
    const headers = await authHeader();
    const res = await fetch(`${FIRESTORE_BASE}/${COLLECTION}/${key}`, {
      method: "PATCH",
      headers: { ...headers, "Content-Type": "application/json" },
      body: JSON.stringify({
        fields: {
          value: { stringValue: JSON.stringify(value) },
          updatedAt: { timestampValue: new Date().toISOString() },
        },
      }),
    });
    if (!res.ok) throw new Error(`set failed: ${res.status}`);
    return true;
  } catch (e) {
    console.error("storage set failed", key, e);
    return false;
  }
}
