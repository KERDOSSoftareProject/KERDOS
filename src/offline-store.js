// KERDOS local snapshot store.
//
// The last successfully loaded workspace is kept in the browser's own
// IndexedDB so the app keeps working with no signal: every screen reads
// the same data it showed a moment ago, and the basket a person is
// building is never lost. This is a copy for reading, keyed per user and
// per organization; the backend stays the source of truth, and a save
// simply overwrites the previous copy. No library, no service, nothing
// leaves the device.

const DB_NAME = "kerdos";
const STORE = "snapshots";

function openDb() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") { reject(new Error("IndexedDB unavailable")); return; }
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

// Best effort: a failed save never interrupts the live app.
export async function saveSnapshot(key, value) {
  try {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, "readwrite");
      tx.objectStore(STORE).put({ savedAt: new Date().toISOString(), value }, key);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    db.close();
  } catch { /* storage unavailable or full: the live app is unaffected */ }
}

// Returns { savedAt, value } or null.
export async function loadSnapshot(key) {
  try {
    const db = await openDb();
    const result = await new Promise((resolve, reject) => {
      const request = db.transaction(STORE).objectStore(STORE).get(key);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return result;
  } catch { return null; }
}
