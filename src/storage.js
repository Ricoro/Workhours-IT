const dbPromise = new Promise((resolve, reject) => {
  const r = indexedDB.open("workhours-it", 1);
  r.onupgradeneeded = () => {
    r.result.createObjectStore("entries", { keyPath: "id" });
    r.result.createObjectStore("settings");
  };
  r.onsuccess = () => resolve(r.result);
  r.onerror = () => reject(r.error);
});
async function request(store, mode, operation) {
  const db = await dbPromise;
  return new Promise((resolve, reject) => {
    const tx = db.transaction(store, mode);
    const result = operation(tx.objectStore(store));
    tx.oncomplete = () => resolve(result?.result);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () =>
      reject(tx.error || new Error("Storage operation cancelled"));
  });
}
export const allEntries = () =>
  request("entries", "readonly", (s) => s.getAll());
export const saveEntry = (e) =>
  request("entries", "readwrite", (s) => s.put(e));
export const removeEntry = (id) =>
  request("entries", "readwrite", (s) => s.delete(id));
export const getSetting = (key) =>
  request("settings", "readonly", (s) => s.get(key));
export const setSetting = (key, value) =>
  request("settings", "readwrite", (s) => s.put(value, key));
export const importEntries = (entries) =>
  request("entries", "readwrite", (s) => {
    for (const e of entries) s.put(e);
  });
