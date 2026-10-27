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

const sameEntries = (a, b) =>
  JSON.stringify([...a].sort((x, y) => x.id.localeCompare(y.id))) ===
  JSON.stringify([...b].sort((x, y) => x.id.localeCompare(y.id)));
export async function commitExcelImport(entries, expectedExisting) {
  const db = await dbPromise;
  return new Promise((resolve, reject) => {
    const tx = db.transaction(["entries", "settings"], "readwrite");
    let failure;
    const store = tx.objectStore("entries");
    const read = store.getAll();
    read.onsuccess = () => {
      if (!sameEntries(read.result, expectedExisting)) {
        failure = new Error(
          "Entries changed while you reviewed the import. Select the file again to review the latest entries.",
        );
        tx.abort();
        return;
      }
      for (const entry of entries) store.add(entry);
      if (entries.length)
        tx.objectStore("settings").put(
          { entries, createdAt: new Date().toISOString() },
          "lastExcelImport",
        );
    };
    tx.oncomplete = () => resolve(entries.length);
    tx.onabort = () =>
      reject(
        failure ||
          tx.error ||
          new Error("Import cancelled; no entries were added."),
      );
    tx.onerror = () => {
      failure ||= tx.error;
    };
  });
}
export async function undoExcelImport() {
  const db = await dbPromise;
  return new Promise((resolve, reject) => {
    const tx = db.transaction(["entries", "settings"], "readwrite");
    let removed = 0,
      kept = 0;
    const settings = tx.objectStore("settings");
    const read = settings.get("lastExcelImport");
    read.onsuccess = () => {
      for (const original of read.result?.entries || []) {
        const store = tx.objectStore("entries");
        const current = store.get(original.id);
        current.onsuccess = () => {
          if (current.result && sameEntries([current.result], [original])) {
            store.delete(original.id);
            removed++;
          } else if (current.result) kept++;
        };
      }
      settings.delete("lastExcelImport");
    };
    tx.oncomplete = () => resolve({ removed, kept });
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error || new Error("Undo cancelled."));
  });
}
