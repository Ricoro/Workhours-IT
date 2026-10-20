import "./style.css";
import { COMPANIES, localDate, validateEntry, totals } from "./model.js";
import {
  allEntries,
  saveEntry,
  removeEntry,
  getSetting,
  setSetting,
  importEntries,
} from "./storage.js";
import { workbook } from "./excel.js";
const $ = (s) => document.querySelector(s);
function preference(key, fallback = null) {
  try {
    return localStorage.getItem(key) ?? fallback;
  } catch {
    return fallback;
  }
}
function remember(key, value) {
  // Optional browser preferences must not turn a committed entry into a failed save.
  try {
    if (value === null) localStorage.removeItem(key);
    else localStorage.setItem(key, value);
    return true;
  } catch {
    return false;
  }
}
const escape = (s) =>
  String(s).replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        c
      ],
  );
let entries = [],
  month = localDate().slice(0, 7),
  editing = null,
  handle = null,
  filter = "all",
  search = "",
  syncQueue = Promise.resolve();
const fmt = (n) =>
  new Intl.NumberFormat("en-GB", { maximumFractionDigits: 2 }).format(n);
const monthLabel = (m) =>
  new Date(`${m}-01T12:00:00`).toLocaleDateString("en-GB", {
    month: "long",
    year: "numeric",
  });

$("#app").innerHTML =
  `<aside class="sidebar"><a class="brand" href="#"><span class="brand-icon">w<span>·</span></span>workhours<span class="brand-it">IT</span></a><div class="workspace-label">YOUR WORKSPACE</div><div class="nav-item"><span>◷</span> Time overview <span class="nav-dot"></span></div><div class="sidebar-note"><span class="lock">⌑</span><strong>Private by design</strong><p>Your entries stay in this browser. No account or server required.</p></div><div class="sidebar-footer"><span class="avatar">IT</span><div><strong>IT workspace</strong><small>Local workspace</small></div></div></aside><main><header><div class="breadcrumb">Workspace <span>/</span> <strong>Time overview</strong></div><button class="icon-button" id="theme" aria-label="Toggle dark mode" title="Toggle dark mode">☾</button></header><div class="content"><div class="heading"><div><div class="eyebrow">MAKE YOUR TIME COUNT</div><h1>Your work, accounted for<span>.</span></h1><p>One place for your IT hours. Every company, every task.</p></div><button class="button secondary" id="export">↓ <span>Export month</span></button></div><section class="toolbar"><div class="month-nav"><button id="previous" aria-label="Previous month">‹</button><label class="month-label"><span id="month-title"></span><input id="month" type="month" aria-label="Select month"></label><button id="next" aria-label="Next month">›</button></div><span class="period-note">Your monthly overview</span></section><section id="stats" class="stats" aria-label="Monthly hours by company"></section><div class="work-grid"><section class="panel entry-panel"><div class="panel-title"><div><h2 id="form-title">Log your time</h2><p>A little record of a job well done.</p></div><span class="small-icon" aria-hidden="true"><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M6 18 18 6M6 6h12v12"/></svg></span></div><form id="entry-form"><div class="form-row"><label>Work date<input type="date" id="date" required min="1900-01-01" max="9999-12-31"></label><label>Employee ID<input id="employee" required maxlength="50" value="12"></label></div><label>What did you work on?<textarea id="notes" required maxlength="10000" rows="4" placeholder="e.g. Set up a laptop and resolved email issues…"></textarea></label><div class="allocation-heading"><label>Hours per company</label><span>Split a task across companies</span></div><div class="allocations">${COMPANIES.map((c) => `<label class="allocation"><span class="company-mark ${c.toLowerCase()}">${c === "SEABOW" ? "SB" : c === "4SOME" ? "4S" : c === "ENAB" ? "EN" : c}</span><span>${c}</span><div><input type="number" id="hours-${c}" aria-label="${c} hours" min="0" max="24" step="0.01" placeholder="0.00"><span>h</span></div></label>`).join("")}</div><div class="entry-total"><span>Entry total</span><strong><span id="entry-hours">0</span> <small>hours</small></strong></div><button class="button primary" id="save" type="submit">+ <span>Save entry</span></button><button class="text-button" id="cancel-edit" type="button" hidden>Cancel edit</button><p class="save-hint" id="draft-status">Entries save on this device. Your unfinished draft is kept too.</p></form></section><section class="panel history-panel"><div class="panel-title"><div><h2>Activity log <span class="count" id="count">0</span></h2><p>The details behind your hours.</p></div></div><div class="log-filters"><label class="search"><span>⌕</span><input id="search" placeholder="Search your work…" aria-label="Search activities"></label><select id="company-filter" aria-label="Filter company"><option value="all">All companies</option>${COMPANIES.map((c) => `<option>${c}</option>`).join("")}</select></div><div id="activity"></div><div class="log-footer"><span id="log-summary"></span><span>Saved on this device <span class="green-dot"></span></span></div></section></div><section class="backup-panel"><div class="backup-icon" aria-hidden="true"><svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 16V3m-5 5 5-5 5 5M5 15v6h14v-6"/></svg></div><div class="backup-copy"><h3>Your hours. Your Excel.</h3><p>Link a folder to update <strong>IT-hours.xlsx</strong> after every saved change. Choose your OneDrive sync folder to sync through the desktop app.</p><span id="sync-status" role="status">No folder linked · manual exports are available</span></div><button class="button secondary" id="link-folder">Link Excel folder</button></section><details class="data-tools"><summary>Backup & restore</summary><p>Browser data can be cleared. Keep a backup before switching devices or browsers. Restoring merges entries by ID; matching IDs are replaced.</p><div><button class="button secondary" id="backup">Download backup</button><label class="button secondary file-label">Restore backup<input type="file" id="restore" accept=".json,application/json"></label></div></details><footer class="footer"><span>workhours <strong>IT</strong><span class="footer-divider">/</span> A clearer view of your time.</span><span>Local storage · No tracking</span></footer></div></main><div class="toast" id="toast" role="status" hidden></div><dialog id="delete-dialog"><h2>Delete this entry?</h2><p>This removes the task and all its company hours.</p><div class="dialog-actions"><button class="button secondary" id="keep">Keep entry</button><button class="button danger" id="confirm-delete">Delete entry</button></div></dialog>`;
function notify(message, error = false) {
  const t = $("#toast");
  t.textContent = message;
  t.classList.toggle("error", error);
  t.hidden = false;
  clearTimeout(notify.timer);
  notify.timer = setTimeout(() => (t.hidden = true), 6500);
}
function render() {
  const selected = entries.filter((e) => e.date.startsWith(month));
  const t = totals(selected);
  const total = Object.values(t).reduce((a, b) => a + b, 0);
  $("#month-title").textContent = monthLabel(month);
  $("#month").value = month;
  $("#stats").innerHTML =
    `<article class="stat total-stat"><span>Hours this month <span>◷</span></span><strong>${fmt(total)}<small>h</small></strong><div>${selected.length} ${selected.length === 1 ? "task" : "tasks"} recorded</div></article>${COMPANIES.map((c) => `<article class="stat"><span><i class="company-dot ${c.toLowerCase()}"></i>${c}</span><strong>${fmt(t[c])}<small>h</small></strong><div class="progress"><span class="${c.toLowerCase()}" style="width:${total ? (t[c] / total) * 100 : 0}%"></span></div></article>`).join("")}`;
  const visible = selected
    .filter(
      (e) =>
        (filter === "all" || e.allocations[filter]) &&
        e.notes.toLowerCase().includes(search.toLowerCase()),
    )
    .sort(
      (a, b) =>
        b.date.localeCompare(a.date) || b.updatedAt.localeCompare(a.updatedAt),
    );
  $("#count").textContent = selected.length;
  $("#activity").innerHTML = visible.length
    ? `<div class="table-wrap"><table><thead><tr><th>DATE / TASK</th><th>COMPANY</th><th class="right">HOURS</th><th><span class="sr-only">Actions</span></th></tr></thead><tbody>${visible
        .map(
          (e) =>
            `<tr><td><span class="task-date">${new Date(`${e.date}T12:00:00`).toLocaleDateString("en-GB", { day: "numeric", month: "short" })}</span><p class="task-notes">${escape(e.notes)}</p><span class="employee-note">Employee ${escape(e.employee)}</span></td><td><div class="badges">${COMPANIES.filter(
              (c) => e.allocations[c],
            )
              .map(
                (c) =>
                  `<span class="badge ${c.toLowerCase()}">${c}<small>${fmt(e.allocations[c])}h</small></span>`,
              )
              .join(
                "",
              )}</div></td><td class="right task-hours">${fmt(Object.values(e.allocations).reduce((a, b) => a + b, 0))}<small>h</small></td><td><div class="row-actions"><button class="edit-button" data-id="${escape(e.id)}" aria-label="Edit entry" title="Edit entry">✎</button><button class="delete-button" data-id="${escape(e.id)}" aria-label="Delete entry" title="Delete entry">×</button></div></td></tr>`,
        )
        .join("")}</tbody></table></div>`
    : `<div class="empty"><div>◷</div><h3>${selected.length ? "No matching entries" : "A fresh month, ready for you"}</h3><p>${selected.length ? "Try another company or search term." : "Log your first task and watch your hours come together."}</p><span>${selected.length ? "" : "Your next entry starts on the left ↖"}</span></div>`;
  $("#log-summary").textContent =
    `${visible.length} ${visible.length === 1 ? "entry" : "entries"} shown`;
}
function formEntry() {
  const allocations = {};
  for (const c of COMPANIES) {
    const value = $(`#hours-${c}`).value;
    if (value !== "" && Number(value) !== 0) allocations[c] = Number(value);
  }
  return {
    id: editing || crypto.randomUUID(),
    date: $("#date").value,
    employee: $("#employee").value.trim(),
    notes: $("#notes").value.trim(),
    allocations,
    updatedAt: new Date().toISOString(),
  };
}
function updateTotal() {
  const n = COMPANIES.reduce(
    (n, c) => n + (Number($(`#hours-${c}`).value) || 0),
    0,
  );
  $("#entry-hours").textContent = fmt(n);
}
function captureDraft() {
  const e = formEntry();
  const saved = remember(
    "workhours-draft",
    JSON.stringify({ ...e, id: editing }),
  );
  $("#draft-status").textContent = saved
    ? "Draft saved on this device"
    : "Draft could not be saved. Save the entry before leaving.";
  updateTotal();
}
function resetForm(clearDraft = true) {
  editing = null;
  $("#entry-form").reset();
  $("#date").value = localDate();
  $("#employee").value = preference("workhours-employee", "12");
  $("#form-title").textContent = "Log your time";
  $("#save span").textContent = "Save entry";
  $("#cancel-edit").hidden = true;
  if (clearDraft) remember("workhours-draft", null);
  $("#draft-status").textContent =
    "Entries save on this device. Your unfinished draft is kept too.";
  updateTotal();
}
function fillForm(e, isDraft = false) {
  editing = e.id || null;
  $("#date").value = e.date;
  $("#employee").value = e.employee;
  $("#notes").value = e.notes;
  for (const c of COMPANIES) $(`#hours-${c}`).value = e.allocations[c] || "";
  $("#form-title").textContent = editing ? "Edit your entry" : "Log your time";
  $("#save span").textContent = editing ? "Save changes" : "Save entry";
  $("#cancel-edit").hidden = !editing;
  updateTotal();
  if (!isDraft) captureDraft();
}
async function refresh() {
  entries = await allEntries();
  render();
}
function syncExcel() {
  syncQueue = syncQueue
    .catch(() => {})
    .then(async () => {
      if (!handle) return;
      $("#sync-status").textContent = "Updating linked Excel file…";
      try {
        if ((await handle.queryPermission({ mode: "readwrite" })) !== "granted")
          throw new Error(
            "Folder permission expired. Click Reconnect Excel folder.",
          );
        const file = await handle.getFileHandle("IT-hours.xlsx", {
          create: true,
        });
        const data = await workbook(await allEntries());
        const writable = await file.createWritable();
        try {
          await writable.write(data);
          await writable.close();
        } catch (e) {
          await writable.abort().catch(() => {});
          throw e;
        }
        $("#sync-status").textContent =
          `Excel updated at ${new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} · IT-hours.xlsx`;
      } catch (e) {
        $("#sync-status").textContent =
          `Excel not updated: ${e.message} Your entries remain saved in this browser.`;
        notify("Entry saved locally; Excel update needs attention.", true);
      }
    });
  return syncQueue;
}
function download(data, name, type) {
  const url = URL.createObjectURL(new Blob([data], { type }));
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
$("#entry-form").addEventListener("input", captureDraft);
$("#entry-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = $("#save");
  button.disabled = true;
  try {
    const e = validateEntry(formEntry());
    await saveEntry(e);
    remember("workhours-employee", e.employee);
    month = e.date.slice(0, 7);
    resetForm();
    await refresh();
    notify("Entry saved on this device.");
    channel?.postMessage("refresh");
    void syncExcel();
  } catch (e) {
    notify(e.message, true);
  } finally {
    button.disabled = false;
  }
});
$("#cancel-edit").onclick = () => resetForm();
$("#activity").onclick = async (event) => {
  const button = event.target.closest("button");
  if (!button) return;
  const e = entries.find((e) => e.id === button.dataset.id);
  if (!e) return;
  if (button.classList.contains("edit-button")) {
    fillForm(e);
    $("#entry-form").scrollIntoView({ behavior: "smooth", block: "center" });
    $("#notes").focus();
  } else {
    const dialog = $("#delete-dialog");
    dialog.showModal();
    $("#keep").onclick = () => dialog.close();
    $("#confirm-delete").onclick = async () => {
      try {
        await removeEntry(e.id);
        if (editing === e.id) resetForm();
        dialog.close();
        await refresh();
        channel?.postMessage("refresh");
        notify("Entry deleted.");
        void syncExcel();
      } catch (err) {
        notify(err.message, true);
      }
    };
  }
};
function shiftMonth(delta) {
  const d = new Date(`${month}-01T12:00:00`);
  d.setMonth(d.getMonth() + delta);
  month = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  render();
}
$("#previous").onclick = () => shiftMonth(-1);
$("#next").onclick = () => shiftMonth(1);
$("#month").onchange = (e) => {
  if (/^\d{4}-\d{2}$/.test(e.target.value)) {
    month = e.target.value;
    render();
  }
};
$("#search").oninput = (e) => {
  search = e.target.value;
  render();
};
$("#company-filter").onchange = (e) => {
  filter = e.target.value;
  render();
};
$("#export").onclick = async () => {
  try {
    download(
      await workbook(await allEntries(), month),
      `IT-hours-${month}.xlsx`,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    notify(`Excel exported for ${monthLabel(month)} (all companies).`);
  } catch (e) {
    notify(`Export failed: ${e.message}`, true);
  }
};
$("#link-folder").onclick = async () => {
  if (!("showDirectoryPicker" in window)) {
    notify(
      "Folder auto-save needs desktop Chrome or Edge. Use Export month or Download backup in this browser.",
      true,
    );
    return;
  }
  try {
    if (handle) {
      if ((await handle.requestPermission({ mode: "readwrite" })) !== "granted")
        throw new Error("Folder permission was not granted.");
    } else {
      const picked = await window.showDirectoryPicker({
        mode: "readwrite",
        id: "workhours-excel",
      });
      let exists = false;
      try {
        await picked.getFileHandle("IT-hours.xlsx");
        exists = true;
      } catch (e) {
        if (e.name !== "NotFoundError") throw e;
      }
      if (
        exists &&
        !confirm(
          "IT-hours.xlsx already exists in this folder. Replace it with this app’s entries? Choose Cancel to keep the existing file.",
        )
      )
        return;
      await setSetting("excel-folder", picked);
      handle = picked;
    }
    $("#link-folder").textContent = "Reconnect Excel folder";
    await syncExcel();
  } catch (e) {
    if (e.name !== "AbortError") notify(e.message, true);
  }
};
$("#backup").onclick = async () => {
  try {
    download(
      JSON.stringify({ version: 1, entries: await allEntries() }, null, 2),
      `workhours-backup-${localDate()}.json`,
      "application/json",
    );
  } catch (e) {
    notify(e.message, true);
  }
};
$("#restore").onchange = async (event) => {
  try {
    const file = event.target.files[0];
    if (!file) return;
    if (file.size > 20 * 1024 * 1024)
      throw new Error("Backup is too large (maximum 20 MB).");
    const data = JSON.parse(await file.text());
    if (data.version !== 1 || !Array.isArray(data.entries))
      throw new Error("Choose a Workhours JSON backup.");
    const restored = data.entries.map((e) => {
      validateEntry(e);
      if (typeof e.id !== "string" || !/^[a-zA-Z0-9-]{1,100}$/.test(e.id))
        throw new Error("Invalid entry ID in backup.");
      return {
        id: e.id,
        date: e.date,
        notes: e.notes,
        employee: String(e.employee),
        allocations: e.allocations,
        updatedAt: new Date().toISOString(),
      };
    });
    if (
      !confirm(
        `Merge ${restored.length} entries? Entries with matching IDs will be replaced.`,
      )
    )
      return;
    await importEntries(restored);
    await refresh();
    channel?.postMessage("refresh");
    void syncExcel();
    notify(`${restored.length} entries restored.`);
  } catch (e) {
    notify(`Restore failed: ${e.message}`, true);
  } finally {
    event.target.value = "";
  }
};
function applyTheme(dark) {
  document.documentElement.dataset.theme = dark ? "dark" : "light";
  $("#theme").textContent = dark ? "☀" : "☾";
  $("#theme").setAttribute("aria-pressed", String(dark));
}
applyTheme(
  preference("workhours-theme") === "dark" ||
    (!preference("workhours-theme") &&
      matchMedia("(prefers-color-scheme: dark)").matches),
);
$("#theme").onclick = () => {
  const dark = document.documentElement.dataset.theme !== "dark";
  applyTheme(dark);
  remember("workhours-theme", dark ? "dark" : "light");
};
const channel =
  "BroadcastChannel" in window ? new BroadcastChannel("workhours") : null;
if (channel)
  channel.onmessage = () => {
    void refresh();
  };
async function initialize() {
  try {
    const draft = preference("workhours-draft");
    resetForm(false);
    if (draft) {
      try {
        const parsed = JSON.parse(draft);
        if (
          !parsed ||
          typeof parsed.notes !== "string" ||
          !parsed.allocations ||
          typeof parsed.allocations !== "object"
        )
          throw new Error("Invalid draft");
        fillForm(parsed, true);
      } catch {
        resetForm(false);
        notify(
          "The unfinished draft could not be restored. Your saved entries are unaffected.",
          true,
        );
      }
    }
    await refresh();
  } catch (e) {
    notify(
      `Local storage unavailable: ${e.message}. Enable browser storage before entering hours.`,
      true,
    );
    $("#save").disabled = true;
    return;
  }
  try {
    handle = await getSetting("excel-folder");
    if (handle) {
      $("#link-folder").textContent = "Reconnect Excel folder";
      $("#sync-status").textContent =
        "Folder remembered. Reconnect to resume automatic Excel updates.";
      if ((await handle.queryPermission({ mode: "readwrite" })) === "granted")
        void syncExcel();
    }
  } catch (e) {
    $("#sync-status").textContent =
      `Excel reconnect needed: ${e.message}. Your saved entries are available.`;
  }
  if (navigator.storage?.persist)
    void navigator.storage.persist().catch(() => {});
}
void initialize();
