import "./excel-import.css";
import {
  allEntries,
  getSetting,
  commitExcelImport,
  undoExcelImport,
} from "./storage.js";
const escape = (value) =>
  String(value ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        c
      ],
  );
const fmt = (n) =>
  new Intl.NumberFormat("en-GB", { maximumFractionDigits: 2 }).format(n);
const $ = (selector) => document.querySelector(selector);
export function setupExcelImport({ onChanged, notify }) {
  document
    .querySelector("#app")
    .insertAdjacentHTML(
      "beforeend",
      `<input type="file" id="excel-file" accept=".xlsx,.xlsb" hidden><dialog id="excel-review" aria-labelledby="excel-review-title"><div class="import-top"><div><h2 id="excel-review-title">Review Excel import</h2><p id="excel-filename"></p></div><button type="button" class="icon-button" id="excel-close" aria-label="Close import review">×</button></div><p class="import-privacy">Files are read on this device. Nothing is uploaded. Existing entries stay in place.</p><label class="import-policy">When a date disagrees with its monthly tab<select id="excel-date-policy"><option value="worksheet">Use the tab’s month/year, keeping the day</option><option value="original">Keep the date as written in Excel</option></select></label><div id="excel-review-body"></div><p id="excel-import-error" class="import-error" role="alert" hidden></p><p class="import-backup-hint">Download a backup before importing. You can also undo the latest Excel import; entries edited afterwards will be kept.</p><div class="import-actions"><button class="button secondary" id="excel-review-backup">Download existing backup</button><button class="button secondary" id="excel-cancel">Cancel</button><button class="button primary" id="excel-confirm">Import entries</button></div></dialog>`,
    );
  const undo = document.createElement("button");
  undo.className = "button secondary";
  undo.id = "excel-undo";
  undo.textContent = "Undo last Excel import";
  undo.hidden = true;
  document.querySelector(".data-tools > div").append(undo);
  const dialog = $("#excel-review");
  let parsed, baseline, plan, parser;
  let busy = false;
  let includes = new Set();
  const setError = (message) => {
    const el = $("#excel-import-error");
    el.textContent = message || "";
    el.hidden = !message;
  };
  async function updateUndo() {
    try {
      undo.hidden = !(await getSetting("lastExcelImport"));
    } catch {
      undo.hidden = true;
    }
  }
  function render() {
    plan = parser.planExcelImport(parsed, baseline, {
      datePolicy: $("#excel-date-policy").value,
      includePossible: includes,
    });
    const hours = plan.entries.reduce(
      (n, e) => n + Object.values(e.allocations).reduce((a, b) => a + b, 0),
      0,
    );
    const openSections = new Set(
      [...$("#excel-review-body").querySelectorAll("details[open]")].map(
        (el) => el.dataset.section,
      ),
    );
    const details = (title, items, renderItem) =>
      items.length
        ? `<details class="import-details" data-section="${escape(title)}" ${openSections.has(title) ? "open" : ""}><summary>${title} (${items.length})</summary><ul>${items.map(renderItem).join("")}</ul></details>`
        : "";
    $("#excel-review-body").innerHTML =
      `<div class="import-counts"><div><strong>${plan.entries.length}</strong><span>entries to add · ${fmt(hours)}h</span></div><div><strong>${plan.exactDuplicates.length}</strong><span>exact / previous matches skipped</span></div><div><strong>${plan.corrections.length}</strong><span>dates adjusted in new entries</span></div></div>${plan.entries.length === 0 && !plan.errors.length ? "<p>No new entries selected. Your existing entries are unchanged.</p>" : ""}${details("Possible duplicates — skipped unless selected", plan.possibleDuplicates, (x) => `<li><label><input type="checkbox" data-include="${escape(x.row.id)}" ${x.included ? "checked" : ""}> Import this row anyway</label><p><strong>${escape(x.row.company)} · ${escape(x.date)} · ${fmt(x.row.hours)}h</strong></p><p>Excel: ${escape(x.row.notes)}</p><p>Existing: ${escape(x.existingNotes)}</p><small>${escape(x.row.sheet)} · row ${x.row.rowNumber}</small></li>`)}${details("Date corrections", plan.corrections, (x) => `<li><strong>${escape(x.row.sheet)} · row ${x.row.rowNumber}</strong><p>${escape(x.from)} → ${escape(x.to)} · ${escape(x.row.company)} · ${fmt(x.row.hours)}h</p></li>`)}${details("Rows requiring correction — nothing will be imported until resolved", plan.errors, (x) => `<li>${escape(x.row.sheet)} · row ${x.row.rowNumber}: ${escape(x.message)}</li>`)}${details("Unfinished rows skipped (no hours or description)", plan.placeholders, (x) => `<li>${escape(x.sheet)} · row ${x.rowNumber}</li>`)}${plan.ignoredSheets.length ? `<p class="import-meta">Non-work sheets ignored: ${plan.ignoredSheets.map(escape).join(", ")}</p>` : ""}`;
    $("#excel-confirm").disabled = !plan.canImport || busy;
    $("#excel-confirm").textContent =
      `Import ${plan.entries.length} ${plan.entries.length === 1 ? "entry" : "entries"}`;
  }
  function close() {
    if (!busy) {
      dialog.close();
      parsed = null;
      baseline = null;
      plan = null;
    }
  }
  $("#excel-close").onclick = close;
  $("#excel-cancel").onclick = close;
  dialog.addEventListener("cancel", (event) => {
    if (busy) event.preventDefault();
  });
  $("#excel-date-policy").onchange = () => {
    setError("");
    render();
  };
  $("#excel-review-body").onchange = (event) => {
    const id = event.target.dataset.include;
    if (!id) return;
    event.target.checked ? includes.add(id) : includes.delete(id);
    render();
  };
  $("#excel-open").onclick = () => $("#excel-file").click();
  $("#excel-file").onchange = async (event) => {
    const file = event.target.files[0];
    if (!file) return;
    const opener = $("#excel-open");
    opener.disabled = true;
    try {
      if (!/\.(xlsx|xlsb)$/i.test(file.name))
        throw new Error("Choose a .xlsx or .xlsb workbook.");
      if (file.size > 20 * 1024 * 1024)
        throw new Error("Choose an Excel file up to 20 MB.");
      notify("Reading Excel locally…");
      parser ||= await import("./excel-import.js");
      parsed = await parser.readExcelImport(await file.arrayBuffer());
      baseline = await allEntries();
      includes = new Set();
      $("#excel-date-policy").value = "worksheet";
      $("#excel-filename").textContent = file.name;
      setError("");
      render();
      dialog.showModal();
    } catch (e) {
      notify(`Excel import failed: ${e.message}`, true);
    } finally {
      opener.disabled = false;
      event.target.value = "";
    }
  };
  $("#excel-review-backup").onclick = () => $("#backup").click();
  $("#excel-confirm").onclick = async () => {
    if (!plan?.canImport || busy) return;
    busy = true;
    dialog
      .querySelectorAll("button,input,select")
      .forEach((el) => (el.disabled = true));
    setError("");
    try {
      const count = await commitExcelImport(plan.entries, baseline);
      dialog.close();
      parsed = null;
      baseline = null;
      plan = null;
      await updateUndo();
      try {
        await onChanged();
      } catch {
        notify(
          `${count} entries saved locally. Reload the page to refresh the view.`,
          true,
        );
        return;
      }
      notify(
        `${count} Excel ${count === 1 ? "entry" : "entries"} imported. Existing entries kept.`,
      );
    } catch (e) {
      setError(e.message);
    } finally {
      busy = false;
      dialog
        .querySelectorAll("button,input,select")
        .forEach((el) => (el.disabled = false));
      if (plan) $("#excel-confirm").disabled = true;
    }
  };
  undo.onclick = async () => {
    if (
      !confirm(
        "Undo the latest Excel import? Only unchanged entries from that import will be removed. Entries you edited afterwards will be kept.",
      )
    )
      return;
    undo.disabled = true;
    try {
      const { removed, kept } = await undoExcelImport();
      await onChanged();
      await updateUndo();
      notify(
        `Excel import undone: ${removed} ${removed === 1 ? "entry" : "entries"} removed.${kept ? ` ${kept} edited ${kept === 1 ? "entry" : "entries"} kept.` : ""}`,
      );
    } catch (e) {
      notify(e.message, true);
    } finally {
      undo.disabled = false;
    }
  };
  void updateUndo();
  return { updateUndo };
}
