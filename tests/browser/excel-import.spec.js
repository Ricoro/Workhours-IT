import { test, expect } from "@playwright/test";
import * as XLSX from "@e965/xlsx";
import ExcelJS from "exceljs";
const appPath = process.env.PLAYWRIGHT_APP_PATH || "/";
const header = [
  "Account",
  "Date",
  "Quantity",
  "Employee",
  "Item",
  "Notes",
  "Project",
];
function excel(sheets, type = "xlsb") {
  const book = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets))
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  return XLSX.write(book, { type: "buffer", bookType: type });
}
async function upload(page, buffer, name = "history.xlsb") {
  await page
    .locator("#excel-file")
    .setInputFiles({ name, mimeType: "application/octet-stream", buffer });
  await expect(page.locator("#excel-review")).toBeVisible();
}
async function add(page, date, notes, company = "EC", hours = "2") {
  await page.locator("#date").fill(date);
  await page.locator("#notes").fill(notes);
  await page.getByLabel(`${company} hours`, { exact: true }).fill(hours);
  await page.locator("#save").click();
  await expect(page.locator("#notes")).toHaveValue("");
}
async function records(page) {
  return page.evaluate(
    () =>
      new Promise((resolve, reject) => {
        const r = indexedDB.open("workhours-it", 1);
        r.onsuccess = () => {
          const db = r.result;
          const tx = db.transaction("entries", "readonly");
          const q = tx.objectStore("entries").getAll();
          tx.oncomplete = () => {
            db.close();
            resolve(q.result);
          };
          tx.onerror = () => reject(tx.error);
        };
      }),
  );
}
async function fileBook(page) {
  const bytes = await page.evaluate(async () => {
    const folder = await (
      await navigator.storage.getDirectory()
    ).getDirectoryHandle("import-excel");
    const file = await (await folder.getFileHandle("IT-hours.xlsx")).getFile();
    return Array.from(new Uint8Array(await file.arrayBuffer()));
  });
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(Buffer.from(bytes));
  return book;
}

test("preview historical XLSB, preserve current records, update linked Excel and safely repeat/undo", async ({
  page,
}) => {
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(appPath);
  await add(page, "2026-10-16", "Installed a laptop");
  await add(page, "2026-10-12", "Spam filters created", "4SOME", "1.5");
  const original = await records(page);
  await page.evaluate(async () => {
    const folder = await (
      await navigator.storage.getDirectory()
    ).getDirectoryHandle("import-excel", { create: true });
    window.showDirectoryPicker = async () => folder;
  });
  await page.locator("#link-folder").click();
  await expect(page.locator("#sync-status")).toContainText("Excel updated");
  const buffer = excel({
    "MAART 2023": [
      [null, ...header],
      [
        null,
        "ENAB",
        new Date("2024-03-12T12:00:00Z"),
        0.5,
        12,
        "WH",
        "Maintenance <img src=x onerror=alert(1)>",
        "WH_ENAB",
      ],
      [null, "EC", new Date("2023-03-20T12:00:00Z")],
    ],
    "October 2026": [
      header,
      ["EC", "2026-10-16", 2, 12, "WH", "installed a laptop", "WH_EC"],
      [
        "4SOME",
        "2026-10-12",
        1.5,
        12,
        "WH",
        "Spam filter users configured",
        "WH_4SOME",
      ],
    ],
  });
  await expect(
    page.getByRole("button", { name: "↑ Import Excel", exact: true }),
  ).toBeVisible();
  await upload(page, buffer);
  await expect(page.locator("#excel-confirm")).toHaveText("Import 1 entry");
  await expect(page.locator(".import-counts strong")).toHaveText([
    "1",
    "1",
    "1",
  ]);
  await page
    .getByText("Possible duplicates — skipped unless selected", {
      exact: false,
    })
    .click();
  await expect(page.locator("#excel-review-body")).toContainText(
    "Spam filters created",
  );
  await expect(page.locator("#excel-review img")).toHaveCount(0);
  await page.locator("#excel-confirm").click();
  await expect(page.locator("#toast")).toHaveText(
    "1 Excel entry imported. Existing entries kept.",
  );
  let stored = await records(page);
  expect(stored).toHaveLength(3);
  for (const e of original)
    expect(stored.find((x) => x.id === e.id)).toEqual(e);
  await expect
    .poll(
      async () =>
        (await fileBook(page)).getWorksheet("Summary").getCell("B6").value,
    )
    .toBe(4);
  expect(
    (await fileBook(page)).getWorksheet("March 2023").getCell("C2").value,
  ).toBe(0.5);
  await page.locator("#month").fill("2023-03");
  await expect(page.locator(".task-notes")).toHaveText(
    "Maintenance <img src=x onerror=alert(1)>",
  );
  await expect(page.locator("#activity img")).toHaveCount(0);
  await upload(page, buffer);
  await expect(page.locator("#excel-confirm")).toHaveText("Import 0 entries");
  await expect(page.locator("#excel-confirm")).toBeDisabled();
  await page.locator("#excel-cancel").click();
  await page.locator(".data-tools summary").click();
  page.once("dialog", (d) => d.accept());
  await page.locator("#excel-undo").click();
  await expect(page.locator("#toast")).toContainText("1 entry removed");
  expect(await records(page)).toEqual(original);
  await expect
    .poll(
      async () =>
        (await fileBook(page)).getWorksheet("Summary").getCell("B6").value,
    )
    .toBe(3.5);
  expect((await fileBook(page)).getWorksheet("March 2023")).toBeUndefined();
  expect(errors).toEqual([]);
});

test("date policy, cancel and explicit possible-duplicate inclusion are reviewed before saving", async ({
  page,
}) => {
  await page.goto(appPath);
  await add(page, "2026-10-16", "Installed a laptop");
  let buffer = excel(
    {
      "February 2025": [
        header,
        ["EC", "2025-03-31", 0.5, 12, "WH", "Valid original date", "WH_EC"],
      ],
    },
    "xlsx",
  );
  await upload(page, buffer, "dates.xlsx");
  await expect(page.locator("#excel-confirm")).toBeDisabled();
  await expect(page.locator("#excel-review-body")).toContainText(
    "Rows requiring correction",
  );
  await page.locator("#excel-date-policy").selectOption("original");
  await expect(page.locator("#excel-confirm")).toBeEnabled();
  await page.locator("#excel-cancel").click();
  expect(await records(page)).toHaveLength(1);
  buffer = excel({
    "October 2026": [
      header,
      [
        "EC",
        "2026-10-16",
        2,
        12,
        "WH",
        "A differently described laptop task",
        "WH_EC",
      ],
    ],
  });
  await upload(page, buffer);
  await expect(page.locator("#excel-confirm")).toBeDisabled();
  await page
    .getByText("Possible duplicates — skipped unless selected", {
      exact: false,
    })
    .click();
  await page.getByLabel("Import this row anyway").check();
  await expect(page.locator("#excel-confirm")).toBeEnabled();
  await page.locator("#excel-confirm").click();
  await expect(page.locator("#toast")).toContainText("1 Excel entry imported");
  expect(await records(page)).toHaveLength(2);
});

test("concurrent entry changes reject the whole import without partial writes", async ({
  page,
  context,
}) => {
  await page.goto(appPath);
  const buffer = excel({
    "October 2026": [
      header,
      ["EC", "2026-10-01", 1, 12, "WH", "Import one", "WH_EC"],
      ["ENAB", "2026-10-02", 1, 12, "WH", "Import two", "WH_ENAB"],
    ],
  });
  await upload(page, buffer);
  const other = await context.newPage();
  await other.goto(appPath);
  await add(other, "2026-10-20", "Saved in another tab");
  await page.locator("#excel-confirm").click();
  await expect(page.locator("#excel-import-error")).toContainText(
    "Entries changed while you reviewed",
  );
  expect(await records(page)).toHaveLength(1);
  await expect(page.locator("#excel-confirm")).toBeDisabled();
  await page.locator("#excel-close").click();
  await upload(page, buffer);
  await page.locator("#excel-confirm").click();
  await expect(page.locator("#toast")).toContainText(
    "2 Excel entries imported",
  );
  expect(await records(page)).toHaveLength(3);
});

test("undo keeps subsequently edited imported entries and mobile review fits the screen", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.addInitScript(() =>
    localStorage.setItem("workhours-theme", "dark"),
  );
  await page.goto(appPath);
  const buffer = excel({
    "March 2023": [
      header,
      ["EC", "2023-03-01", 1, 12, "WH", "Edit this later", "WH_EC"],
      ["ENAB", "2023-03-02", 1, 12, "WH", "Untouched", "WH_ENAB"],
    ],
  });
  await upload(page, buffer);
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBeTruthy();
  await page.screenshot({
    path: "/tmp/workhours-import-mobile.png",
    fullPage: true,
  });
  await page.locator("#excel-confirm").click();
  await expect(page.locator("#toast")).toContainText(
    "2 Excel entries imported",
  );
  await page.locator("#month").fill("2023-03");
  await page
    .locator("tr")
    .filter({ hasText: "Edit this later" })
    .getByRole("button", { name: "Edit entry", exact: true })
    .click();
  await page.locator("#notes").fill("Edited after import");
  await page.locator("#save").click();
  await expect(page.locator("#notes")).toHaveValue("");
  await add(page, "2026-10-20", "New manual task");
  await page.locator(".data-tools summary").click();
  page.once("dialog", (d) => d.accept());
  await page.locator("#excel-undo").click();
  await expect(page.locator("#toast")).toContainText("1 edited entry kept");
  const stored = await records(page);
  expect(stored).toHaveLength(2);
  expect(stored.map((e) => e.notes).sort()).toEqual([
    "Edited after import",
    "New manual task",
  ]);
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBeTruthy();
});
