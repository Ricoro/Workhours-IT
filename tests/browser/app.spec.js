import { test, expect } from "@playwright/test";
import ExcelJS from "exceljs";
const appPath = process.env.PLAYWRIGHT_APP_PATH || "/";
test("save, reload, edit, filter, Excel export, dark mode and delete", async ({
  page,
}) => {
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.goto(appPath);
  await page.locator("#date").fill("2026-10-20");
  await page.locator("#notes").fill("Set up laptop <script>alert(1)</script>");
  await page.getByLabel("EC hours", { exact: true }).fill("1.25");
  await page.getByLabel("ENAB hours", { exact: true }).fill("0.5");
  await expect(page.locator("#entry-hours")).toHaveText("1.75");
  await page.reload();
  await expect(page.locator("#notes")).toHaveValue(
    "Set up laptop <script>alert(1)</script>",
  );
  await page.locator("#save").click();
  await expect(page.locator("#count")).toHaveText("1");
  await expect(page.locator(".task-notes")).toHaveText(
    "Set up laptop <script>alert(1)</script>",
  );
  await page.reload();
  await page.locator("#month").fill("2026-10");
  await expect(page.locator("#count")).toHaveText("1");
  await expect(page.locator(".total-stat>strong")).toHaveText("1.75h");
  await page.getByRole("button", { name: "Edit entry", exact: true }).click();
  await page.getByLabel("EC hours", { exact: true }).fill("2");
  await page.locator("#save").click();
  await expect(page.locator(".total-stat>strong")).toHaveText("2.5h");
  await page.locator("#company-filter").selectOption("SEABOW");
  await expect(page.locator(".empty h3")).toHaveText("No matching entries");
  await page.locator("#company-filter").selectOption("all");
  const downloadPromise = page.waitForEvent("download");
  await page.locator("#export").click();
  const file = await downloadPromise;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(await file.path());
  expect(wb.getWorksheet("2026-10").getCell("C2").value).toBe(2);
  expect(wb.getWorksheet("Summary").getCell("B6").value).toBe(2.5);
  await page.locator("#theme").click();
  const theme = await page.locator("html").getAttribute("data-theme");
  await page.reload();
  await expect(page.locator("html")).toHaveAttribute("data-theme", theme);
  await page.locator("#month").fill("2026-10");
  await page.getByRole("button", { name: "Delete entry", exact: true }).click();
  await page.locator("#confirm-delete").click();
  await expect(page.locator("#count")).toHaveText("0");
  expect(errors).toEqual([]);
});
test("draft, responsive layout, backup restore and invalid input", async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto(appPath);
  await expect(page.locator("h1")).toBeVisible();
  expect(
    await page.evaluate(
      () => document.documentElement.scrollWidth <= innerWidth,
    ),
  ).toBeTruthy();
  await page.locator("#notes").fill("Network maintenance");
  await page.getByLabel("4SOME hours", { exact: true }).fill("1.5");
  await page.reload();
  await expect(page.getByLabel("4SOME hours", { exact: true })).toHaveValue(
    "1.5",
  );
  await page.locator("#save").click();
  await expect(page.locator("#count")).toHaveText("1");
  await page.locator("summary").click();
  const downloaded = page.waitForEvent("download");
  await page.locator("#backup").click();
  const backup = await downloaded;
  await page.getByRole("button", { name: "Delete entry", exact: true }).click();
  await page.locator("#confirm-delete").click();
  await expect(page.locator("#count")).toHaveText("0");
  page.once("dialog", (d) => d.accept());
  await page.locator("#restore").setInputFiles(await backup.path());
  await expect(page.locator("#count")).toHaveText("1");
  await page.locator("#notes").fill("Invalid total");
  await page.getByLabel("EC hours", { exact: true }).fill("20");
  await page.getByLabel("ENAB hours", { exact: true }).fill("10");
  await page.locator("#save").click();
  await expect(page.locator("#toast")).toContainText("no more than 24");
  await expect(page.locator("#count")).toHaveText("1");
  await page.screenshot({ path: "/tmp/workhours-mobile.png", fullPage: true });
});

test("linked Excel workbook updates on save, edit and delete", async ({
  page,
}) => {
  await page.goto(appPath);
  await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const folder = await root.getDirectoryHandle("excel-test", {
      create: true,
    });
    window.showDirectoryPicker = async () => folder;
  });
  await page.locator("#link-folder").click();
  await expect(page.locator("#sync-status")).toContainText("Excel updated");
  await page.locator("#date").fill("2026-10-20");
  await page.locator("#notes").fill("Linked folder task");
  await page.getByLabel("SEABOW hours", { exact: true }).fill("2.5");
  await page.locator("#save").click();
  await expect(page.locator("#sync-status")).toContainText("Excel updated");
  async function readExcel() {
    const bytes = await page.evaluate(async () => {
      const root = await navigator.storage.getDirectory();
      const folder = await root.getDirectoryHandle("excel-test");
      const handle = await folder.getFileHandle("IT-hours.xlsx");
      return Array.from(
        new Uint8Array(await (await handle.getFile()).arrayBuffer()),
      );
    });
    const book = new ExcelJS.Workbook();
    await book.xlsx.load(Buffer.from(bytes));
    return book;
  }
  await expect
    .poll(
      async () =>
        (await readExcel()).getWorksheet("Summary").getCell("B6").value,
    )
    .toBe(2.5);
  let book = await readExcel();
  expect(book.getWorksheet("2026-10").getCell("A2").value).toBe("SEABOW");
  expect(book.getWorksheet("Summary").getCell("B6").value).toBe(2.5);
  await page.getByRole("button", { name: "Edit entry", exact: true }).click();
  await page.getByLabel("SEABOW hours", { exact: true }).fill("3");
  await page.locator("#save").click();
  await expect(page.locator("#sync-status")).toContainText("Excel updated");
  await expect
    .poll(
      async () =>
        (await readExcel()).getWorksheet("Summary").getCell("B6").value,
    )
    .toBe(3);
  book = await readExcel();
  expect(book.getWorksheet("Summary").getCell("B6").value).toBe(3);
  await page.getByRole("button", { name: "Delete entry", exact: true }).click();
  await page.locator("#confirm-delete").click();
  await expect(page.locator("#sync-status")).toContainText("Excel updated");
  await expect
    .poll(
      async () =>
        (await readExcel()).getWorksheet("Summary").getCell("B6").value,
    )
    .toBe(0);
  book = await readExcel();
  expect(book.getWorksheet("Summary").getCell("B6").value).toBe(0);
  expect(book.getWorksheet("2026-10")).toBeUndefined();
});

test("saving entries still works when localStorage preferences and drafts are unavailable", async ({
  page,
}) => {
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.addInitScript(() => {
    Storage.prototype.getItem = () => {
      throw new DOMException("Storage denied", "SecurityError");
    };
    Storage.prototype.setItem = () => {
      throw new DOMException("Storage denied", "SecurityError");
    };
    Storage.prototype.removeItem = () => {
      throw new DOMException("Storage denied", "SecurityError");
    };
  });
  await page.goto(appPath);
  await expect(page.locator("#save")).toBeEnabled();
  await page.locator("#date").fill("2026-10-20");
  await page.locator("#notes").fill("Saved despite unavailable preferences");
  await page.getByLabel("EC hours", { exact: true }).fill("1");
  await expect(page.locator("#draft-status")).toContainText(
    "Draft could not be saved",
  );
  await page.locator("#save").click();
  await expect(page.locator("#count")).toHaveText("1");
  await expect(page.locator("#notes")).toHaveValue("");
  await expect(page.locator("#toast")).toHaveText(
    "Entry saved on this device.",
  );
  await page.locator("#theme").click();
  await page.reload();
  await page.locator("#month").fill("2026-10");
  await expect(page.locator("#count")).toHaveText("1");
  expect(errors).toEqual([]);
});

test("invalid backups are rejected atomically and corrupt drafts do not block saved entries", async ({
  page,
}) => {
  await page.goto(appPath);
  await page.locator("#date").fill("2026-10-20");
  await page.locator("#notes").fill("Keep this task");
  await page.getByLabel("EC hours", { exact: true }).fill("1");
  await page.locator("#save").click();
  await expect(page.locator("#count")).toHaveText("1");
  await page.locator("summary").click();
  const good = {
    id: "new-backup-task",
    date: "2026-10-21",
    notes: "Must not be partially imported",
    employee: "12",
    allocations: { ENAB: 2 },
  };
  await page.locator("#restore").setInputFiles({
    name: "invalid-backup.json",
    mimeType: "application/json",
    buffer: Buffer.from(
      JSON.stringify({
        version: 1,
        entries: [good, { ...good, id: "bad", date: "2026-99-99" }],
      }),
    ),
  });
  await expect(page.locator("#toast")).toContainText(
    "Choose a valid work date",
  );
  await expect(page.locator("#count")).toHaveText("1");
  await page.evaluate(() =>
    localStorage.setItem("workhours-draft", "not valid JSON"),
  );
  await page.reload();
  await page.locator("#month").fill("2026-10");
  await expect(page.locator("#count")).toHaveText("1");
  await expect(page.locator("#save")).toBeEnabled();
  await expect(page.locator("#toast")).toContainText(
    "draft could not be restored",
  );
});

test("Excel write failure preserves entries and reconnect retries the file", async ({
  page,
}) => {
  await page.goto(appPath);
  await page.evaluate(async () => {
    const folder = await (
      await navigator.storage.getDirectory()
    ).getDirectoryHandle("retry-excel", { create: true });
    window.showDirectoryPicker = async () => folder;
    window.originalCreateWritable =
      FileSystemFileHandle.prototype.createWritable;
    FileSystemFileHandle.prototype.createWritable = async () => {
      throw new Error("Workbook is locked");
    };
  });
  await page.locator("#link-folder").click();
  await expect(page.locator("#sync-status")).toContainText(
    "Workbook is locked",
  );
  await page.locator("#date").fill("2026-10-20");
  await page.locator("#notes").fill("Saved while Excel is locked");
  await page.getByLabel("ENAB hours", { exact: true }).fill("1.25");
  await page.locator("#save").click();
  await expect(page.locator("#count")).toHaveText("1");
  await expect(page.locator("#sync-status")).toContainText(
    "Your entries remain saved",
  );
  await page.evaluate(
    () =>
      (FileSystemFileHandle.prototype.createWritable =
        window.originalCreateWritable),
  );
  await page.locator("#link-folder").click();
  await expect(page.locator("#sync-status")).toContainText("Excel updated");
  const bytes = await page.evaluate(async () => {
    const root = await navigator.storage.getDirectory();
    const folder = await root.getDirectoryHandle("retry-excel");
    const file = await (await folder.getFileHandle("IT-hours.xlsx")).getFile();
    return Array.from(new Uint8Array(await file.arrayBuffer()));
  });
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(Buffer.from(bytes));
  expect(book.getWorksheet("Summary").getCell("B6").value).toBe(1.25);
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  await page.addInitScript(() => {
    FileSystemDirectoryHandle.prototype.queryPermission = async () => {
      throw new Error("Folder access temporarily unavailable");
    };
    navigator.storage.persist = async () => {
      throw new Error("Persistence denied");
    };
  });
  await page.reload();
  await expect(page.locator("#sync-status")).toContainText(
    "Excel reconnect needed",
  );
  await page.locator("#month").fill("2026-10");
  await expect(page.locator("#count")).toHaveText("1");
  await expect(page.locator("#save")).toBeEnabled();
  expect(errors).toEqual([]);
});

test("linking a folder preserves an existing workbook when overwrite is declined", async ({
  page,
}) => {
  await page.goto(appPath);
  await page.evaluate(async () => {
    const folder = await (
      await navigator.storage.getDirectory()
    ).getDirectoryHandle("existing-excel", { create: true });
    const file = await folder.getFileHandle("IT-hours.xlsx", { create: true });
    const writer = await file.createWritable();
    await writer.write("existing confidential workbook");
    await writer.close();
    window.showDirectoryPicker = async () => folder;
  });
  page.once("dialog", (d) => d.dismiss());
  await page.locator("#link-folder").click();
  await expect(page.locator("#sync-status")).toContainText("No folder linked");
  expect(
    await page.evaluate(async () => {
      const folder = await (
        await navigator.storage.getDirectory()
      ).getDirectoryHandle("existing-excel");
      return (
        await (await folder.getFileHandle("IT-hours.xlsx")).getFile()
      ).text();
    }),
  ).toBe("existing confidential workbook");
});
