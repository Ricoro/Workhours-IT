import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { validateEntry, totals, rows, validDate } from "../src/model.js";
import { workbook } from "../src/excel.js";
const entry = {
  id: "test",
  date: "2026-10-20",
  notes: "Laptop & email support",
  employee: "12",
  allocations: { EC: 1.25, ENAB: 0.5 },
  updatedAt: "2026-10-20T12:00:00Z",
};
test("allocations are validated and accumulated per company", () => {
  validateEntry(entry);
  assert.deepEqual(totals([entry, entry]), {
    EC: 2.5,
    ENAB: 1,
    "4SOME": 0,
    SEABOW: 0,
  });
  assert.equal(rows([entry]).length, 2);
  for (const allocations of [
    { EC: -1 },
    { EC: 0 },
    { EC: NaN },
    { BAD: 1 },
    { EC: 20, ENAB: 5 },
    { EC: 0.001 },
    {},
  ])
    assert.throws(() => validateEntry({ ...entry, allocations }));
  assert.equal(validDate("2026-02-30"), false);
  assert.throws(() => validateEntry({ ...entry, notes: " " }));
});
test("Excel roundtrip preserves original columns, dates, hours and untrusted text", async () => {
  const other = {
    ...entry,
    id: "other",
    date: "2026-11-01",
    allocations: { SEABOW: 2 },
    notes: '=HYPERLINK("https://example.com")',
  };
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(await workbook([entry, other]));
  assert.deepEqual(
    book.worksheets.map((s) => s.name),
    ["Summary", "2026-10", "2026-11"],
  );
  const sheet = book.getWorksheet("2026-10");
  assert.deepEqual(sheet.getRow(1).values.slice(1, 8), [
    "Account",
    "Date",
    "Quantity",
    "Employee",
    "Item",
    "Notes",
    "Project",
  ]);
  assert.equal(sheet.getCell("C2").value, 1.25);
  assert.equal(
    sheet.getCell("B2").value.toISOString().slice(0, 10),
    "2026-10-20",
  );
  assert.equal(sheet.getCell("G2").value, "WH_EC");
  assert.equal(sheet.getCell("J6").value, 1.75);
  assert.equal(book.getWorksheet("Summary").getCell("B6").value, 3.75);
  assert.equal(
    book.getWorksheet("2026-11").getCell("F2").type,
    ExcelJS.ValueType.String,
  );
  const monthly = new ExcelJS.Workbook();
  await monthly.xlsx.load(await workbook([entry, other], "2026-10"));
  assert.deepEqual(
    monthly.worksheets.map((s) => s.name),
    ["Summary", "2026-10"],
  );
  assert.equal(monthly.getWorksheet("Summary").getCell("B6").value, 1.75);
});

test("malformed dates fail safely and decimal allocations accept an exact 24 hours", () => {
  for (const date of [
    "2026-99-99",
    "2026-00-20",
    "2026-02-29",
    "1900-02-29",
    null,
    12,
  ])
    assert.equal(validDate(date), false);
  assert.equal(validDate("2024-02-29"), true);
  assert.equal(validDate("9999-12-31"), true);
  assert.throws(() => validateEntry(null), /Choose a valid work date/);
  assert.throws(
    () => validateEntry({ ...entry, date: "2026-99-99" }),
    /Choose a valid work date/,
  );
  validateEntry({
    ...entry,
    allocations: { EC: 0.01, ENAB: 23.98, "4SOME": 0.01 },
  });
  for (const employee of [{}, [], NaN, Infinity, "   "])
    assert.throws(() => validateEntry({ ...entry, employee }), /employee ID/);
});
