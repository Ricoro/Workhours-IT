import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "@e965/xlsx";
import { readExcelImport, planExcelImport } from "../src/excel-import.js";
const headers = [
  "Account",
  "Date",
  "Quantity",
  "Employee",
  "Item",
  "Notes",
  "Project",
];
function file(sheets, type = "xlsx", props = {}) {
  const book = XLSX.utils.book_new();
  book.Workbook = { WBProps: props };
  for (const [name, rows] of Object.entries(sheets))
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(rows), name);
  return XLSX.write(book, { type: "buffer", bookType: type });
}
const current = {
  id: "existing-1",
  date: "2026-10-16",
  employee: 12,
  notes: "Installed a laptop",
  allocations: { EC: 2 },
  updatedAt: "2026-10-20T12:00:00Z",
};
test("reads XLSB with shifted columns, Dutch months, serial dates, and unfinished rows", async () => {
  const input = file(
    {
      "MAART 2023": [
        [null, ...headers],
        [
          null,
          "EC",
          new Date("2024-03-12T12:00:00Z"),
          0.5,
          12,
          "WH",
          "Onderhoud é & <test>",
          "WH_EC",
        ],
        [null, "ENAB", new Date("2023-03-15T12:00:00Z")],
      ],
      Summary: [
        ["Company", "Hours"],
        ["EC", 0.5],
      ],
    },
    "xlsb",
  );
  const parsed = await readExcelImport(input);
  assert.equal(parsed.rows.length, 1);
  assert.equal(parsed.placeholders.length, 1);
  assert.equal(parsed.rows[0].originalDate, "2024-03-12");
  const plan = planExcelImport(parsed, [], { datePolicy: "worksheet" });
  assert.equal(plan.entries[0].date, "2023-03-12");
  assert.equal(plan.entries[0].notes, "Onderhoud é & <test>");
  assert.equal(plan.entries[0].employee, 12);
  assert.equal(plan.corrections.length, 1);
  assert.equal(plan.errors.length, 0);
  assert.equal(
    planExcelImport(parsed, [], { datePolicy: "original" }).entries[0].date,
    "2024-03-12",
  );
});
test("skips exact matches, previews reworded matches and allows explicitly importing them", async () => {
  const parsed = await readExcelImport(
    file({
      "October 2026": [
        headers,
        ["EC", "2026-10-16", 2, 12, "WH", "installed a laptop", "WH_EC"],
        [
          "4SOME",
          "2026-10-12",
          1.5,
          12,
          "WH",
          "Configured spam filtering",
          "WH_4SOME",
        ],
      ],
    }),
  );
  const other = {
    ...current,
    id: "existing-2",
    date: "2026-10-12",
    notes: "Spam filter users created",
    allocations: { "4SOME": 1.5 },
  };
  const before = JSON.stringify([current, other]);
  const plan = planExcelImport(parsed, [current, other]);
  assert.equal(plan.exactDuplicates.length, 1);
  assert.equal(plan.possibleDuplicates.length, 1);
  assert.equal(plan.entries.length, 0);
  assert.equal(JSON.stringify([current, other]), before);
  const include = new Set([plan.possibleDuplicates[0].row.id]);
  assert.equal(
    planExcelImport(parsed, [current, other], { includePossible: include })
      .entries.length,
    1,
  );
});
test("reimport is idempotent and preserves legitimate repeated source rows", async () => {
  const input = file({
    "2026-10": [
      headers,
      ["EC", "2026-10-16", 2, 12, "WH", "Installed a laptop", "WH_EC"],
      ["EC", "2026-10-16", 2, 12, "WH", "Installed a laptop", "WH_EC"],
    ],
  });
  const parsed = await readExcelImport(input);
  const plan = planExcelImport(parsed, [current]);
  assert.equal(plan.entries.length, 1);
  const again = planExcelImport(await readExcelImport(input), [
    current,
    ...plan.entries,
  ]);
  assert.equal(again.entries.length, 0);
  const edited = { ...plan.entries[0], notes: "Later edited description" };
  assert.equal(planExcelImport(parsed, [current, edited]).entries.length, 0);
});
test("invalid source rows and impossible date corrections block the entire plan", async () => {
  const parsed = await readExcelImport(
    file({
      "February 2025": [
        headers,
        ["EC", "2025-03-31", 1, 12, "WH", "Valid original date", "WH_EC"],
        ["ENAB", "2025-02-12", -2, 12, "WH", "Invalid hours", "WH_ENAB"],
      ],
    }),
  );
  const plan = planExcelImport(parsed, []);
  assert.equal(plan.errors.length, 2);
  assert.equal(plan.canImport, false);
  const original = planExcelImport(parsed, [], { datePolicy: "original" });
  assert.equal(original.errors.length, 1);
  assert.equal(original.canImport, false);
});
test("supports split-company app exports and infers year for old month-only sheets", async () => {
  const parsed = await readExcelImport(
    file({
      MAART: [
        headers,
        ["EC", "2022-03-10", 1, 12, "WH", "Shared maintenance", "WH_EC"],
        ["ENAB", "2022-03-10", 0.5, 12, "WH", "Shared maintenance", "WH_ENAB"],
      ],
    }),
  );
  const existing = {
    ...current,
    date: "2022-03-10",
    notes: "Shared maintenance",
    allocations: { EC: 1, ENAB: 0.5 },
  };
  const plan = planExcelImport(parsed, [existing]);
  assert.equal(plan.entries.length, 0);
  assert.equal(plan.exactDuplicates.length, 2);
  assert.equal(plan.corrections.length, 0);
});
test("supports 1904 date systems and rejects workbooks without work-hour headers", async () => {
  const parsed = await readExcelImport(
    file(
      {
        "January 1904": [
          headers,
          ["EC", 0, 0.5, 12, "WH", "Historical date", "WH_EC"],
        ],
      },
      "xlsx",
      { date1904: true },
    ),
  );
  assert.equal(parsed.rows[0].originalDate, "1904-01-01");
  await assert.rejects(
    () =>
      readExcelImport(
        file({
          Other: [
            ["Something", "Else"],
            [1, 2],
          ],
        }),
      ),
    /Account.*Date.*Quantity.*Notes/,
  );
});
