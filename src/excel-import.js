import * as XLSX from "@e965/xlsx";
import { COMPANIES, EMPLOYEE_ID, validDate, validateEntry } from "./model.js";

const MONTHS = {
  januari: 1,
  january: 1,
  februari: 2,
  february: 2,
  maart: 3,
  march: 3,
  april: 4,
  mei: 5,
  may: 5,
  juni: 6,
  june: 6,
  juli: 7,
  july: 7,
  augustus: 8,
  august: 8,
  september: 9,
  oktober: 10,
  october: 10,
  november: 11,
  december: 12,
};
const REQUIRED = ["account", "date", "quantity", "notes"];
const blank = (value) =>
  value === null || value === undefined || String(value).trim() === "";
const normalized = (value) =>
  String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("en-GB");
function excelDate(value, date1904) {
  if (value instanceof Date && Number.isFinite(value.getTime()))
    return value.toISOString().slice(0, 10);
  if (typeof value === "number" && Number.isFinite(value)) {
    const d = XLSX.SSF.parse_date_code(value, { date1904 });
    if (d)
      return `${String(d.y).padStart(4, "0")}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  if (typeof value === "string") {
    const s = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const match = s.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})$/);
    if (match)
      return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  }
  return null;
}
function sheetMonth(name, rows) {
  const iso = name.match(/^(\d{4})-(0[1-9]|1[0-2])$/);
  if (iso) return iso[0];
  const parts = name.trim().toLowerCase().split(/\s+/);
  const month = MONTHS[parts[0]];
  if (!month) return null;
  let year = /^\d{4}$/.test(parts[1] || "") ? parts[1] : null;
  if (!year && parts.length === 1) {
    const years = new Map();
    for (const r of rows)
      if (validDate(r.originalDate)) {
        const y = r.originalDate.slice(0, 4);
        years.set(y, (years.get(y) || 0) + 1);
      }
    year = [...years].sort((a, b) => b[1] - a[1])[0]?.[0];
  }
  return year ? `${year}-${String(month).padStart(2, "0")}` : null;
}
export async function readExcelImport(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  if (!bytes.byteLength || bytes.byteLength > 20 * 1024 * 1024)
    throw new Error("Choose an Excel file up to 20 MB.");
  let book;
  try {
    book = XLSX.read(bytes, {
      type: "array",
      cellDates: false,
      sheetRows: 10002,
      cellHTML: false,
      cellFormula: false,
      bookVBA: false,
    });
  } catch {
    throw new Error(
      "Could not read this workbook. Choose an unencrypted .xlsx or .xlsb file.",
    );
  }
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  const hash = [...new Uint8Array(digest)]
    .map((n) => n.toString(16).padStart(2, "0"))
    .join("");
  const result = { rows: [], placeholders: [], ignoredSheets: [] };
  let recognized = 0;
  const date1904 =
    book.Workbook?.WBProps?.date1904 === true ||
    book.Workbook?.WBProps?.date1904 === 1;
  for (let index = 0; index < book.SheetNames.length; index++) {
    const name = book.SheetNames[index];
    const sheet = book.Sheets[name];
    if (!sheet?.["!ref"]) {
      result.ignoredSheets.push(name);
      continue;
    }
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const grid = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: true,
    });
    const headerIndex = grid
      .slice(0, 20)
      .findIndex((row) =>
        REQUIRED.every((k) => row.some((v) => normalized(v) === k)),
      );
    if (headerIndex < 0) {
      result.ignoredSheets.push(name);
      continue;
    }
    recognized++;
    if (
      sheet["!fullref"] &&
      XLSX.utils.decode_range(sheet["!fullref"]).e.r > range.e.r
    )
      throw new Error(
        `${name}: more than 10,000 rows. Split the workbook before importing.`,
      );
    const header = Object.fromEntries(
      grid[headerIndex].map((v, i) => [normalized(v), i]),
    );
    const rows = [];
    for (let i = headerIndex + 1; i < grid.length; i++) {
      const values = grid[i];
      const rowNumber = range.s.r + i + 1;
      if (REQUIRED.every((k) => blank(values[header[k]]))) continue;
      const location = { sheet: name, rowNumber };
      if (blank(values[header.quantity]) && blank(values[header.notes])) {
        result.placeholders.push(location);
        continue;
      }
      const quantity = values[header.quantity];
      const hours = blank(quantity)
        ? NaN
        : typeof quantity === "number"
          ? quantity
          : Number(String(quantity).trim().replace(",", "."));
      rows.push({
        ...location,
        id: `excel-${hash}-${index}-${rowNumber}`,
        company: String(values[header.account] ?? "")
          .trim()
          .toUpperCase(),
        originalDate: excelDate(values[header.date], date1904),
        hours,
        notes: String(values[header.notes] ?? "").trim(),
      });
      if (result.rows.length + rows.length > 30000)
        throw new Error(
          "This import exceeds 30,000 work rows. Split it into smaller files.",
        );
    }
    const month = sheetMonth(name, rows);
    result.rows.push(...rows.map((r) => ({ ...r, sheetMonth: month })));
  }
  if (!recognized)
    throw new Error(
      "No work-hour sheets found. The file needs Account, Date, Quantity and Notes columns.",
    );
  return result;
}
function matchKey(company, date, hours, notes) {
  return JSON.stringify([
    company,
    date,
    Math.round(hours * 100),
    normalized(notes),
  ]);
}
export function planExcelImport(
  parsed,
  existing,
  { datePolicy = "worksheet", includePossible = new Set() } = {},
) {
  const slots = existing.flatMap((e) =>
    COMPANIES.filter((c) => e.allocations[c]).map((c) => ({
      id: e.id,
      company: c,
      date: e.date,
      hours: e.allocations[c],
      key: matchKey(c, e.date, e.allocations[c], e.notes),
      notes: e.notes,
      used: false,
    })),
  );
  const byID = new Set(existing.map((e) => e.id));
  const plan = {
    entries: [],
    exactDuplicates: [],
    possibleDuplicates: [],
    corrections: [],
    errors: [],
    placeholders: parsed.placeholders,
    ignoredSheets: parsed.ignoredSheets,
  };
  const timestamp = new Date().toISOString();
  for (const row of parsed.rows) {
    let date = row.originalDate;
    if (datePolicy === "worksheet" && row.sheetMonth && validDate(date))
      date = `${row.sheetMonth}-${date.slice(8)}`;
    const entry = {
      id: row.id,
      date,
      employee: EMPLOYEE_ID,
      notes: row.notes,
      allocations: { [row.company]: row.hours },
      updatedAt: timestamp,
    };
    try {
      if (!COMPANIES.includes(row.company)) throw new Error("Unknown company.");
      validateEntry(entry);
    } catch (e) {
      plan.errors.push({ row, message: e.message });
      continue;
    }
    const correction =
      date !== row.originalDate
        ? { row, from: row.originalDate, to: date }
        : null;
    if (byID.has(row.id)) {
      slots.filter((s) => s.id === row.id).forEach((s) => (s.used = true));
      plan.exactDuplicates.push({ row, reason: "Previously imported row" });
      continue;
    }
    const exact = slots.find(
      (s) =>
        !s.used && s.key === matchKey(row.company, date, row.hours, row.notes),
    );
    if (exact) {
      exact.used = true;
      plan.exactDuplicates.push({
        row,
        reason: "Same company, date, hours and description",
      });
      continue;
    }
    const possible = slots.find(
      (s) =>
        !s.used &&
        s.company === row.company &&
        s.date === date &&
        Math.round(s.hours * 100) === Math.round(row.hours * 100),
    );
    if (possible) {
      possible.used = true;
      plan.possibleDuplicates.push({
        row,
        existingNotes: possible.notes,
        date,
        included: includePossible.has(row.id),
      });
      if (!includePossible.has(row.id)) continue;
    }
    plan.entries.push(entry);
    if (correction) plan.corrections.push(correction);
  }
  plan.canImport = plan.errors.length === 0 && plan.entries.length > 0;
  return plan;
}
