import ExcelJS from "exceljs";
import { COMPANIES, totals, rows } from "./model.js";
export async function workbook(entries, month) {
  const book = new ExcelJS.Workbook();
  book.creator = "Workhours IT";
  book.created = new Date();
  const selected = month
    ? entries.filter((e) => e.date.startsWith(month))
    : entries;
  const summary = book.addWorksheet("Summary");
  summary.addRow(["Company", "Hours"]);
  const t = totals(selected);
  for (const c of COMPANIES) summary.addRow([c, t[c]]);
  summary.addRow(["TOTAL", Object.values(t).reduce((a, b) => a + b, 0)]);
  summary.getColumn(1).width = 24;
  summary.getColumn(2).width = 18;
  summary.getColumn(2).numFmt = "0.00";
  const months = month
    ? [month]
    : [...new Set(selected.map((e) => e.date.slice(0, 7)))].sort();
  for (const m of months) {
    const sheet = book.addWorksheet(m);
    sheet.addRow([
      "Account",
      "Date",
      "Quantity",
      "Employee",
      "Item",
      "Notes",
      "Project",
    ]);
    for (const row of rows(selected.filter((e) => e.date.startsWith(m))))
      sheet.addRow(row);
    [15, 16, 14, 14, 12, 80, 20].forEach(
      (width, i) => (sheet.getColumn(i + 1).width = width),
    );
    sheet.getColumn(2).numFmt = "dd/mm/yyyy";
    sheet.getColumn(3).numFmt = "0.00";
    sheet.getColumn(6).alignment = { wrapText: true, vertical: "top" };
    sheet.autoFilter = { from: "A1", to: `G${Math.max(1, sheet.rowCount)}` };
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    const mt = totals(selected.filter((e) => e.date.startsWith(m)));
    sheet.getCell("I1").value = "Company";
    sheet.getCell("J1").value = "Hours";
    COMPANIES.forEach((c, i) => {
      sheet.getCell(`I${i + 2}`).value = c;
      sheet.getCell(`J${i + 2}`).value = mt[c];
    });
    sheet.getCell("I6").value = "TOTAL";
    sheet.getCell("J6").value = Object.values(mt).reduce((a, b) => a + b, 0);
    sheet.getColumn(9).width = 18;
    sheet.getColumn(10).numFmt = "0.00";
  }
  for (const sheet of book.worksheets) {
    sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
    sheet.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF173F35" },
    };
    sheet.getRow(1).height = 28;
  }
  return book.xlsx.writeBuffer();
}
