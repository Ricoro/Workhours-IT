export const EMPLOYEE_ID = 12;
export const COMPANIES = ["EC", "ENAB", "4SOME", "SEABOW"];
export function localDate() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
export function validDate(value) {
  if (
    typeof value !== "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(value) ||
    value < "1900-01-01" ||
    value > "9999-12-31"
  )
    return false;
  const date = new Date(`${value}T00:00:00Z`);
  return (
    Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value
  );
}
export function validateEntry(entry) {
  if (!entry || !validDate(entry.date))
    throw new Error("Choose a valid work date.");
  if (
    typeof entry.notes !== "string" ||
    !entry.notes.trim() ||
    entry.notes.length > 10000
  )
    throw new Error("Describe the work (up to 10,000 characters).");
  if (
    !entry.allocations ||
    Object.keys(entry.allocations).some((c) => !COMPANIES.includes(c))
  )
    throw new Error("Unknown company.");
  const hours = Object.values(entry.allocations);
  if (
    !hours.length ||
    hours.some(
      (h) =>
        typeof h !== "number" ||
        !Number.isFinite(h) ||
        h <= 0 ||
        h > 24 ||
        Math.abs(h * 100 - Math.round(h * 100)) > 0.00001,
    ) ||
    hours.reduce((a, b) => a + Math.round(b * 100), 0) > 2400
  )
    throw new Error(
      "Enter hours for at least one company. Use up to 2 decimals and no more than 24 hours per entry.",
    );
  return entry;
}
export function totals(entries) {
  return Object.fromEntries(
    COMPANIES.map((c) => [
      c,
      Math.round(
        entries.reduce((n, e) => n + (e.allocations[c] || 0), 0) * 100,
      ) / 100,
    ]),
  );
}
export function rows(entries) {
  return [...entries]
    .sort((a, b) => a.date.localeCompare(b.date))
    .flatMap((e) =>
      COMPANIES.filter((c) => e.allocations[c]).map((c) => [
        c,
        new Date(`${e.date}T12:00:00Z`),
        e.allocations[c],
        EMPLOYEE_ID,
        "WH",
        e.notes,
        `WH_${c}`,
      ]),
    );
}
