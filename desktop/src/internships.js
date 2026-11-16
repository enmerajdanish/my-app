const ExcelJS = require("exceljs");
const path = require("node:path");
const D = require("./domain");
const fields = {
  name: "اسم",
  father: "ولد",
  department: "دیپارتمنت",
  faculty: "پوهنځی",
  semester: "سمستر",
  referral: "مرجع معرفی شده",
  internship: "کارآموزی",
  practical: "کارعملی",
  teaching: "تدریس آموزی",
  startSolar: "تاریخ شروع",
  endSolar: "تاریخ ختم",
};
const flags = ["internship", "practical", "teaching"];
const digits = (v) =>
  String(v ?? "")
    .replace(/[۰-۹]/g, (c) => "۰۱۲۳۴۵۶۷۸۹".indexOf(c))
    .replace(/[٠-٩]/g, (c) => "٠١٢٣٤٥٦٧٨٩".indexOf(c));
const norm = (v) => D.normalize(digits(v)).replace(/[\s‌ـ]/g, "");
function solar(date) {
  const parts = new Intl.DateTimeFormat("en-US", {
    calendar: "persian",
    timeZone: "UTC",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(date);
  return ["year", "month", "day"]
    .map((k) => parts.find((p) => p.type === k).value)
    .join("/");
}
function date(v) {
  if (v instanceof Date) {
    if (Number.isNaN(v.getTime())) D.fail("تاریخ شمسی معتبر نیست.");
    return solar(v);
  }
  v = digits(v).trim();
  if (!v) return "";
  const m = v.match(/^(1[2345]\d{2})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (!m) D.fail("تاریخ شمسی را به شکل ۱۴۰۵/۱۲/۱۴ وارد کنید.");
  const year = +m[1],
    month = +m[2],
    day = +m[3];
  // Find solar New Year with Intl, including the actual Esfand length.
  const base = new Date(Date.UTC(year + 621, 2, 18));
  let start;
  for (let i = 0; i < 8; i++) {
    const d = new Date(+base + i * 86400000);
    if (solar(d) === `${year}/01/01`) {
      start = d;
      break;
    }
  }
  if (!start || month < 1 || month > 12 || day < 1 || day > 31)
    D.fail("تاریخ شمسی معتبر نیست.");
  const offset = month <= 6 ? (month - 1) * 31 : 186 + (month - 7) * 30;
  const target = `${year}/${String(month).padStart(2, "0")}/${String(day).padStart(2, "0")}`;
  if (solar(new Date(+start + (offset + day - 1) * 86400000)) !== target)
    D.fail("تاریخ شمسی معتبر نیست.");
  return target;
}
function flag(v, key) {
  if (v === true || v === 1) return true;
  if (v === false || v === 0 || v == null || String(v).trim() === "")
    return false;
  const value = norm(v);
  if (
    [
      "1",
      "true",
      "بلی",
      "بله",
      "هو",
      "yes",
      "x",
      "✓",
      "✔",
      "√",
      norm(fields[key]),
    ].includes(value)
  )
    return true;
  if (["0", "false", "نخیر", "نه", "no", "-", "—"].includes(value))
    return false;
  D.fail("علامت نوع کارآموزی معتبر نیست.");
}
function data(b) {
  const r = {};
  for (const k of Object.keys(fields))
    r[k] = flags.includes(k)
      ? flag(b[k], k)
      : D.text(b[k], k === "referral" ? 500 : 100);
  if (!r.name || !r.father) D.fail("اسم و ولد ضروری است.");
  r.startSolar = date(b.startSolar);
  r.endSolar = date(b.endSolar);
  if (r.startSolar && r.endSolar && r.startSolar > r.endSolar)
    D.fail("تاریخ ختم باید پس از شروع باشد.");
  return r;
}
const key = (r) =>
  JSON.stringify(
    Object.keys(fields).map((k) =>
      typeof r[k] === "string" ? norm(r[k]) : r[k],
    ),
  );
function save(s, b, id) {
  const r = data(b),
    existing = id ? s.internships.find((x) => x.id === id) : null;
  if (id && !existing) D.fail("مورد پیدا نشد.", 404);
  if (s.internships.some((x) => x.id !== id && key(x) === key(r)))
    D.fail("این ردیف کارآموزی قبلاً ثبت شده است.");
  const record = {
    ...(existing || { id: D.id(), createdAt: new Date().toISOString() }),
    ...r,
  };
  if (existing) Object.assign(existing, record);
  else s.internships.push(record);
  return record;
}
function text(cell) {
  const v = cell.value;
  if (v == null) return "";
  if (v instanceof Date) return v;
  if (v && typeof v === "object" && "formula" in v)
    return v.result == null ? "" : String(v.result);
  return String(cell.text || "").trim();
}
async function preview(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.worksheets[0];
  if (!sheet) D.fail("فایل خالی است.");
  if (sheet.rowCount > 10000) D.fail("حداکثر ۱۰۰۰۰ سطر در یک فایل مجاز است.");
  let headerRow = 0,
    columns = {};
  sheet.eachRow((row, n) => {
    if (headerRow || n > 30) return;
    const cells = [];
    row.eachCell((c, i) => cells.push([norm(text(c)), i]));
    if (
      cells.some(([v]) => v === norm("اسم")) &&
      cells.some(([v]) => v === norm("مرجع معرفی شده"))
    ) {
      headerRow = n;
      for (const [k, title] of Object.entries(fields)) {
        const match = cells.find(([v]) => v === norm(title));
        if (match) columns[k] = match[1];
      }
    }
  });
  if (Object.keys(columns).length !== Object.keys(fields).length)
    D.fail("ستون‌های فایل کارآموزی مطابق نمونه نیست.");
  const rows = [];
  sheet.eachRow((row, n) => {
    if (n <= headerRow) return;
    const raw = { sourceRow: n };
    for (const [k, i] of Object.entries(columns)) raw[k] = text(row.getCell(i));
    if (!raw.name && !raw.father) return;
    if (
      ["بااحترام", "آمریت انسجام خدمات محصلان"].some(
        (v) => norm(raw.name) === norm(v),
      )
    )
      return;
    if (norm(raw.name) === norm("اسم") && norm(raw.father) === norm("ولد"))
      return;
    try {
      rows.push({ ...data(raw), sourceRow: n, error: "" });
    } catch (e) {
      rows.push({
        ...Object.fromEntries(
          Object.entries(raw).map(([k, v]) => [
            k,
            v instanceof Date ? solar(v) : v,
          ]),
        ),
        error: e.message,
      });
    }
  });
  if (!rows.length) D.fail("فایل خالی است.");
  return { rows, headerRow };
}
function filter(rows, q = {}) {
  const from = date(q.from),
    to = date(q.to);
  if (from && to && from > to) D.fail("تاریخ ختم باید پس از شروع باشد.");
  const result = rows.filter(
    (r) =>
      (!q.search ||
        norm(Object.values(r).join(" ")).includes(norm(q.search))) &&
      (!q.faculty || r.faculty === q.faculty) &&
      (!q.referral || norm(r.referral).includes(norm(q.referral))) &&
      (!q.type || (flags.includes(q.type) && r[q.type])) &&
      (!from || (r.startSolar && r.startSolar >= from)) &&
      (!to || (r.startSolar && r.startSolar <= to)),
  );
  return result.sort((a, b) =>
    q.order === "asc"
      ? (a.startSolar || "").localeCompare(b.startSolar || "")
      : (b.startSolar || "").localeCompare(a.startSolar || ""),
  );
}
async function workbook(rows, settings = {}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, "../templates/internships.xlsx"));
  const sh = wb.worksheets[0];
  sh.name = "کارآموزی";
  sh.getCell("A2").value =
    `لست محصلین ${settings.university || "پوهنتون کندوز"} جهت سپری نمودن کارعملی، تدریس‌آموزی و کارآموزی`;
  const style = Array.from({ length: 12 }, (_, i) =>
    structuredClone(sh.getRow(4).getCell(i + 1).style),
  );
  rows.forEach((r, i) => {
    const row = sh.getRow(i + 4);
    row.height = 30;
    [
      i + 1,
      ...Object.keys(fields).map((k) =>
        flags.includes(k) ? (r[k] ? fields[k] : "") : r[k] || "",
      ),
    ].forEach((v, j) => {
      const c = row.getCell(j + 1);
      c.value = v;
      c.style = structuredClone(style[j]);
      c.alignment = {
        ...c.alignment,
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
    });
  });
  if (!rows.length) for (let j = 1; j <= 12; j++) sh.getCell(4, j).value = null;
  const footer = Math.max(5, rows.length + 5);
  sh.mergeCells(footer, 1, footer, 12);
  sh.getCell(footer, 1).value = "بااحترام";
  sh.mergeCells(footer + 1, 1, footer + 1, 12);
  sh.getCell(footer + 1, 1).value =
    settings.office || "آمریت انسجام خدمات محصلان";
  for (const r of [footer, footer + 1]) {
    sh.getRow(r).height = 24;
    sh.getCell(r, 1).alignment = { horizontal: "center" };
    sh.getCell(r, 1).font = { name: "B Nazanin", size: 12, bold: true };
  }
  sh.views = [{ rightToLeft: true, state: "frozen", ySplit: 3 }];
  sh.pageSetup = {
    ...sh.pageSetup,
    paperSize: 9,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printTitlesRow: "1:3",
    printArea: `A1:L${footer + 1}`,
  };
  wb.creator = "Hamrah Mohasel";
  wb.lastModifiedBy = "Hamrah Mohasel";
  return wb.xlsx.writeBuffer();
}
module.exports = {
  fields,
  flags,
  date,
  data,
  key,
  save,
  preview,
  filter,
  workbook,
};
