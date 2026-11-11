const ExcelJS = require("exceljs");
const { normalize, fail } = require("./domain");
const headers = {
  name: "اسم",
  father: "ولد",
  department: "دیپارتمنت",
  faculty: "پوهنځی",
  semester: "سمستر",
  province: "ولایت",
  phone: "شماره تماس",
  notes: "ملاحظات",
  card: "شماره کارت محصلی",
};
const cellText = (c) => (c.value == null ? "" : String(c.text || "").trim());
async function preview(buffer, mapping) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.worksheets[0];
  if (!sheet) fail("فایل خالی است.");
  if (sheet.rowCount > 10000) fail("حداکثر ۱۰۰۰۰ سطر در یک فایل مجاز است.");
  let headerRow = 0,
    columns = {};
  if (mapping) {
    for (const k of Object.keys(headers))
      if (mapping[k]) columns[k] = Number(mapping[k]);
    headerRow = Number(mapping.headerRow) || 3;
  } else
    sheet.eachRow((row, n) => {
      if (headerRow || n > 30) return;
      const cells = [];
      row.eachCell((c, i) => cells.push([normalize(cellText(c)), i]));
      if (
        cells.some(([v]) => v === "اسم") &&
        cells.some(([v]) => v === "ولد")
      ) {
        headerRow = n;
        for (const [key, title] of Object.entries(headers)) {
          const match = cells.find(([v]) => v === normalize(title));
          if (match) columns[key] = match[1];
        }
      }
    });
  if (!columns.name || !columns.father)
    fail("ستون‌های اسم و ولد پیدا نشد؛ شماره ستون‌ها را مشخص کنید.");
  const rows = [];
  sheet.eachRow((row, n) => {
    if (n <= headerRow) return;
    const r = { sourceRow: n };
    for (const [k, i] of Object.entries(columns))
      r[k] = cellText(row.getCell(i));
    if (!r.name && !r.father) return;
    if (
      ["بااحترام", "آمریت انسجام خدمات محصلان"].some((v) =>
        (r.name || "").includes(v),
      )
    )
      return;
    r.error = !r.name || !r.father ? "اسم یا ولد خالی است." : "";
    rows.push(r);
  });
  return { rows, columns, headerRow };
}
async function workbook(title, columns, rows) {
  const wb = new ExcelJS.Workbook();
  const sh = wb.addWorksheet("گزارش", {
    views: [{ rightToLeft: true }],
    pageSetup: {
      paperSize: 9,
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
    },
  });
  sh.mergeCells(1, 1, 1, columns.length);
  sh.getCell(1, 1).value = "پوهنتون کندوز — آمریت انسجام خدمات محصلان";
  sh.mergeCells(2, 1, 2, columns.length);
  sh.getCell(2, 1).value = title;
  sh.addRow(columns);
  rows.forEach((r) => sh.addRow(r));
  sh.columns.forEach((c, i) => (c.width = i === 0 ? 9 : 24));
  sh.eachRow((r, n) => {
    r.height = n <= 3 ? 28 : 24;
    r.eachCell((c) => {
      c.font = { name: "Arial", size: 12, bold: n <= 3 };
      c.alignment = {
        horizontal: "center",
        vertical: "middle",
        wrapText: true,
      };
      c.border = { bottom: { style: "thin", color: { argb: "FFD5DDE3" } } };
      if (n === 3)
        c.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFE1ECE8" },
        };
    });
  });
  sh.pageSetup.printTitlesRow = "1:3";
  return wb.xlsx.writeBuffer();
}

module.exports = { headers, preview, workbook };
