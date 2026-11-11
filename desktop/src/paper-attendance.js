const ExcelJS = require("exceljs");
const D = require("./domain");
const { translate } = require("../public/i18n");
const aliases = {
  name: ["اسم", "نام", "نام محصل", "نوم", "student name", "name"],
  father: ["ولد", "نام پدر", "د پلار نوم", "father", "father name"],
  card: ["شماره کارت محصلی", "شماره کارت", "کارت", "student id"],
  sessions: [
    "جلسات",
    "تعداد جلسات",
    "مجموع جلسات",
    "کل جلسات",
    "مجموع",
    "ټولې ناستې",
    "total",
    "sessions",
  ],
  present: ["حاضر", "روزهای حاضر", "حاضری", "حاضران", "حاضرې ورځې", "present"],
  absent: [
    "غیرحاضر",
    "غیر حاضر",
    "غایب",
    "غیابت",
    "غیرحاضری",
    "غیر حاضری",
    "تعداد غیرحاضری",
    "غیرحاضرې ورځې",
    "ناسوب",
    "absent",
  ],
  leave: ["رخصت", "رخصتی", "روزهای رخصت", "رخصت ورځې", "leave"],
};
const digits = (v) =>
  String(v ?? "").replace(/[۰-۹٠-٩]/g, (c) =>
    String(
      "۰۱۲۳۴۵۶۷۸۹".indexOf(c) >= 0
        ? "۰۱۲۳۴۵۶۷۸۹".indexOf(c)
        : "٠١٢٣٤٥٦٧٨٩".indexOf(c),
    ),
  );
const cellText = (c) => (c.value == null ? "" : String(c.text || "").trim());
function count(v, label, required = true) {
  if (v === "" || v == null) {
    if (required) D.fail(`${label} خالی است.`);
    return null;
  }
  const clean = digits(v).trim();
  if (!/^\d+$/.test(clean)) D.fail(`${label} باید عدد صحیح و غیرمنفی باشد.`);
  return D.integer(Number(clean), 0, 10000, label);
}
function counts(row, defaultSessions) {
  const sessions = count(
    row.sessions === "" || row.sessions == null
      ? defaultSessions
      : row.sessions,
    "جلسات",
  );
  if (!sessions) D.fail("تعداد جلسات باید بیشتر از صفر باشد.");
  const absent = count(row.absent, "غیرحاضری");
  const leave = count(row.leave, "رخصت", false) ?? 0;
  const present =
    count(row.present, "حاضری", false) ?? sessions - absent - leave;
  if (present < 0 || present + absent + leave !== sessions)
    D.fail("مجموع حاضر، غیرحاضر و رخصت با تعداد جلسات برابر نیست.");
  return { sessions, present, absent, leave };
}
function analyze(s, cid, rows, defaultSessions) {
  if (!Array.isArray(rows) || rows.length > 10000)
    D.fail("فهرست نتایج معتبر نیست.");
  const enrolled = s.enrollments
    .filter((e) => e.courseId === cid)
    .map((e) => s.students.find((st) => st.id === e.studentId));
  const seen = new Set();
  return rows.map((r, i) => {
    const out = {
      sourceRow: r.sourceRow || i + 1,
      name: D.text(r.name, 100),
      father: D.text(r.father, 100),
      card: D.text(r.card),
      sessions: r.sessions ?? defaultSessions,
      present: r.present ?? "",
      absent: r.absent ?? "",
      leave: r.leave ?? "",
      error: "",
      studentId: null,
    };
    try {
      if (r.parseError) D.fail(D.text(r.parseError));
      if (!out.name || !out.father) D.fail("اسم و ولد ضروری است.");
      const matches = enrolled.filter((st) =>
        out.card
          ? st.card === out.card &&
            D.normalize(st.name) === D.normalize(out.name) &&
            D.normalize(st.father) === D.normalize(out.father)
          : D.normalize(st.name) === D.normalize(out.name) &&
            D.normalize(st.father) === D.normalize(out.father),
      );
      if (matches.length !== 1)
        D.fail(
          matches.length
            ? "اسم و ولد مبهم است؛ شماره کارت را مشخص کنید."
            : "محصل مطابق اسم و ولد در این کورس ثبت نیست.",
        );
      Object.assign(out, counts(out, defaultSessions));
      if (seen.has(matches[0].id))
        D.fail("نتیجهٔ تکراری برای همین محصل در فایل.");
      seen.add(matches[0].id);
      out.studentId = matches[0].id;
      out.eligible = out.absent < s.settings[0].absenceLimit;
    } catch (e) {
      out.error = e.message;
      out.eligible = false;
    }
    return out;
  });
}
async function parse(buffer, mapping = {}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.worksheets[0];
  if (!sheet || sheet.rowCount > 10000 || sheet.columnCount > 1100)
    D.fail("فایل خالی یا بزرگ‌تر از حد مجاز است.");
  let headerRow = Number(mapping.headerRow) || 0;
  const columns = {};
  if (!headerRow && Object.keys(aliases).some((k) => mapping[k]))
    D.fail("برای تنظیم دستی، شمارهٔ سطر عنوان را هم وارد کنید.");
  if (Boolean(mapping.dailyStart) !== Boolean(mapping.dailyEnd))
    D.fail("شروع و ختم ستون روزها را با هم تعیین کنید.");
  if (headerRow) {
    if (
      !Number.isInteger(headerRow) ||
      headerRow < 1 ||
      headerRow > sheet.rowCount
    )
      D.fail("شمارهٔ سطر عنوان معتبر نیست.");
    sheet.getRow(headerRow).eachCell((cell, i) => {
      const value = D.normalize(cellText(cell));
      for (const [key, names] of Object.entries(aliases))
        if (names.some((a) => D.normalize(a) === value)) columns[key] = i;
    });
    for (const k of Object.keys(aliases))
      if (mapping[k]) columns[k] = D.integer(mapping[k], 1, 1100, "شمارهٔ ستون");
  } else {
    sheet.eachRow((row, n) => {
      if (headerRow || n > 30) return;
      const found = {};
      row.eachCell((cell, i) => {
        const value = D.normalize(cellText(cell));
        for (const [key, names] of Object.entries(aliases))
          if (names.some((a) => D.normalize(a) === value)) found[key] = i;
      });
      if (found.name && found.father) {
        Object.assign(columns, found);
        headerRow = n;
      }
    });
  }
  if (!columns.name || !columns.father)
    D.fail("ستون‌های اسم و ولد پیدا نشد؛ شماره ستون‌ها را دستی مشخص کنید.");
  const daily = [];
  const dateColumns = [];
  if (headerRow > 1)
    sheet.getRow(headerRow - 1).eachCell((cell, i) => {
      if (
        !Object.values(columns).includes(i) &&
        /^\d{4}[/-]\d{1,2}[/-]\d{1,2}$/.test(digits(cellText(cell)))
      )
        dateColumns.push(i);
    });
  const knownDays = [
    "شنبه",
    "یکشنبه",
    "دوشنبه",
    "سه شنبه",
    "سه‌شنبه",
    "چهارشنبه",
    "پنجشنبه",
    "جمعه",
    "جمعه‌",
  ];
  const calendarLayout =
    dateColumns.length >= 5 &&
    dateColumns.every((i) =>
      knownDays.includes(cellText(sheet.getCell(headerRow, i))),
    );
  const countif = (col, marker) => {
    let found = false;
    if (col)
      sheet.eachRow((row, n) => {
        const formula = row.getCell(col).value?.formula;
        if (
          n > headerRow &&
          typeof formula === "string" &&
          new RegExp(
            'COUNTIF\\([^,;]+[,;]\\s*"' + marker + '"\\s*\\)',
            "i",
          ).test(formula)
        )
          found = true;
      });
    return found;
  };
  const calendarProfile =
    calendarLayout &&
    countif(columns.present, "a") &&
    countif(columns.absent, "r");
  const holidayColumns = calendarProfile
    ? dateColumns.filter(
        (i) =>
          D.normalize(cellText(sheet.getCell(headerRow, i))) === "جمعه" &&
          ["FFFF0000", "FFD8DDD8", "FF000000"].includes(
            (
              sheet.getCell(headerRow, i).fill?.fgColor?.argb || ""
            ).toUpperCase(),
          ),
      )
    : [];
  const hasExplicitDaily = Boolean(mapping.dailyStart && mapping.dailyEnd);
  if (calendarProfile && !hasExplicitDaily) {
    for (const i of dateColumns) if (!holidayColumns.includes(i)) daily.push(i);
  }

  if (mapping.dailyStart && mapping.dailyEnd) {
    const start = D.integer(mapping.dailyStart, 1, 1100, "شروع ستون روزها"),
      end = D.integer(mapping.dailyEnd, start, 1100, "ختم ستون روزها");
    for (let i = start; i <= end; i++) daily.push(i);
  } else if (!columns.absent && !daily.length) {
    sheet.getRow(headerRow).eachCell((cell, i) => {
      const value = digits(cellText(cell));
      if (
        !Object.values(columns).includes(i) &&
        (/^(جلسه|روز|ناسته)\s*\d+/.test(value) ||
          /^\d{1,3}$/.test(value) ||
          /^\d{1,4}[\/.-]\d{1,2}(?:[\/.-]\d{1,4})?$/.test(value) ||
          cell.value instanceof Date)
      )
        daily.push(i);
    });
  }
  if (!columns.absent && !daily.length)
    D.fail(
      "ستون غیرحاضری یا ستون‌های روزها پیدا نشد؛ تنظیم دستی ستون‌ها را تکمیل کنید.",
    );
  if (daily.some((i) => Object.values(columns).includes(i)))
    D.fail("ستون روزها با ستون مشخصات یا مجموع تداخل دارد.");
  const marks = {
    ح: "present",
    حاضر: "present",
    p: "present",
    present: "present",
    "✓": "present",
    "+": "present",
    غ: "absent",
    غایب: "absent",
    غیرحاضر: "absent",
    "غیر حاضر": "absent",
    ناسوب: "absent",
    a: "absent",
    absent: "absent",
    "×": "absent",
    "-": "absent",
    ر: "leave",
    رخصت: "leave",
    l: "leave",
    leave: "leave",
  };
  if (calendarProfile) {
    marks.a = "present";
    marks.r = "absent";
    marks.l = "leave";
  }
  const inferredSessionTotals = new Set();
  const rows = [];
  sheet.eachRow((row, n) => {
    if (n <= headerRow) return;
    const r = { sourceRow: n };
    for (const [k, col] of Object.entries(columns))
      r[k] = cellText(row.getCell(col));
    if (!r.name && !r.father) return;
    if (daily.length) {
      if (
        holidayColumns.some(
          (i) => !daily.includes(i) && cellText(row.getCell(i)),
        )
      )
        r.parseError =
          "روز تعطیل دارای علامت حاضری است؛ وضعیت روز را اصلاح کنید.";
      r.sessions = daily.length;
      r.present = r.absent = r.leave = 0;
      for (const i of daily) {
        const value = D.normalize(cellText(row.getCell(i)));
        const mark = marks[value];
        if (!mark) {
          r.parseError =
            "وضعیت یک روز خالی یا ناشناخته است؛ همهٔ روزها را تکمیل کنید.";
          break;
        }
        r[mark]++;
      }
    }
    if (
      !daily.length &&
      !columns.sessions &&
      r.present !== "" &&
      r.present != null &&
      r.absent !== "" &&
      r.absent != null
    ) {
      try {
        r.sessions =
          count(r.present, "حاضر") +
          count(r.absent, "غیرحاضر") +
          (count(r.leave, "رخصت", false) || 0);
        inferredSessionTotals.add(r.sessions);
      } catch (e) {
        r.parseError = e.message;
      }
    }
    rows.push(r);
  });
  return {
    rows,
    columns,
    headerRow,
    mode: daily.length ? "daily" : "summary",
    dailyColumns: daily,
    profile: calendarProfile
      ? "kunduz-calendar"
      : daily.length
        ? "generic-daily"
        : "final-summary",
    inferredSessionTotals: [...inferredSessionTotals],
    excludedHolidayColumns: calendarProfile
      ? dateColumns.filter((i) => !daily.includes(i))
      : [],
  };
}
async function template(s, c, lang = "fa") {
  const wb = new ExcelJS.Workbook();
  const sh = wb.addWorksheet("نتایج حاضری", { views: [{ rightToLeft: true }] });
  sh.addRow(["پوهنتون کندوز — آمریت انسجام خدمات محصلان"]);
  sh.addRow([c.name + " — نتایج حاضری"]);
  sh.addRow(["شماره", "اسم", "ولد", "جلسات", "حاضر", "غیرحاضر", "رخصت"]);
  s.enrollments
    .filter((e) => e.courseId === c.id)
    .forEach((e, i) => {
      const st = s.students.find((st) => st.id === e.studentId);
      sh.addRow([i + 1, st.name, st.father, c.sessions, null, null, 0]);
    });
  sh.columns.forEach((col, i) => (col.width = [8, 26, 26, 12, 12, 12, 12][i]));
  sh.getRow(3).font = { bold: true };
  sh.eachRow((row, n) => {
    if (n <= 3)
      row.eachCell((cell) => {
        if (typeof cell.value === "string")
          cell.value = translate(cell.value, lang);
      });
  });
  return wb.xlsx.writeBuffer();
}
const { daysFor, paperWorkbook, paperHTML } = require("./calendar-attendance");
module.exports = {
  parse,
  analyze,
  counts,
  template,
  daysFor,
  paperWorkbook,
  paperHTML,
  digits,
};
