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
  if (mapping.dailyStart && mapping.dailyEnd) {
    const start = D.integer(mapping.dailyStart, 1, 1100, "شروع ستون روزها"),
      end = D.integer(mapping.dailyEnd, start, 1100, "ختم ستون روزها");
    for (let i = start; i <= end; i++) daily.push(i);
  } else if (!columns.absent) {
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
  const rows = [];
  sheet.eachRow((row, n) => {
    if (n <= headerRow) return;
    const r = { sourceRow: n };
    for (const [k, col] of Object.entries(columns))
      r[k] = cellText(row.getCell(col));
    if (!r.name && !r.father) return;
    if (daily.length) {
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
    rows.push(r);
  });
  return {
    rows,
    columns,
    headerRow,
    mode: daily.length ? "daily" : "summary",
    dailyColumns: daily,
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
async function paperWorkbook(s, c, sessions, lang = "fa") {
  const wb = new ExcelJS.Workbook();
  const students = s.enrollments
    .filter((e) => e.courseId === c.id)
    .map((e) => s.students.find((st) => st.id === e.studentId));
  for (let start = 1; start <= sessions; start += 15) {
    const end = Math.min(sessions, start + 14);
    const days = Array.from({ length: end - start + 1 }, (_, i) => start + i);
    const sh = wb.addWorksheet("حاضری " + start + "-" + end, {
      views: [{ rightToLeft: true }],
      pageSetup: {
        paperSize: 9,
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
      },
    });
    sh.mergeCells(1, 1, 1, 3 + days.length);
    sh.getCell("A1").value =
      s.settings[0].university + " — " + s.settings[0].office;
    sh.mergeCells(2, 1, 2, 3 + days.length);
    sh.getCell("A2").value =
      c.name +
      " | استاد: " +
      (s.users.find((u) => u.id === c.teacherId)?.name || "") +
      " | تاریخ: __________";
    sh.addRow(["شماره", "اسم", "ولد", ...days.map((d) => "جلسه " + d)]);
    students.forEach((st, i) =>
      sh.addRow([i + 1, st.name, st.father, ...days.map(() => "")]),
    );
    sh.columns.forEach((col, i) => (col.width = i === 0 ? 7 : i < 3 ? 24 : 5));
    sh.eachRow((row, n) => {
      row.height = n <= 3 ? 28 : 30;
      for (let i = 1; i <= 3 + days.length; i++) {
        const cell = row.getCell(i);
        cell.font = { name: "Arial", size: 11, bold: n <= 3 };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        cell.border = {
          top: { style: "thin" },
          bottom: { style: "thin" },
          left: { style: "thin" },
          right: { style: "thin" },
        };
      }
    });
    sh.pageSetup.printTitlesRow = "1:3";
    sh.eachRow((row, n) => {
      if (n <= 3)
        row.eachCell((cell) => {
          if (typeof cell.value === "string")
            cell.value = translate(cell.value, lang);
        });
    });
  }
  return wb.xlsx.writeBuffer();
}
const esc = (v) =>
  String(v ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        c
      ],
  );
function paperHTML(s, c, sessions, lang = "fa") {
  const ps = lang === "ps";
  const text = ps
    ? {
        title: "د حاضري پاڼه",
        number: "شمېره",
        name: "نوم",
        father: "د پلار نوم",
        date: "نېټه",
        instructor: "استاد",
        session: "ناسته",
        print: "چاپ",
        note: "حاضري په قلم ولیکئ. د ناستو د نېټو لپاره پورته تش ځایونه وکاروئ.",
      }
    : {
        title: "فورم حاضری قلمی",
        number: "شماره",
        name: "اسم",
        father: "ولد",
        date: "تاریخ",
        instructor: "استاد",
        session: "جلسه",
        print: "چاپ",
        note: "حاضری را با قلم درج کنید. تاریخ هر جلسه در خانهٔ خالی بالای آن نوشته شود.",
      };
  const students = s.enrollments
    .filter((e) => e.courseId === c.id)
    .map((e) => s.students.find((st) => st.id === e.studentId));
  const pages = [];
  for (let start = 1; start <= sessions; start += 15)
    for (let offset = 0; offset < students.length; offset += 20) {
      const days = Array.from(
        { length: Math.min(15, sessions - start + 1) },
        (_, i) => start + i,
      );
      pages.push(
        `<article><header><img src="/assets/logo1.jpg"><div><b>${esc(s.settings[0].university)}</b><br>${esc(s.settings[0].office)}<h2>${text.title} — ${esc(c.name)}</h2>${text.instructor}: ${esc(s.users.find((u) => u.id === c.teacherId)?.name)} · ${esc(c.days)} · ${esc(c.startTime)} — ${esc(c.endTime)}</div><img src="/assets/logo2.jpg"></header><table><thead><tr><th rowspan="2">${text.number}</th><th rowspan="2" class="name">${text.name}</th><th rowspan="2" class="name">${text.father}</th>${days.map((d) => `<th>${text.session}<br>${d}</th>`).join("")}</tr><tr>${days.map(() => '<th class="date">____<br>____</th>').join("")}</tr></thead><tbody>${students
          .slice(offset, offset + 20)
          .map(
            (st, i) =>
              `<tr><td>${offset + i + 1}</td><td>${esc(st.name)}</td><td>${esc(st.father)}</td>${days.map(() => "<td></td>").join("")}</tr>`,
          )
          .join(
            "",
          )}</tbody></table><footer>${text.note}<span>${offset + 1} — ${Math.min(students.length, offset + 20)} / ${students.length}</span></footer></article>`,
      );
    }
  return `<!doctype html><html lang="${lang}" dir="rtl"><meta charset="utf-8"><title>${text.title}</title><link href="/assets/fonts.css" rel="stylesheet"><style>@page{size:A4 landscape;margin:9mm}*{box-sizing:border-box}body{font-family:Vazirmatn,Arial;background:#eee;margin:0}button{margin:15px;padding:10px 24px}article{width:279mm;min-height:190mm;padding:4mm;background:white;margin:10px auto;break-after:page}header{display:flex;justify-content:space-between;align-items:center;text-align:center;font-size:11px;margin-bottom:12px}header img{width:62px;height:62px;object-fit:contain}h2{margin:5px;font-size:17px}table{border-collapse:collapse;width:100%;table-layout:fixed;font-size:10px;text-align:center}th,td{border:1px solid #555;padding:3px}th:first-child{width:32px}.name{width:120px}.date{height:34px;font-size:8px}tbody td{height:23px}footer{font-size:9px;margin-top:10px;display:flex;justify-content:space-between}@media print{button{display:none}body{background:white}article{margin:0;padding:0;min-height:0}article:last-child{break-after:auto}}</style><button onclick="print()">${text.print}</button>${pages.join("")}</html>`;
}
module.exports = {
  parse,
  analyze,
  counts,
  template,
  paperWorkbook,
  paperHTML,
  digits,
};
