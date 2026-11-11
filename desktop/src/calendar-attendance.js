const fs = require("node:fs");
const path = require("node:path");
const ExcelJS = require("exceljs");
const D = require("./domain");
const { translate } = require("../public/i18n");
const templateFile = path.join(__dirname, "../templates/attendance.xlsx");
const esc = (v) =>
  String(v ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        c
      ],
  );
function daysFor(c, count, options = {}) {
  D.integer(count, 1, 366, "تعداد روزها");
  const start = D.date(options.start || c.start, true),
    closed = options.fridayClosed !== false;
  const names = [
    "یکشنبه",
    "دوشنبه",
    "سه شنبه",
    "چهارشنبه",
    "پنجشنبه",
    "جمعه",
    "شنبه",
  ];
  const format = new Intl.DateTimeFormat("en-US", {
    calendar: "persian",
    timeZone: "UTC",
    year: "numeric",
    month: "numeric",
    day: "numeric",
  });
  return Array.from({ length: count }, (_, i) => {
    const date = new Date(start + "T12:00:00Z");
    date.setUTCDate(date.getUTCDate() + i);
    const p = format.formatToParts(date);
    return {
      date: date.toISOString().slice(0, 10),
      solar: ["year", "month", "day"]
        .map((k) => p.find((x) => x.type === k).value)
        .join("/"),
      weekday: names[date.getUTCDay()],
      off: closed && date.getUTCDay() === 5,
    };
  });
}
function roster(s, c) {
  return s.enrollments
    .filter((e) => e.courseId === c.id)
    .map((e) => s.students.find((st) => st.id === e.studentId));
}
async function paperWorkbook(s, c, count, lang = "fa", options = {}) {
  const days = daysFor(c, count, options),
    students = roster(s, c);
  const protoBook = new ExcelJS.Workbook();
  await protoBook.xlsx.load(fs.readFileSync(templateFile));
  const proto = protoBook.worksheets[0],
    protoModel = proto.model;
  const wb = new ExcelJS.Workbook();
  const model = wb.model;
  model.themes = protoBook.model.themes;
  wb.model = model;
  wb.creator = "Hamrah Mohasel";
  for (let offset = 0; offset < days.length; offset += 31) {
    const part = days.slice(offset, offset + 31);
    const name = "حاضری " + (offset + 1) + "-" + (offset + part.length);
    const sh = wb.addWorksheet(name);
    sh.model = { ...structuredClone(protoModel), name, id: sh.id };
    sh.views = [{ rightToLeft: true }];
    sh.getCell("A2").value = translate("جدول حاضری کورس", lang) + " " + c.name;
    sh.getCell("AJ2").value =
      translate("استاد", lang) +
      ": " +
      (s.users.find((u) => u.id === c.teacherId)?.name || "");
    sh.getCell("D2").value = translate("تاریخ", lang);
    sh.getCell("D3").value = translate("روز", lang);
    ["شماره", "اسم", "ولد"].forEach(
      (h, i) => (sh.getCell(3, i + 1).value = translate(h, lang)),
    );
    sh.getCell("AJ3").value = translate("حاضر", lang);
    sh.getCell("AK3").value = translate("غیر حاضر", lang);
    for (let i = 0; i < 31; i++) {
      const day = part[i],
        col = i + 5;
      for (const n of [2, 3]) {
        const cell = sh.getCell(n, col);
        cell.style = structuredClone(proto.getCell(n, day?.off ? 11 : 5).style);
        if (!day)
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFE8E8E8" },
          };
      }
      sh.getCell(2, col).value = day?.solar || null;
      sh.getCell(3, col).value = day ? translate(day.weekday, lang) : null;
    }
    const active = part.map((d, i) => (d.off ? null : i + 5)).filter(Boolean);
    const ranges = [];
    for (const col of active) {
      const last = ranges.at(-1);
      if (last && last[1] === col - 1) last[1] = col;
      else ranges.push([col, col]);
    }
    students.forEach((st, i) => {
      const row = i + 4;
      sh.getRow(row).height = proto.getRow(4).height || 15;
      for (let col = 1; col <= 37; col++) {
        const source =
          col >= 5 && col <= 35
            ? !part[col - 5] || part[col - 5].off
              ? 11
              : 5
            : col;
        sh.getCell(row, col).style = structuredClone(
          proto.getCell(4, source).style,
        );
        sh.getCell(row, col).value = null;
      }
      sh.getCell(row, 1).value = i + 1;
      sh.getCell(row, 2).value = st.name;
      sh.getCell(row, 3).value = st.father;
      const refs = ranges.map(
        ([a, b]) =>
          `${sh.getColumn(a).letter}${row}:${sh.getColumn(b).letter}${row}`,
      );
      const any = refs.map((ref) => `COUNTA(${ref})`).join(",") || "0";
      for (const [col, marker] of [
        [36, "a"],
        [37, "r"],
      ]) {
        const terms =
          refs.map((ref) => `COUNTIF(${ref},"${marker}")`).join(",") || "0";
        sh.getCell(row, col).value = {
          formula: `IF(SUM(${any})=0,"",SUM(${terms}))`,
          result: "",
        };
      }
      if ((i + 1) % 25 === 0 && i + 1 < students.length)
        sh.getRow(row).addPageBreak();
    });
    if (students.length > 1) sh.mergeCells(4, 4, students.length + 3, 4);
    sh.pageSetup = {
      ...sh.pageSetup,
      paperSize: 9,
      orientation: "landscape",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      printTitlesRow: "2:3",
      printArea: `A2:AK${students.length + 3}`,
      margins: {
        left: 0.25,
        right: 0.25,
        top: 0.3,
        bottom: 0.3,
        header: 0.1,
        footer: 0.1,
      },
    };
  }
  return wb.xlsx.writeBuffer();
}
function paperHTML(s, c, count, lang = "fa", options = {}) {
  const all = daysFor(c, count, options),
    students = roster(s, c),
    t = (v) => translate(v, lang);
  const widths = [
    5.140625,
    9.140625,
    13.5703125,
    5.85546875,
    ...Array(31).fill(3.7109375),
    14.42578125,
    13.140625,
  ];
  const total = widths.reduce((a, b) => a + b, 0);
  const pages = [];
  for (let start = 0; start < all.length; start += 31)
    for (let at = 0; at < students.length; at += 25) {
      const days = all.slice(start, start + 31),
        people = students.slice(at, at + 25);
      pages.push(
        `<article><table><colgroup>${widths.map((w) => `<col style="width:${(w / total) * 100}%">`).join("")}</colgroup><thead><tr><th colspan="3" class="identity title">${t("جدول حاضری کورس")} ${esc(c.name)}</th><th class="day">${t("تاریخ")}</th>${Array.from(
          { length: 31 },
          (_, i) => {
            const d = days[i];
            return `<th class="day ${!d ? "unused" : d.off ? "holiday" : ""}"><span class="vertical" dir="ltr">${d?.solar || ""}</span></th>`;
          },
        ).join(
          "",
        )}<th colspan="2" class="teacher">${t("استاد")}: ${esc(s.users.find((u) => u.id === c.teacherId)?.name || "")}</th></tr><tr><th class="identity">${t("شماره")}</th><th class="identity">${t("اسم")}</th><th class="identity">${t("ولد")}</th><th class="day">${t("روز")}</th>${Array.from({ length: 31 }, (_, i) => `<th class="day ${!days[i] ? "unused" : days[i].off ? "holiday" : ""}"><span class="vertical weekday">${days[i] ? t(days[i].weekday) : ""}</span></th>`).join("")}<th class="day">${t("حاضر")}</th><th class="day">${t("غیر حاضر")}</th></tr></thead><tbody>${people.map((st, i) => `<tr><td>${at + i + 1}</td><td class="name">${esc(st.name)}</td><td class="name">${esc(st.father)}</td>${i === 0 ? `<td rowspan="${people.length}"></td>` : ""}${Array.from({ length: 31 }, (_, j) => `<td class="${!days[j] ? "unused" : days[j].off ? "closed" : ""}"></td>`).join("")}<td></td><td></td></tr>`).join("")}</tbody></table></article>`,
      );
    }
  return `<!doctype html><html lang="${lang}" dir="rtl"><meta charset="utf-8"><title>${t("فورم حاضری")}</title><link rel="stylesheet" href="/assets/fonts.css"><style>@page{size:A4 landscape;margin:8mm}*{box-sizing:border-box}body{font-family:Calibri,Vazirmatn,Arial;background:#eee;margin:0}button{margin:15px;padding:10px 24px}article{width:281mm;min-height:192mm;padding:3mm;background:#fff;margin:10px auto;break-after:page}table{border-collapse:collapse;width:100%;table-layout:fixed;text-align:center;font-size:11px;color:#000}th,td{border:1px solid #111;padding:2px}th{font-weight:700}.identity{background:#bdd7ee}.title,.teacher{font-size:14px}.teacher{background:#fff}.day{background:#c6e0b4}.holiday{background:#f00}.unused{background:#e8e8e8}.closed{border-top:0;border-bottom:0}tbody td{height:21px}td.name{font-family:'B Nazanin',Vazirmatn,Arial;font-weight:bold;font-size:12px;white-space:nowrap;line-height:15px}.vertical{display:inline-block;writing-mode:vertical-rl;transform:rotate(180deg);white-space:nowrap;font-size:10px;min-height:70px}.weekday{min-height:44px}th{-webkit-print-color-adjust:exact;print-color-adjust:exact}@media print{button{display:none}body{background:#fff}article{padding:0;margin:0;min-height:0}article:last-child{break-after:auto}}</style><button onclick="document.fonts.ready.then(fitNames).then(()=>print())">${t("چاپ")}</button>${pages.join("")}<script>function fitNames(){const context=document.createElement('canvas').getContext('2d');for(const el of document.querySelectorAll('td.name')){const style=getComputedStyle(el);context.font=style.fontWeight+' 12px '+style.fontFamily;const width=context.measureText(el.textContent).width;if(width)el.style.fontSize=Math.max(6,Math.min(12,12*(el.clientWidth-5)/width))+'px';}}document.fonts.ready.then(fitNames);</script></html>`;
}
module.exports = { daysFor, paperWorkbook, paperHTML };
