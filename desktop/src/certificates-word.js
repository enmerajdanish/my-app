const fs = require("node:fs");
const path = require("node:path");
const sharp = require("sharp");
const {
  Document,
  Packer,
  Header,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  ImageRun,
  AlignmentType,
  WidthType,
  BorderStyle,
  PageOrientation,
  VerticalAlign,
} = require("docx");
const root = path.join(__dirname, "../public/assets");
const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const borders = {
  top: none,
  bottom: none,
  left: none,
  right: none,
  insideHorizontal: none,
  insideVertical: none,
};
const faDate = (v, lang) => {
  const parts = new Intl.DateTimeFormat(lang === "ps" ? "ps-AF" : "fa-AF", {
    calendar: "persian",
    timeZone: "Asia/Kabul",
    numberingSystem: "arabext",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date(v + "T12:00:00Z"));
  return ["year", "month", "day"]
    .map((k) => parts.find((p) => p.type === k)?.value)
    .join("/");
};
let background;
async function getBackground() {
  if (!background) {
    const paths = [];
    for (let i = 0; i < 7; i++)
      paths.push(
        `<rect x="${21 + i * 4}" y="${21 + i * 4}" width="${1081 - i * 8}" height="${752 - i * 8}" rx="${i % 2 ? 5 : 0}" fill="none" stroke="${i % 2 ? "#c29c56" : "#174f43"}" stroke-width="${i === 0 ? 3 : 0.55}"/>`,
      );
    for (const [x, y, rotate] of [
      [60, 60, 0],
      [1063, 60, 90],
      [1063, 734, 180],
      [60, 734, 270],
    ])
      paths.push(
        `<g transform="translate(${x},${y}) rotate(${rotate})" fill="none" stroke="#b59049" stroke-width="1.4"><path d="M0 78V0H78M8 64V8H64M0 20Q20 20 20 0M15 45Q45 45 45 15M0 55Q15 45 32 32Q45 15 55 0"/><path d="M5 45C40 10 48 45 13 32C-4 25 30 0 37 10"/></g>`,
      );
    paths.push(
      '<g opacity=".035" transform="translate(561 390)" fill="none" stroke="#175843" stroke-width="4">' +
        Array.from(
          { length: 16 },
          (_, i) =>
            `<ellipse rx="165" ry="90" transform="rotate(${i * 22.5})"/>`,
        ).join("") +
        "</g>",
    );
    background = await sharp(
      Buffer.from(
        `<svg xmlns="http://www.w3.org/2000/svg" width="1123" height="794"><rect width="1123" height="794" fill="#fffef9"/>${paths.join("")}<path d="M75 228H1048M75 627H1048" stroke="#c7a664" stroke-width=".7"/></svg>`,
      ),
    )
      .png()
      .toBuffer();
  }
  return background;
}
const run = (text, options = {}) =>
  new TextRun({
    text: String(text ?? ""),
    font: "Arial",
    size: 26,
    rightToLeft: true,
    color: "173F36",
    ...options,
  });
const p = (children, options = {}) =>
  new Paragraph({
    children: typeof children === "string" ? [run(children)] : children,
    alignment: AlignmentType.CENTER,
    bidirectional: true,
    spacing: { after: 90, line: 300 },
    ...options,
  });
const cell = (children, width) =>
  new TableCell({
    children,
    width: { size: width, type: WidthType.PERCENTAGE },
    borders,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 0, bottom: 0, left: 40, right: 40 },
  });
async function build(certs, settings, lang = "fa") {
  const ps = lang === "ps";
  const bg = await getBackground();
  const sections = certs.map((c) => {
    const picture = new ImageRun({
      type: "png",
      data: bg,
      transformation: { width: 1123, height: 794 },
      floating: {
        horizontalPosition: { relative: "page", offset: 0 },
        verticalPosition: { relative: "page", offset: 0 },
        behindDocument: true,
        allowOverlap: true,
      },
    });
    const header = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders,
      rows: [
        new TableRow({
          children: [
            cell(
              [
                p([
                  new ImageRun({
                    type: "jpg",
                    data: fs.readFileSync(path.join(root, "logo1.jpg")),
                    transformation: { width: 83, height: 83 },
                  }),
                ]),
              ],
              17,
            ),
            cell(
              [
                p(ps ? "د افغانستان اسلامي امارت" : "امارت اسلامی افغانستان"),
                p(ps ? "د لوړو زده کړو وزارت" : "وزارت تحصیلات عالی"),
                p(settings.university, { spacing: { after: 50 } }),
                p(settings.office, { spacing: { after: 30 } }),
              ],
              66,
            ),
            cell(
              [
                p([
                  new ImageRun({
                    type: "jpg",
                    data: fs.readFileSync(path.join(root, "logo2.jpg")),
                    transformation: { width: 83, height: 83 },
                  }),
                ]),
              ],
              17,
            ),
          ],
          cantSplit: true,
        }),
      ],
    });
    const meta = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders,
      rows: [
        new TableRow({
          children: [
            cell([p((ps ? "د ثبت ګڼه: " : "شماره ثبت: ") + c.number)], 50),
            cell(
              [p((ps ? "نېټه: " : "تاریخ صدور: ") + faDate(c.issued, lang))],
              50,
            ),
          ],
        }),
      ],
    });
    const body = ps
      ? [
          run("ښاغلی "),
          run(c.studentName, { bold: true, size: 35 }),
          run(" د "),
          run(c.father, { bold: true, size: 32 }),
          run(" زوی، د "),
          run(c.courseName, { bold: true, size: 32 }),
          run(
            " زده‌کړې کورس له " +
              faDate(c.start, lang) +
              " نېټې څخه تر " +
              faDate(c.end, lang) +
              " نېټې پورې په بریالیتوب بشپړ کړی دی. د هغه د هڅو او منظم ګډون په ستاینه دا تصدیق‌پاڼه ورکول کېږي. د لوی څښتن تعالی له درباره ورته د لا زیاتو بریالیتوبونو او پرمختګونو هیله کوو.",
          ),
        ]
      : [
          run("تصدیق می‌گردد که محترم "),
          run(c.studentName, { bold: true, size: 35 }),
          run(" ولد "),
          run(c.father, { bold: true, size: 32 }),
          run("، کورس آموزشی "),
          run(c.courseName, { bold: true, size: 32 }),
          run(
            " را از تاریخ " +
              faDate(c.start, lang) +
              " الی " +
              faDate(c.end, lang) +
              " با موفقیت به پایان رسانیده است. این تصدیق‌نامه به پاس تلاش، اشتراک منظم و تکمیل دورهٔ آموزشی برای ایشان اعطا می‌گردد. از بارگاه خداوند متعال برای ایشان موفقیت‌های بیشتر آرزو می‌نماییم.",
          ),
        ];
    const signatures = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders,
      rows: [
        new TableRow({
          children: [
            cell(
              [
                p(settings.director),
                p(
                  ps
                    ? "د محصلانو خدمتونو انسجام آمر"
                    : "آمر انسجام خدمات محصلان",
                ),
                p("________________________", { spacing: { before: 180 } }),
              ],
              38,
            ),
            cell(
              [
                p("❖", { spacing: { after: 0 } }),
                p(c.number, { spacing: { after: 0 } }),
              ],
              24,
            ),
            cell(
              [
                p(settings.deputy),
                p(ps ? "د محصلانو چارو معاون" : "معاون امور محصلان"),
                p("________________________", { spacing: { before: 180 } }),
              ],
              38,
            ),
          ],
        }),
      ],
    });
    return {
      properties: {
        page: {
          size: {
            width: 11906,
            height: 16838,
            orientation: PageOrientation.LANDSCAPE,
          },
          margin: { top: 700, right: 1050, bottom: 650, left: 1050 },
          borders: {
            pageBorders: { offsetFrom: "page", zOrder: "back" },
            ...Object.fromEntries(
              ["Top", "Right", "Bottom", "Left"].map((side) => [
                "pageBorder" + side,
                {
                  style: BorderStyle.DOUBLE,
                  size: 12,
                  color: "A98A46",
                  space: 18,
                },
              ]),
            ),
          },
        },
      },
      headers: {
        default: new Header({
          children: [
            p([picture], {
              alignment: AlignmentType.LEFT,
              bidirectional: false,
              spacing: { after: 0, line: 1 },
            }),
          ],
        }),
      },
      children: [
        header,
        meta,
        p(
          [
            run("❖  ━━━━━━━━━━━━━━━  ❖  ━━━━━━━━━━━━━━━  ❖", {
              color: "B08D47",
              size: 20,
            }),
          ],
          { spacing: { before: 100, after: 120 } },
        ),
        p(
          [
            run(ps ? "تصدیق‌پاڼه" : "تصدیق‌نامهٔ فراغت", {
              bold: true,
              size: 60,
              color: "175843",
            }),
          ],
          { spacing: { after: 80 } },
        ),
        p(
          [
            run("CERTIFICATE OF COMPLETION", {
              size: 17,
              color: "A17B37",
              rightToLeft: false,
              characterSpacing: 55,
            }),
          ],
          { spacing: { after: 220 } },
        ),
        p(body, {
          alignment: AlignmentType.JUSTIFIED,
          spacing: { after: 130, line: 460 },
        }),
        ...(c.description
          ? [
              p(
                [
                  run(String(c.description).slice(0, 180), {
                    size: 21,
                    color: "607369",
                  }),
                ],
                {
                  spacing: { after: 90 },
                },
              ),
            ]
          : []),
        p(
          [
            run((ps ? "استاد: " : "استاد دوره: ") + c.teacherName, {
              size: 23,
            }),
          ],
          { spacing: { after: 170 } },
        ),
        signatures,
      ],
    };
  });
  return Packer.toBuffer(
    new Document({
      creator: "Hamrah Mohasel",
      title: ps ? "تصدیق‌پاڼه" : "تصدیق‌نامه",
      description: "A4 landscape course completion certificates",
      styles: {
        default: {
          document: {
            run: { font: "Arial", size: 26 },
            paragraph: { bidirectional: true },
          },
        },
      },
      sections,
    }),
  );
}
module.exports = { build };
