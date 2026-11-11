const fs = require("node:fs");
const path = require("node:path");
const crypto = require("node:crypto");
const JSZip = require("jszip");
const templatePath = path.join(__dirname, "../templates/certificate.docx");
const escapeXML = (v) =>
  String(v ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&apos;",
      })[c],
  );
function solar(v) {
  const p = new Intl.DateTimeFormat("en-US", {
    calendar: "persian",
    timeZone: "Asia/Kabul",
    year: "numeric",
    month: "numeric",
    day: "numeric",
  }).formatToParts(new Date(v + "T12:00:00Z"));
  return Object.fromEntries(
    ["year", "month", "day"].map((k) => [k, p.find((x) => x.type === k).value]),
  );
}
function officialName(value, original) {
  const simple = (v) => String(v || "").replace(/[«»"\s]/g, "");
  return !value || simple(value) === simple(original) ? original : value;
}
function values(c, settings) {
  const start = solar(c.start),
    end = solar(c.end),
    issued = solar(c.issued);
  return {
    STUDENT: c.studentName,
    FATHER: c.father,
    COURSE: c.courseName,
    START_DAY: start.day,
    START_MONTH: start.month,
    START_YEAR: start.year,
    END_DAY: end.day,
    END_MONTH: end.month,
    END_YEAR: end.year,
    ISSUED: `${issued.year}/${issued.month}/${issued.day}`,
    DIRECTOR: officialName(settings.director, 'صفی الله "الکوزی"'),
    DEPUTY: officialName(settings.deputy, 'مولوی کمال الدین " منیب"'),
  };
}
function fill(xml, c, settings) {
  const data = values(c, settings);
  return xml.replace(/\{([A-Z_]+)\}/g, (_m, key) => {
    if (!Object.hasOwn(data, key))
      throw new Error("Unknown certificate field: " + key);
    return escapeXML(data[key]);
  });
}
function uniqueDrawingIds(xml, index) {
  if (index === 0) return xml;
  const offset = index * 10000;
  xml = xml.replace(
    /_x0000_([sit])(\d+)/g,
    (_m, type, id) => `_x0000_${type}${Number(id) + offset}`,
  );
  xml = xml.replace(/<(?:wp:docPr|pic:cNvPr|wps:cNvPr)\b[^>]*>/g, (tag) =>
    tag.replace(/\bid="(\d+)"/, (_m, id) => `id="${Number(id) + offset}"`),
  );
  xml = xml.replace(
    /(?:wp14:anchorId|wp14:editId|w14:paraId|w14:textId)="[A-Fa-f0-9]+"/g,
    (m) =>
      m.replace(
        /".*"/,
        `"${crypto.randomBytes(4).toString("hex").toUpperCase()}"`,
      ),
  );
  xml = xml.replace(
    /<w:bookmarkStart\b[^>]*\bw:name="_GoBack"[^>]*\/>|<w:bookmarkEnd\b[^>]*\bw:id="0"[^>]*\/>/g,
    "",
  );
  return xml;
}
async function build(certs, settings) {
  if (!certs.length) throw new Error("No certificates selected");
  const zip = await JSZip.loadAsync(fs.readFileSync(templatePath));
  const source = await zip.file("word/document.xml").async("string");
  const match = source.match(
    /<w:body>([\s\S]*?)(<w:sectPr\b[\s\S]*?<\/w:sectPr>)<\/w:body>/,
  );
  if (!match) throw new Error("Invalid certificate template");
  const pages = certs.map((c, index) => {
    let page = uniqueDrawingIds(fill(match[1], c, settings), index);
    if (index) page = page.replace(/<w:pPr>/, "<w:pPr><w:pageBreakBefore/>");
    return page;
  });
  zip.file(
    "word/document.xml",
    source.replace(
      match[0],
      () => "<w:body>" + pages.join("") + match[2] + "</w:body>",
    ),
  );
  // Internal certificate IDs remain in the database only. The registration textbox is blank in the template.
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
module.exports = { build };
