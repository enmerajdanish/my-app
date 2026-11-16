const internshipFields = [
  ["name", "اسم"],
  ["father", "ولد"],
  ["department", "دیپارتمنت"],
  ["faculty", "پوهنځی"],
  ["semester", "سمستر"],
  ["referral", "مرجع معرفی شده"],
  ["internship", "کارآموزی", "checkbox"],
  ["practical", "کارعملی", "checkbox"],
  ["teaching", "تدریس آموزی", "checkbox"],
  ["startSolar", "تاریخ شروع (شمسی)"],
  ["endSolar", "تاریخ ختم (شمسی)"],
];
let internshipRows = [];
const internshipFilters = {
  faculty: "",
  referral: "",
  type: "",
  from: "",
  to: "",
  order: "desc",
};
function internshipQuery() {
  return new URLSearchParams({ ...internshipFilters, search }).toString();
}
const internshipNormalize = (v) =>
  String(v || "")
    .replace(/[۰-۹]/g, (c) => "۰۱۲۳۴۵۶۷۸۹".indexOf(c))
    .replace(/[٠-٩]/g, (c) => "٠١٢٣٤٥٦٧٨٩".indexOf(c))
    .replace(/ي/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/[\s‌ـ]/g, "")
    .toLowerCase();
function internshipDateFilter(v) {
  v = internshipNormalize(v);
  if (!v) return "";
  const m = v.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  return m ? `${m[1]}/${m[2].padStart(2, "0")}/${m[3].padStart(2, "0")}` : v;
}
function internshipsPanel() {
  const f = internshipFilters,
    from = internshipDateFilter(f.from),
    to = internshipDateFilter(f.to);
  const rows = (state.internships || []).filter(
    (r) =>
      (!search ||
        internshipNormalize(Object.values(r).join(" ")).includes(
          internshipNormalize(search),
        )) &&
      (!f.faculty || r.faculty === f.faculty) &&
      (!f.referral ||
        internshipNormalize(r.referral).includes(
          internshipNormalize(f.referral),
        )) &&
      (!f.type || r[f.type]) &&
      (!from || (r.startSolar && r.startSolar >= from)) &&
      (!to || (r.startSolar && r.startSolar <= to)),
  );
  rows.sort((a, b) =>
    f.order === "asc"
      ? (a.startSolar || "").localeCompare(b.startSolar || "")
      : (b.startSolar || "").localeCompare(a.startSolar || ""),
  );
  const faculties = [
    ...new Set((state.internships || []).map((r) => r.faculty).filter(Boolean)),
  ].sort();
  return `<section class="panel"><div class="panelhead"><h2>کارآموزی، کارعملی و تدریس‌آموزی</h2><div>${button("وارد کردن Excel", "internshipImport")} <a class="btn btn-outline-primary" href="/api/internships/export?${esc(internshipQuery())}">دریافت Excel</a></div></div><p class="sub">مطابق جدول نمونه؛ هر سه نوع مستقل‌اند. تاریخ‌ها شمسی است و خانه‌های اختیاری می‌توانند خالی باشند.</p><div class="toolbar">${searchbox("جستجوی اسم، ولد یا مرجع…")}<select class="form-select" data-internship-filter="faculty" aria-label="پوهنځی"><option value="">همهٔ پوهنځی‌ها</option>${faculties.map((v) => `<option value="${esc(v)}" ${f.faculty === v ? "selected" : ""}>${esc(v)}</option>`).join("")}</select><input class="form-control" data-internship-filter="referral" placeholder="مرجع معرفی شده" value="${esc(f.referral)}"><select class="form-select" data-internship-filter="type" aria-label="نوع کارآموزی"><option value="">همهٔ انواع</option>${internshipFields
    .filter((x) => x[2] === "checkbox")
    .map(
      ([k, t]) =>
        `<option value="${k}" ${f.type === k ? "selected" : ""}>${t}</option>`,
    )
    .join(
      "",
    )}</select><input class="form-control" data-internship-filter="from" aria-label="از تاریخ شمسی" placeholder="از تاریخ: ۱۴۰۵/۰۱/۰۱" value="${esc(f.from)}"><input class="form-control" data-internship-filter="to" aria-label="تا تاریخ شمسی" placeholder="تا تاریخ: ۱۴۰۵/۱۲/۲۹" value="${esc(f.to)}"><select class="form-select" data-internship-filter="order" aria-label="ترتیب"><option value="desc" ${f.order === "desc" ? "selected" : ""}>جدیدتر نخست</option><option value="asc" ${f.order === "asc" ? "selected" : ""}>قدیمی‌تر نخست</option></select></div><span class="muted">${num(rows.length)} مورد</span>${
    rows.length
      ? table(
          ["شماره", ...internshipFields.map((x) => x[1]), "عملیات"],
          rows.map((r, i) => [
            num(i + 1),
            ...internshipFields.map(([k, t, type]) =>
              esc(type === "checkbox" ? (r[k] ? t : "—") : r[k] || "—"),
            ),
            button("ویرایش", "editEntity", "internships:" + r.id) +
              button(
                "حذف",
                "deleteEntity",
                "internships:" + r.id,
                "btn-outline-danger",
              ),
          ]),
        )
      : empty("موردی پیدا نشد.")
  }</section>`;
}
function internshipImport() {
  openModal(
    `<h2>وارد کردن کارآموزی از Excel</h2><p class="sub">فایل xlsx مطابق نمونه، حداکثر ۵ مگابایت. پیش از ثبت نهایی تمام ستون‌ها قابل بررسی و اصلاح‌اند.</p><form id="internshipPreview"><input class="form-control" type="file" name="file" accept=".xlsx" required><div id="formError"></div><div class="actions">${button("انصراف", "closeModal")}<button class="btn btn-primary">نمایش پیش‌نمایش</button></div></form>`,
  );
}
function showInternshipPreview() {
  openModal(
    `<h2>بررسی و تأیید کارآموزی</h2><p class="sub">در ستون نوع، نام همان نوع یا ✓ بنویسید؛ خانهٔ خالی یعنی انتخاب نشده. تاریخ شمسی: ۱۴۰۵/۱۲/۱۴. موارد تکراری هنگام ثبت رد می‌شوند.</p><div style="max-height:55vh;overflow:auto">${table(
      ["سطر", ...internshipFields.map((x) => x[1]), "بررسی"],
      internshipRows.map((r, i) => [
        num(r.sourceRow),
        ...internshipFields.map(
          ([k, t, type]) =>
            `<input style="min-width:100px" aria-label="${k}" data-internship-row="${i}" data-key="${k}" value="${esc(type === "checkbox" && typeof r[k] === "boolean" ? (r[k] ? t : "") : (r[k] ?? ""))}">`,
        ),
        esc(r.error || "آمادهٔ بررسی"),
      ]),
      "import-table",
    )}</div><div id="formError"></div><div class="actions">${button("انصراف", "closeModal")}${button("تأیید و ثبت نهایی", "confirmInternshipImport", "", "btn-primary")}</div>`,
  );
}
async function confirmInternshipImport(el) {
  el.disabled = true;
  try {
    const result = await post("/internships/import", { rows: internshipRows });
    await refresh();
    openModal(
      `<h2>نتیجهٔ وارد کردن فایل</h2><p>${num(result.identified)} شناسایی‌شده · ${num(result.registered)} ثبت‌شده · ${num(result.errors.length)} مورد ثبت‌نشده</p>${table(
        ["سطر", "اسم", "علت"],
        result.errors.map((e) => [num(e.row), esc(e.name), esc(e.error)]),
      )}${button("بستن", "closeModal")}`,
    );
  } catch (e) {
    el.disabled = false;
    throw e;
  }
}
document.addEventListener("input", (e) => {
  if (e.target.dataset.internshipRow !== undefined)
    internshipRows[Number(e.target.dataset.internshipRow)][
      e.target.dataset.key
    ] = e.target.value;
});
document.addEventListener("change", (e) => {
  const k = e.target.dataset.internshipFilter;
  if (k && Object.hasOwn(internshipFilters, k)) {
    internshipFilters[k] = e.target.value;
    render();
  }
});
