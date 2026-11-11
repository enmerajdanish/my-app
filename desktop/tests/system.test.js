const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const request = require("supertest");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
const { Store } = require("../src/store");
const { createApp } = require("../src/server");
const D = require("../src/domain");
const Paper = require("../src/paper-attendance");
function binary(agent, url) {
  return agent
    .get(url)
    .buffer(true)
    .parse((res, cb) => {
      const parts = [];
      res.on("data", (c) => parts.push(c));
      res.on("end", () => cb(null, Buffer.concat(parts)));
    });
}
async function fixture(t) {
  const folder = fs.mkdtempSync(path.join(os.tmpdir(), "hamrah-11-"));
  process.env.DB_KIND = "sqlite";
  process.env.SQLITE_PATH = path.join(folder, "courses.sqlite");
  process.env.SETUP_TOKEN = "test-token";
  const store = new Store();
  await store.connect();
  const { app } = await createApp({
    store,
    secret: "test-session-secret-at-least-32-characters-long",
  });
  t.after(async () => {
    await store.close();
    fs.rmSync(folder, { recursive: true, force: true });
  });
  const agent = request.agent(app);
  let csrf = (await agent.get("/api/session")).body.csrf;
  const post = (url, body, status = 200) =>
    agent
      .post("/api" + url)
      .set("X-CSRF-Token", csrf)
      .send(body)
      .expect(status);
  await post("/setup", {
    token: "test-token",
    name: "مدیر",
    username: "admin",
    password: "Password12345",
  });
  const login = await post("/login", {
    username: "admin",
    password: "Password12345",
  });
  csrf = login.body.csrf;
  const teacher = (
    await post("/entity/users", {
      name: "استاد",
      username: "teacher",
      password: "Password12345",
      active: true,
    })
  ).body;
  const data = {
    name: "کمپیوتر",
    code: "ICDL",
    teacherId: teacher.id,
    start: "2026-01-01",
    end: D.today(),
    startTime: "14:00",
    endTime: "15:00",
    sessions: 30,
    capacity: 100,
    status: "active",
  };
  const course = (await post("/entity/courses", data)).body;
  const a = (
    await post(`/courses/${course.id}/enroll`, {
      name: "محصل اول",
      father: "پدر اول",
    })
  ).body;
  const b = (
    await post(`/courses/${course.id}/enroll`, {
      name: "محصل دوم",
      father: "پدر دوم",
    })
  ).body;
  return {
    store,
    agent,
    post,
    getCsrf: () => csrf,
    course,
    data,
    a,
    b,
    teacher,
    app,
    folder,
  };
}
async function resultFile(rows, daily = false) {
  const w = new ExcelJS.Workbook();
  const s = w.addWorksheet("نتایج");
  s.addRow(["پوهنتون"]);
  s.addRow(["نمونه ساختگی"]);
  s.addRow(
    daily
      ? ["شماره", "اسم", "ولد", "جلسه 1", "جلسه 2", "جلسه 3"]
      : ["شماره", "اسم", "ولد", "جلسات", "حاضر", "غیرحاضر", "رخصت"],
  );
  rows.forEach((r, i) =>
    s.addRow([
      i + 1,
      r.name,
      r.father,
      ...(daily ? r.marks : [r.sessions, r.present, r.absent, r.leave]),
    ]),
  );
  return w.xlsx.writeBuffer();
}
test("حاضری کاغذی، نتایج معتبر و تصدیق‌نامه Word با کنترل دسترسی", async (t) => {
  const f = await fixture(t);
  const { agent, post, store, course, a, b } = f;
  await agent.post(`/api/courses/${course.id}/results`).send({}).expect(403);
  await post(`/courses/${course.id}/attendance`, {}, 410);
  await post(
    `/courses/${course.id}/enroll`,
    { name: a.name, father: a.father },
    400,
  );
  const before = (await agent.get(`/api/courses/${course.id}/report`)).body;
  assert.ok(before.every((r) => !r.eligible && r.source === "none"));
  const paper = await agent
    .get(`/api/courses/${course.id}/paper?sessions=30`)
    .expect(200);
  assert.match(paper.text, /A4 landscape/);
  assert.match(paper.text, /محصل اول/);
  assert.ok(!paper.text.includes("شماره تماس"));
  assert.match(paper.text, /<td><\/td>/);
  await agent
    .get(`/api/courses/${course.id}/paper?format=xlsx&sessions=30`)
    .expect(200)
    .expect("Content-Type", /spreadsheetml/);
  const rows = [
    { ...a, sessions: 30, present: 21, absent: 9, leave: 0 },
    { ...b, sessions: 30, present: 20, absent: 10, leave: 0 },
  ];
  const preview = await agent
    .post(`/api/courses/${course.id}/results/preview`)
    .set("X-CSRF-Token", f.getCsrf())
    .field("sessions", "30")
    .attach("file", Buffer.from(await resultFile(rows)), "results.xlsx")
    .expect(200);
  assert.equal(preview.body.rows[0].eligible, true);
  assert.equal(preview.body.rows[1].eligible, false);
  await post(
    `/courses/${course.id}/results`,
    {
      rows: [
        ...rows,
        { name: "نام ناشناخته", father: "ولد", sessions: 30, absent: 0 },
      ],
    },
    400,
  );
  assert.equal((await store.load()).attendanceResults.length, 0);
  await post(`/courses/${course.id}/results`, { rows: [rows[0]] }, 400);
  await post(
    `/courses/${course.id}/results`,
    { rows: [rows[0], { ...rows[1], absent: "" }] },
    400,
  );
  await post(
    `/courses/${course.id}/results`,
    { rows: [rows[0], rows[0]] },
    400,
  );
  await post(
    `/courses/${course.id}/results`,
    { rows: [rows[0], { ...rows[1], sessions: 31, present: 21 }] },
    400,
  );
  await post(`/courses/${course.id}/results`, { rows });
  await post(`/courses/${course.id}/results`, { rows });
  assert.equal((await store.load()).attendanceResults.length, 2);
  const report = (await agent.get(`/api/courses/${course.id}/report`)).body;
  assert.equal(report[0].eligible, true);
  assert.equal(report[1].eligible, false);
  await post(`/courses/${course.id}/certificates`, {}, 400);
  await post("/entity/courses/" + course.id, {
    ...f.data,
    status: "completed",
  });
  await post(`/courses/${course.id}/certificates`, { studentId: b.id }, 400);
  const cert = (await post(`/courses/${course.id}/certificates`, {})).body;
  assert.equal(cert.length, 1);
  const doc = await binary(agent, "/api/certificates/word?ids=" + cert[0].id)
    .expect(200)
    .expect("Content-Type", /wordprocessingml/);
  const zip = await JSZip.loadAsync(doc.body);
  const xml = await zip.file("word/document.xml").async("string");
  assert.match(xml, /w:orient="landscape"/);
  assert.match(xml, /w:w="16838"/);
  assert.match(xml, /w:h="11906"/);
  assert.match(xml, /محصل اول/);
  assert.ok(!xml.includes("محصل دوم"));
  assert.ok(
    Object.keys(zip.files).filter((k) => k.startsWith("word/media/")).length >=
      2,
  );
  const ps = await binary(
    agent,
    "/api/certificates/word?ids=" + cert[0].id + "&lang=ps",
  ).expect(200);
  const psXml = await (
    await JSZip.loadAsync(ps.body)
  )
    .file("word/document.xml")
    .async("string");
  assert.match(psXml.replace(/<[^>]+>/g, ""), /تائید پاڼه/);
  const ta = request.agent(f.app);
  const token = (await ta.get("/api/session")).body.csrf;
  const teacherLogin = await ta
    .post("/api/login")
    .set("X-CSRF-Token", token)
    .send({ username: "teacher", password: "Password12345" })
    .expect(200);
  await ta
    .post(`/api/courses/${course.id}/results`)
    .set("X-CSRF-Token", teacherLogin.body.csrf)
    .send({ rows })
    .expect(403);
  await ta.get("/api/certificates/word?ids=" + cert[0].id).expect(403);
  const settings = (await store.load()).settings[0];
  await post("/settings", { ...settings, language: "ps", theme: "dark" });
  assert.equal((await agent.get("/api/state")).body.settings[0].language, "ps");
  assert.equal((await agent.get("/api/state")).body.settings[0].theme, "dark");
  // A later correction that makes the formerly eligible student ineligible blocks old downloads too.
  await post(`/courses/${course.id}/results`, {
    rows: [{ ...rows[0], present: 20, absent: 10 }, rows[1]],
  });
  await agent.get("/api/certificates/word?ids=" + cert[0].id).expect(400);
  const backup = (await agent.get("/api/backup")).body;
  await post("/restore", backup);
  await agent.get("/api/state").expect(401);
});
test("قالب روزانه، اعداد فارسی و جلوگیری از نتیجه خالی یا مبهم", async () => {
  const daily = await Paper.parse(
    await resultFile(
      [{ name: "الف", father: "ب", marks: ["ح", "غ", "ر"] }],
      true,
    ),
  );
  assert.equal(daily.mode, "daily");
  assert.deepEqual(
    [daily.rows[0].present, daily.rows[0].absent, daily.rows[0].leave],
    [1, 1, 1],
  );
  const missing = await Paper.parse(
    await resultFile(
      [{ name: "الف", father: "ب", marks: ["ح", "", "ر"] }],
      true,
    ),
  );
  assert.ok(missing.rows[0].parseError);
  assert.deepEqual(
    Paper.counts({ sessions: "۳۰", present: "۲۰", absent: "۹", leave: "۱" }),
    { sessions: 30, present: 20, absent: 9, leave: 1 },
  );
  assert.throws(() => Paper.counts({ sessions: 30, absent: "" }), /خالی/);
  assert.throws(() => Paper.counts({ sessions: 30, absent: -1 }), /غیرمنفی/);
  assert.throws(
    () => Paper.counts({ sessions: 30, present: 30, absent: 9, leave: 0 }),
    /برابر/,
  );
  const state = {
    settings: [{ absenceLimit: 10 }],
    students: [
      { id: "1", name: "الف", father: "ب" },
      { id: "2", name: "الف", father: "ب" },
    ],
    enrollments: [
      { courseId: "c", studentId: "1" },
      { courseId: "c", studentId: "2" },
    ],
  };
  assert.match(
    Paper.analyze(
      state,
      "c",
      [{ name: "الف", father: "ب", sessions: 30, absent: 0 }],
      30,
    )[0].error,
    /مبهم/,
  );
});
test("ارتقای دیتابیس قدیمی و بازیابی پشتیبان نسخه قبل بدون حذف حساب‌ها", async (t) => {
  const f = await fixture(t);
  const old = (await f.agent.get("/api/backup")).body;
  delete old.data.attendanceResults;
  await f.post("/restore", old);
  assert.equal((await f.store.load()).users.length, 2);
  assert.deepEqual((await f.store.load()).attendanceResults, []);
  // Drop only the newly introduced empty table to reproduce a v1 schema.
  f.store.db.exec("DROP TABLE attendanceResults");
  await f.store.close();
  const reopened = new Store();
  await reopened.connect();
  f.store.db = reopened.db;
  const state = await f.store.load();
  assert.equal(state.students.length, 2);
  assert.equal(state.users.length, 2);
  assert.deepEqual(state.attendanceResults, []);
});

test("حفظ قالب اصلی Word و هر دو لوگو، شماره ثبت خالی و صدور چندصفحه‌ای", async () => {
  const Word = require("../src/certificates-word");
  const base = await JSZip.loadAsync(
    fs.readFileSync(path.join(__dirname, "../templates/certificate.docx")),
  );
  const original = await base.file("word/document.xml").async("string");
  const row = {
    number: "KU-DO-NOT-PRINT-123456789",
    studentName: "نام نمونه & $& الف",
    father: "پدر نمونه",
    courseName: "ICDL",
    start: "2026-08-23",
    end: "2026-11-12",
    issued: "2026-11-12",
  };
  const zip = await JSZip.loadAsync(await Word.build([row], {}));
  const xml = await zip.file("word/document.xml").async("string");
  const strip = (s) => s.replace(/(<w:t\b[^>]*>)[\s\S]*?<\/w:t>/g, "$1</w:t>");
  assert.equal(
    strip(xml),
    strip(original),
    "تمام موقعیت‌ها، کادرها، فونت‌ها و ساختار قالب باید یکسان بماند",
  );
  for (const name of [
    "word/styles.xml",
    "word/theme/theme1.xml",
    "word/media/image1.jpg",
    "word/media/image2.jpg",
  ])
    assert.deepEqual(
      await zip.file(name).async("nodebuffer"),
      await base.file(name).async("nodebuffer"),
    );
  const boxes = [
    ...xml.matchAll(/<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/g),
  ].map((m) => m[0].replace(/<[^>]+>/g, ""));
  const registration = boxes.filter((t) => t.startsWith("د ثبت "));
  assert.equal(registration.length, 2);
  registration.forEach((t) => {
    assert.ok(!/[0-9۰-۹]/.test(t));
    assert.ok(!t.includes("عـ"));
  });
  assert.ok(!xml.includes(row.number));
  assert.ok(xml.includes("نام نمونه &amp; $&amp; الف"));
  assert.ok(!/\{[A-Z_]+\}/.test(xml));
  const batch = await JSZip.loadAsync(
    await Word.build(
      [row, { ...row, studentName: "محصل دوم", father: "پدر دوم" }],
      {},
    ),
  );
  const multi = await batch.file("word/document.xml").async("string");
  assert.equal((multi.match(/<w:pageBreakBefore\/>/g) || []).length, 1);
  assert.equal((multi.match(/<w:pgSz\b/g) || []).length, 1);
  assert.match(multi, /محصل دوم/);
  const ids = [...multi.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map(
    (m) => m[1],
  );
  assert.equal(new Set(ids).size, ids.length);
});
test("فورم حاضری مطابق نمونه: تاریخ و روزها، جمعه سرخ، ۳۱ ستون و روزهای داینامیک", async () => {
  const c = {
    id: "c",
    name: "ICDL",
    teacherId: "t",
    start: "2026-03-21",
    end: "2026-06-01",
  };
  const state = {
    users: [{ id: "t", name: "استاد نمونه" }],
    settings: [{ absenceLimit: 10 }],
    students: [{ id: "s", name: "محصل نمونه", father: "پدر نمونه" }],
    enrollments: [{ courseId: "c", studentId: "s" }],
  };
  const w = new ExcelJS.Workbook();
  await w.xlsx.load(await Paper.paperWorkbook(state, c, 40));
  assert.equal(w.worksheets.length, 2);
  const sh = w.worksheets[0];
  assert.equal(sh.getCell("E2").value, "1405/1/1");
  assert.equal(sh.getCell("E3").value, "شنبه");
  assert.equal(sh.getCell("K3").value, "جمعه");
  assert.equal(sh.getCell("K2").fill.fgColor.argb, "FFFF0000");
  assert.equal(sh.getCell("AF2").value, "1405/1/28");
  assert.equal(sh.getCell("B4").value, "محصل نمونه");
  assert.equal(sh.getCell("E4").value, null);
  assert.ok(sh.getCell("AJ4").value.formula.includes('"a"'));
  assert.ok(sh.getCell("AK4").value.formula.includes('"r"'));
  assert.equal(sh.getCell("B4").font.name, "B Nazanin");
  assert.equal(sh.getCell("E4").font.name, "Webdings");
  assert.equal(w.worksheets[1].getCell("E2").value, "1405/2/1");
  assert.equal(w.worksheets[1].getCell("N2").value, null);
  const html = Paper.paperHTML(state, c, 40);
  assert.equal((html.match(/<article>/g) || []).length, 2);
  assert.match(html, /1405\/1\/28/);
  assert.ok(!html.includes("شماره تماس"));
  assert.match(html, /جمعه/);
  const daily = new ExcelJS.Workbook();
  await daily.xlsx.load(await Paper.paperWorkbook(state, c, 31));
  const ds = daily.worksheets[0];
  for (let col = 5; col <= 35; col++)
    if (ds.getCell(3, col).value !== "جمعه") ds.getCell(4, col).value = "a";
  ds.getCell("F4").value = "r";
  const parsed = await Paper.parse(await daily.xlsx.writeBuffer());
  assert.equal(parsed.profile, "kunduz-calendar");
  assert.equal(parsed.rows[0].sessions, 27);
  assert.equal(parsed.rows[0].present, 26);
  assert.equal(parsed.rows[0].absent, 1);
  assert.equal(parsed.excludedHolidayColumns.length, 4);
  ds.getCell("G4").value = "ar";
  const bad = await Paper.parse(await daily.xlsx.writeBuffer());
  assert.ok(bad.rows[0].parseError);
  assert.equal(Paper.analyze(state, "c", bad.rows, 27)[0].eligible, false);
});
test("تعداد جلسات در فایل پنج‌ستونی داینامیک است و به ۴۰ محدود نیست", async () => {
  for (const total of [12, 27, 40, 65]) {
    const w = new ExcelJS.Workbook();
    const sh = w.addWorksheet("نتایج");
    sh.mergeCells("A1:E2");
    sh.getCell("A1").value = "نتایج نهایی حاضری";
    sh.getRow(3).values = ["شماره", "اسم", "ولد", "حاضر", "غیر حاضر"];
    sh.addRow([1, "محصل نمونه", "پدر نمونه", total - 9, 9]);
    const parsed = await Paper.parse(await w.xlsx.writeBuffer());
    assert.equal(parsed.headerRow, 3);
    assert.equal(parsed.rows[0].sessions, total);
    assert.deepEqual(parsed.inferredSessionTotals, [total]);
    assert.deepEqual(Paper.counts(parsed.rows[0], 999), {
      sessions: total,
      present: total - 9,
      absent: 9,
      leave: 0,
    });
  }
});
