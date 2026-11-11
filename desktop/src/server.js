const express = require("express");
const session = require("express-session");
const helmet = require("helmet");
const { rateLimit } = require("express-rate-limit");
const bcrypt = require("bcryptjs");
const crypto = require("node:crypto");
const path = require("node:path");
const multer = require("multer");
const Paper = require("./paper-attendance");
const Word = require("./certificates-word");
const { Store } = require("./store");
const D = require("./domain");
const Docs = require("./documents");
class DatabaseSessions extends session.Store {
  constructor(store) {
    super();
    this.store = store;
  }
  get(sid, cb) {
    this.store
      .load()
      .then((s) => {
        const r = s.sessions.find(
          (x) => x.id === sid && x.expires > Date.now(),
        );
        cb(null, r?.value || null);
      })
      .catch(cb);
  }
  set(sid, value, cb = () => {}) {
    this.store
      .run((s) => {
        s.sessions = s.sessions.filter(
          (x) => x.id !== sid && x.expires > Date.now(),
        );
        s.sessions.push({ id: sid, value, expires: Date.now() + 8 * 3600000 });
      })
      .then(() => cb())
      .catch(cb);
  }
  destroy(sid, cb = () => {}) {
    this.store
      .run((s) => {
        s.sessions = s.sessions.filter((x) => x.id !== sid);
      })
      .then(() => cb())
      .catch(cb);
  }
}
async function createApp(options = {}) {
  const store = options.store || new Store();
  if (!options.store) await store.connect();
  await store.run((s) => {
    if (!s.settings.length)
      s.settings.push({
        id: "main",
        university: "پوهنتون کندوز",
        office: "آمریت انسجام خدمات محصلان",
        absenceLimit: 10,
        director: "صفی الله «الکوزی»",
        deputy: "مولوی کمال الدین «منیب»",
        language: "fa",
        theme: "light",
      });
  });
  const app = express();
  app.disable("x-powered-by");
  app.set("trust proxy", process.env.TRUST_PROXY === "1" ? 1 : false);
  app.use(
    helmet({
      contentSecurityPolicy: {
        directives: {
          defaultSrc: ["'self'"],
          scriptSrc: ["'self'", "'unsafe-inline'"],
          styleSrc: ["'self'", "'unsafe-inline'"],
          imgSrc: ["'self'", "data:"],
          fontSrc: ["'self'"],
          upgradeInsecureRequests: null,
        },
      },
      crossOriginEmbedderPolicy: false,
    }),
  );
  app.use(express.json({ limit: "3mb" }));
  const secret = process.env.SESSION_SECRET || options.secret;
  if (!secret || secret.length < 32)
    throw new Error(
      "SESSION_SECRET باید حداقل ۳۲ حرف باشد؛ npm run setup را اجرا کنید.",
    );
  app.use(
    session({
      name: "course.sid",
      secret,
      store: new DatabaseSessions(store),
      resave: false,
      saveUninitialized: false,
      cookie: {
        httpOnly: true,
        sameSite: "strict",
        secure: process.env.COOKIE_SECURE === "1",
        maxAge: 8 * 3600000,
      },
    }),
  );
  app.get("/api/session", async (req, res) => {
    req.session.csrf ||= crypto.randomBytes(32).toString("hex");
    const s = await store.load();
    const user = s.users.find((u) => u.id === req.session.userId && u.active);
    res.json({
      csrf: req.session.csrf,
      user: user ? publicUser(user) : null,
      needsSetup: s.users.length === 0,
      preferences: {
        language: s.settings[0].language || "fa",
        theme: s.settings[0].theme || "light",
      },
    });
  });
  app.use("/api", (req, res, next) => {
    res.set("Cache-Control", "no-store");
    if (
      !["GET", "HEAD"].includes(req.method) &&
      (!req.session.csrf || req.get("X-CSRF-Token") !== req.session.csrf)
    )
      return res
        .status(403)
        .json({ error: "نشست نامعتبر است؛ صفحه را تازه کنید." });
    next();
  });
  const loginLimit = rateLimit({
    windowMs: 15 * 60000,
    limit: 15,
    standardHeaders: "draft-8",
    legacyHeaders: false,
    message: { error: "تلاش‌های ورود زیاد است؛ ۱۵ دقیقه بعد کوشش کنید." },
  });
  app.post("/api/setup", loginLimit, async (req, res) => {
    await store.run(async (s) => {
      if (s.users.length) D.fail("حساب مدیر قبلاً ساخته شده است.", 403);
      if (
        !process.env.SETUP_TOKEN ||
        req.body.token !== process.env.SETUP_TOKEN
      )
        D.fail("کد راه‌اندازی درست نیست.", 403);
      if (String(req.body.password || "").length < 10)
        D.fail("رمز حداقل ۱۰ حرف باشد.");
      if (!D.text(req.body.username) || !D.text(req.body.name))
        D.fail("نام و نام کاربری ضروری است.");
      s.users.push({
        id: D.id(),
        name: D.text(req.body.name),
        username: D.text(req.body.username),
        password: await bcrypt.hash(req.body.password, 12),
        role: "admin",
        active: true,
      });
    });
    res.json({ ok: true });
  });
  app.post("/api/login", loginLimit, async (req, res) => {
    const s = await store.load();
    const u = s.users.find(
      (u) => u.username === D.text(req.body.username) && u.active,
    );
    if (
      !u ||
      !(await bcrypt.compare(String(req.body.password || ""), u.password))
    )
      D.fail("نام کاربری یا رمز عبور درست نیست.", 401);
    await new Promise((resolve, reject) =>
      req.session.regenerate((e) => (e ? reject(e) : resolve())),
    );
    req.session.userId = u.id;
    req.session.csrf = crypto.randomBytes(32).toString("hex");
    res.json({ user: publicUser(u), csrf: req.session.csrf });
  });
  app.use("/api", async (req, res, next) => {
    const s = await store.load();
    const u = s.users.find((u) => u.id === req.session.userId && u.active);
    if (!u) return res.status(401).json({ error: "لطفاً وارد سیستم شوید." });
    req.user = u;
    req.state = s;
    next();
  });
  app.post("/api/logout", (req, res, next) =>
    req.session.destroy((e) => (e ? next(e) : res.json({ ok: true }))),
  );
  app.post("/api/password", async (req, res) => {
    if (
      !(await bcrypt.compare(String(req.body.current || ""), req.user.password))
    )
      D.fail("رمز فعلی درست نیست.");
    if (String(req.body.password || "").length < 10)
      D.fail("رمز جدید حداقل ۱۰ حرف باشد.");
    const hash = await bcrypt.hash(req.body.password, 12);
    await store.run((s) => {
      s.users.find((x) => x.id === req.user.id).password = hash;
      s.sessions = s.sessions.filter(
        (x) => x.value.userId !== req.user.id || x.id === req.sessionID,
      );
    });
    res.json({ ok: true });
  });
  app.get("/api/state", (req, res) => {
    const s = structuredClone(req.state);
    s.users = s.users.map(publicUser);
    delete s.sessions;
    if (req.user.role !== "admin") {
      s.courses = s.courses.filter((c) => c.teacherId === req.user.id);
      const ids = new Set(s.courses.map((c) => c.id));
      s.enrollments = s.enrollments.filter((e) => ids.has(e.courseId));
      const students = new Set(s.enrollments.map((e) => e.studentId));
      s.students = s.students.filter((x) => students.has(x.id));
      s.attendance = s.attendance.filter((a) => ids.has(a.courseId));
      s.attendanceResults = s.attendanceResults.filter((a) =>
        ids.has(a.courseId),
      );
      s.certificates = [];
      s.seminars = [];
      s.events = [];
      s.users = s.users.filter((x) => x.id === req.user.id);
    }
    res.json(s);
  });
  app.post("/api/entity/:kind{/:id}", async (req, res) => {
    D.admin(req.user);
    const result = await store.run(async (s) => {
      const result = await D.saveEntity(
        s,
        req.user,
        req.params.kind,
        req.body,
        req.params.id,
      );
      if (req.params.kind === "users" && req.params.id && req.body.password)
        s.sessions = s.sessions.filter(
          (x) => x.value.userId !== req.params.id || x.id === req.sessionID,
        );
      return result;
    });
    res.json(result);
  });
  app.delete("/api/entity/:kind/:id", async (req, res) => {
    D.admin(req.user);
    const { kind, id } = req.params;
    if (!["courses", "users", "students", "seminars", "events"].includes(kind))
      D.fail("بخش نامعتبر است.");
    await store.run((s) => {
      if (!s[kind].some((x) => x.id === id)) D.fail("مورد پیدا نشد.", 404);
      if (
        kind === "users" &&
        (id === req.user.id || s.courses.some((c) => c.teacherId === id))
      )
        D.fail("این حساب حذف نمی‌شود؛ آن را غیرفعال کنید.");
      if (kind === "courses" && s.enrollments.some((e) => e.courseId === id))
        D.fail("کورس دارای محصل است؛ وضعیت آن را تغییر دهید.");
      if (kind === "students" && s.enrollments.some((e) => e.studentId === id))
        D.fail("ابتدا ثبت‌نام محصل را حذف کنید.");
      s[kind] = s[kind].filter((x) => x.id !== id);
    });
    res.json({ ok: true });
  });
  app.post("/api/courses/:cid/enroll", async (req, res) => {
    D.admin(req.user);
    res.json(
      await store.run((s) => {
        D.course(s, req.user, req.params.cid);
        return D.enroll(s, req.params.cid, req.body);
      }),
    );
  });
  app.delete("/api/courses/:cid/enroll/:sid", async (req, res) => {
    D.admin(req.user);
    await store.run((s) => {
      D.course(s, req.user, req.params.cid);
      if (
        s.attendanceResults.some(
          (a) =>
            a.courseId === req.params.cid && a.studentId === req.params.sid,
        ) ||
        s.attendance.some(
          (a) =>
            a.courseId === req.params.cid && a.studentId === req.params.sid,
        ) ||
        s.certificates.some(
          (a) =>
            a.courseId === req.params.cid && a.studentId === req.params.sid,
        )
      )
        D.fail("محصل دارای حاضری یا تصدیق‌نامه است؛ ثبت‌نام حذف نمی‌شود.");
      s.enrollments = s.enrollments.filter(
        (e) =>
          !(e.courseId === req.params.cid && e.studentId === req.params.sid),
      );
    });
    res.json({ ok: true });
  });
  app.post("/api/courses/:cid/attendance", (_req, res) =>
    res.status(410).json({
      error:
        "حاضری روزانه حذف شده است؛ فورم چاپی و وارد کردن نتایج حاضری را استفاده کنید.",
    }),
  );
  app.get("/api/courses/:cid/report", (req, res) => {
    D.course(req.state, req.user, req.params.cid);
    const from = D.date(req.query.from),
      to = D.date(req.query.to);
    res.json(D.report(req.state, req.params.cid, from, to));
  });
  const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 5 * 1024 * 1024, files: 1 },
  });
  app.get("/api/courses/:cid/paper", async (req, res) => {
    const c = D.course(req.state, req.user, req.params.cid);
    const sessions = D.integer(
      req.query.days || req.query.sessions || 31,
      1,
      366,
      "تعداد روزها",
    );
    const options = {
      start: D.date(req.query.start || c.start, true),
      fridayClosed: req.query.fridayClosed !== "false",
    };
    if (req.query.format === "xlsx") {
      const b = await Paper.paperWorkbook(
        req.state,
        c,
        sessions,
        req.query.lang === "ps" ? "ps" : "fa",
        options,
      );
      res
        .set({
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "Content-Disposition": 'attachment; filename="paper-attendance.xlsx"',
        })
        .send(Buffer.from(b));
    } else
      res
        .type("html")
        .send(
          Paper.paperHTML(
            req.state,
            c,
            sessions,
            req.query.lang === "ps" ? "ps" : "fa",
            options,
          ),
        );
  });
  app.get("/api/courses/:cid/results/template", async (req, res) => {
    D.admin(req.user);
    const c = D.course(req.state, req.user, req.params.cid);
    const b = await Paper.template(
      req.state,
      c,
      req.query.lang === "ps" ? "ps" : "fa",
    );
    res
      .set({
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="attendance-results.xlsx"',
      })
      .send(Buffer.from(b));
  });
  app.post(
    "/api/courses/:cid/results/preview",
    upload.single("file"),
    async (req, res) => {
      D.admin(req.user);
      D.course(req.state, req.user, req.params.cid);
      if (!req.file || !req.file.originalname.toLowerCase().endsWith(".xlsx"))
        D.fail("فایل xlsx انتخاب کنید.");
      let mapping = {};
      try {
        mapping = JSON.parse(req.body.mapping || "{}");
      } catch {
        D.fail("تنظیم ستون‌ها معتبر نیست.");
      }
      const parsed = await Paper.parse(req.file.buffer, mapping);
      res.json({
        ...parsed,
        rows: Paper.analyze(
          req.state,
          req.params.cid,
          parsed.rows,
          req.body.sessions,
        ),
      });
    },
  );
  app.post("/api/courses/:cid/results", async (req, res) => {
    D.admin(req.user);
    const result = await store.run((s) => {
      D.course(s, req.user, req.params.cid);
      const rows = Paper.analyze(
        s,
        req.params.cid,
        req.body.rows,
        req.body.sessions,
      );
      const errors = rows.filter((r) => r.error);
      if (errors.length)
        D.fail("فایل دارای خطا است؛ تمام خطاهای پیش‌نمایش را اصلاح کنید.");
      if (
        !rows.length ||
        rows.length !==
          s.enrollments.filter((e) => e.courseId === req.params.cid).length
      )
        D.fail("فایل نتایج باید تمام محصلین ثبت‌شدهٔ کورس را شامل شود.");
      if (new Set(rows.map((r) => r.sessions)).size !== 1)
        D.fail("تعداد جلسات تمام محصلین باید یکسان باشد.");
      let saved = 0;
      for (const r of rows.filter((r) => !r.error)) {
        const old = s.attendanceResults.find(
          (x) => x.courseId === req.params.cid && x.studentId === r.studentId,
        );
        const data = {
          id: old?.id || D.id(),
          courseId: req.params.cid,
          studentId: r.studentId,
          sessions: r.sessions,
          present: r.present,
          absent: r.absent,
          leave: r.leave,
          importedAt: new Date().toISOString(),
          recordedBy: req.user.id,
        };
        if (old) Object.assign(old, data);
        else s.attendanceResults.push(data);
        saved++;
      }
      return { identified: rows.length, saved, errors };
    });
    res.json(result);
  });
  app.post("/api/import/preview", upload.single("file"), async (req, res) => {
    D.admin(req.user);
    if (!req.file || !req.file.originalname.toLowerCase().endsWith(".xlsx"))
      D.fail("فایل xlsx با حجم کمتر از ۵ مگابایت انتخاب کنید.");
    let mapping;
    if (req.body.mapping) {
      try {
        mapping = JSON.parse(req.body.mapping);
      } catch {
        D.fail("تنظیم ستون‌ها معتبر نیست.");
      }
    }
    res.json(await Docs.preview(req.file.buffer, mapping));
  });
  app.post("/api/courses/:cid/import", async (req, res) => {
    D.admin(req.user);
    if (!Array.isArray(req.body.rows) || req.body.rows.length > 10000)
      D.fail("سطرها معتبر نیست.");
    const result = await store.run((s) => {
      D.course(s, req.user, req.params.cid);
      let registered = 0;
      const errors = [];
      req.body.rows.forEach((r, i) => {
        try {
          D.enroll(s, req.params.cid, r);
          registered++;
        } catch (e) {
          errors.push({
            row: r.sourceRow || i + 1,
            name: D.text(r.name),
            error: e.message,
          });
        }
      });
      return { identified: req.body.rows.length, registered, errors };
    });
    res.json(result);
  });
  app.get("/api/courses/:cid/export", async (req, res) => {
    const s = req.state,
      c = D.course(s, req.user, req.params.cid);
    let columns, rows;
    if (req.query.type === "attendance") {
      const dates = [
        ...new Set(
          s.attendance.filter((a) => a.courseId === c.id).map((a) => a.date),
        ),
      ].sort();
      columns = ["شماره", "اسم", "ولد", ...dates.map(DocsDate)];
      rows = D.report(s, c.id).map((r, i) => [
        i + 1,
        r.name,
        r.father,
        ...dates.map(
          (date) =>
            ({ present: "حاضر", absent: "غیرحاضر", leave: "رخصت" })[
              s.attendance.find(
                (a) =>
                  a.courseId === c.id &&
                  a.studentId === r.id &&
                  a.date === date,
              )?.status
            ] || "",
        ),
      ]);
    } else {
      columns = ["شماره", ...Object.values(Docs.headers)];
      rows = s.enrollments
        .filter((e) => e.courseId === c.id)
        .map((e, i) => {
          const r = s.students.find((x) => x.id === e.studentId);
          return [i + 1, ...Object.keys(Docs.headers).map((k) => r[k] || "")];
        });
    }
    const buffer = await Docs.workbook(c.name, columns, rows);
    res
      .set({
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="course.xlsx"',
      })
      .send(Buffer.from(buffer));
  });
  app.post("/api/courses/:cid/certificates", async (req, res) => {
    D.admin(req.user);
    const issued = await store.run((s) => {
      const c = D.course(s, req.user, req.params.cid);
      if (c.status !== "completed")
        D.fail("ابتدا وضعیت کورس را تکمیل‌شده تعیین کنید.");
      const reports = D.report(s, c.id);
      const targets = req.body.studentId
        ? reports.filter((r) => r.id === req.body.studentId)
        : reports.filter((r) => r.eligible);
      if (!targets.length) D.fail("محصل واجد شرایط پیدا نشد.");
      const out = [];
      for (const r of targets) {
        if (!r.eligible) D.fail("محصل واجد شرایط نیست.");
        let cert = s.certificates.find(
          (x) => x.courseId === c.id && x.studentId === r.id,
        );
        if (!cert) {
          cert = {
            id: D.id(),
            number: "KU-" + crypto.randomBytes(8).toString("hex").toUpperCase(),
            courseId: c.id,
            studentId: r.id,
            studentName: r.name,
            father: r.father,
            courseName: c.name,
            description: c.description,
            teacherName: s.users.find((u) => u.id === c.teacherId)?.name || "",
            start: c.start,
            end: c.end,
            issued: D.today(),
          };
          s.certificates.push(cert);
        }
        out.push(cert);
      }
      return out;
    });
    res.json(issued);
  });
  app.get("/api/certificates/word", async (req, res) => {
    D.admin(req.user);
    const ids = String(req.query.ids || "").split(",");
    const certs = req.state.certificates.filter((c) => ids.includes(c.id));
    if (!certs.length) D.fail("تصدیق‌نامه پیدا نشد.", 404);
    for (const cert of certs) {
      const c = D.course(req.state, req.user, cert.courseId);
      if (
        c.status !== "completed" ||
        !D.report(req.state, c.id).some(
          (r) => r.id === cert.studentId && r.eligible,
        )
      )
        D.fail("نتایج حاضری این محصل معتبر یا واجد شرایط نیست.");
    }
    const b = await Word.build(
      certs,
      req.state.settings[0],
      req.query.lang === "ps" ? "ps" : "fa",
    );
    res
      .set({
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": 'attachment; filename="certificates.docx"',
      })
      .send(b);
  });
  app.get("/api/certificates/print", (_req, res) =>
    res.status(410).json({ error: "خروجی تصدیق‌نامه به Word تغییر کرده است." }),
  );
  app.post("/api/settings", async (req, res) => {
    D.admin(req.user);
    await store.run((s) => {
      Object.assign(s.settings[0], {
        university: D.text(req.body.university),
        office: D.text(req.body.office),
        absenceLimit: D.integer(req.body.absenceLimit, 1, 365, "حد غیرحاضری"),
        director: D.text(req.body.director),
        deputy: D.text(req.body.deputy),
        language: ["fa", "ps"].includes(req.body.language)
          ? req.body.language
          : s.settings[0].language || "fa",
        theme: ["light", "dark", "system"].includes(req.body.theme)
          ? req.body.theme
          : s.settings[0].theme || "light",
      });
    });
    res.json({ ok: true });
  });
  app.get("/api/backup", (req, res) => {
    D.admin(req.user);
    const data = structuredClone(req.state);
    data.sessions = [];
    res
      .set("Content-Disposition", 'attachment; filename="courses-backup.json"')
      .json({
        format: "kunduz-courses-v1",
        exportedAt: new Date().toISOString(),
        data,
      });
  });
  app.post("/api/restore", async (req, res) => {
    D.admin(req.user);
    const b = req.body;
    if (b.format !== "kunduz-courses-v1" || !b.data)
      D.fail("فایل پشتیبان معتبر نیست.");
    await store.run((s) => {
      if (b.data.attendanceResults === undefined) b.data.attendanceResults = [];
      validateBackup(b.data);
      for (const key of Object.keys(s))
        s[key] = key === "sessions" ? [] : b.data[key];
    });
    res.json({ ok: true });
  });
  app.get("/api/activities/export", async (req, res) => {
    D.admin(req.user);
    const kind = req.query.kind === "seminars" ? "seminars" : "events";
    const rows = req.state[kind]
      .filter(
        (x) =>
          (!req.query.from || x.date >= req.query.from) &&
          (!req.query.to || x.date <= req.query.to),
      )
      .sort((a, b) => a.date.localeCompare(b.date));
    const b = await Docs.workbook(
      kind === "seminars" ? "سمینارها" : "برنامه‌ها",
      ["عنوان", "تاریخ", "مکان", "برگزارکننده", "تعداد", "وضعیت"],
      rows.map((x) => [
        x.title,
        DocsDate(x.date),
        x.location,
        x.organizer,
        x.participants,
        { planned: "برنامه‌ریزی‌شده", held: "برگزارشده", cancelled: "لغوشده" }[
          x.status
        ],
      ]),
    );
    res
      .set({
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="activities.xlsx"',
      })
      .send(Buffer.from(b));
  });
  app.use(
    "/assets/bootstrap",
    express.static(path.join(__dirname, "../node_modules/bootstrap/dist")),
  );
  app.use(
    "/assets/font",
    express.static(
      path.join(__dirname, "../node_modules/@fontsource/vazirmatn/files"),
    ),
  );
  app.use(express.static(path.join(__dirname, "../public")));
  app.use((err, req, res, next) => {
    if (!err.status) console.error(err.message);
    res.status(err.status || 400).json({
      error: err.status
        ? err.message
        : err.code === "LIMIT_FILE_SIZE"
          ? "حجم فایل بیشتر از ۵ مگابایت است."
          : "ذخیره یا خواندن اطلاعات انجام نشد؛ اطلاعات ورودی و اتصال دیتابیس را بررسی کنید.",
    });
  });
  if (process.env.BACKUP_DIR)
    require("./backups").scheduleBackups(store, process.env.BACKUP_DIR);
  return { app, store };
}
function publicUser({ password, ...u }) {
  return u;
}
function DocsDate(v) {
  return new Intl.DateTimeFormat("fa-AF", {
    calendar: "persian",
    timeZone: "Asia/Kabul",
  }).format(new Date(v + "T12:00:00Z"));
}
function validateBackup(s) {
  const { tables } = require("./store");
  for (const t of tables) {
    if (!Array.isArray(s[t])) D.fail("ساختار پشتیبان نادرست است.");
    const ids = s[t].map((x) => x.id);
    if (
      ids.some((x) => typeof x !== "string") ||
      new Set(ids).size !== ids.length
    )
      D.fail("شناسه‌های پشتیبان معتبر نیست.");
  }
  if (
    !s.users.some(
      (u) =>
        u.role === "admin" &&
        u.active &&
        typeof u.password === "string" &&
        /^\$2[aby]\$/.test(u.password),
    )
  )
    D.fail("پشتیبان مدیر معتبر ندارد.");
  if (
    s.settings.length !== 1 ||
    !Number.isInteger(s.settings[0].absenceLimit) ||
    s.settings[0].absenceLimit < 1
  )
    D.fail("تنظیمات پشتیبان معتبر نیست.");
  if (
    s.settings[0].language !== undefined &&
    !["fa", "ps"].includes(s.settings[0].language)
  )
    D.fail("تنظیمات پشتیبان معتبر نیست.");
  if (
    s.settings[0].theme !== undefined &&
    !["light", "dark", "system"].includes(s.settings[0].theme)
  )
    D.fail("تنظیمات پشتیبان معتبر نیست.");
  for (const c of s.courses) {
    if (
      !s.users.some((u) => u.id === c.teacherId) ||
      !Number.isInteger(c.capacity)
    )
      D.fail("ارتباط کورس معتبر نیست.");
    D.date(c.start, true);
    D.date(c.end, true);
  }
  for (const e of s.enrollments)
    if (
      !s.courses.some((c) => c.id === e.courseId) ||
      !s.students.some((x) => x.id === e.studentId)
    )
      D.fail("ثبت‌نام نامعتبر در پشتیبان.");
  for (const a of [...s.attendance, ...s.attendanceResults, ...s.certificates])
    if (
      !s.enrollments.some(
        (e) => e.courseId === a.courseId && e.studentId === a.studentId,
      )
    )
      D.fail("ارتباط حاضری یا تصدیق‌نامه نامعتبر است.");
  for (const r of s.attendanceResults) {
    const v = Paper.counts(r, r.sessions);
    if (
      JSON.stringify([v.sessions, v.present, v.absent, v.leave]) !==
      JSON.stringify([r.sessions, r.present, r.absent, r.leave])
    )
      D.fail("نتایج پشتیبان معتبر نیست.");
  }
  for (const a of s.attendance) {
    D.date(a.date, true);
    if (!["present", "absent", "leave"].includes(a.status))
      D.fail("حاضری نامعتبر است.");
  }
  for (const [t, keys] of [
    ["users", ["username"]],
    ["courses", ["code"]],
    ["enrollments", ["courseId", "studentId"]],
    ["attendance", ["courseId", "studentId", "date"]],
    ["attendanceResults", ["courseId", "studentId"]],
    ["certificates", ["courseId", "studentId"]],
  ]) {
    const values = s[t].map((x) => keys.map((k) => x[k]).join("|"));
    if (new Set(values).size !== values.length)
      D.fail("اطلاعات تکراری در پشتیبان است.");
  }
}
if (require.main === module) {
  require("./config");
  process.env.BACKUP_DIR ||= path.join(__dirname, "../backups");
  createApp()
    .then(({ app }) => {
      const port = Number(process.env.PORT) || 3000;
      app.listen(port, process.env.HOST || "127.0.0.1", () =>
        console.log(`http://127.0.0.1:${port}`),
      );
    })
    .catch((e) => {
      console.error(e.message);
      process.exitCode = 1;
    });
}
module.exports = { createApp };
