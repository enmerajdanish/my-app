const crypto = require("node:crypto");
const bcrypt = require("bcryptjs");
const id = () => crypto.randomUUID();
const fail = (message, status = 400) => {
  throw Object.assign(new Error(message), { status });
};
const text = (v, max = 300) =>
  String(v ?? "")
    .trim()
    .slice(0, max);
const normalize = (v) =>
  text(v)
    .replace(/ي/g, "ی")
    .replace(/ك/g, "ک")
    .replace(/\s+/g, " ")
    .toLowerCase();
const today = () =>
  new Date(Date.now() + 270 * 60000).toISOString().slice(0, 10);
function date(v, required = false) {
  if (!v && !required) return "";
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(v) ||
    Number.isNaN(Date.parse(v)) ||
    new Date(v).toISOString().slice(0, 10) !== v
  )
    fail("تاریخ معتبر انتخاب کنید.");
  return v;
}
function integer(v, min, max, label) {
  const n = Number(v);
  if (!Number.isInteger(n) || n < min || n > max) fail(`${label} معتبر نیست.`);
  return n;
}
function course(s, user, cid) {
  const c = s.courses.find((x) => x.id === cid);
  if (!c) fail("کورس پیدا نشد.", 404);
  if (user.role !== "admin" && c.teacherId !== user.id)
    fail("به این کورس دسترسی ندارید.", 403);
  return c;
}
function admin(user) {
  if (user.role !== "admin") fail("این بخش مخصوص مدیر است.", 403);
}
function report(s, cid) {
  const c = s.courses?.find((c) => c.id === cid);
  return s.enrollments
    .filter((e) => e.courseId === cid)
    .map((e) => {
      const student = s.students.find((st) => st.id === e.studentId);
      const result = (s.attendanceResults || []).find(
        (r) => r.courseId === cid && r.studentId === e.studentId,
      );
      if (result) {
        return {
          ...student,
          ...result,
          id: student.id,
          source: "import",
          unmarked: 0,
          percent: Math.round((result.present / result.sessions) * 100),
          eligible:
            result.sessions > 0 &&
            result.present + result.absent + result.leave === result.sessions &&
            result.absent < Number(s.settings[0].absenceLimit),
        };
      }
      return {
        ...student,
        sessions: c?.sessions || 0,
        present: 0,
        absent: 0,
        leave: 0,
        unmarked: c?.sessions || 0,
        percent: 0,
        eligible: false,
        source: "none",
      };
    });
}
function studentData(b) {
  const r = { name: text(b.name, 100), father: text(b.father, 100) };
  if (!r.name || !r.father) fail("اسم و ولد ضروری است.");
  for (const k of [
    "department",
    "faculty",
    "semester",
    "province",
    "phone",
    "notes",
    "card",
  ])
    r[k] = text(b[k]);
  return r;
}
function enroll(s, cid, b) {
  const c = s.courses.find((x) => x.id === cid);
  if (!c) fail("کورس پیدا نشد.");
  if (c.status !== "active") fail("ثبت‌نام فقط در کورس فعال ممکن است.");
  const data = studentData(b);
  let st = data.card ? s.students.find((x) => x.card === data.card) : null;
  st ||= s.students.find(
    (x) =>
      normalize(x.name) === normalize(data.name) &&
      normalize(x.father) === normalize(data.father) &&
      (!data.phone || !x.phone || x.phone === data.phone),
  );
  if (
    st &&
    s.enrollments.some((e) => e.courseId === cid && e.studentId === st.id)
  )
    fail("این محصل قبلاً در کورس ثبت شده است.");
  if (s.enrollments.filter((e) => e.courseId === cid).length >= c.capacity)
    fail("ظرفیت کورس تکمیل شده است.");
  if (!st) {
    st = { id: id(), ...data, createdAt: today() };
    s.students.push(st);
  }
  s.enrollments.push({
    id: id(),
    courseId: cid,
    studentId: st.id,
    date: today(),
  });
  return st;
}
async function saveEntity(s, user, kind, b, rowId) {
  admin(user);
  const existing = rowId ? s[kind]?.find((x) => x.id === rowId) : null;
  if (rowId && !existing) fail("مورد پیدا نشد.", 404);
  let data;
  if (kind === "courses") {
    data = {
      name: text(b.name, 120),
      code: text(b.code, 40),
      teacherId: text(b.teacherId),
      start: date(b.start, true),
      end: date(b.end, true),
      startTime: text(b.startTime),
      endTime: text(b.endTime),
      days: text(b.days),
      sessions: integer(b.sessions, 1, 1000, "تعداد جلسات"),
      capacity: integer(b.capacity, 1, 10000, "ظرفیت"),
      description: text(b.description, 2000),
      status: b.status,
    };
    if (
      !data.name ||
      !data.code ||
      !["active", "completed", "cancelled"].includes(data.status)
    )
      fail("نام، کد و وضعیت کورس را بررسی کنید.");
    if (data.start > data.end) fail("تاریخ ختم باید پس از شروع باشد.");
    if (
      !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(data.startTime) ||
      !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(data.endTime) ||
      data.startTime >= data.endTime
    )
      fail("ساعت شروع و ختم را بررسی کنید.");
    if (
      !s.users.some(
        (u) => u.id === data.teacherId && u.role === "teacher" && u.active,
      )
    )
      fail("استاد فعال انتخاب کنید.");
    if (s.courses.some((c) => c.id !== rowId && c.code === data.code))
      fail("کد کورس تکراری است.");
    if (
      existing &&
      s.enrollments.filter((e) => e.courseId === rowId).length > data.capacity
    )
      fail("ظرفیت کمتر از محصلین ثبت‌شده است.");
  } else if (kind === "users") {
    data = {
      name: text(b.name, 100),
      username: text(b.username, 60),
      phone: text(b.phone),
      specialty: text(b.specialty),
      role: b.role === "admin" ? "admin" : "teacher",
      active: b.active !== false,
      canEditAttendance: b.canEditAttendance === true,
    };
    if (!data.name || !data.username) fail("نام و نام کاربری ضروری است.");
    if (s.users.some((u) => u.id !== rowId && u.username === data.username))
      fail("نام کاربری تکراری است.");
    if (
      existing?.role === "teacher" &&
      data.role !== "teacher" &&
      s.courses.some((c) => c.teacherId === rowId)
    )
      fail("ابتدا کورس‌های این استاد را به استاد دیگری بسپارید.");
    if (
      existing?.role === "admin" &&
      (!data.active || data.role !== "admin") &&
      s.users.filter((u) => u.role === "admin" && u.active).length <= 1
    )
      fail("آخرین مدیر را نمی‌توان غیرفعال کرد.");
    if (b.password) {
      if (String(b.password).length < 10) fail("رمز باید حداقل ۱۰ حرف باشد.");
      data.password = await bcrypt.hash(b.password, 12);
    } else if (!existing) fail("رمز عبور ضروری است.");
  } else if (kind === "students") {
    data = studentData(b);
    if (
      data.card &&
      s.students.some((x) => x.id !== rowId && x.card === data.card)
    )
      fail("شماره کارت تکراری است.");
    if (
      s.students.some(
        (x) =>
          x.id !== rowId &&
          normalize(x.name) === normalize(data.name) &&
          normalize(x.father) === normalize(data.father),
      )
    )
      fail("محصل با همین اسم و ولد قبلاً ثبت است؛ مشخصات را بررسی کنید.");
  } else if (["seminars", "events"].includes(kind)) {
    data = {
      title: text(b.title, 200),
      topic: text(b.topic, 500),
      presenter: text(b.presenter),
      specialty: text(b.specialty),
      organizer: text(b.organizer),
      coordinator: text(b.coordinator),
      date: date(b.date, true),
      startTime: text(b.startTime),
      endTime: text(b.endTime),
      location: text(b.location),
      participants:
        b.participants === ""
          ? null
          : integer(b.participants, 0, 1000000, "تعداد اشتراک‌کنندگان"),
      description: text(b.description, 3000),
      status: b.status,
      type: kind === "events" ? text(b.type) : "seminar",
    };
    if (!data.title || !["planned", "held", "cancelled"].includes(data.status))
      fail("عنوان و وضعیت را بررسی کنید.");
    if (
      !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(data.startTime) ||
      !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(data.endTime)
    )
      fail("ساعت شروع و ختم معتبر وارد کنید.");
    if (data.startTime && data.endTime && data.startTime >= data.endTime)
      fail("ساعت ختم باید پس از شروع باشد.");
  } else fail("بخش معتبر نیست.", 404);
  const record = {
    ...(existing || { id: id(), createdAt: new Date().toISOString() }),
    ...data,
  };
  if (existing) Object.assign(existing, record);
  else s[kind].push(record);
  const { password, ...safe } = record;
  return safe;
}
module.exports = {
  id,
  fail,
  text,
  normalize,
  today,
  date,
  integer,
  course,
  admin,
  report,
  studentData,
  enroll,
  saveEntity,
};
