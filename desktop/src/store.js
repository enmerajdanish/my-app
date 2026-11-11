const fs = require("node:fs");
const path = require("node:path");
const tables = [
  "users",
  "courses",
  "students",
  "enrollments",
  "attendance",
  "attendanceResults",
  "certificates",
  "seminars",
  "events",
  "settings",
  "sessions",
];
class Store {
  constructor() {
    this.queue = Promise.resolve();
  }
  async connect() {
    if (process.env.DB_KIND === "sqlite") {
      const { DatabaseSync } = require("node:sqlite");
      const file =
        process.env.SQLITE_PATH ||
        path.join(process.cwd(), "data", "courses.sqlite");
      fs.mkdirSync(path.dirname(file), { recursive: true });
      this.db = new DatabaseSync(file);
      this.db.exec(
        "PRAGMA journal_mode=WAL; PRAGMA foreign_keys=ON; PRAGMA busy_timeout=5000;",
      );
      const relations = {
        courses:
          ", teacher_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.teacherId')) STORED REFERENCES users(id) DEFERRABLE INITIALLY DEFERRED",
        enrollments:
          ", course_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.courseId')) STORED REFERENCES courses(id) DEFERRABLE INITIALLY DEFERRED, student_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.studentId')) STORED REFERENCES students(id) DEFERRABLE INITIALLY DEFERRED, UNIQUE(course_id,student_id)",
        attendance:
          ", course_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.courseId')) STORED, student_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.studentId')) STORED, FOREIGN KEY(course_id,student_id) REFERENCES enrollments(course_id,student_id) DEFERRABLE INITIALLY DEFERRED",
        attendanceResults:
          ", course_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.courseId')) STORED, student_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.studentId')) STORED, FOREIGN KEY(course_id,student_id) REFERENCES enrollments(course_id,student_id) DEFERRABLE INITIALLY DEFERRED",
        certificates:
          ", course_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.courseId')) STORED, student_id TEXT GENERATED ALWAYS AS (json_extract(data,'$.studentId')) STORED, FOREIGN KEY(course_id,student_id) REFERENCES enrollments(course_id,student_id) DEFERRABLE INITIALLY DEFERRED",
      };
      for (const t of tables)
        this.db.exec(
          `CREATE TABLE IF NOT EXISTS ${t}(id TEXT PRIMARY KEY, data TEXT NOT NULL CHECK(json_valid(data)) ${relations[t] || ""})`,
        );
      for (const [t, name, expr] of [
        ["users", "username", "json_extract(data,'$.username')"],
        ["courses", "code", "json_extract(data,'$.code')"],
        [
          "enrollments",
          "student_course",
          "json_extract(data,'$.studentId'),json_extract(data,'$.courseId')",
        ],
        [
          "attendance",
          "student_course_date",
          "json_extract(data,'$.studentId'),json_extract(data,'$.courseId'),json_extract(data,'$.date')",
        ],
        [
          "attendanceResults",
          "student_course",
          "json_extract(data,'$.studentId'),json_extract(data,'$.courseId')",
        ],
        [
          "certificates",
          "student_course",
          "json_extract(data,'$.studentId'),json_extract(data,'$.courseId')",
        ],
      ])
        this.db.exec(
          `CREATE UNIQUE INDEX IF NOT EXISTS ${t}_${name} ON ${t}(${expr})`,
        );
    } else {
      const { MongoClient } = require("mongodb");
      this.client = new MongoClient(
        process.env.MONGODB_URI || "mongodb://127.0.0.1:27017/?replicaSet=rs0",
      );
      await this.client.connect();
      this.mongo = this.client.db(process.env.MONGODB_DB || "kunduz_courses");
      for (const t of tables)
        await this.mongo.collection(t).createIndex({ id: 1 }, { unique: true });
      for (const [t, keys] of [
        ["users", { username: 1 }],
        ["courses", { code: 1 }],
        ["enrollments", { studentId: 1, courseId: 1 }],
        ["attendance", { studentId: 1, courseId: 1, date: 1 }],
        ["certificates", { studentId: 1, courseId: 1 }],
        ["attendanceResults", { studentId: 1, courseId: 1 }],
      ])
        await this.mongo.collection(t).createIndex(keys, { unique: true });
    }
  }
  async load(session) {
    const state = {};
    for (const t of tables) {
      state[t] = this.db
        ? this.db
            .prepare(`SELECT data FROM ${t}`)
            .all()
            .map((r) => JSON.parse(r.data))
        : await this.mongo
            .collection(t)
            .find({}, { session, projection: { _id: 0 } })
            .toArray();
    }
    return state;
  }
  async persist(before, after, session) {
    for (const t of tables) {
      const old = new Map(before[t].map((r) => [r.id, r]));
      const next = new Map(after[t].map((r) => [r.id, r]));
      for (const id of old.keys())
        if (!next.has(id)) {
          if (this.db) this.db.prepare(`DELETE FROM ${t} WHERE id=?`).run(id);
          else await this.mongo.collection(t).deleteOne({ id }, { session });
        }
      for (const [id, row] of next)
        if (JSON.stringify(old.get(id)) !== JSON.stringify(row)) {
          if (this.db)
            this.db
              .prepare(
                `INSERT INTO ${t}(id,data) VALUES(?,?) ON CONFLICT(id) DO UPDATE SET data=excluded.data`,
              )
              .run(id, JSON.stringify(row));
          else
            await this.mongo
              .collection(t)
              .replaceOne({ id }, row, { upsert: true, session });
        }
    }
  }
  run(fn) {
    const operation = this.queue.then(async () => {
      if (this.db) {
        this.db.exec("BEGIN IMMEDIATE");
        try {
          const before = await this.load();
          const after = structuredClone(before);
          const result = await fn(after);
          await this.persist(before, after);
          this.db.exec("COMMIT");
          return result;
        } catch (e) {
          this.db.exec("ROLLBACK");
          throw e;
        }
      }
      const session = this.client.startSession();
      try {
        let result;
        await session.withTransaction(async () => {
          const before = await this.load(session);
          const after = structuredClone(before);
          result = await fn(after);
          await this.persist(before, after, session);
        });
        return result;
      } finally {
        await session.endSession();
      }
    });
    this.queue = operation.catch(() => {});
    return operation;
  }
  async close() {
    if (this.db) this.db.close();
    if (this.client) await this.client.close();
  }
}
module.exports = { Store, tables };
