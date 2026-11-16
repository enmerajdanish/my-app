const fs = require("node:fs");
const path = require("node:path");
const net = require("node:net");
const { spawn } = require("node:child_process");
const { MongoClient } = require("mongodb");
const root = path.join(__dirname, "..");
const localURI = "mongodb://127.0.0.1:27018/?replicaSet=rs0";
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
function managedURI(uri) {
  try {
    const u = new URL(uri);
    return (
      u.protocol === "mongodb:" &&
      ["127.0.0.1", "localhost"].includes(u.hostname) &&
      u.port === "27018" &&
      !u.username &&
      !u.password &&
      u.searchParams.get("replicaSet") === "rs0"
    );
  } catch {
    return false;
  }
}
function mongodPath(env = process.env, platform = process.platform) {
  if (env.MONGOD_PATH) {
    if (!fs.existsSync(env.MONGOD_PATH))
      throw new Error(
        "مسیر MONGOD_PATH معتبر نیست؛ مسیر کامل mongod.exe را بررسی کنید.",
      );
    return env.MONGOD_PATH;
  }
  const filename = platform === "win32" ? "mongod.exe" : "mongod";
  for (const dir of (env.PATH || "").split(
    platform === "win32" ? ";" : path.delimiter,
  )) {
    const file = path.join(dir.replace(/^"|"$/g, ""), filename);
    if (dir && fs.existsSync(file)) return file;
  }
  if (platform === "win32") {
    const folder = path.join(
      env.ProgramFiles || "C:\\Program Files",
      "MongoDB",
      "Server",
    );
    if (fs.existsSync(folder))
      for (const version of fs
        .readdirSync(folder)
        .sort((a, b) => b.localeCompare(a, undefined, { numeric: true }))) {
        const file = path.join(folder, version, "bin", filename);
        if (fs.existsSync(file)) return file;
      }
  }
  throw new Error(
    "MongoDB Server پیدا نشد. نصب Compass به‌تنهایی کافی نیست؛ MongoDB Community Server را نصب کنید یا مسیر کامل mongod.exe را در MONGOD_PATH فایل .env بنویسید.",
  );
}
function portOpen(port) {
  return new Promise((resolve) => {
    const socket = net.createConnection({ host: "127.0.0.1", port });
    const done = (value) => {
      socket.destroy();
      resolve(value);
    };
    socket.setTimeout(1000);
    socket.once("connect", () => done(true));
    socket.once("error", () => done(false));
    socket.once("timeout", () => done(false));
  });
}
async function withClient(uri, fn) {
  const client = new MongoClient(uri, {
    serverSelectionTimeoutMS: 2500,
    connectTimeoutMS: 1500,
  });
  try {
    await client.connect();
    return await fn(client);
  } finally {
    await client.close();
  }
}
function standaloneError() {
  return new Error(
    "MongoDB متصل است اما Replica Set نیست. برای دیتابیس تمرینی مستقل، MONGODB_URI فایل .env را به " +
      localURI +
      " تغییر دهید و npm run dev را دوباره اجرا کنید. این کار اطلاعات دیتابیس قبلی را انتقال نمی‌دهد.",
  );
}
async function ready(uri) {
  return withClient(uri, async (c) => {
    const hello = await c.db("admin").command({ hello: 1 });
    if (!hello.setName) throw standaloneError();
    if (!hello.isWritablePrimary)
      throw new Error("منتظر آماده‌شدن MongoDB Primary هستیم.");
    return hello;
  });
}
async function ensureMongo(uri = localURI) {
  if (!managedURI(uri)) {
    await ready(uri);
    return;
  }
  const direct = new URL(uri);
  direct.searchParams.delete("replicaSet");
  direct.searchParams.set("directConnection", "true");
  if (await portOpen(27018)) {
    await ready(uri);
    return;
  }
  const executable = mongodPath();
  const folder = path.join(root, "data", "mongodb-local");
  fs.mkdirSync(folder, { recursive: true });
  const args = [
    "--dbpath",
    folder,
    "--bind_ip",
    "127.0.0.1",
    "--port",
    "27018",
    "--replSet",
    "rs0",
    "--logpath",
    path.join(folder, "mongod.log"),
    "--logappend",
  ];
  const child = spawn(executable, args, {
    detached: true,
    stdio: "ignore",
    windowsHide: true,
  });
  let failure;
  child.once("error", (e) => {
    failure = e;
  });
  child.once("exit", (code) => {
    failure ||= new Error(
      "mongod متوقف شد؛ فایل data/mongodb-local/mongod.log را بررسی کنید. کد: " +
        code,
    );
  });
  child.unref();
  try {
    let initialized = false,
      last;
    for (let i = 0; i < 35; i++) {
      if (failure) throw failure;
      try {
        await withClient(direct.toString(), async (c) => {
          const hello = await c.db("admin").command({ hello: 1 });
          if (!hello.setName) {
            if (!hello.isreplicaset) throw standaloneError();
            const cmd = await c.db("admin").command({ getCmdLineOpts: 1 });
            if (
              path.resolve(cmd.parsed?.storage?.dbPath || "") !==
              path.resolve(folder)
            )
              throw new Error(
                "دیتابیس روی این پورت مربوط این پروژه نیست؛ پیکربندی آن تغییر داده نشد.",
              );
            try {
              await c
                .db("admin")
                .command({
                  replSetInitiate: {
                    _id: "rs0",
                    members: [{ _id: 0, host: "127.0.0.1:27018" }],
                  },
                });
            } catch (e) {
              if (e.code !== 23) throw e;
            }
          }
        });
        initialized = true;
        break;
      } catch (e) {
        last = e;
        await sleep(500);
      }
    }
    if (!initialized) throw last || new Error("MongoDB آماده نشد.");
    let primary = false;
    for (let i = 0; i < 35; i++) {
      if (failure) throw failure;
      try {
        await ready(uri);
        primary = true;
        break;
      } catch (e) {
        last = e;
        await sleep(500);
      }
    }
    if (!primary) throw last;
    console.log(
      "MongoDB محلی پروژه روی پورت 27018 آماده است؛ Docker استفاده نمی‌شود.",
    );
  } catch (e) {
    child.kill();
    throw e;
  }
}
async function main() {
  if (!fs.existsSync(path.join(root, ".env"))) require("../src/setup");
  require("../src/config");
  await ensureMongo(process.env.MONGODB_URI || localURI);
  console.log("آدرس برنامه: http://127.0.0.1:" + (process.env.PORT || 3000));
}
if (require.main === module)
  main().catch((e) => {
    console.error(e.message);
    process.exitCode = 1;
  });
module.exports = { localURI, managedURI, mongodPath, ensureMongo };
