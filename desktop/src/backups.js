const fs = require("node:fs");
const path = require("node:path");
function scheduleBackups(store, folder) {
  const save = async () => {
    try {
      const data = await store.run((s) => structuredClone(s));
      if (!data.users.length) return;
      data.sessions = [];
      fs.mkdirSync(folder, { recursive: true, mode: 0o700 });
      const name =
        "courses-" + new Date().toISOString().replace(/[:.]/g, "-") + ".json";
      fs.writeFileSync(
        path.join(folder, name),
        JSON.stringify({
          format: "kunduz-courses-v1",
          exportedAt: new Date().toISOString(),
          data,
        }),
        { mode: 0o600 },
      );
      const files = fs
        .readdirSync(folder)
        .filter((x) => /^courses-.*\.json$/.test(x))
        .sort()
        .reverse();
      for (const old of files.slice(30)) fs.unlinkSync(path.join(folder, old));
    } catch (e) {
      console.error("پشتیبان خودکار ذخیره نشد:", e.message);
    }
  };
  const initial = setTimeout(save, 30000);
  initial.unref();
  const timer = setInterval(save, 3600000);
  timer.unref();
  return () => {
    clearTimeout(initial);
    clearInterval(timer);
  };
}
module.exports = { scheduleBackups };
