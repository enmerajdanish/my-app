const fs = require("node:fs");
const path = require("node:path");
const file = path.join(__dirname, "../.env");
if (fs.existsSync(file)) {
  const { parseEnv } = require("node:util");
  for (const [k, v] of Object.entries(parseEnv(fs.readFileSync(file, "utf8"))))
    if (process.env[k] === undefined) process.env[k] = v;
}
