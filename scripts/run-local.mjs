import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const executable = process.platform === "win32"
  ? path.join(rootDir, "node_modules", ".bin", "vinext.cmd")
  : path.join(rootDir, "node_modules", ".bin", "vinext");
const env = { ...process.env, PATH: `${path.dirname(process.execPath)}${path.delimiter}${process.env.PATH || ""}` };
const uiPort = String(process.env.UI_PORT || "3100");

const api = spawn(process.execPath, [path.join(rootDir, "server", "index.mjs")], { cwd: rootDir, stdio: "inherit", env });
const ui = spawn(executable, ["dev", "--hostname", "0.0.0.0", "--port", uiPort], { cwd: rootDir, stdio: "inherit", env, shell: process.platform === "win32" });

function stop(code = 0) {
  api.kill("SIGTERM");
  ui.kill("SIGTERM");
  setTimeout(() => process.exit(code), 250).unref();
}

api.on("exit", (code) => { if (code) stop(code); });
ui.on("exit", (code) => { if (code) stop(code); });
process.on("SIGINT", () => stop(0));
process.on("SIGTERM", () => stop(0));
