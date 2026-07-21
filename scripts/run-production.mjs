import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const executable = process.platform === "win32"
  ? path.join(rootDir, "node_modules", ".bin", "vinext.cmd")
  : path.join(rootDir, "node_modules", ".bin", "vinext");
const env = { ...process.env, PATH: `${path.dirname(process.execPath)}${path.delimiter}${process.env.PATH || ""}` };

const api = spawn(process.execPath, [path.join(rootDir, "server", "index.mjs")], { cwd: rootDir, stdio: "inherit", env });
const ui = spawn(executable, ["start", "--hostname", "0.0.0.0"], { cwd: rootDir, stdio: "inherit", env, shell: process.platform === "win32" });

let stopping = false;
function stop(code = 0) {
  if (stopping) return;
  stopping = true;
  api.kill("SIGTERM");
  ui.kill("SIGTERM");
  setTimeout(() => process.exit(code), 500).unref();
}

api.on("exit", (code) => { if (!stopping) stop(code || 0); });
ui.on("exit", (code) => { if (!stopping) stop(code || 0); });
process.on("SIGINT", () => stop(0));
process.on("SIGTERM", () => stop(0));
