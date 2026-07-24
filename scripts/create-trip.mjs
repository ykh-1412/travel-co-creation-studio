import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function argumentsMap(values) {
  const result = new Map();
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (!value.startsWith("--")) continue;
    const [rawKey, inlineValue] = value.slice(2).split(/=(.*)/s, 2);
    if (inlineValue !== undefined) result.set(rawKey, inlineValue);
    else if (values[index + 1] && !values[index + 1].startsWith("--")) result.set(rawKey, values[++index]);
    else result.set(rawKey, "true");
  }
  return result;
}

function positiveInteger(value, fallback) {
  const number = Math.round(Number(value));
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

function safeSlug(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48);
}

function updateEnv(text, values) {
  const lines = text.split(/\r?\n/);
  const seen = new Set();
  const updated = lines.map((line) => {
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=/);
    if (!match || !Object.prototype.hasOwnProperty.call(values, match[1])) return line;
    seen.add(match[1]);
    return `${match[1]}=${values[match[1]]}`;
  });
  for (const [key, value] of Object.entries(values)) {
    if (!seen.has(key)) updated.push(`${key}=${value}`);
  }
  return `${updated.join("\n").replace(/\n+$/g, "")}\n`;
}

const args = argumentsMap(process.argv.slice(2));
const destination = String(args.get("destination") || "").trim();
if (!destination) {
  console.error("请填写目的地，例如：npm run trip:new -- --destination=苏州 --slug=suzhou");
  process.exit(1);
}

const people = positiveInteger(args.get("people"), 6);
const days = positiveInteger(args.get("days"), 2);
const nights = positiveInteger(args.get("nights"), Math.max(1, days));
const slug = safeSlug(args.get("slug") || destination) || `trip-${Date.now()}`;
const title = String(args.get("title") || `${people} 人${destination}出行共创`).trim();
const schedule = String(args.get("schedule") || "抵达与返程时间待团队确认").trim();
const force = args.get("force") === "true";
const dataRelative = `data/trips/${slug}.json`;
const outputRelative = `outputs/${slug}`;
const dataPath = path.join(rootDir, dataRelative);

try {
  await fs.access(dataPath);
  if (!force) {
    console.error(`项目 ${dataRelative} 已存在；如确认覆盖，请追加 --force。`);
    process.exit(1);
  }
} catch {
  // A new trip project can be created.
}

const template = JSON.parse(await fs.readFile(path.join(rootDir, "data", "trip-template.json"), "utf8"));
template.project = { ...template.project, name: title, destination, days, people };
template.tripProfile = { ...template.tripProfile, schedule, groupSize: people, nights };
template.finalPlan = {
  ...template.finalPlan,
  title,
  destination,
  schedule,
  people,
  nights,
  stay: { ...template.finalPlan.stay, capacity: `目标 ${people} 人，待住宿确认` },
  updatedAt: new Date().toISOString(),
};
template.itinerary.day0[0].title = `抵达${destination} · 集合`;
template.finalPlan.itinerary = [
  ...template.itinerary.day0.map((item) => ({ day: "周五晚上", sourceUrl: "", ...item })),
  ...template.itinerary.day1.map((item) => ({ day: "周六", sourceUrl: "", ...item })),
  ...template.itinerary.day2.map((item) => ({ day: "周日", sourceUrl: "", ...item })),
];
template.finalPlan.reservations = template.reservations.map((item) => ({ ...item }));

await fs.mkdir(path.dirname(dataPath), { recursive: true });
await fs.mkdir(path.join(rootDir, outputRelative), { recursive: true });
await fs.writeFile(dataPath, `${JSON.stringify(template, null, 2)}\n`, "utf8");

const envPath = path.join(rootDir, ".env");
let envText;
try {
  envText = await fs.readFile(envPath, "utf8");
} catch {
  envText = await fs.readFile(path.join(rootDir, ".env.example"), "utf8");
}
envText = updateEnv(envText, {
  TRIP_DATA_FILE: dataRelative,
  TRIP_OUTPUT_DIR: outputRelative,
  WORKBOOK_FILE_NAME: `${destination}出行共创.xlsx`,
});
await fs.writeFile(envPath, envText, "utf8");

console.log(`已创建：${title}`);
console.log(`项目数据：${dataRelative}`);
console.log(`Excel 目录：${outputRelative}`);
console.log("重新启动网站后，新项目即可生效；原项目文件不会被删除。");
