import http from "node:http";
import { createHash, randomUUID, timingSafeEqual } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const dataFile = path.join(rootDir, "data", "store.json");
const outputDir = path.join(rootDir, "outputs", "019f7eda-a998-7660-a73d-e43b3af67965");
const workbookPath = path.join(outputDir, "扬州团队旅行攻略.xlsx");
const port = Number(process.env.API_PORT || 8787);
const host = "0.0.0.0";
let writeInProgress = false;
let lastKnownWorkbookMtime = 0;
let processingChain = Promise.resolve();

await fs.mkdir(outputDir, { recursive: true });

try {
  const envText = await fs.readFile(path.join(rootDir, ".env"), "utf8");
  for (const line of envText.split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)\s*$/);
    if (!match || match[2] === "") continue;
    const value = match[2].replace(/^(["'])(.*)\1$/, "$2");
    process.env[match[1]] ??= value;
  }
} catch {
  // .env is optional; demo mode works without it.
}

const defaultDetails = {
  address: "待核实",
  capacity: "待核实",
  rooms: "待核实",
  beds: "待核实",
  bathrooms: "待核实",
  twoNightTotal: "待核实",
  environment: "待核实",
  entireRental: "待核实",
  kitchen: "待核实",
  barbecue: "待核实",
  bbqEquipment: "待核实",
  breakfast: "待核实",
  parking: "待核实",
  transport: "待核实",
  checkIn: "待核实",
  checkOut: "待核实",
  bookingStatus: "待确认",
  reservation: "待核实",
  openingHours: "待核实",
  difficulty: "待核实",
  horrorLevel: "待核实",
  usage: "待分配",
};

const defaultTripProfile = {
  planVersion: "weekend-v2",
  dates: "待团队确认",
  schedule: "周五晚抵达 · 周日傍晚返程",
  groupSize: 6,
  nights: 2,
  stayPreference: "环境好、整租优先、至少 3 个独立睡眠空间",
  barbecue: "周六晚在民宿烧烤",
  breakfasts: ["周六早茶", "周日早餐"],
  activity: "6 人密室 / 团队活动",
  accommodationBudget: "待团队确认",
};

const defaultReservations = [
  { id: "reserve-stay", item: "6 人住宿（周五、周六两晚）", type: "住宿", targetTime: "周五入住，周日退房", status: "待补充真实链接", owner: "待认领", deadline: "日期确认后立即", note: "整租优先；核对 3 个睡眠空间、两晚总价与取消政策" },
  { id: "reserve-bbq", item: "周六晚民宿烧烤", type: "餐饮", targetTime: "周六 18:30", status: "待确认", owner: "待认领", deadline: "订房前", note: "确认能否烧烤、设备、食材和清洁费用" },
  { id: "reserve-breakfast-1", item: "周六早茶", type: "早餐", targetTime: "周六 08:00", status: "待补充餐厅链接", owner: "待认领", deadline: "出发前 7 天", note: "6 人同桌；核对是否可预约及排队时间" },
  { id: "reserve-breakfast-2", item: "周日早餐", type: "早餐", targetTime: "周日 08:30", status: "待补充餐厅链接", owner: "待认领", deadline: "出发前 7 天", note: "优先选择民宿附近，避免影响退房" },
  { id: "reserve-escape", item: "6 人密室 / 团队活动", type: "活动", targetTime: "周日 10:30", status: "待补充真实链接", owner: "待认领", deadline: "出发前 7 天", note: "核对主题、难度、恐怖程度、时长与取消政策" },
  { id: "reserve-attraction", item: "景点 / 博物馆预约", type: "门票", targetTime: "周六至周日", status: "待确认", owner: "待认领", deadline: "开放预约后", note: "以官方开放时间和预约规则为准" },
];

const defaultFridayPlan = [
  { time: "20:00", endTime: "21:00", title: "抵达扬州 · 集合", subtitle: "六人会合后前往住宿，晚到成员可直接在民宿集合。", transport: "车站 / 自驾点 → 民宿", cost: 30, category: "抵达", note: "按最终车次调整", address: "住宿地址待定", bookingStatus: "待确认", sourceId: "" },
  { time: "21:00", endTime: "22:30", title: "入住两晚 · 夜宵碰头", subtitle: "核对房间分配、周六烧烤安排和次日集合时间。", transport: "民宿内", cost: 0, category: "住宿", note: "住宿尚未下单", address: "待补充真实民宿链接", bookingStatus: "未预订", sourceId: "place-005" },
];

const defaultSaturdayPlan = [
  { time: "08:00", endTime: "09:30", title: "周六早茶", subtitle: "六人同桌，用一壶茶和几笼点心正式打开周末。", transport: "从民宿步行 / 打车", cost: 80, category: "早餐", note: "门店和预约方式待选", address: "待补充餐厅链接", bookingStatus: "未预订", sourceId: "place-006" },
  { time: "10:00", endTime: "13:00", title: "瘦西湖慢游", subtitle: "把扬州地标放在上午，按体力选择完整或精简路线。", transport: "约 10–20 分钟车程", cost: 0, category: "景点", note: "票价和开放时间待核实", address: "蜀冈—瘦西湖（详细入口待确认）", bookingStatus: "待确认", sourceId: "place-001" },
  { time: "14:30", endTime: "17:00", title: "个园 · 东关街散步", subtitle: "园林、老城和采购一次串联，为晚上回民宿留足时间。", transport: "打车至老城后步行", cost: 0, category: "老城", note: "门票与动线待核实", address: "东关街历史文化街区", bookingStatus: "待确认", sourceId: "place-003" },
  { time: "18:30", endTime: "21:30", title: "民宿烧烤夜", subtitle: "六人一起采购、烧烤和聊天，是本次行程的固定核心。", transport: "回民宿后不再移动", cost: 120, category: "聚餐", note: "订房前确认允许烧烤、设备和清洁费", address: "住宿地址待定", bookingStatus: "未预订", sourceId: "place-005" },
];

const defaultSundayPlan = [
  { time: "08:30", endTime: "09:30", title: "周日早餐 · 整理退房", subtitle: "优先选择民宿附近，早餐后完成行李整理。", transport: "步行优先", cost: 40, category: "早餐", note: "门店待选", address: "民宿附近待定", bookingStatus: "未预订", sourceId: "" },
  { time: "10:30", endTime: "12:30", title: "6 人密室 / 团队活动", subtitle: "主题以六人可玩、难度适中和交通顺路为优先。", transport: "携带行李打车 / 先寄存", cost: 150, category: "活动", note: "主题、恐怖程度和预约规则待补充", address: "扬州市区待定", bookingStatus: "未预订", sourceId: "place-007" },
  { time: "13:00", endTime: "14:00", title: "淮扬菜午餐", subtitle: "根据密室位置选择顺路餐厅，控制用餐时间。", transport: "短途步行 / 打车", cost: 100, category: "餐饮", note: "餐厅待选", address: "待补充餐厅链接", bookingStatus: "未预订", sourceId: "" },
  { time: "14:30", endTime: "17:00", title: "中国大运河博物馆", subtitle: "用室内展览收束扬州的古今线索，也作为雨天方案。", transport: "前往运河三湾", cost: 0, category: "博物馆", note: "必须核实预约与闭馆日", address: "运河三湾（入口待核实）", bookingStatus: "待确认", sourceId: "place-004" },
  { time: "17:30", endTime: "19:00", title: "弹性返程", subtitle: "按六人的车次分批前往车站，预留行李和晚高峰时间。", transport: "打车 / 自驾返程", cost: 60, category: "返程", note: "按最终车次调整", address: "扬州东站 / 扬州站待定", bookingStatus: "待确认", sourceId: "" },
];

function itineraryWithDay(itinerary) {
  return [
    ...(itinerary.day0 || []).map((item) => ({ day: "周五晚上", sourceUrl: "", ...item })),
    ...(itinerary.day1 || []).map((item) => ({ day: "周六", sourceUrl: "", ...item })),
    ...(itinerary.day2 || []).map((item) => ({ day: "周日", sourceUrl: "", ...item })),
  ];
}

function splitFinalItinerary(items) {
  const result = { day0: [], day1: [], day2: [] };
  for (const item of Array.isArray(items) ? items : []) {
    const key = String(item.day || "").includes("周五") ? "day0" : String(item.day || "").includes("周六") ? "day1" : "day2";
    const itineraryItem = { ...item };
    delete itineraryItem.day;
    delete itineraryItem.sourceUrl;
    result[key].push(itineraryItem);
  }
  return result;
}

function finalValueStatus(value) {
  const text = String(value ?? "").trim();
  if (!text || /待|未选择|未确定|未知|尚未|占位|空位|暂无/.test(text)) return "待补充";
  return "已确定";
}

function limitedText(value, fallback = "", maximum = 500) {
  const text = String(value ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").trim();
  return (text || String(fallback ?? "").trim()).slice(0, maximum);
}

function submittedText(object, key, fallback = "", maximum = 500, allowBlank = true) {
  if (!Object.prototype.hasOwnProperty.call(object, key)) return limitedText(fallback, "", maximum);
  const text = String(object[key] ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").trim().slice(0, maximum);
  return allowBlank ? text : (text || limitedText(fallback, "", maximum));
}

function safeWebUrl(value, fallback = "") {
  const text = limitedText(value, fallback, 2_000);
  if (!text) return "";
  try {
    const parsed = new URL(text);
    return ["http:", "https:"].includes(parsed.protocol) ? parsed.toString() : "";
  } catch {
    return "";
  }
}

function boundedInteger(value, fallback, minimum, maximum) {
  const number = Math.round(Number(value));
  return Number.isFinite(number) ? Math.min(maximum, Math.max(minimum, number)) : fallback;
}

function boundedCost(value, fallback = 0) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.min(1_000_000, Math.max(0, number)) : fallback;
}

function normalizePlanDay(value) {
  const day = limitedText(value, "周六", 20);
  if (day.includes("周五") || day.includes("星期五")) return "周五晚上";
  if (day.includes("周日") || day.includes("星期日") || day.includes("周天")) return "周日";
  return "周六";
}

function sanitizeFinalPlan(input, current) {
  const source = input && typeof input === "object" ? input : {};
  const previous = current && typeof current === "object" ? current : {};
  const stay = source.stay && typeof source.stay === "object" ? source.stay : {};
  const previousStay = previous.stay && typeof previous.stay === "object" ? previous.stay : {};
  const itinerary = Array.isArray(source.itinerary) ? source.itinerary.slice(0, 100) : (previous.itinerary || []);
  const reservations = Array.isArray(source.reservations) ? source.reservations.slice(0, 100) : (previous.reservations || []);

  return {
    ...previous,
    version: "excel-home-v1",
    title: submittedText(source, "title", previous.title || "6 人扬州周末旅行", 120, false),
    destination: submittedText(source, "destination", previous.destination || "扬州", 80, false),
    dates: submittedText(source, "dates", previous.dates || "待团队确认", 120),
    schedule: submittedText(source, "schedule", previous.schedule || "周五晚抵达 · 周日傍晚返程", 160),
    people: boundedInteger(source.people, boundedInteger(previous.people, 6, 1, 50), 1, 50),
    nights: boundedInteger(source.nights, boundedInteger(previous.nights, 2, 1, 30), 1, 30),
    perPersonBudget: submittedText(source, "perPersonBudget", previous.perPersonBudget || "待团队确认", 80),
    summary: submittedText(source, "summary", previous.summary || "", 1_000),
    stay: {
      name: submittedText(stay, "name", previousStay.name || "待选择真实民宿", 160),
      address: submittedText(stay, "address", previousStay.address || "待补充民宿详细地址", 300),
      capacity: submittedText(stay, "capacity", previousStay.capacity || "目标 6 人", 160),
      roomsBeds: submittedText(stay, "roomsBeds", previousStay.roomsBeds || "待确认", 240),
      twoNightTotal: submittedText(stay, "twoNightTotal", previousStay.twoNightTotal || "待确认", 120),
      checkInOut: submittedText(stay, "checkInOut", previousStay.checkInOut || "待确认", 200),
      barbecue: submittedText(stay, "barbecue", previousStay.barbecue || "待确认", 300),
      bbqEquipment: submittedText(stay, "bbqEquipment", previousStay.bbqEquipment || "待确认", 400),
      breakfast: submittedText(stay, "breakfast", previousStay.breakfast || "待确认", 400),
      sourceUrl: Object.prototype.hasOwnProperty.call(stay, "sourceUrl") ? safeWebUrl(stay.sourceUrl) : safeWebUrl(previousStay.sourceUrl || ""),
    },
    itinerary: itinerary.map((raw, index) => {
      const item = raw && typeof raw === "object" ? raw : {};
      return {
        day: normalizePlanDay(item.day),
        time: submittedText(item, "time", "", 30),
        endTime: submittedText(item, "endTime", "", 30),
        category: submittedText(item, "category", "安排", 60, false),
        title: submittedText(item, "title", `新增安排 ${index + 1}`, 160, false),
        subtitle: submittedText(item, "subtitle", "", 800),
        address: submittedText(item, "address", "待补充", 300),
        transport: submittedText(item, "transport", "待补充", 300),
        cost: boundedCost(item.cost),
        bookingStatus: submittedText(item, "bookingStatus", "待确认", 80),
        sourceUrl: safeWebUrl(item.sourceUrl),
        note: submittedText(item, "note", "", 500),
        sourceId: submittedText(item, "sourceId", "", 120),
      };
    }),
    reservations: reservations.map((raw) => {
      const item = raw && typeof raw === "object" ? raw : {};
      return {
        id: limitedText(item.id, `reserve-${randomUUID()}`, 120),
        item: submittedText(item, "item", "新增待办", 200, false),
        type: submittedText(item, "type", "待分类", 60),
        targetTime: submittedText(item, "targetTime", "待确认", 120),
        status: submittedText(item, "status", "待确认", 80),
        owner: submittedText(item, "owner", "待认领", 80),
        deadline: submittedText(item, "deadline", "待确认", 120),
        note: submittedText(item, "note", "", 500),
      };
    }),
    updatedAt: new Date().toISOString(),
  };
}

function normalizeState(raw) {
  const state = raw && typeof raw === "object" ? raw : {};
  const hasWeekendPlan = state.tripProfile?.planVersion === defaultTripProfile.planVersion;
  state.project = { name: "扬州周末共创攻略", destination: "扬州", days: 2, people: 6, budget: 6000, status: "方案共创中", tagline: "周五晚集合，周末一起住、一起吃、一起玩。", ...(state.project || {}) };
  if (state.project.name === "扬州两日慢游") state.project.name = "扬州周末共创攻略";
  if (state.project.tagline === "一半烟火，一半园林。把散落的灵感，整理成一起出发的路线。") state.project.tagline = "周五晚集合，周末一起住、一起吃、一起玩。";
  state.project.people = 6;
  state.project.days = 2;
  state.tripProfile = { ...defaultTripProfile, ...(state.tripProfile || {}) };
  state.tripProfile.breakfasts = Array.isArray(state.tripProfile.breakfasts) ? state.tripProfile.breakfasts : defaultTripProfile.breakfasts;
  state.links = Array.isArray(state.links) ? state.links.map((item) => {
    const completed = item.status === "已写入Excel";
    const processing = ["等待处理", "正在读取", "AI分析中"].includes(item.status);
    const blocked = item.status === "需要人工补充";
    return {
      ...item,
      readStatus: item.readStatus || (completed ? "成功读取" : blocked ? "读取受限" : item.status === "处理失败" ? "读取失败" : "等待读取"),
      organizedStatus: item.organizedStatus || (completed ? "已整理" : processing ? "整理中" : "未整理"),
      factsFound: Array.isArray(item.factsFound) ? item.factsFound : (completed ? ["页面标题", "内容分类", "核心摘要"] : []),
      missingFields: Array.isArray(item.missingFields) ? item.missingFields : (completed ? ["价格或预约等动态信息仍需核实"] : ["网页正文未完整读取"]),
      resultNote: item.resultNote || (completed ? "已生成候选资料并写入 Excel" : item.error || "等待后台处理"),
    };
  }) : [];
  const seedDetails = {
    "place-005": { address: "东关街周边（具体门牌待真实链接）", capacity: "目标 6 人，实际容量待核实", rooms: "目标至少 3 个独立睡眠空间", beds: "床型与床数待核实", bathrooms: "待核实", twoNightTotal: "日期和房源确认后计算", environment: "环境好、安静、公共空间充足", entireRental: "整租优先，待核实", kitchen: "希望可用，待核实", barbecue: "必须确认允许，当前未知", bbqEquipment: "设备、炭火和清洁费待核实", breakfast: "不强求含早，周边早餐需方便", parking: "待核实", transport: "老城步行 / 打车便利优先", checkIn: "周五晚，具体时间待核实", checkOut: "周日，具体时间待核实", bookingStatus: "未预订", reservation: "等待团队提交真实民宿链接" },
    "place-006": { address: "老城 / 瘦西湖周边待选", capacity: "6 人同桌", openingHours: "早餐时段待核实", reservation: "核对能否提前排号或预约", bookingStatus: "未预订", usage: "周六早茶" },
    "place-007": { address: "扬州市区待选", capacity: "6 人", openingHours: "待核实", reservation: "通常需预约，具体规则待真实链接", bookingStatus: "未预订", difficulty: "中等优先", horrorLevel: "团队确认", usage: "周日上午团队活动" },
  };
  state.places = Array.isArray(state.places) ? state.places.map((item) => ({
    ...item,
    tags: Array.isArray(item.tags) ? item.tags : [],
    pros: Array.isArray(item.pros) ? item.pros : [],
    cons: Array.isArray(item.cons) ? item.cons : [],
    votes: item.votes && typeof item.votes === "object" && !Array.isArray(item.votes) ? item.votes : {},
    details: { ...defaultDetails, ...(seedDetails[item.id] || {}), ...(item.details || {}), address: item.details?.address || seedDetails[item.id]?.address || item.area || "待核实" },
  })) : [];
  state.itinerary = state.itinerary || {};
  if (!hasWeekendPlan) {
    state.itinerary.day0 = defaultFridayPlan;
    state.itinerary.day1 = defaultSaturdayPlan;
    state.itinerary.day2 = defaultSundayPlan;
  }
  state.itinerary.day0 = Array.isArray(state.itinerary.day0) && state.itinerary.day0.length ? state.itinerary.day0 : defaultFridayPlan;
  for (const day of ["day0", "day1", "day2"]) {
    state.itinerary[day] = (Array.isArray(state.itinerary[day]) ? state.itinerary[day] : []).map((item) => ({ address: "待确认", bookingStatus: "待确认", sourceId: "", ...item }));
  }
  state.reservations = Array.isArray(state.reservations) && state.reservations.length ? state.reservations : defaultReservations;
  const defaultStay = state.places.find((item) => item.id === "place-005") || state.places.find((item) => item.category === "住宿") || {};
  const defaultFinalPlan = {
    version: "excel-home-v1",
    title: "6 人扬州周末旅行",
    destination: state.project.destination,
    dates: state.tripProfile.dates,
    schedule: state.tripProfile.schedule,
    people: state.tripProfile.groupSize,
    nights: state.tripProfile.nights,
    perPersonBudget: "待团队确认",
    summary: "周五晚抵达，周六早茶与园林、晚上民宿烧烤，周日安排早餐、密室和返程。",
    stay: {
      name: "待选择真实民宿",
      address: "待补充民宿详细地址",
      capacity: "目标 6 人，待房源确认",
      roomsBeds: "至少 3 个独立睡眠空间，床型待确认",
      twoNightTotal: "待确认日期和房源后计算",
      checkInOut: "周五晚上入住 · 周日退房",
      barbecue: "待确认民宿允许周六晚烧烤",
      bbqEquipment: "设备、食材和清洁费用待确认",
      breakfast: "周六早茶、周日早餐；具体门店待选择",
      sourceUrl: defaultStay.sourceUrl || "",
    },
    itinerary: itineraryWithDay(state.itinerary).map((item) => ({ ...item, sourceUrl: state.places.find((place) => place.id === item.sourceId)?.sourceUrl || item.sourceUrl || "" })),
    reservations: state.reservations.map((item) => ({ ...item })),
    updatedAt: new Date().toISOString(),
  };
  state.finalPlan = {
    ...defaultFinalPlan,
    ...(state.finalPlan || {}),
    stay: { ...defaultFinalPlan.stay, ...(state.finalPlan?.stay || {}) },
    itinerary: Array.isArray(state.finalPlan?.itinerary) && state.finalPlan.itinerary.length ? state.finalPlan.itinerary : defaultFinalPlan.itinerary,
    reservations: Array.isArray(state.finalPlan?.reservations) && state.finalPlan.reservations.length ? state.finalPlan.reservations : defaultFinalPlan.reservations,
  };
  if (state.finalPlan.version === "excel-home-v1") {
    state.itinerary = splitFinalItinerary(state.finalPlan.itinerary);
    state.reservations = state.finalPlan.reservations.map((item) => ({ ...item }));
    state.project.name = state.finalPlan.title;
    state.project.destination = state.finalPlan.destination;
    state.project.people = Number(state.finalPlan.people) || state.project.people;
    state.tripProfile.dates = state.finalPlan.dates;
    state.tripProfile.schedule = state.finalPlan.schedule;
    state.tripProfile.groupSize = Number(state.finalPlan.people) || state.tripProfile.groupSize;
    state.tripProfile.nights = Number(state.finalPlan.nights) || state.tripProfile.nights;
  }
  state.settings = { provider: "演示分析", workbookPath: "", lastExcelSync: "", ...(state.settings || {}) };
  return state;
}

async function readState() {
  return normalizeState(JSON.parse(await fs.readFile(dataFile, "utf8")));
}

async function writeState(state) {
  const temporary = `${dataFile}.tmp`;
  await fs.writeFile(temporary, `${JSON.stringify(state, null, 2)}\n`, "utf8");
  await fs.rename(temporary, dataFile);
}

function providerLabel() {
  const provider = (process.env.AI_PROVIDER || "demo").toLowerCase();
  if (provider === "deepseek") return `DeepSeek · ${process.env.DEEPSEEK_MODEL || "deepseek-v4-flash"}`;
  if (provider === "doubao") return "豆包";
  if (provider === "ollama") return `Ollama · ${process.env.OLLAMA_MODEL || "qwen2.5:7b"}`;
  return "演示分析";
}

function sendJson(response, status, payload) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type",
    "Access-Control-Allow-Methods": "GET,POST,DELETE,OPTIONS",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(payload));
}

function safeCredentialMatch(actual, expected) {
  const actualBytes = Buffer.from(actual);
  const expectedBytes = Buffer.from(expected);
  return actualBytes.length === expectedBytes.length && timingSafeEqual(actualBytes, expectedBytes);
}

function isPublicTunnelRequest(request) {
  if (/^(1|true|yes)$/i.test(process.env.PUBLIC_REQUIRE_PASSWORD || "")) return true;
  const hostHeader = String(request.headers.host || "");
  return Boolean(request.headers["cf-connecting-ip"]) || hostHeader.endsWith(".trycloudflare.com");
}

function publicAccessToken() {
  const password = process.env.PUBLIC_ACCESS_PASSWORD || "";
  return password ? createHash("sha256").update(`yangzhou-trip:${password}`).digest("hex") : "";
}

function hasPublicAccess(request) {
  const expected = publicAccessToken();
  const cookies = String(request.headers.cookie || "").split(";");
  const supplied = cookies.map((item) => item.trim()).find((item) => item.startsWith("yz_access="))?.slice("yz_access=".length) || "";
  return Boolean(expected) && safeCredentialMatch(supplied, expected);
}

function sendPasswordPage(response, invalid = false) {
  const error = invalid ? '<p class="error" role="alert">密码不正确，请重新输入。</p>' : "";
  const html = `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>进入下扬州 · 团队旅行共创台</title>
<style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;background:#f6f1e7;color:#173d3b;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC",sans-serif}body:before{content:"";position:fixed;inset:0;background:radial-gradient(circle at 15% 15%,rgba(15,105,99,.16),transparent 34%),radial-gradient(circle at 85% 85%,rgba(217,112,62,.15),transparent 32%);pointer-events:none}.card{position:relative;width:min(440px,100%);padding:38px;border:1px solid rgba(23,61,59,.16);border-radius:28px;background:rgba(255,252,246,.94);box-shadow:0 24px 70px rgba(23,61,59,.14)}.mark{display:grid;place-items:center;width:52px;height:52px;border-radius:16px;background:#0f6963;color:white;font:700 24px serif}.eyebrow{margin:28px 0 8px;color:#d9703e;font-size:12px;font-weight:800;letter-spacing:.18em}.card h1{margin:0;font:700 clamp(28px,7vw,38px)/1.15 Georgia,"Songti SC",serif}.intro{margin:14px 0 28px;color:#5e7471;line-height:1.7}label{display:block;margin-bottom:9px;font-size:13px;font-weight:800}input{width:100%;height:52px;padding:0 16px;border:1px solid #cfdad5;border-radius:14px;background:white;color:#173d3b;font-size:17px;outline:none}input:focus{border-color:#0f6963;box-shadow:0 0 0 4px rgba(15,105,99,.1)}button{width:100%;height:52px;margin-top:14px;border:0;border-radius:14px;background:#0f6963;color:white;font-size:16px;font-weight:800;cursor:pointer}button:hover{background:#0b5752}.error{margin:0 0 12px;padding:10px 12px;border-radius:12px;background:#f5dfcd;color:#8a3f20;font-size:13px}.note{margin:18px 0 0;color:#81908d;font-size:12px;text-align:center}
</style></head><body><main class="card"><div class="mark">扬</div><p class="eyebrow">TEAM TRIP · YANGZHOU</p><h1>朋友，输入密码<br>一起下扬州。</h1><p class="intro">这是团队内部的旅行共创空间。只需要输入共享密码，不需要注册账号。</p>${error}<form method="post" action="/__team-login"><label for="password">团队访问密码</label><input id="password" name="password" type="password" autocomplete="current-password" required autofocus placeholder="请输入密码"><button type="submit">进入共创台</button></form><p class="note">密码由旅行发起人提供</p></main></body></html>`;
  response.writeHead(invalid ? 401 : 200, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(html),
  });
  response.end(html);
}

async function authorizePublicRequest(request, response, url) {
  if (!isPublicTunnelRequest(request)) return true;
  if (request.method === "POST" && url.pathname === "/__team-login") {
    const chunks = [];
    let size = 0;
    for await (const chunk of request) {
      size += chunk.length;
      if (size > 10_000) throw new Error("登录内容过大");
      chunks.push(chunk);
    }
    const form = new URLSearchParams(Buffer.concat(chunks).toString("utf8"));
    const supplied = form.get("password") || "";
    const expected = process.env.PUBLIC_ACCESS_PASSWORD || "";
    if (expected && safeCredentialMatch(supplied, expected)) {
      response.writeHead(303, {
        Location: "/",
        "Set-Cookie": `yz_access=${publicAccessToken()}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=2592000`,
        "Cache-Control": "no-store",
      });
      response.end();
      return false;
    }
    sendPasswordPage(response, true);
    return false;
  }
  if (hasPublicAccess(request)) return true;
  sendPasswordPage(response, false);
  return false;
}

async function readBody(request) {
  const chunks = [];
  let size = 0;
  for await (const chunk of request) {
    size += chunk.length;
    if (size > 1_000_000) throw new Error("提交内容过大");
    chunks.push(chunk);
  }
  return chunks.length ? JSON.parse(Buffer.concat(chunks).toString("utf8")) : {};
}

function cleanUrl(value) {
  const url = new URL(value);
  url.hash = "";
  for (const key of [...url.searchParams.keys()]) {
    if (/^(utm_|spm|from|source)/i.test(key)) url.searchParams.delete(key);
  }
  return url.toString();
}

function decodeEntities(value) {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
}

function pageText(html) {
  return decodeEntities(html)
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 18_000);
}

function extractTitle(html, fallback) {
  const og = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)/i)
    || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:title["']/i);
  const title = og?.[1] || html.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1] || fallback;
  return decodeEntities(title).replace(/\s+/g, " ").trim().slice(0, 160);
}

function classify(text, requested) {
  if (requested && requested !== "自动识别") {
    if (requested === "密室/活动") return "活动";
    if (requested === "攻略文章") return "攻略";
    return requested;
  }
  if (/民宿|酒店|公寓|客栈|房型|入住|住宿/i.test(text)) return "住宿";
  if (/密室|剧本杀|逃脱|沉浸式|团建|活动/i.test(text)) return "活动";
  if (/餐厅|早茶|美食|人均|菜单|淮扬菜/i.test(text)) return "餐饮";
  if (/景区|园林|博物馆|门票|景点|游览/i.test(text)) return "景点";
  return "攻略";
}

function parseJsonFromModel(text) {
  const trimmed = text.trim().replace(/^```json\s*/i, "").replace(/```$/i, "").trim();
  return JSON.parse(trimmed);
}

const systemPrompt = `你是扬州 6 人周末旅行的资料整理助手。团队周五晚抵达、周日返程，住两晚，周六晚希望在民宿烧烤，并需要两顿早餐和一次密室或团队活动。只根据网页正文输出严格 JSON，不得猜测或编造。
顶层字段：name, category, area, price(number或null), priceLabel, duration, summary, tags(string数组), pros(string数组), cons(string数组), score(0到5), dataStatus, factsFound(string数组), missingFields(string数组), details(object)。
category 只能是住宿、活动、景点、餐饮、攻略。
details 必须包含：address, capacity, rooms, beds, bathrooms, twoNightTotal, environment, entireRental, kitchen, barbecue, bbqEquipment, breakfast, parking, transport, checkIn, checkOut, bookingStatus, reservation, openingHours, difficulty, horrorLevel, usage。
住宿重点提取完整地址、6人容量、房间/床/卫浴、两晚总价、环境、是否整租、厨房、是否允许烧烤及设备、早餐、停车、交通、入住退房和取消/预订要求；活动重点提取地址、适合人数、时长、难度、恐怖程度和预约要求；餐饮重点提取地址、营业时间、6人预约与适合周六早茶/周日早餐/正餐的用途。网页没有明确写出的字段一律写“待核实”，并放入 missingFields。`;

async function modelAnalysis(content, fallback) {
  const provider = (process.env.AI_PROVIDER || "demo").toLowerCase();
  if (provider === "deepseek") {
    if (!process.env.DEEPSEEK_API_KEY) throw new Error("DeepSeek 模式缺少 DEEPSEEK_API_KEY");
    const base = (process.env.DEEPSEEK_BASE_URL || "https://api.deepseek.com").replace(/\/$/, "");
    const model = process.env.DEEPSEEK_MODEL || "deepseek-v4-flash";
    const response = await fetch(`${base}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${process.env.DEEPSEEK_API_KEY}` },
      body: JSON.stringify({
        model,
        thinking: { type: "disabled" },
        temperature: 0.2,
        max_tokens: 1_800,
        response_format: { type: "json_object" },
        messages: [{ role: "system", content: systemPrompt }, { role: "user", content }],
      }),
      signal: AbortSignal.timeout(90_000),
    });
    if (!response.ok) throw new Error(`DeepSeek 请求失败：${response.status}`);
    const data = await response.json();
    return parseJsonFromModel(data.choices?.[0]?.message?.content || "{}");
  }
  if (provider === "doubao") {
    if (!process.env.DOUBAO_API_KEY || !process.env.DOUBAO_MODEL) throw new Error("豆包模式缺少 DOUBAO_API_KEY 或 DOUBAO_MODEL");
    const base = (process.env.DOUBAO_BASE_URL || "https://ark.cn-beijing.volces.com/api/v3").replace(/\/$/, "");
    const response = await fetch(`${base}/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${process.env.DOUBAO_API_KEY}` },
      body: JSON.stringify({ model: process.env.DOUBAO_MODEL, temperature: 0.2, response_format: { type: "json_object" }, messages: [{ role: "system", content: systemPrompt }, { role: "user", content }] }),
      signal: AbortSignal.timeout(90_000),
    });
    if (!response.ok) throw new Error(`豆包请求失败：${response.status}`);
    const data = await response.json();
    return parseJsonFromModel(data.choices?.[0]?.message?.content || "{}");
  }
  if (provider === "ollama") {
    const base = (process.env.OLLAMA_BASE_URL || "http://127.0.0.1:11434").replace(/\/$/, "");
    const response = await fetch(`${base}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: process.env.OLLAMA_MODEL || "qwen2.5:7b", stream: false, format: "json", messages: [{ role: "system", content: systemPrompt }, { role: "user", content }] }),
      signal: AbortSignal.timeout(120_000),
    });
    if (!response.ok) throw new Error(`Ollama 请求失败：${response.status}`);
    const data = await response.json();
    return parseJsonFromModel(data.message?.content || "{}");
  }
  return fallback;
}

function demoAnalysis(title, text, category, url) {
  const priceMatch = text.match(/(?:¥|￥|人均|价格)[^\d]{0,6}(\d{2,5})/i);
  const price = priceMatch ? Number(priceMatch[1]) : null;
  const hostname = new URL(url).hostname.replace(/^www\./, "");
  return {
    name: title || hostname,
    category,
    area: text.includes("扬州") ? "扬州（具体区域待核实）" : "地点待核实",
    price,
    priceLabel: price ? `参考 ¥${price}` : "价格待核实",
    duration: "时长待核实",
    summary: `已读取 ${hostname} 的网页标题与公开正文，并按关键词归入“${category}”。当前为演示分析，关键事实仍需人工确认。`,
    tags: [category, "网页已读取", "事实待核实"],
    pros: ["链接和网页摘要已进入统一资料库，便于团队继续比较"],
    cons: ["未启用大模型深度分析，价格、时间与预订要求需要人工核实"],
    score: 3.5,
    dataStatus: "演示规则分析，等待人工核实",
    factsFound: ["页面标题", "公开正文", "内容分类"],
    missingFields: ["详细地址", "动态价格", "预订要求"],
    details: { ...defaultDetails },
  };
}

function stringList(value, fallback = []) {
  return Array.isArray(value) ? value.map((item) => String(item).trim()).filter(Boolean) : fallback;
}

function normalizeDetails(value) {
  const supplied = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  return Object.fromEntries(Object.keys(defaultDetails).map((key) => {
    const detail = supplied[key];
    return [key, detail === null || detail === undefined || detail === "" ? defaultDetails[key] : String(detail).slice(0, 500)];
  }));
}

async function updateLink(id, patch) {
  const state = await readState();
  const index = state.links.findIndex((item) => item.id === id);
  if (index < 0) return null;
  state.links[index] = { ...state.links[index], ...patch, updatedAt: new Date().toISOString() };
  await writeState(state);
  return state.links[index];
}

async function processLink(id) {
  const record = await updateLink(id, { status: "正在读取", readStatus: "等待读取", organizedStatus: "整理中", resultNote: "正在读取网页", model: providerLabel(), error: "" });
  if (!record) return;
  try {
    const response = await fetch(record.url, {
      headers: { "User-Agent": "Mozilla/5.0 (compatible; YangzhouTripStudio/1.0; local team research)" },
      redirect: "follow",
      signal: AbortSignal.timeout(20_000),
    });
    if (!response.ok) throw new Error(`网页返回 ${response.status}`);
    const html = (await response.text()).slice(0, 1_500_000);
    const title = extractTitle(html, new URL(record.url).hostname);
    const text = pageText(html);
    if (text.length < 40) throw new Error("网页正文过少，可能需要登录或验证码");
    const category = classify(`${title} ${text}`, record.category);
    await updateLink(id, { status: "AI分析中", title, category, readStatus: "成功读取", organizedStatus: "整理中", resultNote: "网页已读取，正在提取事实" });
    const fallback = demoAnalysis(title, text, category, record.url);
    const analysis = await modelAnalysis(`链接：${record.url}\n标题：${title}\n正文：${text}`, fallback);
    const state = await readState();
    const linkIndex = state.links.findIndex((item) => item.id === id);
    if (linkIndex < 0) return;
    const factsFound = stringList(analysis.factsFound, fallback.factsFound).slice(0, 12);
    const missingFields = stringList(analysis.missingFields, fallback.missingFields).slice(0, 12);
    state.links[linkIndex] = {
      ...state.links[linkIndex],
      title,
      category: analysis.category || category,
      summary: analysis.summary || fallback.summary,
      status: "已写入Excel",
      readStatus: "成功读取",
      organizedStatus: "已整理",
      factsFound,
      missingFields,
      resultNote: missingFields.length ? `已整理；仍有 ${missingFields.length} 项待核实` : "已完整整理并写入 Excel",
      model: providerLabel(),
      updatedAt: new Date().toISOString(),
      error: "",
    };
    const existingIndex = state.places.findIndex((item) => item.sourceId === id);
    const place = {
      id: existingIndex >= 0 ? state.places[existingIndex].id : `place-${randomUUID()}`,
      sourceId: id,
      name: analysis.name || title,
      category: analysis.category || category,
      area: analysis.area || "地点待核实",
      price: Number.isFinite(analysis.price) ? analysis.price : null,
      priceLabel: analysis.priceLabel || "价格待核实",
      duration: analysis.duration || "时长待核实",
      score: Math.max(0, Math.min(5, Number(analysis.score) || 3.5)),
      tags: stringList(analysis.tags, [category]).slice(0, 5),
      pros: stringList(analysis.pros, ["等待进一步分析"]).slice(0, 4),
      cons: stringList(analysis.cons, ["关键信息待核实"]).slice(0, 4),
      selected: existingIndex >= 0 ? state.places[existingIndex].selected : false,
      sourceUrl: record.url,
      dataStatus: analysis.dataStatus || "AI总结，等待人工核实",
      details: normalizeDetails(analysis.details),
    };
    if (existingIndex >= 0) state.places[existingIndex] = place;
    else state.places.unshift(place);
    await writeState(state);
    await syncToExcel(state);
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    const blocked = /登录|验证码|正文过少|403|401/.test(message);
    await updateLink(id, {
      status: blocked ? "需要人工补充" : "处理失败",
      readStatus: blocked ? "读取受限" : "读取失败",
      organizedStatus: "未整理",
      factsFound: [],
      missingFields: ["网页正文", "名称", "地址", "价格", "特点与预订信息"],
      resultNote: `未整理：${message}`,
      error: message,
      summary: "网页未能自动读取，系统没有生成未经证实的内容。",
    });
    const state = await readState();
    await syncToExcel(state).catch(() => {});
  }
}

function enqueueLink(id) {
  processingChain = processingChain.then(() => processLink(id)).catch((error) => {
    console.error("Link task failed:", error);
  });
}

const headers = {
  links: ["ID", "提交时间", "提交人", "原始链接", "页面标题", "自动分类", "读取结果", "整理结果", "处理状态", "已提取信息", "缺失信息", "AI摘要", "分析模型", "错误原因", "备注"],
  report: ["ID", "提交人", "页面标题", "分类", "读取结果", "整理结果", "结果说明", "已提取信息", "缺失信息", "错误原因", "原始链接", "更新时间"],
  stays: ["ID", "名称", "区域", "详细地址", "每晚价格", "两晚总价", "适合人数", "房间", "床位", "卫浴", "环境特点", "是否整租", "厨房", "能否烧烤", "烧烤设备/费用", "早餐", "停车", "交通", "入住时间", "退房时间", "预订状态", "优点", "缺点", "推荐分", "想去票", "可以票", "不考虑票", "投票详情", "是否入选", "原始链接", "数据状态"],
  activities: ["ID", "名称", "活动类型", "区域", "详细地址", "价格", "适合人数", "时长", "难度", "恐怖程度", "营业时间", "预约要求", "预订状态", "优点", "缺点", "推荐分", "想去票", "可以票", "不考虑票", "投票详情", "是否入选", "原始链接", "数据状态"],
  food: ["ID", "名称", "类型", "适合安排", "区域", "详细地址", "人均价格", "营业时间", "预约要求", "预订状态", "推荐菜/特点", "优点", "缺点", "推荐分", "想去票", "可以票", "不考虑票", "投票详情", "是否入选", "原始链接", "数据状态"],
  guides: ["ID", "标题", "涉及区域", "AI摘要", "避坑信息", "交通建议", "已提取信息", "缺失信息", "想去票", "可以票", "不考虑票", "投票详情", "是否入选", "原始链接", "数据状态"],
  itinerary: ["日期", "开始时间", "结束时间", "类型", "地点", "活动安排", "详细地址", "交通", "预计费用/人", "预订状态", "注意事项", "来源ID", "来源链接"],
  reservations: ["ID", "待办事项", "类型", "计划时间", "当前状态", "负责人", "完成期限", "核对说明"],
  settings: ["设置项", "当前内容", "说明"],
  overview: ["项目指标", "当前值", "说明"],
};

const sheetDescriptions = {
  "链接汇总": "每一条团队链接的完整台账；读取失败也会保留，不会静默丢失。",
  "处理报告": "快速查看哪些链接已成功整理、哪些未整理，以及缺失了什么。",
  "住宿候选": "查看团队投票并比较 6 人住两晚、烧烤与房型；黄色“是否入选”可直接修改。",
  "活动候选": "查看团队投票、密室和景点条件；黄色“是否入选”可直接修改。",
  "餐饮候选": "查看团队投票和餐饮候选；黄色“是否入选”可直接修改。",
  "攻略文章": "查看攻略摘要、团队投票和缺失信息；黄色“是否入选”可直接修改。",
  "两日行程": "包含周五晚抵达，以及周六、周日两天的初版安排。",
  "预订清单": "所有需要团队确认或下单的事项；网站不会代替你付款。",
  "项目设置": "这次扬州周末旅行的需求约束与运行设置。",
  "项目总览": "当前资料完成度、候选数量和下一步重点。",
};

function columnWidth(header) {
  if (["ID", "来源ID"].includes(header)) return 20;
  if (/原始链接|来源链接/.test(header)) return 36;
  if (/摘要|优点|缺点|说明|信息|特点|交通/.test(header)) return 30;
  if (/标题|名称|地点|待办事项|活动安排/.test(header)) return 24;
  if (/时间|日期|状态|结果|分类|类型|区域|价格|费用|人数|分|房间|床位|卫浴|厨房|早餐|停车|负责人|期限/.test(header)) return 16;
  return 18;
}

function placeVoteColumns(place) {
  const votes = place.votes && typeof place.votes === "object" ? place.votes : {};
  const entries = Object.entries(votes);
  return [
    entries.filter(([, choice]) => choice === "想去").length,
    entries.filter(([, choice]) => choice === "可以").length,
    entries.filter(([, choice]) => choice === "不考虑").length,
    entries.map(([name, choice]) => `${name}：${choice}`).join("；"),
  ];
}

function ensureSheet(workbook, name, sheetHeaders) {
  let sheet = workbook.getWorksheet(name);
  const created = !sheet;
  if (!sheet) {
    sheet = workbook.addWorksheet(name, { views: [{ showGridLines: false, state: "frozen", ySplit: 3 }] });
  }
  if (sheet.getCell(1, 1).isMerged) sheet.unMergeCells(1, 1, 1, 1);
  if (sheet.getCell(2, 1).isMerged) sheet.unMergeCells(2, 1, 2, 1);
  sheet.mergeCells(1, 1, 1, sheetHeaders.length);
  sheet.mergeCells(2, 1, 2, sheetHeaders.length);
  sheet.views = [{ showGridLines: false, state: "frozen", ySplit: 3 }];
  sheet.getCell(1, 1).value = name;
  sheet.getCell(2, 1).value = sheetDescriptions[name] || "团队旅行资料工作表";
  sheet.getRow(3).values = sheetHeaders;
  sheet.getRow(1).height = 36;
  sheet.getRow(2).height = 34;
  sheet.getRow(3).height = 34;
  for (let column = 1; column <= sheetHeaders.length; column += 1) {
    const titleCell = sheet.getRow(1).getCell(column);
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "173D3B" } };
    titleCell.font = { name: "PingFang SC", color: { argb: "FFF8EE" }, size: 15, bold: true };
    titleCell.alignment = { vertical: "middle", horizontal: column === 1 ? "left" : "center" };
    const descriptionCell = sheet.getRow(2).getCell(column);
    descriptionCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "E9E1D4" } };
    descriptionCell.font = { name: "PingFang SC", color: { argb: "496462" }, size: 10, italic: column === 1 };
    descriptionCell.alignment = { vertical: "middle", wrapText: true };
    const headerCell = sheet.getRow(3).getCell(column);
    headerCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F6963" } };
    headerCell.font = { name: "PingFang SC", color: { argb: "FFFFFF" }, size: 10, bold: true };
    headerCell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    headerCell.border = { bottom: { style: "thin", color: { argb: "D9703E" } } };
    sheet.getColumn(column).width = columnWidth(sheetHeaders[column - 1]);
  }
  if (created) sheet.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  sheet.properties.defaultRowHeight = 24;
  return sheet;
}

function recreateSheet(workbook, name, sheetHeaders) {
  const existing = workbook.getWorksheet(name);
  if (existing) workbook.removeWorksheet(existing.id);
  return ensureSheet(workbook, name, sheetHeaders);
}

function replaceRows(sheet, rows, columnCount) {
  const previousLastRow = sheet.rowCount;
  for (let rowNumber = 4; rowNumber <= previousLastRow; rowNumber += 1) {
    sheet.getRow(rowNumber).values = [];
  }
  for (const [index, values] of rows.entries()) {
    const row = sheet.getRow(4 + index);
    row.values = values.map((value) => value === "" ? null : value);
    row.height = 34;
    for (let column = 1; column <= columnCount; column += 1) {
      row.getCell(column).alignment = { vertical: "middle", wrapText: true };
      row.getCell(column).font = { name: "PingFang SC", color: { argb: "173D3B" }, size: 9 };
      row.getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: index % 2 ? "FFF9EF" : "F7F1E7" } };
      row.getCell(column).border = { bottom: { style: "hair", color: { argb: "D7DED8" } } };
    }
  }
  sheet.autoFilter = { from: { row: 3, column: 1 }, to: { row: Math.max(4, 3 + rows.length), column: columnCount } };
}

function styleCandidateDecisions(sheet, sheetHeaders, rowCount) {
  const decisionColumn = sheetHeaders.indexOf("是否入选") + 1;
  if (!decisionColumn) return;
  for (let rowNumber = 4; rowNumber < 4 + Math.max(rowCount, 1); rowNumber += 1) {
    const cell = sheet.getCell(rowNumber, decisionColumn);
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4D6" } };
    cell.font = { name: "PingFang SC", color: { argb: "8A542C" }, size: 9, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.dataValidation = { type: "list", allowBlank: false, formulae: ['"是,否"'] };
  }
}

function homeStatusStyle(status) {
  if (status === "已确定" || status === "已完成" || status === "已预订") {
    return { fill: "DCECE4", font: "23634C" };
  }
  return { fill: "F5DFCD", font: "9B4D28" };
}

function writeFinalPlanSheet(workbook, state) {
  const sheet = workbook.addWorksheet("行程首页", {
    views: [{ showGridLines: false, state: "frozen", ySplit: 3 }],
    properties: { tabColor: { argb: "D9703E" } },
  });
  const fp = state.finalPlan;
  const stay = fp.stay;
  const widths = [18, 18, 18, 14, 20, 25, 24, 20, 14, 16, 26, 36];
  widths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });

  sheet.mergeCells("A1:L1");
  sheet.getCell("A1").value = "扬州旅行 · 最终方案首页";
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "173D3B" } };
  sheet.getCell("A1").font = { name: "Songti SC", color: { argb: "FFF8EE" }, size: 22, bold: true };
  sheet.getCell("A1").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getRow(1).height = 48;

  sheet.mergeCells("A2:L2");
  sheet.getCell("A2").value = "网站只读取这一页作为最终攻略；后面的工作表都是候选资料和处理记录。";
  sheet.getCell("A2").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "DCEBE6" } };
  sheet.getCell("A2").font = { name: "PingFang SC", color: { argb: "173D3B" }, size: 11, bold: true };
  sheet.getCell("A2").alignment = { vertical: "middle", horizontal: "center" };
  sheet.getRow(2).height = 34;

  const basics = [
    ["方案标题", fp.title, "显示为网站主标题"],
    ["目的地", fp.destination, "城市或主要目的地"],
    ["出行日期", fp.dates, "确定日期后直接在黄色单元格修改"],
    ["行程结构", fp.schedule, "例如：周五晚抵达 · 周日返程"],
    ["同行人数", fp.people, "用于住宿、餐饮和活动人数判断"],
    ["住宿晚数", fp.nights, "本次为周五、周六两晚"],
    ["人均预算", fp.perPersonBudget, "可填写数字或预算区间"],
    ["方案说明", fp.summary, "网站首页的行程摘要"],
  ];
  const stayFields = [
    ["民宿名称", stay.name, "确定住宿后填写真实名称"],
    ["详细地址", stay.address, "尽量填写完整门牌或平台可见地址"],
    ["适合人数", stay.capacity, "确认房源允许 6 人入住"],
    ["房间 / 床位", stay.roomsBeds, "写清房间数、床型和床数"],
    ["两晚总价", stay.twoNightTotal, "填写含清洁费、服务费后的总价"],
    ["入住 / 退房", stay.checkInOut, "填写具体时间和延迟入住限制"],
    ["能否烧烤", stay.barbecue, "周六晚核心条件"],
    ["烧烤设备 / 费用", stay.bbqEquipment, "烤炉、炭火、食材、清洁费与限制"],
    ["早餐安排", stay.breakfast, "周六早茶和周日早餐的具体门店"],
    ["民宿原始链接", stay.sourceUrl, "保留预订平台或介绍页链接"],
  ];
  const completionValues = [...basics, ...stayFields].map((item) => item[1]);
  const confirmed = completionValues.filter((item) => finalValueStatus(item) === "已确定").length;
  sheet.mergeCells("A3:L3");
  sheet.getCell("A3").value = `黄色单元格可直接修改；保存 Excel 后约 4 秒同步网站。当前已确定 ${confirmed}/${completionValues.length} 项，待补充 ${completionValues.length - confirmed} 项。`;
  sheet.getCell("A3").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4D6" } };
  sheet.getCell("A3").font = { name: "PingFang SC", color: { argb: "8A542C" }, size: 10 };
  sheet.getCell("A3").alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.getRow(3).height = 34;

  function section(row, title, description) {
    sheet.mergeCells(row, 1, row, 12);
    const cell = sheet.getCell(row, 1);
    cell.value = `${title}  ·  ${description}`;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "0F6963" } };
    cell.font = { name: "PingFang SC", color: { argb: "FFFFFF" }, size: 11, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    sheet.getRow(row).height = 30;
  }

  function fieldRow(row, label, fieldValue, note) {
    sheet.getCell(row, 1).value = label;
    sheet.mergeCells(row, 2, row, 3);
    sheet.getCell(row, 2).value = fieldValue === "" ? null : fieldValue;
    const status = finalValueStatus(fieldValue);
    sheet.getCell(row, 4).value = status;
    sheet.mergeCells(row, 5, row, 12);
    sheet.getCell(row, 5).value = note;
    sheet.getCell(row, 1).font = { name: "PingFang SC", color: { argb: "496462" }, size: 10, bold: true };
    sheet.getCell(row, 2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4D6" } };
    sheet.getCell(row, 2).font = { name: "PingFang SC", color: { argb: "1B5E7A" }, size: 10 };
    sheet.getCell(row, 4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: homeStatusStyle(status).fill } };
    sheet.getCell(row, 4).font = { name: "PingFang SC", color: { argb: homeStatusStyle(status).font }, size: 9, bold: true };
    sheet.getCell(row, 4).alignment = { vertical: "middle", horizontal: "center" };
    sheet.getCell(row, 5).font = { name: "PingFang SC", color: { argb: "758785" }, size: 9 };
    for (let column = 1; column <= 12; column += 1) {
      sheet.getCell(row, column).alignment = { ...sheet.getCell(row, column).alignment, vertical: "middle", wrapText: true };
      sheet.getCell(row, column).border = { bottom: { style: "hair", color: { argb: "D7DED8" } } };
    }
    sheet.getRow(row).height = label === "方案说明" || label === "烧烤设备 / 费用" || label === "早餐安排" ? 46 : 34;
  }

  section(5, "一、基础信息", "网站顶部与旅行概览");
  basics.forEach((item, index) => fieldRow(6 + index, ...item));
  section(15, "二、住宿确认", "民宿未确定前保持橙色“待补充”");
  stayFields.forEach((item, index) => fieldRow(16 + index, ...item));

  section(27, "三、周末行程", "直接增删或修改下面的行程，网站按这里显示");
  const itineraryHeaders = ["日期", "开始时间", "结束时间", "类型", "安排", "详细说明", "地址", "交通", "费用/人", "预订状态", "来源链接", "备注"];
  sheet.getRow(28).values = itineraryHeaders;
  itineraryHeaders.forEach((header, index) => {
    const cell = sheet.getCell(28, index + 1);
    cell.value = header;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "173D3B" } };
    cell.font = { name: "PingFang SC", color: { argb: "FFFFFF" }, size: 9, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });
  sheet.getRow(28).height = 32;
  fp.itinerary.forEach((item, index) => {
    const rowNumber = 29 + index;
    const row = sheet.getRow(rowNumber);
    row.values = [item.day, item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost, item.bookingStatus, item.sourceUrl || null, item.note || null];
    row.height = 42;
    for (let column = 1; column <= 12; column += 1) {
      const cell = row.getCell(column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: index % 2 ? "FFF9EF" : "FFF4D6" } };
      cell.font = { name: "PingFang SC", color: { argb: column === 11 ? "1B5E7A" : "173D3B" }, size: 9 };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "hair", color: { argb: "D7DED8" } } };
    }
    row.getCell(9).numFmt = "#,##0";
    row.getCell(10).dataValidation = { type: "list", allowBlank: true, formulae: ['"待确认,未预订,已预订,已完成,无需预订"'] };
  });

  const reservationSectionRow = 30 + fp.itinerary.length;
  section(reservationSectionRow, "四、预订清单", "负责人和状态也会同步到网站");
  const reservationHeaderRow = reservationSectionRow + 1;
  const reservationHeaders = ["ID", "待办事项", "类型", "计划时间", "当前状态", "负责人", "完成期限", "核对说明"];
  sheet.getRow(reservationHeaderRow).values = reservationHeaders;
  reservationHeaders.forEach((header, index) => {
    const cell = sheet.getCell(reservationHeaderRow, index + 1);
    cell.value = header;
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "173D3B" } };
    cell.font = { name: "PingFang SC", color: { argb: "FFFFFF" }, size: 9, bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });
  sheet.getRow(reservationHeaderRow).height = 32;
  fp.reservations.forEach((item, index) => {
    const row = sheet.getRow(reservationHeaderRow + 1 + index);
    row.values = [item.id, item.item, item.type, item.targetTime, item.status, item.owner, item.deadline, item.note || null];
    row.height = 38;
    for (let column = 1; column <= 8; column += 1) {
      const cell = row.getCell(column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: index % 2 ? "FFF9EF" : "FFF4D6" } };
      cell.font = { name: "PingFang SC", color: { argb: "173D3B" }, size: 9 };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "hair", color: { argb: "D7DED8" } } };
    }
    row.getCell(5).dataValidation = { type: "list", allowBlank: true, formulae: ['"待补充真实链接,待补充餐厅链接,待确认,未预订,已预订,已完成"'] };
  });

  sheet.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, printArea: `A1:L${reservationHeaderRow + fp.reservations.length}` };
  return sheet;
}

async function syncToExcel(state) {
  state = normalizeState(state);
  writeInProgress = true;
  try {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "下扬州 · 团队旅行共创台";
    workbook.created = new Date();
    writeFinalPlanSheet(workbook, state);
    const linkSheet = ensureSheet(workbook, "链接汇总", headers.links);
    replaceRows(linkSheet, state.links.map((item) => [item.id, item.createdAt, item.submitter, item.url, item.title, item.category, item.readStatus, item.organizedStatus, item.status, item.factsFound.join("；"), item.missingFields.join("；"), item.summary, item.model, item.error || "", item.note]), headers.links.length);

    const reportSheet = ensureSheet(workbook, "处理报告", headers.report);
    replaceRows(reportSheet, state.links.map((item) => [item.id, item.submitter, item.title || "标题未读取", item.category, item.readStatus, item.organizedStatus, item.resultNote, item.factsFound.join("；"), item.missingFields.join("；"), item.error || "", item.url, item.updatedAt]), headers.report.length);

    const stays = state.places.filter((item) => item.category === "住宿");
    const staySheet = ensureSheet(workbook, "住宿候选", headers.stays);
    replaceRows(staySheet, stays.map((item) => {
      const d = item.details;
      return [item.id, item.name, item.area, d.address, item.price, d.twoNightTotal, d.capacity, d.rooms, d.beds, d.bathrooms, d.environment, d.entireRental, d.kitchen, d.barbecue, d.bbqEquipment, d.breakfast, d.parking, d.transport, d.checkIn, d.checkOut, d.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, ...placeVoteColumns(item), item.selected ? "是" : "否", item.sourceUrl, item.dataStatus];
    }), headers.stays.length);
    styleCandidateDecisions(staySheet, headers.stays, stays.length);

    const activities = state.places.filter((item) => ["活动", "景点"].includes(item.category));
    const activitySheet = ensureSheet(workbook, "活动候选", headers.activities);
    replaceRows(activitySheet, activities.map((item) => [item.id, item.name, item.category, item.area, item.details.address, item.price, item.details.capacity, item.duration, item.details.difficulty, item.details.horrorLevel, item.details.openingHours, item.details.reservation, item.details.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, ...placeVoteColumns(item), item.selected ? "是" : "否", item.sourceUrl, item.dataStatus]), headers.activities.length);
    styleCandidateDecisions(activitySheet, headers.activities, activities.length);

    const food = state.places.filter((item) => item.category === "餐饮");
    const foodSheet = ensureSheet(workbook, "餐饮候选", headers.food);
    replaceRows(foodSheet, food.map((item) => [item.id, item.name, "餐饮", item.details.usage, item.area, item.details.address, item.price, item.details.openingHours, item.details.reservation, item.details.bookingStatus, item.tags.join("；"), item.pros.join("；"), item.cons.join("；"), item.score, ...placeVoteColumns(item), item.selected ? "是" : "否", item.sourceUrl, item.dataStatus]), headers.food.length);
    styleCandidateDecisions(foodSheet, headers.food, food.length);

    const guides = state.places.filter((item) => item.category === "攻略");
    const guideSheet = ensureSheet(workbook, "攻略文章", headers.guides);
    replaceRows(guideSheet, guides.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      return [item.id, item.name, item.area, link?.summary || item.pros.join("；"), item.cons.join("；"), item.details.transport, link?.factsFound?.join("；") || "", link?.missingFields?.join("；") || "", ...placeVoteColumns(item), item.selected ? "是" : "否", item.sourceUrl, item.dataStatus];
    }), headers.guides.length);
    styleCandidateDecisions(guideSheet, headers.guides, guides.length);

    const itinerarySheet = ensureSheet(workbook, "两日行程", headers.itinerary);
    const itineraryRows = [
      ...state.itinerary.day0.map((item) => ["周五晚上", item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost, item.bookingStatus, item.note, item.sourceId, state.places.find((place) => place.id === item.sourceId)?.sourceUrl || ""]),
      ...state.itinerary.day1.map((item) => ["周六", item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost, item.bookingStatus, item.note, item.sourceId, state.places.find((place) => place.id === item.sourceId)?.sourceUrl || ""]),
      ...state.itinerary.day2.map((item) => ["周日", item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost, item.bookingStatus, item.note, item.sourceId, state.places.find((place) => place.id === item.sourceId)?.sourceUrl || ""]),
    ];
    replaceRows(itinerarySheet, itineraryRows, headers.itinerary.length);

    const reservationSheet = ensureSheet(workbook, "预订清单", headers.reservations);
    replaceRows(reservationSheet, state.reservations.map((item) => [item.id, item.item, item.type, item.targetTime, item.status, item.owner, item.deadline, item.note]), headers.reservations.length);

    state.settings.provider = providerLabel();
    const settingsSheet = recreateSheet(workbook, "项目设置", headers.settings);
    replaceRows(settingsSheet, [
      ["目的地", state.project.destination, "本次主题为扬州"],
      ["出行结构", state.tripProfile.schedule, "具体日期待团队确认"],
      ["同行人数", state.tripProfile.groupSize, "按 6 人统一比较住宿与活动"],
      ["住宿晚数", state.tripProfile.nights, "周五、周六两晚"],
      ["住宿偏好", state.tripProfile.stayPreference, "重点检查房间、床位、卫浴与环境"],
      ["住宿预算", state.tripProfile.accommodationBudget, "确认日期后再定预算"],
      ["周六晚安排", state.tripProfile.barbecue, "订房前必须确认允许烧烤"],
      ["早餐", state.tripProfile.breakfasts.join("；"), "两顿都纳入预订清单"],
      ["团队活动", state.tripProfile.activity, "密室或同类活动"],
      ["AI 分析模型", providerLabel(), "所有 AI 结论仍需人工核实"],
      ["访问方式", "本地后台 + 密码保护的公网网址", "朋友无需账号、无需同一 Wi-Fi"],
    ], headers.settings.length);

    const completedLinks = state.links.filter((item) => item.organizedStatus === "已整理").length;
    const unorganizedLinks = state.links.filter((item) => item.organizedStatus === "未整理").length;
    const overviewSheet = recreateSheet(workbook, "项目总览", headers.overview);
    replaceRows(overviewSheet, [
      ["旅行框架", "周五晚 + 周六周日", "6 人，住宿两晚"],
      ["已收集链接", state.links.length, "所有链接均保留处理记录"],
      ["已整理链接", completedLinks, "已形成候选资料并写入分类表"],
      ["未整理链接", unorganizedLinks, "请在处理报告查看原因并补充信息"],
      ["候选资料", state.places.length, "含住宿、活动、餐饮、景点和攻略"],
      ["已入选候选", state.places.filter((item) => item.selected).length, "用于当前行程草案"],
      ["待确认预订", state.reservations.filter((item) => !/已完成|已预订/.test(item.status)).length, "付款前由团队最终确认"],
      ["当前重点", "补充真实民宿、两顿早餐和密室链接", "先核对住宿能否烧烤"],
    ], headers.overview.length);

    for (const sheet of workbook.worksheets) {
      sheet.pageSetup = { ...(sheet.pageSetup || {}), orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
      sheet.headerFooter = { ...(sheet.headerFooter || {}), oddFooter: "&L扬州团队旅行攻略&R第 &P / &N 页" };
    }
    state.settings.lastExcelSync = new Date().toISOString();
    await writeState(state);
    const temporary = `${workbookPath}.tmp.xlsx`;
    await workbook.xlsx.writeFile(temporary);
    await fs.rename(temporary, workbookPath);
    lastKnownWorkbookMtime = (await fs.stat(workbookPath)).mtimeMs;
  } finally {
    writeInProgress = false;
  }
}

function headerIndex(sheet) {
  const map = new Map();
  sheet.getRow(3).eachCell((cell, column) => map.set(String(cell.value || "").trim(), column));
  return map;
}

function value(row, map, name) {
  const column = map.get(name);
  const raw = column ? row.getCell(column).value : null;
  if (raw && typeof raw === "object" && "text" in raw) return raw.text;
  if (raw && typeof raw === "object" && "result" in raw) return raw.result ?? "";
  return raw ?? "";
}

function scalarCell(cell) {
  const raw = cell?.value;
  if (raw && typeof raw === "object" && "text" in raw) return raw.text;
  if (raw && typeof raw === "object" && "result" in raw) return raw.result ?? "";
  return raw ?? "";
}

function importFinalPlanHome(workbook, state) {
  const sheet = workbook.getWorksheet("行程首页");
  if (!sheet) return false;
  const labelRows = new Map();
  let itineraryHeaderRow = 0;
  let reservationHeaderRow = 0;
  sheet.eachRow((row, number) => {
    const first = String(scalarCell(row.getCell(1)) || "").trim();
    const second = String(scalarCell(row.getCell(2)) || "").trim();
    if (first) labelRows.set(first, number);
    if (first === "日期" && second === "开始时间") itineraryHeaderRow = number;
    if (first === "ID" && second === "待办事项") reservationHeaderRow = number;
  });
  const field = (label, fallback = "") => {
    const row = labelRows.get(label);
    if (!row) return fallback;
    const edited = scalarCell(sheet.getCell(row, 2));
    return edited === "" || edited === null || edited === undefined ? fallback : edited;
  };
  const existing = state.finalPlan;
  const itinerary = [];
  if (itineraryHeaderRow) {
    const lastItineraryRow = reservationHeaderRow ? reservationHeaderRow - 2 : sheet.rowCount;
    for (let number = itineraryHeaderRow + 1; number <= lastItineraryRow; number += 1) {
      const row = sheet.getRow(number);
      const day = String(scalarCell(row.getCell(1)) || "").trim();
      const title = String(scalarCell(row.getCell(5)) || "").trim();
      if (!day || !title || day.startsWith("四、")) continue;
      const numericCost = Number(scalarCell(row.getCell(9)));
      itinerary.push({
        day,
        time: String(scalarCell(row.getCell(2)) || ""),
        endTime: String(scalarCell(row.getCell(3)) || ""),
        category: String(scalarCell(row.getCell(4)) || "安排"),
        title,
        subtitle: String(scalarCell(row.getCell(6)) || ""),
        address: String(scalarCell(row.getCell(7)) || "待补充"),
        transport: String(scalarCell(row.getCell(8)) || "待补充"),
        cost: Number.isFinite(numericCost) ? numericCost : 0,
        bookingStatus: String(scalarCell(row.getCell(10)) || "待确认"),
        sourceUrl: String(scalarCell(row.getCell(11)) || ""),
        note: String(scalarCell(row.getCell(12)) || ""),
        sourceId: "",
      });
    }
  }
  const reservations = [];
  if (reservationHeaderRow) {
    for (let number = reservationHeaderRow + 1; number <= sheet.rowCount; number += 1) {
      const row = sheet.getRow(number);
      const item = String(scalarCell(row.getCell(2)) || "").trim();
      if (!item) continue;
      reservations.push({
        id: String(scalarCell(row.getCell(1)) || `reserve-${randomUUID()}`),
        item,
        type: String(scalarCell(row.getCell(3)) || "待分类"),
        targetTime: String(scalarCell(row.getCell(4)) || "待确认"),
        status: String(scalarCell(row.getCell(5)) || "待确认"),
        owner: String(scalarCell(row.getCell(6)) || "待认领"),
        deadline: String(scalarCell(row.getCell(7)) || "待确认"),
        note: String(scalarCell(row.getCell(8)) || ""),
      });
    }
  }
  state.finalPlan = {
    ...existing,
    version: "excel-home-v1",
    title: String(field("方案标题", existing.title)),
    destination: String(field("目的地", existing.destination)),
    dates: String(field("出行日期", existing.dates)),
    schedule: String(field("行程结构", existing.schedule)),
    people: Number(field("同行人数", existing.people)) || existing.people,
    nights: Number(field("住宿晚数", existing.nights)) || existing.nights,
    perPersonBudget: field("人均预算", existing.perPersonBudget),
    summary: String(field("方案说明", existing.summary)),
    stay: {
      ...existing.stay,
      name: String(field("民宿名称", existing.stay.name)),
      address: String(field("详细地址", existing.stay.address)),
      capacity: String(field("适合人数", existing.stay.capacity)),
      roomsBeds: String(field("房间 / 床位", existing.stay.roomsBeds)),
      twoNightTotal: field("两晚总价", existing.stay.twoNightTotal),
      checkInOut: String(field("入住 / 退房", existing.stay.checkInOut)),
      barbecue: String(field("能否烧烤", existing.stay.barbecue)),
      bbqEquipment: String(field("烧烤设备 / 费用", existing.stay.bbqEquipment)),
      breakfast: String(field("早餐安排", existing.stay.breakfast)),
      sourceUrl: String(field("民宿原始链接", existing.stay.sourceUrl)),
    },
    itinerary: itineraryHeaderRow ? itinerary : existing.itinerary,
    reservations: reservationHeaderRow ? reservations : existing.reservations,
    updatedAt: new Date().toISOString(),
  };
  state.itinerary = splitFinalItinerary(state.finalPlan.itinerary);
  state.reservations = state.finalPlan.reservations.map((item) => ({ ...item }));
  state.project.name = state.finalPlan.title;
  state.project.destination = state.finalPlan.destination;
  state.project.people = state.finalPlan.people;
  state.tripProfile.dates = state.finalPlan.dates;
  state.tripProfile.schedule = state.finalPlan.schedule;
  state.tripProfile.groupSize = state.finalPlan.people;
  state.tripProfile.nights = state.finalPlan.nights;
  return true;
}

async function importFromExcel() {
  if (writeInProgress) return;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const state = await readState();
  const hasFinalPlanHome = importFinalPlanHome(workbook, state);
  const configs = [
    ["住宿候选", "住宿", "每晚价格", { address: "详细地址", capacity: "适合人数", rooms: "房间", beds: "床位", bathrooms: "卫浴", twoNightTotal: "两晚总价", environment: "环境特点", entireRental: "是否整租", kitchen: "厨房", barbecue: "能否烧烤", bbqEquipment: "烧烤设备/费用", breakfast: "早餐", parking: "停车", transport: "交通", checkIn: "入住时间", checkOut: "退房时间", bookingStatus: "预订状态" }],
    ["活动候选", null, "价格", { address: "详细地址", capacity: "适合人数", difficulty: "难度", horrorLevel: "恐怖程度", openingHours: "营业时间", reservation: "预约要求", bookingStatus: "预订状态" }],
    ["餐饮候选", "餐饮", "人均价格", { usage: "适合安排", address: "详细地址", openingHours: "营业时间", reservation: "预约要求", bookingStatus: "预订状态" }],
  ];
  for (const [sheetName, fixedCategory, priceColumn, detailColumns] of configs) {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) continue;
    const map = headerIndex(sheet);
    sheet.eachRow((row, number) => {
      if (number < 4) return;
      const id = String(value(row, map, "ID") || "").trim();
      if (!id) return;
      const index = state.places.findIndex((item) => item.id === id);
      if (index < 0) return;
      const current = state.places[index];
      const numericPrice = Number(value(row, map, priceColumn));
      const details = { ...defaultDetails, ...(current.details || {}) };
      for (const [key, columnName] of Object.entries(detailColumns)) {
        const edited = String(value(row, map, columnName) || "").trim();
        if (edited) details[key] = edited;
      }
      state.places[index] = {
        ...current,
        name: String(value(row, map, sheetName === "餐饮候选" ? "名称" : "名称") || current.name),
        category: fixedCategory || String(value(row, map, "活动类型") || current.category),
        area: String(value(row, map, "区域") || current.area),
        price: Number.isFinite(numericPrice) && numericPrice > 0 ? numericPrice : current.price,
        priceLabel: Number.isFinite(numericPrice) && numericPrice > 0 ? `参考 ¥${numericPrice}` : current.priceLabel,
        score: Number(value(row, map, "推荐分")) || current.score,
        selected: String(value(row, map, "是否入选")).trim() === "是",
        pros: String(value(row, map, "优点") || current.pros.join("；")).split(/[；;]/).filter(Boolean),
        cons: String(value(row, map, "缺点") || current.cons.join("；")).split(/[；;]/).filter(Boolean),
        dataStatus: String(value(row, map, "数据状态") || current.dataStatus),
        details,
      };
    });
  }

  const guideSheet = workbook.getWorksheet("攻略文章");
  if (guideSheet) {
    const map = headerIndex(guideSheet);
    guideSheet.eachRow((row, number) => {
      if (number < 4) return;
      const id = String(value(row, map, "ID") || "").trim();
      if (!id) return;
      const index = state.places.findIndex((item) => item.id === id);
      if (index < 0) return;
      state.places[index] = {
        ...state.places[index],
        selected: String(value(row, map, "是否入选")).trim() === "是",
      };
    });
  }

  const itinerarySheet = workbook.getWorksheet("两日行程");
  if (itinerarySheet && !hasFinalPlanHome) {
    const map = headerIndex(itinerarySheet);
    const imported = { day0: [], day1: [], day2: [] };
    itinerarySheet.eachRow((row, number) => {
      if (number < 4) return;
      const dayLabel = String(value(row, map, "日期") || "").trim();
      const title = String(value(row, map, "地点") || "").trim();
      if (!dayLabel || !title) return;
      const dayKey = dayLabel.includes("周五") ? "day0" : dayLabel.includes("周六") || dayLabel === "Day 1" ? "day1" : "day2";
      const numericCost = Number(value(row, map, "预计费用/人"));
      imported[dayKey].push({
        time: String(value(row, map, "开始时间") || ""),
        endTime: String(value(row, map, "结束时间") || ""),
        category: String(value(row, map, "类型") || "安排"),
        title,
        subtitle: String(value(row, map, "活动安排") || ""),
        address: String(value(row, map, "详细地址") || "待确认"),
        transport: String(value(row, map, "交通") || "待确认"),
        cost: Number.isFinite(numericCost) ? numericCost : 0,
        bookingStatus: String(value(row, map, "预订状态") || "待确认"),
        note: String(value(row, map, "注意事项") || ""),
        sourceId: String(value(row, map, "来源ID") || ""),
      });
    });
    for (const day of ["day0", "day1", "day2"]) if (imported[day].length) state.itinerary[day] = imported[day];
  }

  const reservationSheet = workbook.getWorksheet("预订清单");
  if (reservationSheet && !hasFinalPlanHome) {
    const map = headerIndex(reservationSheet);
    reservationSheet.eachRow((row, number) => {
      if (number < 4) return;
      const id = String(value(row, map, "ID") || "").trim();
      if (!id) return;
      const index = state.reservations.findIndex((item) => item.id === id);
      if (index < 0) return;
      state.reservations[index] = {
        ...state.reservations[index],
        item: String(value(row, map, "待办事项") || state.reservations[index].item),
        type: String(value(row, map, "类型") || state.reservations[index].type),
        targetTime: String(value(row, map, "计划时间") || state.reservations[index].targetTime),
        status: String(value(row, map, "当前状态") || state.reservations[index].status),
        owner: String(value(row, map, "负责人") || state.reservations[index].owner),
        deadline: String(value(row, map, "完成期限") || state.reservations[index].deadline),
        note: String(value(row, map, "核对说明") || state.reservations[index].note),
      };
    });
  }
  state.settings.lastExcelSync = new Date().toISOString();
  await writeState(state);
  await syncToExcel(state);
}

async function stateForClient() {
  const state = await readState();
  state.settings.provider = providerLabel();
  state.settings.workbookPath = path.relative(rootDir, workbookPath);
  return state;
}

async function proxyToUi(request, response, url) {
  const upstreamUrl = new URL(`${url.pathname}${url.search}`, "http://127.0.0.1:3000");
  const upstreamHeaders = new Headers();
  for (const [name, value] of Object.entries(request.headers)) {
    if (value == null || ["host", "connection", "content-length"].includes(name.toLowerCase())) continue;
    upstreamHeaders.set(name, Array.isArray(value) ? value.join(", ") : value);
  }
  upstreamHeaders.set("x-forwarded-host", request.headers.host || "localhost:8787");
  upstreamHeaders.set("x-forwarded-proto", request.headers["x-forwarded-proto"] || "http");

  const init = { method: request.method || "GET", headers: upstreamHeaders, redirect: "manual" };
  if (!['GET', 'HEAD'].includes(init.method)) {
    const chunks = [];
    for await (const chunk of request) chunks.push(chunk);
    init.body = Buffer.concat(chunks);
  }
  const upstream = await fetch(upstreamUrl, init);
  const bytes = Buffer.from(await upstream.arrayBuffer());
  const responseHeaders = {};
  for (const [name, value] of upstream.headers.entries()) {
    if (["connection", "content-encoding", "content-length", "transfer-encoding"].includes(name.toLowerCase())) continue;
    responseHeaders[name] = value;
  }
  responseHeaders["Content-Length"] = String(bytes.length);
  response.writeHead(upstream.status, responseHeaders);
  if (request.method === "HEAD") return response.end();
  return response.end(bytes);
}

const server = http.createServer(async (request, response) => {
  const url = new URL(request.url || "/", `http://${request.headers.host || "localhost"}`);
  if (request.method === "OPTIONS") return sendJson(response, 204, {});
  try {
    if (!(await authorizePublicRequest(request, response, url))) return;
    if (request.method === "GET" && url.pathname === "/api/health") {
      return sendJson(response, 200, { ok: true, provider: providerLabel(), workbook: path.relative(rootDir, workbookPath), now: new Date().toISOString() });
    }
    if (request.method === "GET" && url.pathname === "/api/state") {
      return sendJson(response, 200, await stateForClient());
    }
    if (request.method === "POST" && url.pathname === "/api/final-plan") {
      const body = await readBody(request);
      const state = await readState();
      const baseUpdatedAt = limitedText(body.baseUpdatedAt, "", 80);
      if (baseUpdatedAt && baseUpdatedAt !== state.finalPlan.updatedAt) {
        return sendJson(response, 409, {
          error: "这份方案刚刚被其他人或 Excel 更新过。请刷新后再编辑，避免覆盖最新内容。",
          state: await stateForClient(),
        });
      }
      state.finalPlan = sanitizeFinalPlan(body.finalPlan, state.finalPlan);
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, state: await stateForClient() });
    }
    const voteMatch = url.pathname.match(/^\/api\/places\/([^/]+)\/vote$/);
    if (request.method === "POST" && voteMatch) {
      const id = decodeURIComponent(voteMatch[1]);
      const body = await readBody(request);
      const nickname = limitedText(body.nickname, "", 20);
      const choice = limitedText(body.choice, "", 20);
      if (!nickname || nickname === "团队成员") return sendJson(response, 400, { error: "请先填写团队昵称" });
      if (!["想去", "可以", "不考虑"].includes(choice)) return sendJson(response, 400, { error: "请选择有效意见" });
      const state = await readState();
      const index = state.places.findIndex((item) => item.id === id);
      if (index < 0) return sendJson(response, 404, { error: "候选项目不存在" });
      const votes = { ...(state.places[index].votes || {}) };
      const removed = votes[nickname] === choice;
      if (removed) delete votes[nickname];
      else votes[nickname] = choice;
      state.places[index] = {
        ...state.places[index],
        votes,
      };
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, removed, state: await stateForClient() });
    }
    if (request.method === "GET" && url.pathname === "/api/download/excel") {
      const bytes = await fs.readFile(workbookPath);
      response.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent("扬州团队旅行攻略.xlsx")}`,
        "Access-Control-Allow-Origin": "*",
        "Content-Length": bytes.length,
      });
      return response.end(bytes);
    }
    if (request.method === "POST" && url.pathname === "/api/links") {
      const body = await readBody(request);
      const incoming = Array.isArray(body.urls) ? body.urls : [];
      if (!incoming.length) return sendJson(response, 400, { error: "请提交至少一个链接" });
      const state = await readState();
      const existing = new Set(state.links.map((item) => item.url));
      const created = [];
      for (const raw of incoming.slice(0, 30)) {
        let normalized;
        try { normalized = cleanUrl(raw); } catch { continue; }
        if (existing.has(normalized)) continue;
        existing.add(normalized);
        const now = new Date().toISOString();
        const id = `link-${createHash("sha1").update(`${normalized}-${now}`).digest("hex").slice(0, 12)}`;
        state.links.unshift({
          id,
          url: normalized,
          title: "",
          category: body.category || "自动识别",
          status: "等待处理",
          readStatus: "等待读取",
          organizedStatus: "整理中",
          factsFound: [],
          missingFields: [],
          resultNote: "已进入后台处理队列",
          submitter: String(body.submitter || "团队成员").slice(0, 40),
          note: String(body.note || "").slice(0, 300),
          createdAt: now,
          updatedAt: now,
          summary: "",
          model: "",
          error: "",
        });
        created.push(id);
      }
      await writeState(state);
      created.forEach((id, index) => setTimeout(() => enqueueLink(id), 250 + index * 80));
      return sendJson(response, 202, { created: created.length, duplicates: incoming.length - created.length, ids: created });
    }
    const retryMatch = url.pathname.match(/^\/api\/links\/([^/]+)\/retry$/);
    if (request.method === "POST" && retryMatch) {
      const id = decodeURIComponent(retryMatch[1]);
      await updateLink(id, { status: "等待处理", readStatus: "等待读取", organizedStatus: "整理中", resultNote: "已重新加入后台处理队列", error: "" });
      setTimeout(() => enqueueLink(id), 200);
      return sendJson(response, 202, { ok: true });
    }
    const deleteMatch = url.pathname.match(/^\/api\/links\/([^/]+)$/);
    if (request.method === "DELETE" && deleteMatch) {
      const id = decodeURIComponent(deleteMatch[1]);
      const state = await readState();
      state.links = state.links.filter((item) => item.id !== id);
      state.places = state.places.filter((item) => item.sourceId !== id);
      await writeState(state);
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true });
    }
    if (request.method === "POST" && url.pathname === "/api/sync/from-excel") {
      await importFromExcel();
      return sendJson(response, 200, { ok: true, state: await stateForClient() });
    }
    if (request.method === "POST" && url.pathname === "/api/sync/to-excel") {
      const state = await readState();
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, state: await stateForClient() });
    }
    if (!url.pathname.startsWith("/api/")) return proxyToUi(request, response, url);
    return sendJson(response, 404, { error: "接口不存在" });
  } catch (error) {
    return sendJson(response, 500, { error: error instanceof Error ? error.message : "服务器错误" });
  }
});

server.listen(port, host, () => {
  console.log(`Yangzhou Trip API: http://localhost:${port}`);
});

try { lastKnownWorkbookMtime = (await fs.stat(workbookPath)).mtimeMs; } catch { lastKnownWorkbookMtime = 0; }
setInterval(async () => {
  if (writeInProgress) return;
  try {
    const mtime = (await fs.stat(workbookPath)).mtimeMs;
    if (lastKnownWorkbookMtime && mtime > lastKnownWorkbookMtime + 100) await importFromExcel();
    else if (!lastKnownWorkbookMtime) lastKnownWorkbookMtime = mtime;
  } catch { /* The workbook is generated during setup. */ }
}, 4_000).unref();

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => server.close(() => process.exit(0)));
}
