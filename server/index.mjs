import http from "node:http";
import { createHash, randomUUID, timingSafeEqual } from "node:crypto";
import { lookup as dnsLookup } from "node:dns/promises";
import fs from "node:fs/promises";
import { isIP } from "node:net";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { brotliCompressSync, constants as zlibConstants, gzipSync } from "node:zlib";
import ExcelJS from "exceljs";

const rootDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

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

function projectPath(value, fallback) {
  const resolved = path.resolve(rootDir, value || fallback);
  if (resolved !== rootDir && !resolved.startsWith(`${rootDir}${path.sep}`)) {
    throw new Error("旅行项目文件必须位于当前仓库内");
  }
  return resolved;
}

const dataFile = projectPath(process.env.TRIP_DATA_FILE, "data/store.json");
const outputDir = projectPath(process.env.TRIP_OUTPUT_DIR, "outputs/019f7eda-a998-7660-a73d-e43b3af67965");
const requestedWorkbookName = path.basename(process.env.WORKBOOK_FILE_NAME || "出行共创项目.xlsx");
const workbookFileName = requestedWorkbookName.toLowerCase().endsWith(".xlsx") ? requestedWorkbookName : `${requestedWorkbookName}.xlsx`;
const workbookPath = path.join(outputDir, workbookFileName);
const templateDataFile = path.join(rootDir, "data", "trip-template.json");
const localPort = Number(process.env.API_PORT || 8787);
const publicPort = Number(process.env.PUBLIC_API_PORT || localPort + 1);
const uiPort = Number(process.env.UI_PORT || 3100);
const localHost = process.env.API_HOST || "127.0.0.1";
const publicHost = process.env.PUBLIC_API_HOST || "127.0.0.1";
const accessModeSymbol = Symbol("travel-access-mode");
const responseRequestSymbol = Symbol("travel-response-request");
let writeInProgress = false;
let lastKnownWorkbookMtime = 0;
let processingChain = Promise.resolve();
let mutationChain = Promise.resolve();
const budgetAdviceCache = new Map();
const budgetAdvicePending = new Map();
const uiAssetCache = new Map();

await fs.mkdir(path.dirname(dataFile), { recursive: true });
await fs.mkdir(outputDir, { recursive: true });
try {
  await fs.access(dataFile);
} catch {
  await fs.copyFile(templateDataFile, dataFile);
}

const defaultDetails = {
  address: "待核实",
  locationHighlights: "待核实",
  distanceToCore: "待核实",
  capacity: "待核实",
  roomType: "待核实",
  rooms: "待核实",
  beds: "待核实",
  bedTypes: "待核实",
  bathrooms: "待核实",
  twoNightTotal: "待核实",
  extraFees: "待核实",
  deposit: "待核实",
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
  cancellationPolicy: "待核实",
  evidence: "待核实",
  cuisineType: "待核实",
  signatureDishes: "待核实",
  sixPersonTotal: "待核实",
  privateRoom: "待核实",
  queueInfo: "待核实",
  groupSuitability: "待核实",
  themeName: "待核实",
  escapeStyle: "待核实",
  venueSize: "待核实",
  roomCount: "待核实",
  minPlayers: "待核实",
  maxPlayers: "待核实",
  sixPersonSession: "待核实",
  npcInteraction: "待核实",
  physicalIntensity: "待核实",
  costume: "待核实",
  leisureFacilities: "待核实",
  packageInfo: "待核实",
  overnight: "待核实",
  includedMeals: "待核实",
  restArea: "待核实",
  genderArrangement: "待核实",
  serviceRestrictions: "待核实",
  attractionType: "待核实",
  indoorOutdoor: "待核实",
  weatherImpact: "待核实",
  ticketInfo: "待核实",
  recommendedDuration: "待核实",
};

const supportedCategories = ["住宿", "餐饮", "密室", "休闲娱乐", "景点", "攻略"];
const subCategoryOptions = {
  住宿: ["民宿", "酒店", "公寓", "客栈", "度假村", "其他住宿"],
  餐饮: ["烧烤", "火锅", "早茶早餐", "炒菜正餐", "自助餐", "夜宵", "甜品饮品", "咖啡", "酒吧", "其他餐饮"],
  密室: ["推理密室", "机关密室", "剧情密室", "恐怖密室", "未分类密室"],
  休闲娱乐: ["汗蒸", "桑拿", "洗浴中心", "温泉", "SPA", "足疗按摩", "KTV", "桌游", "电竞", "电玩城", "沉浸式剧场", "其他休闲"],
  景点: ["博物馆", "园林景区", "历史街区", "公园", "古镇", "露营", "户外运动", "演出展览", "其他景点"],
  攻略: ["美食攻略", "住宿攻略", "行程攻略", "综合攻略"],
};

function validSubCategory(category, subCategory) {
  return (subCategoryOptions[category] || []).includes(String(subCategory || ""));
}

function normalizeCategory(value, context = "") {
  const explicit = String(value || "");
  if (supportedCategories.includes(explicit)) return explicit;
  const text = `${value || ""} ${context || ""}`;
  if (/民宿|酒店|公寓|客栈|住宿|度假村/i.test(text)) return "住宿";
  if (/密室|剧本杀|逃脱/i.test(text)) return "密室";
  if (/餐饮|餐厅|美食|烧烤|烤肉|火锅|早茶|早餐|炒菜|淮扬菜|夜宵/i.test(text)) return "餐饮";
  if (/汗蒸|桑拿|洗浴|温泉|SPA|足疗|按摩|KTV|桌游|电竞|电玩城|休闲娱乐|室内休闲/i.test(text)) return "休闲娱乐";
  if (/景点|景区|园林|博物馆|历史街区|公园|古镇|露营|户外|Citywalk|演出|展览/i.test(text)) return "景点";
  if (/攻略|文章|游记/i.test(text)) return "攻略";
  if (String(value || "") === "活动") return "休闲娱乐";
  return "攻略";
}

function inferSubCategory(text, category) {
  const source = String(text || "");
  const rules = {
    住宿: [[/民宿/i, "民宿"], [/酒店/i, "酒店"], [/公寓/i, "公寓"], [/客栈/i, "客栈"], [/度假村/i, "度假村"]],
    餐饮: [[/烧烤|烤串|烤肉/i, "烧烤"], [/火锅|涮肉|牛肉锅/i, "火锅"], [/早茶|早餐|包子|面馆/i, "早茶早餐"], [/炒菜|淮扬菜|家常菜|川菜|湘菜|私房菜/i, "炒菜正餐"], [/自助餐/i, "自助餐"], [/夜宵/i, "夜宵"], [/甜品|奶茶|饮品/i, "甜品饮品"], [/咖啡/i, "咖啡"], [/酒吧/i, "酒吧"]],
    密室: [[/推理|硬核/i, "推理密室"], [/机关|解谜/i, "机关密室"], [/剧情|沉浸/i, "剧情密室"], [/恐怖|追逐|微恐|中恐|重恐/i, "恐怖密室"]],
    休闲娱乐: [[/汗蒸/i, "汗蒸"], [/桑拿/i, "桑拿"], [/洗浴/i, "洗浴中心"], [/温泉/i, "温泉"], [/SPA/i, "SPA"], [/足疗|按摩/i, "足疗按摩"], [/KTV/i, "KTV"], [/桌游/i, "桌游"], [/电竞/i, "电竞"], [/电玩城/i, "电玩城"], [/沉浸式剧场/i, "沉浸式剧场"]],
    景点: [[/博物馆/i, "博物馆"], [/园林|景区/i, "园林景区"], [/历史街区|老街|古街|Citywalk/i, "历史街区"], [/公园/i, "公园"], [/古镇/i, "古镇"], [/露营/i, "露营"], [/户外运动/i, "户外运动"], [/演出|展览/i, "演出展览"]],
    攻略: [[/美食/i, "美食攻略"], [/住宿|民宿|酒店/i, "住宿攻略"], [/两日|周末|路线|行程/i, "行程攻略"]],
  };
  const fallback = { 住宿: "其他住宿", 餐饮: "其他餐饮", 密室: "未分类密室", 休闲娱乐: "其他休闲", 景点: "其他景点", 攻略: "综合攻略" }[category] || "待分类";
  return (rules[category] || []).find(([pattern]) => pattern.test(source))?.[1] || fallback;
}

function inferFeatureTags(text, category) {
  const source = String(text || "");
  const groupSize = source.match(/(?:适合|容纳|支持|我们|团队)?\s*(\d{1,2})\s*人/i)?.[1];
  const rules = [
    [/烧烤|烤串|烤肉/i, "烧烤"], [/火锅|涮肉/i, "火锅"], [/早茶|早餐/i, "早茶早餐"], [/包间/i, "有包间"],
    [/微恐/i, "微恐"], [/中恐/i, "中恐"], [/重恐/i, "重恐"], [/无恐/i, "无恐"], [/NPC|真人互动/i, "真人互动"],
    [/汗蒸/i, "汗蒸"], [/桑拿/i, "桑拿"], [/洗浴/i, "洗浴"], [/温泉/i, "温泉"], [/过夜|24小时/i, "可过夜"],
    [/室内/i, "室内"], [/室外|户外/i, "室外"], [/预约/i, "需要预约"], [/停车/i, "可停车"],
  ];
  return [...new Set([category, ...rules.filter(([pattern]) => pattern.test(source)).map(([, label]) => label), ...(groupSize ? [`适合${groupSize}人`] : [])])].slice(0, 12);
}

const defaultTripProfile = {
  planVersion: "three-day-v1",
  dates: "待团队确认",
  schedule: "3 天 2 晚 · 国际航班直达济州",
  groupSize: 6,
  nights: 2,
  stayPreference: "济州市区交通方便、环境舒适，适合 6 人入住",
  barbecue: "待团队确认",
  breakfasts: ["第 2 天早餐待选", "第 3 天早餐待选"],
  activity: "东线自然景观、海岸步道与济州美食",
  accommodationBudget: "待团队确认",
};

const defaultReservations = [
  { id: "reserve-dates-flight", item: "确定日期与国际直达济州往返航班", type: "机票", targetTime: "第 1 天抵达 · 第 3 天返程", status: "待确认", owner: "待认领", deadline: "优先完成", note: "不要购买经首尔转韩国国内线的方案；六人的去回航班与行李额需统一" },
  { id: "reserve-entry", item: "逐人核对济州免签、K-ETA 与入境材料", type: "入境", targetTime: "订票前与出发前各核对一次", status: "待确认", owner: "待认领", deadline: "订票前", note: "向航司、韩国出入境 1345 或官方渠道确认最新要求；入境由边检最终判断" },
  { id: "reserve-stay", item: "6 人济州市区住宿（2 晚）", type: "住宿", targetTime: "第 1 天入住 · 第 3 天退房", status: "待补充真实链接", owner: "待认领", deadline: "日期确认后立即", note: "比较房型、床型、含税总价、厨房、位置和取消政策" },
  { id: "reserve-charter", item: "第 2 天东线公交 + 必要时短途拼车", type: "交通", targetTime: "第 2 天约 08:00–20:00", status: "待确认公交班次", owner: "待认领", deadline: "出发前 14 天", note: "先查 111/112 急行或 201 干线；接驳不便时再由六人分摊短途车费" },
  { id: "reserve-food", item: "两顿特色正餐 + 民宿做饭 / 宵夜", type: "餐饮", targetTime: "第 1/2 天", status: "待补充真实链接", owner: "待认领", deadline: "住宿和公交路线确定后", note: "特色正餐按人均约 ¥100，早餐约 ¥20–30；其余可买菜做饭或吃简餐宵夜" },
  { id: "reserve-attractions", item: "景点开放、天气与牛岛分支", type: "景点", targetTime: "第 2 天", status: "待确认", owner: "待认领", deadline: "出发前 3 天与当天早晨", note: "优先免费海岸、市场和步道；贵景点不强制保留" },
];

const defaultDay1Plan = [
  { time: "14:00", endTime: "16:00", title: "国际航班直达济州 · 入境集合", subtitle: "六人按最终航班抵达济州国际机场，完成入境后前往济州市区住宿。", transport: "机场公交优先；晚到时分乘两辆出租车并由六人分摊", cost: 15, category: "抵达", note: "示意时段，必须按真实直达航班调整", address: "Jeju International Airport", bookingStatus: "待确认", sourceId: "guide-entry" },
  { time: "16:30", endTime: "17:30", title: "入住济州市区两晚", subtitle: "优先选择机场和市区餐饮交通方便、适合 6 人连住且能简单做饭的酒店或公寓。", transport: "公交 / 步行优先", cost: 0, category: "住宿", note: "房型、含税总价与取消政策待确认", address: "济州市区待选", bookingStatus: "未预订", sourceId: "requirement-stay" },
  { time: "18:00", endTime: "20:00", title: "东门传统市场 · 落地轻松逛吃", subtitle: "抵达日只安排市区活动，逛市场、吃小食并采购饮水零食。", transport: "公交 / 步行", cost: 30, category: "市场", note: "若航班晚到则直接取消", address: "济州市东门传统市场", bookingStatus: "无需预约", sourceId: "place-dongmun" },
  { time: "20:00", endTime: "21:30", title: "黑猪烤肉晚餐", subtitle: "按人均约 ¥100 控制点单，也可以买菜回民宿做饭。", transport: "步行优先", cost: 100, category: "晚餐", note: "具体餐厅待团队投递", address: "济州市区待选", bookingStatus: "未预订", sourceId: "" },
];

const defaultDay2Plan = [
  { time: "08:00", endTime: "09:30", title: "济州市区出发 · 东线公交", subtitle: "优先乘急行或干线公交前往城山，接驳不便时再短途拼车。", transport: "111/112 急行或 201 干线", cost: 35, category: "交通", note: "出发前核对当天班次和天气", address: "济州市区 → 城山邑", bookingStatus: "待确认", sourceId: "guide-transport" },
  { time: "10:00", endTime: "12:00", title: "城山日出峰", subtitle: "选择免费海岸段或登顶路线，看火山地貌与海景。", transport: "公交到城山后步行", cost: 25, category: "自然景观", note: "贵或体力不合适就走免费段", address: "Seongsan Ilchulbong", bookingStatus: "出发前复核", sourceId: "place-seongsan" },
  { time: "12:15", endTime: "13:30", title: "城山东线海鲜午餐", subtitle: "带鱼、海鲜或鲍鱼粥三选一，按人均约 ¥100 找公交站附近餐厅。", transport: "步行 / 短途公交", cost: 100, category: "午餐", note: "具体餐厅待团队投递", address: "城山邑待选", bookingStatus: "未预订", sourceId: "" },
  { time: "14:00", endTime: "16:30", title: "涉地可支", subtitle: "沿东部海岸散步；如改去牛岛，当天不再硬塞这个景点。", transport: "公交 + 步行；必要时短途拼车", cost: 10, category: "海岸散步", note: "风大或雨天缩短", address: "Seopjikoji", bookingStatus: "出发前复核", sourceId: "place-seopjikoji" },
  { time: "17:00", endTime: "20:00", title: "返回济州市区 · 自由晚餐", subtitle: "晚餐可买菜回民宿做饭，或六人一起吃简餐宵夜。", transport: "公交返回", cost: 70, category: "返程与晚餐", note: "不按正式聚餐上限计算", address: "济州市区待选", bookingStatus: "待确认", sourceId: "" },
];

const defaultDay3Plan = [
  { time: "08:30", endTime: "09:30", title: "早餐 · 退房 · 行李安排", subtitle: "民宿简单做早餐或选附近小店，早餐后完成退房。", transport: "步行", cost: 25, category: "早餐", note: "住宿与航班确定后再选", address: "住宿附近待选", bookingStatus: "待确认", sourceId: "" },
  { time: "10:00", endTime: "13:00", title: "涯月汉潭海岸散步路", subtitle: "海岸散步加咖啡，作为返程日前半天的轻量路线。", transport: "公交优先", cost: 20, category: "海岸与咖啡", note: "按返程航班和天气缩短或取消", address: "Aewol-eup, Jeju-si, Jeju", bookingStatus: "无需预约", sourceId: "place-aewol" },
  { time: "13:00", endTime: "14:00", title: "涯月简餐或咖啡店午餐", subtitle: "不安排正式大餐，确保留足回机场和取行李时间。", transport: "步行", cost: 60, category: "午餐", note: "具体店铺待团队投递", address: "涯月邑待选", bookingStatus: "待确认", sourceId: "" },
  { time: "14:30", endTime: "17:00", title: "前往济州国际机场 · 国际航班返程", subtitle: "按航司要求预留值机、行李和出境时间。", transport: "机场公交优先；时间紧时分乘两辆出租车", cost: 15, category: "返程", note: "必须按真实航班倒推", address: "Jeju International Airport", bookingStatus: "待确认", sourceId: "guide-entry" },
];

function itineraryWithDay(itinerary) {
  return [
    ...(itinerary.day0 || []).map((item) => ({ day: "第1天", sourceUrl: "", ...item })),
    ...(itinerary.day1 || []).map((item) => ({ day: "第2天", sourceUrl: "", ...item })),
    ...(itinerary.day2 || []).map((item) => ({ day: "第3天", sourceUrl: "", ...item })),
  ];
}

function splitFinalItinerary(items) {
  const result = { day0: [], day1: [], day2: [] };
  for (const item of Array.isArray(items) ? items : []) {
    const label = String(item.day || "");
    const key = /第\s*1\s*天|周五|Day\s*1/i.test(label) ? "day0" : /第\s*2\s*天|周六|Day\s*2/i.test(label) ? "day1" : "day2";
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
  const day = limitedText(value, "第2天", 20);
  if (/第\s*1\s*天|周五|星期五|Day\s*1/i.test(day)) return "第1天";
  if (/第\s*3\s*天|周日|星期日|周天|Day\s*3/i.test(day)) return "第3天";
  return "第2天";
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
    title: submittedText(source, "title", previous.title || "我的出行共创项目", 120, false),
    destination: submittedText(source, "destination", previous.destination || "待确定目的地", 80, false),
    dates: submittedText(source, "dates", previous.dates || "待团队确认", 120),
    schedule: submittedText(source, "schedule", previous.schedule || "3 天 2 晚 · 国际航班直达济州", 160),
    people: boundedInteger(source.people, boundedInteger(previous.people, 4, 1, 50), 1, 50),
    nights: boundedInteger(source.nights, boundedInteger(previous.nights, 2, 1, 30), 1, 30),
    perPersonBudget: submittedText(source, "perPersonBudget", previous.perPersonBudget || "待团队确认", 80),
    roundTripFlightPerPerson: submittedText(source, "roundTripFlightPerPerson", previous.roundTripFlightPerPerson || "待填写实际含税票价（含托运行李）", 120),
    summary: submittedText(source, "summary", previous.summary || "", 1_000),
    stay: {
      name: submittedText(stay, "name", previousStay.name || "待选择真实民宿", 160),
      address: submittedText(stay, "address", previousStay.address || "待补充民宿详细地址", 300),
      capacity: submittedText(stay, "capacity", previousStay.capacity || `目标 ${boundedInteger(source.people, boundedInteger(previous.people, 4, 1, 50), 1, 50)} 人`, 160),
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

const candidateTypes = ["place", "requirement", "guide"];

function normalizeCandidateType(value, { sourceType = "链接", category = "攻略", dataStatus = "" } = {}) {
  if (sourceType === "文字" || /团队文字需求|团队诉求/.test(String(dataStatus || ""))) return "requirement";
  const explicit = String(value || "").trim().toLowerCase();
  if (candidateTypes.includes(explicit)) return explicit;
  if (category === "攻略") return "guide";
  return "place";
}

function candidateTypeLabel(value) {
  return { place: "真实候选", requirement: "团队需求", guide: "攻略文章" }[value] || "真实候选";
}

function normalizeVerificationStatus(value, candidateType, dataStatus = "") {
  if (candidateType === "requirement") return "团队诉求";
  const explicit = String(value || "").trim();
  if (["待核实", "部分核实", "已核实"].includes(explicit)) return explicit;
  const context = String(dataStatus || "");
  if (/已核实|官方已确认/.test(context)) return "已核实";
  if (/partial|部分/.test(context)) return "部分核实";
  return "待核实";
}

function normalizeState(raw) {
  const state = raw && typeof raw === "object" ? raw : {};
  const hasWeekendPlan = state.tripProfile?.planVersion === defaultTripProfile.planVersion;
  state.project = { name: "济州岛三日旅行共创", destination: "韩国济州岛", days: 3, people: 6, budget: 0, status: "方案共创中", tagline: "六个人一起把济州岛的链接和想法整理成可执行计划。", ...(state.project || {}) };
  state.project.people = boundedInteger(state.project.people, 6, 1, 50);
  state.project.days = boundedInteger(state.project.days, 3, 1, 30);
  state.tripProfile = { ...defaultTripProfile, ...(state.tripProfile || {}) };
  state.tripProfile.breakfasts = Array.isArray(state.tripProfile.breakfasts) ? state.tripProfile.breakfasts : defaultTripProfile.breakfasts;
  state.links = Array.isArray(state.links) ? state.links.map((item) => {
    const sourceType = item.sourceType === "文字" || (!item.url && item.inputText) ? "文字" : "链接";
    const completed = item.status === "已写入Excel";
    const processing = ["等待处理", "正在读取", "AI分析中"].includes(item.status);
    const blocked = item.status === "需要人工补充";
    const inputText = String(item.inputText || "").slice(0, 4_000);
    const context = `${item.title || ""} ${item.summary || ""} ${item.note || ""} ${inputText}`;
    const guideLike = /攻略|游记|指南|百科/i.test(context) || /wikipedia\.org|example\.com\/.*guide/i.test(item.url || "");
    const category = item.category === "自动识别" ? "自动识别" : normalizeCategory(guideLike ? "攻略" : item.category, context);
    return {
      ...item,
      sourceType,
      inputText,
      url: String(item.url || ""),
      category,
      subCategory: category === "自动识别" ? "等待识别" : /wikipedia\.org/i.test(item.url || "") ? "综合攻略" : validSubCategory(category, item.subCategory) ? item.subCategory : inferSubCategory(context, category),
      readStatus: item.readStatus || (sourceType === "文字" ? "文字已接收" : completed ? "成功读取" : blocked ? "读取受限" : item.status === "处理失败" ? "读取失败" : "等待读取"),
      organizedStatus: item.organizedStatus || (completed ? "已整理" : processing ? "整理中" : "未整理"),
      factsFound: Array.isArray(item.factsFound) ? item.factsFound : (completed ? (sourceType === "文字" ? ["团队诉求", "内容分类", "关键偏好"] : ["页面标题", "内容分类", "核心摘要"]) : []),
      missingFields: Array.isArray(item.missingFields) ? item.missingFields : (completed ? ["价格或预约等动态信息仍需核实"] : [sourceType === "文字" ? "具体商户与动态信息" : "网页正文未完整读取"]),
      resultNote: item.resultNote || (completed ? "已生成候选资料并写入 Excel" : item.error || "等待后台处理"),
    };
  }) : [];
  const seedDetails = {};
  state.places = Array.isArray(state.places) ? state.places.map((item) => {
    const sourceLink = state.links.find((link) => link.id === item.sourceId);
    const context = `${item.name || ""} ${(item.tags || []).join(" ")} ${item.details?.usage || ""}`;
    const guideLike = item.category === "攻略" || /攻略|游记|指南|百科/i.test(context);
    const category = normalizeCategory(guideLike ? "攻略" : item.category, context);
    const placeholderRequirement = !item.sourceUrl && /候选/.test(String(item.name || "")) && /示例|空位|等待|待投递/.test(String(item.dataStatus || ""));
    const candidateType = placeholderRequirement ? "requirement" : normalizeCandidateType(item.candidateType, {
      sourceType: sourceLink?.sourceType,
      category,
      dataStatus: item.dataStatus,
    });
    const details = { ...defaultDetails, ...(seedDetails[item.id] || {}), ...(item.details || {}), address: item.details?.address || seedDetails[item.id]?.address || item.area || "待核实" };
    const subCategory = validSubCategory(category, item.subCategory) ? item.subCategory : inferSubCategory(context, category);
    const featureTags = Array.isArray(item.featureTags) && item.featureTags.includes(category) ? item.featureTags : inferFeatureTags(`${context} ${subCategory}`, category);
    const aiPriceBasis = limitedText(item.aiPriceBasis, "", 500);
    const aiScoreReason = limitedText(item.aiScoreReason, "", 500);
    const planningSuggestions = stringList(item.planningSuggestions).slice(0, 3);
    const candidateLeads = normalizeCandidateLeads(item.candidateLeads);
    const hasAiMetadata = Boolean(aiPriceBasis || aiScoreReason || planningSuggestions.length || candidateLeads.length);
    const aiCompleteness = normalizedAiCompleteness(item.aiCompleteness, item.missingFields);
    return {
      ...item,
      category,
      candidateType,
      verificationStatus: normalizeVerificationStatus(item.verificationStatus, candidateType, item.dataStatus),
      subCategory,
      featureTags,
      tags: Array.isArray(item.tags) ? item.tags : [],
      pros: Array.isArray(item.pros) ? item.pros : [],
      cons: Array.isArray(item.cons) ? item.cons : [],
      votes: item.votes && typeof item.votes === "object" && !Array.isArray(item.votes) ? item.votes : {},
      decisionStatus: item.decisionStatus || (item.selected ? "拟定" : "待比较"),
      manualNote: String(item.manualNote || ""),
      manualOverrides: item.manualOverrides && typeof item.manualOverrides === "object" && !Array.isArray(item.manualOverrides) ? item.manualOverrides : {},
      aiCompleteness: aiCompleteness === 0 && !hasAiMetadata ? undefined : aiCompleteness ?? undefined,
      aiPriceBasis: aiPriceBasis || undefined,
      aiPriceConfidence: aiPriceBasis || /^AI估价：/.test(String(item.priceLabel || "")) ? normalizeAiConfidence(item.aiPriceConfidence) : undefined,
      aiScoreReason: aiScoreReason || undefined,
      planningSuggestions: planningSuggestions.length ? planningSuggestions : undefined,
      candidateLeads: candidateLeads.length ? candidateLeads : undefined,
      details,
    };
  }) : [];
  state.links = state.links.map((item) => ({
    ...item,
    candidateCount: state.places.filter((place) => place.sourceId === item.id).length,
  }));
  state.itinerary = state.itinerary || {};
  if (!hasWeekendPlan) {
    state.itinerary.day0 = defaultDay1Plan;
    state.itinerary.day1 = defaultDay2Plan;
    state.itinerary.day2 = defaultDay3Plan;
  }
  state.itinerary.day0 = Array.isArray(state.itinerary.day0) && state.itinerary.day0.length ? state.itinerary.day0 : defaultDay1Plan;
  state.itinerary.day1 = Array.isArray(state.itinerary.day1) && state.itinerary.day1.length ? state.itinerary.day1 : defaultDay2Plan;
  state.itinerary.day2 = Array.isArray(state.itinerary.day2) && state.itinerary.day2.length ? state.itinerary.day2 : defaultDay3Plan;
  for (const day of ["day0", "day1", "day2"]) {
    state.itinerary[day] = (Array.isArray(state.itinerary[day]) ? state.itinerary[day] : []).map((item) => ({ address: "待确认", bookingStatus: "待确认", sourceId: "", ...item }));
  }
  state.reservations = Array.isArray(state.reservations) && state.reservations.length ? state.reservations : defaultReservations;
  const defaultStay = state.places.find((item) => item.id === "requirement-stay") || state.places.find((item) => item.category === "住宿") || {};
  const defaultFinalPlan = {
    version: "excel-home-v1",
    title: `${state.project.people} 人${state.project.destination}出行共创`,
    destination: state.project.destination,
    dates: state.tripProfile.dates,
    schedule: state.tripProfile.schedule,
    people: state.tripProfile.groupSize,
    nights: state.tripProfile.nights,
    perPersonBudget: "¥6,000 / 人（包含往返济州机票）",
    roundTripFlightPerPerson: "待填写实际含税票价（含托运行李）",
    summary: "六人总预算约 ¥36,000（¥6,000/人，包含往返济州机票）；机票实际价格确认后，再由 AI 分配住宿、餐饮、岛内交通、游玩和机动金。餐饮仍按正餐约 ¥100/人并可穿插民宿做饭，岛内交通公交优先、必要时短途拼车。",
    stay: {
      name: "待选择真实民宿",
      address: "待补充民宿详细地址",
      capacity: `目标 ${state.project.people} 人，待房源确认`,
      roomsBeds: "优先整租 2–3 间卧室，并确认厨房可用",
      twoNightTotal: "六人两晚约 ¥1,800–3,000（以日期和真实房源为准）",
      checkInOut: "第 1 天入住 · 第 3 天退房",
      barbecue: "非本次硬性条件，按团队投递再确认",
      bbqEquipment: "如选择可烧烤住宿，再核对设备、食材和清洁费用",
      breakfast: "两顿早餐具体门店待团队选择",
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
  state.aiReviews = enforceAiReviewSafety(state, normalizeAiReviews(state.aiReviews));
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

function preferredContentEncoding(request, contentType, size) {
  if (!request || request.method === "HEAD" || size < 1_024 || !/(?:json|javascript|css|html|svg|text)/i.test(contentType || "")) return "";
  const accepted = String(request.headers["accept-encoding"] || "").toLowerCase();
  if (/\bbr\b/.test(accepted)) return "br";
  if (/\bgzip\b/.test(accepted)) return "gzip";
  return "";
}

function sendBuffer(response, status, headers, input, { request = response[responseRequestSymbol], allowCompression = true } = {}) {
  const raw = Buffer.isBuffer(input) ? input : Buffer.from(input || "");
  const responseHeaders = { ...headers };
  const contentType = String(responseHeaders["Content-Type"] || responseHeaders["content-type"] || "");
  const encoding = allowCompression ? preferredContentEncoding(request, contentType, raw.length) : "";
  let bytes = raw;
  if (encoding === "br") {
    bytes = brotliCompressSync(raw, { params: { [zlibConstants.BROTLI_PARAM_QUALITY]: 4 } });
    responseHeaders["Content-Encoding"] = "br";
  } else if (encoding === "gzip") {
    bytes = gzipSync(raw, { level: 6 });
    responseHeaders["Content-Encoding"] = "gzip";
  }
  if (encoding) responseHeaders.Vary = [responseHeaders.Vary, "Accept-Encoding"].filter(Boolean).join(", ");
  responseHeaders["Content-Length"] = String(request?.method === "HEAD" || status === 204 || status === 304 ? 0 : bytes.length);
  response.writeHead(status, responseHeaders);
  if (request?.method === "HEAD" || status === 204 || status === 304) return response.end();
  return response.end(bytes);
}

function sendJson(response, status, payload) {
  return sendBuffer(response, status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, If-None-Match",
    "Access-Control-Allow-Methods": "GET,POST,PATCH,DELETE,OPTIONS",
    "Cache-Control": "no-store",
  }, Buffer.from(JSON.stringify(payload)));
}

function requestHostname(request) {
  const raw = String(request?.headers?.host || "").trim().toLowerCase();
  if (!raw) return "";
  if (raw.startsWith("[")) return raw.slice(1, raw.indexOf("]") > 0 ? raw.indexOf("]") : undefined);
  if (raw === "::1") return raw;
  return raw.split(":")[0].replace(/\.$/, "");
}

function requestAccessMode(request) {
  return request?.[accessModeSymbol] || request?.accessMode || "public";
}

function isLocalManagerRequest(request) {
  return requestAccessMode(request) === "local";
}

function isAllowedBrowserOrigin(request) {
  const rawOrigin = String(request?.headers?.origin || "").trim();
  if (!rawOrigin) return true;
  let origin;
  try {
    origin = new URL(rawOrigin);
  } catch {
    return false;
  }
  if (requestAccessMode(request) === "local") {
    return ["localhost", "127.0.0.1", "::1"].includes(origin.hostname)
      && ["", String(uiPort), String(localPort)].includes(origin.port);
  }
  return origin.protocol === "https:"
    && origin.hostname.endsWith(".ts.net")
    && origin.hostname === requestHostname(request);
}

function requireManagement(request, response) {
  if (isLocalManagerRequest(request) && isAllowedBrowserOrigin(request)) return true;
  sendJson(response, 403, { error: "此操作仅能在发起人的本机管理界面完成" });
  return false;
}

function safeCredentialMatch(actual, expected) {
  const actualBytes = Buffer.from(actual);
  const expectedBytes = Buffer.from(expected);
  return actualBytes.length === expectedBytes.length && timingSafeEqual(actualBytes, expectedBytes);
}

function isPublicTunnelRequest(request) {
  return requestAccessMode(request) === "public";
}

function publicAccessToken() {
  const password = process.env.PUBLIC_ACCESS_PASSWORD || "";
  return password ? createHash("sha256").update(`travel-co-creation:${password}`).digest("hex") : "";
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
<title>进入出行共创台</title>
<style>
*{box-sizing:border-box}body{margin:0;min-height:100vh;display:grid;place-items:center;padding:24px;background:#f5f6f3;color:#171717;font-family:Pretendard,"Apple SD Gothic Neo","Noto Sans KR","PingFang SC",sans-serif}body:before{content:"JEJU";position:fixed;left:-2vw;bottom:-7vw;color:#ff5a36;font-size:min(31vw,420px);font-weight:950;letter-spacing:-.09em;line-height:.8;opacity:.12;pointer-events:none}.card{position:relative;width:min(470px,100%);padding:42px;border:2px solid #171717;border-radius:30px;background:#fff;box-shadow:16px 16px 0 #2f62ff}.card:before{content:"";position:absolute;top:0;left:42px;right:42px;height:7px;background:linear-gradient(90deg,#ff5a36 0 25%,#ffc928 25% 50%,#2f62ff 50% 75%,#ef77a9 75%)}.mark{display:grid;place-items:center;width:56px;height:56px;border-radius:50%;background:#ff5a36;color:#fff;font-size:27px;font-weight:900}.eyebrow{margin:28px 0 10px;color:#2f62ff;font-size:11px;font-weight:900;letter-spacing:.2em}.card h1{margin:0;font-size:clamp(32px,8vw,45px);font-weight:950;line-height:1.08;letter-spacing:-.055em}.intro{margin:17px 0 30px;color:#656565;line-height:1.75;font-size:14px}label{display:block;margin-bottom:9px;font-size:12px;font-weight:850}input{width:100%;height:54px;padding:0 17px;border:1.5px solid #c8c8c8;border-radius:14px;background:#f8f8f6;color:#171717;font-size:17px;outline:none}input:focus{border-color:#2f62ff;box-shadow:0 0 0 4px rgba(47,98,255,.12)}button{width:100%;height:54px;margin-top:14px;border:2px solid #171717;border-radius:14px;background:#171717;color:#fff;font-size:15px;font-weight:900;cursor:pointer;box-shadow:5px 5px 0 #ff5a36}button:hover{transform:translate(-1px,-1px);box-shadow:7px 7px 0 #ff5a36}.error{margin:0 0 12px;padding:11px 13px;border-radius:12px;background:#fff0e9;color:#c43e20;font-size:13px}.note{margin:20px 0 0;color:#8a8a8a;font-size:11px;text-align:center}@media(max-width:520px){.card{padding:34px 24px;box-shadow:9px 9px 0 #2f62ff}.card:before{left:24px;right:24px}}
</style></head><body><main class="card"><div class="mark">ㅈ</div><p class="eyebrow">JEJU · 같이 가요</p><h1>JEJU,<br>TOGETHER.</h1><p class="intro">我们的济州岛旅行共创空间。输入团队共享密码，一起投递种草、比较候选、把三天行程定下来。</p>${error}<form method="post" action="/__team-login"><label for="password">团队访问密码</label><input id="password" name="password" type="password" autocomplete="current-password" required autofocus placeholder="请输入密码"><button type="submit">进入济州共创台 →</button></form><p class="note">06 FRIENDS · 03 DAYS · ONE JEJU NOTE</p></main></body></html>`;
  return sendBuffer(response, invalid ? 401 : 200, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
  }, Buffer.from(html));
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

function normalizedHostname(value) {
  return String(value || "").trim().toLowerCase().replace(/^\[|\]$/g, "").replace(/\.$/, "");
}

function isBlockedNetworkAddress(value, { allowProxySynthetic = false } = {}) {
  const address = normalizedHostname(value);
  const family = isIP(address);
  if (family === 4) {
    const [a, b, c] = address.split(".").map(Number);
    if (a === 0 || a === 10 || a === 127 || a >= 224) return true;
    if (a === 100 && b >= 64 && b <= 127) return true;
    if (a === 169 && b === 254) return true;
    if (a === 172 && b >= 16 && b <= 31) return true;
    if (a === 192 && b === 168) return true;
    if (a === 192 && b === 0 && (c === 0 || c === 2)) return true;
    if (a === 198 && (b === 18 || b === 19)) return !allowProxySynthetic;
    if (a === 198 && b === 51 && c === 100) return true;
    if (a === 203 && b === 0 && c === 113) return true;
    return false;
  }
  if (family === 6) {
    if (address === "::" || address === "::1") return true;
    if (/^(?:fc|fd)/.test(address) || /^fe[89ab]/.test(address) || /^ff/.test(address) || /^2001:db8/.test(address)) return true;
    if (/^(?:0*:)*ffff:/.test(address)) {
      const mapped = address.split(":").at(-1);
      return mapped && isIP(mapped) === 4 ? isBlockedNetworkAddress(mapped, { allowProxySynthetic }) : true;
    }
    return false;
  }
  return false;
}

function assertPublicHttpUrl(value) {
  const url = new URL(value);
  if (!["http:", "https:"].includes(url.protocol)) throw new Error("仅支持 http 或 https 公网链接");
  if (url.username || url.password) throw new Error("链接不能包含账号或密码");
  const hostname = normalizedHostname(url.hostname);
  if (!hostname || hostname === "localhost" || hostname.endsWith(".localhost") || hostname.endsWith(".local")
    || hostname.endsWith(".internal") || hostname.endsWith(".lan") || hostname.endsWith(".home.arpa")) {
    throw new Error("不能读取本机或局域网地址");
  }
  if (isIP(hostname) && isBlockedNetworkAddress(hostname)) throw new Error("不能读取本机或局域网地址");
  return url;
}

function cleanUrl(value) {
  const url = assertPublicHttpUrl(value);
  url.hash = "";
  for (const key of [...url.searchParams.keys()]) {
    if (/^(utm_|spm|from|source)/i.test(key)) url.searchParams.delete(key);
  }
  return url.toString();
}

async function ensurePublicDestination(url) {
  const checked = assertPublicHttpUrl(url);
  const hostname = normalizedHostname(checked.hostname);
  if (isIP(hostname)) return;
  let addresses;
  try {
    addresses = await dnsLookup(hostname, { all: true, verbatim: true });
  } catch {
    throw new Error("域名暂时无法解析");
  }
  if (!addresses.length) throw new Error("域名没有可用地址");
  // Clash/TUN 的 fake-ip 模式会把公网域名映射到 198.18.0.0/15；只对域名解析结果兼容，用户直接填写该网段仍会被拦截。
  if (addresses.some((item) => isBlockedNetworkAddress(item.address, { allowProxySynthetic: true }))) {
    throw new Error("链接解析到了本机或局域网地址，已停止读取");
  }
}

async function fetchExternalPage(value) {
  let target = assertPublicHttpUrl(value);
  for (let redirectCount = 0; redirectCount <= 5; redirectCount += 1) {
    await ensurePublicDestination(target);
    const response = await fetch(target, {
      headers: { "User-Agent": "Mozilla/5.0 (compatible; TravelCoCreationStudio/1.0; local team research)" },
      redirect: "manual",
      signal: AbortSignal.timeout(20_000),
    });
    if (![301, 302, 303, 307, 308].includes(response.status)) return response;
    const location = response.headers.get("location");
    if (!location) return response;
    if (redirectCount === 5) throw new Error("网页跳转次数过多");
    await response.body?.cancel();
    target = assertPublicHttpUrl(new URL(location, target));
  }
  throw new Error("网页跳转次数过多");
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
    if (requested === "攻略文章") return "攻略";
    if (requested === "室内休闲") return "休闲娱乐";
    if (requested === "景点户外") return "景点";
    return normalizeCategory(requested, text);
  }
  return normalizeCategory("", text);
}

function parseJsonFromModel(text) {
  const trimmed = text.trim().replace(/^```json\s*/i, "").replace(/```$/i, "").trim();
  return JSON.parse(trimmed);
}

const categoryFieldInstructions = {
  住宿: "details 提取 address, locationHighlights, capacity, roomType, rooms, beds, bedTypes, bathrooms, twoNightTotal, extraFees, deposit, environment, entireRental, kitchen, barbecue, bbqEquipment, breakfast, parking, checkIn, checkOut, bookingStatus, reservation, cancellationPolicy, evidence。只记录团队人数、住宿总价和偏好之间的匹配点与冲突点，不得替团队下‘适合/不适合’结论。",
  餐饮: "details 提取 address, cuisineType, signatureDishes, sixPersonTotal, privateRoom, queueInfo, groupSuitability, environment, parking, openingHours, reservation, cancellationPolicy, bookingStatus, usage, evidence。groupSuitability 只写客观容量、包间或排队条件，不输出最终适合性结论；subCategory 优先使用烧烤、火锅、早茶早餐、炒菜正餐、自助餐、夜宵、甜品饮品、咖啡、酒吧；price 表示人均价格。",
  密室: "details 提取 address, themeName, escapeStyle, venueSize, roomCount, capacity, minPlayers, maxPlayers, sixPersonSession, horrorLevel, difficulty, npcInteraction, physicalIntensity, costume, openingHours, reservation, cancellationPolicy, parking, bookingStatus, sixPersonTotal, evidence。恐怖程度规范为无恐、微恐、中恐、重恐之一；没有原文依据则待核实。",
  休闲娱乐: "details 提取 address, leisureFacilities, packageInfo, sixPersonTotal, capacity, openingHours, overnight, includedMeals, restArea, privateRoom, genderArrangement, serviceRestrictions, environment, parking, reservation, cancellationPolicy, bookingStatus, evidence。subCategory 优先使用汗蒸、桑拿、洗浴中心、温泉、SPA、足疗按摩、KTV、桌游、电竞、电玩城、沉浸式剧场。",
  景点: "details 提取 address, attractionType, ticketInfo, recommendedDuration, indoorOutdoor, weatherImpact, openingHours, reservation, cancellationPolicy, parking, bookingStatus, evidence。当前不计算路程，必须尽量确定具体地点。",
  攻略: "提取文章中明确提到的地点、餐饮、住宿或活动线索；subCategory 使用美食攻略、住宿攻略、行程攻略或综合攻略；details 至少提取 address, evidence。",
};

function systemPromptFor(category, sourceType = "链接", trip = {}) {
  const destination = String(trip.project?.destination || trip.finalPlan?.destination || "待确定目的地");
  const people = boundedInteger(trip.tripProfile?.groupSize || trip.project?.people || trip.finalPlan?.people, 4, 1, 50);
  const schedule = String(trip.tripProfile?.schedule || trip.finalPlan?.schedule || "日期与行程待确认");
  const stayPreference = String(trip.tripProfile?.stayPreference || "住宿条件待团队确认");
  const activity = String(trip.tripProfile?.activity || "团队活动待确认");
  const sourceRule = sourceType === "文字"
    ? "本次输入是团队成员直接写下的个人诉求，不是商家页面。只返回一个 candidateType=requirement 的候选，把明确表达的预算、位置、人数、类型、环境和偏好提取为需求条件；不得把愿望写成已经核实的商家事实。name 写成简短的需求名称，dataStatus 写明‘团队文字需求，具体商户待匹配’，factsFound 记录已表达的偏好，missingFields 记录仍需用真实链接或商户信息核实的内容。"
    : "本次输入是网页链接。先判断它是单一商户/地点页面，还是攻略、榜单、合集。单一商户只返回一个 candidateType=place；攻略、榜单、合集要把正文中每个名称明确的商户、住宿、景点或活动拆成独立的 candidateType=place 候选，不能把十家店合成一张卡。若正文只是泛泛攻略、没有足够信息形成具体地点，则返回一个 candidateType=guide 的文章候选。只根据网页正文提取事实，无法读取或正文未写明的内容必须标为待核实。";
  return `你是“出行共创台”的旅行资料整理与初步策划助手。本次目的地是“${destination}”，同行 ${people} 人，行程结构是“${schedule}”，住宿偏好是“${stayPreference}”，团队活动偏好是“${activity}”。你的任务是把团队投递整理成可在 Excel 横向比较的数据，并结合输入中的旅行上下文提出初步建议。${sourceRule}输出严格 JSON。
顶层必须是对象并包含 candidates 数组；数组每项必须包含：candidateType(place|requirement|guide), name, category, subCategory, featureTags(string数组), area, price(number或null), priceLabel, duration, summary, tags(string数组), pros(string数组), cons(string数组), score(0到5), scoreReason, aiCompleteness(0到100), aiPriceBasis, aiPriceConfidence(高|中|低), planningSuggestions(string数组), candidateLeads(数组，每项含 name, area, reason, confidence), dataStatus, verificationStatus(待核实|部分核实|已核实), factsFound(string数组), missingFields(string数组), details(object)。最多返回 20 个候选。另可在顶层提供 summary、factsFound、missingFields，概括整篇来源。
category 只能是住宿、餐饮、密室、休闲娱乐、景点、攻略。当前预分类是“${category}”，只有正文明确证明分类错误时才调整。featureTags 可多选，例如海景、黑猪烤肉、海鲜、咖啡、徒步、雨天备选、适合${people}人。
所有类别都要提取具体地点、价格、营业或入住时间、预约/取消规则、人数、优缺点和证据。允许为了初步规划预测价格区间、推荐分、优缺点和资料完整度，但必须把预测写在 priceLabel、aiPriceBasis、aiPriceConfidence、scoreReason 中，并明确标注“AI估价”或“AI推测”，不得冒充来源事实。details.sixPersonTotal 与 details.sixPersonSession 是兼容旧数据的内部字段，分别表示当前 ${people} 人团队总价与当前团队能否独立成团。${categoryFieldInstructions[category] || categoryFieldInstructions.攻略}
planningSuggestions 可以建议“第2天下午”“雨天备用”“晚餐后”等一个到三个时段，并说明与当前行程的衔接或冲突；这只是建议，不得输出或修改 selected、decisionStatus、正式行程日期与是否淘汰。不得替团队输出“适合我们/不适合我们”的最终结论，只能列出匹配点、冲突点和待确认问题。
candidateLeads 可以给出最多三个可能的地点线索；若输入和上下文没有可核实来源，必须把 reason 写明“AI线索待核实”，不得标成真实候选或已核实。
evidence 用简短文字概括输入中明确表达的事实或偏好，不编造引文。输入没有明确写出的事实字段写“待核实”，并放入 missingFields。无关字段可以省略，系统会自动补齐。`;
}

function normalizeAnalysisCandidates(analysis, fallback, { sourceType = "链接" } = {}) {
  const envelope = analysis && typeof analysis === "object" ? analysis : {};
  const supplied = Array.isArray(envelope)
    ? envelope
    : Array.isArray(envelope.candidates)
      ? envelope.candidates
      : [envelope];
  const valid = supplied.filter((item) => item && typeof item === "object" && !Array.isArray(item)).slice(0, 20);
  const candidates = valid.length ? valid : [fallback];
  return candidates.slice(0, sourceType === "文字" ? 1 : 20).map((item, index) => {
    const base = index === 0 ? fallback : {
      ...fallback,
      name: "",
      summary: "",
      factsFound: [],
      missingFields: [],
      details: { ...defaultDetails },
    };
    const category = normalizeCategory(item.category || base.category, `${item.name || ""} ${item.summary || ""}`);
    const candidateType = sourceType === "文字"
      ? "requirement"
      : normalizeCandidateType(item.candidateType, { sourceType, category, dataStatus: item.dataStatus });
    return {
      ...base,
      ...item,
      candidateType,
      category: candidateType === "guide" ? "攻略" : category,
      details: { ...(base.details || {}), ...(item.details && typeof item.details === "object" ? item.details : {}) },
      factsFound: stringList(item.factsFound, base.factsFound || []),
      missingFields: stringList(item.missingFields, base.missingFields || []),
    };
  });
}

async function modelJsonResponse(systemPrompt, content, { maxTokens = 8_000 } = {}) {
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
        max_tokens: maxTokens,
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
  return null;
}

async function modelAnalysis(content, fallback, category, sourceType = "链接", trip = {}) {
  const systemPrompt = systemPromptFor(category, sourceType, trip);
  return await modelJsonResponse(systemPrompt, content) || fallback;
}

function normalizeAiConfidence(value) {
  const normalized = String(value || "").trim();
  return ["高", "中", "低"].includes(normalized) ? normalized : "低";
}

function normalizeCandidateLeads(value) {
  return (Array.isArray(value) ? value : []).filter((item) => item && typeof item === "object" && !Array.isArray(item)).slice(0, 3).map((item) => ({
    name: limitedText(item.name, "待核实地点线索", 120),
    area: limitedText(item.area, "区域待核实", 120),
    reason: limitedText(item.reason, "AI线索待核实", 400),
    confidence: normalizeAiConfidence(item.confidence),
  }));
}

function normalizedAiCompleteness(value, missingFields = []) {
  if (value === null || value === undefined || value === "") return null;
  const proposed = Number(value);
  if (!Number.isFinite(proposed)) return null;
  const missingPenalty = Math.min(70, (Array.isArray(missingFields) ? missingFields.length : 0) * 10);
  return Math.min(100 - missingPenalty, Math.max(0, Math.round(proposed)));
}

function buildTripReviewSnapshot(state) {
  const finalPlan = state.finalPlan || {};
  const places = Array.isArray(state.places) ? state.places : [];
  const links = Array.isArray(state.links) ? state.links : [];
  return {
    trip: {
      destination: finalPlan.destination || state.project?.destination,
      dates: finalPlan.dates || state.tripProfile?.dates,
      schedule: finalPlan.schedule || state.tripProfile?.schedule,
      people: finalPlan.people || state.project?.people,
      perPersonBudget: finalPlan.perPersonBudget || "待确认",
      roundTripFlightPerPerson: finalPlan.roundTripFlightPerPerson || "待确认",
      transportPreference: "公交优先，必要时短途拼车",
      stayPreference: state.tripProfile?.stayPreference || "待确认",
      activityPreference: state.tripProfile?.activity || "待确认",
    },
    stay: finalPlan.stay || {},
    itinerary: (Array.isArray(finalPlan.itinerary) ? finalPlan.itinerary : []).slice(0, 40).map((item, index) => ({
      id: `itinerary-${index}`,
      day: item.day,
      time: item.time,
      endTime: item.endTime,
      category: item.category,
      title: item.title,
      address: item.address,
      transport: item.transport,
      costPerPerson: item.cost,
      bookingStatus: item.bookingStatus,
      note: item.note,
      sourceId: item.sourceId,
    })),
    candidates: places.slice(0, 80).map((place) => {
      const link = links.find((item) => item.id === place.sourceId);
      const quality = candidateQuality(place, link);
      const voteCounts = placeVoteColumns(place);
      return {
        id: place.id,
        type: place.candidateType,
        name: place.name,
        category: place.category,
        subCategory: place.subCategory,
        area: place.area,
        price: place.price,
        priceLabel: place.priceLabel,
        duration: place.duration,
        aiScore: place.score,
        completeness: place.aiCompleteness ?? quality.percent,
        missingFields: [...new Set([...(place.missingFields || []), ...quality.missing])].slice(0, 12),
        verificationStatus: place.verificationStatus,
        votes: { want: voteCounts[0], okay: voteCounts[1], reject: voteCounts[2] },
        selected: Boolean(place.selected),
        humanDecision: place.decisionStatus,
        humanNote: limitedText(place.manualNote, "", 300),
        lockedFields: Object.keys(place.manualOverrides || {}),
      };
    }),
    recentRequirements: links.filter((item) => item.sourceType === "文字").slice(0, 20).map((item) => ({
      id: item.id,
      submitter: item.submitter,
      text: limitedText(item.inputText, "", 1_000),
      status: item.status,
    })),
    budget: buildBudgetSnapshot(state),
    authority: {
      allowed: ["AI推荐分", "AI估价", "优缺点", "资料完整度", "缺失信息", "建议时段", "待核实候选线索"],
      forbidden: ["自动入选", "自动淘汰", "替团队判断是否适合", "直接修改正式行程日期", "覆盖投票", "覆盖人工锁定字段"],
    },
  };
}

const reviewPatchFields = new Set([
  "score", "priceLabel", "duration", "summary", "pros", "cons", "missingFields", "featureTags",
  "aiCompleteness", "aiPriceBasis", "aiPriceConfidence", "aiScoreReason", "planningSuggestions",
]);

function normalizeReviewSuggestion(item, index = 0) {
  const source = item && typeof item === "object" && !Array.isArray(item) ? item : {};
  const targetType = ["candidate", "itinerary", "stay", "plan", "information"].includes(source.targetType) ? source.targetType : "information";
  const rawField = limitedText(source.field, "", 80);
  const field = ({ aiScore: "score", recommendedScore: "score", completeness: "aiCompleteness", aiEstimatedPrice: "priceLabel", priceEstimate: "priceLabel", suggestedTime: "planningSuggestions" })[rawField] || rawField;
  const canApplyDirectly = targetType === "candidate" && reviewPatchFields.has(field) && Boolean(source.targetId);
  const proposedValue = Array.isArray(source.proposedValue)
    ? stringList(source.proposedValue).slice(0, 12)
    : typeof source.proposedValue === "number" ? source.proposedValue : limitedText(source.proposedValue, "", 1_000);
  return {
    id: limitedText(source.id, `suggestion-${index + 1}-${randomUUID().slice(0, 8)}`, 120),
    category: limitedText(source.category, "资料完善", 80),
    targetType,
    targetId: limitedText(source.targetId, "", 160),
    targetLabel: limitedText(source.targetLabel, "当前方案", 160),
    field,
    currentValue: limitedText(source.currentValue, "", 1_000),
    proposedValue,
    reason: limitedText(source.reason, "等待补充建议理由", 1_000),
    confidence: normalizeAiConfidence(source.confidence),
    evidenceType: limitedText(source.evidenceType, "AI推测", 80),
    status: ["待处理", "已采纳", "已忽略"].includes(source.status) ? source.status : "待处理",
    applyMode: canApplyDirectly ? "direct" : "advice",
    appliedAt: limitedText(source.appliedAt, "", 80),
    appliedResult: limitedText(source.appliedResult, "", 500),
  };
}

function normalizeAiReviews(value) {
  return (Array.isArray(value) ? value : []).filter((item) => item && typeof item === "object" && !Array.isArray(item)).slice(0, 10).map((item) => ({
    id: limitedText(item.id, `review-${randomUUID()}`, 120),
    generatedAt: limitedText(item.generatedAt, new Date().toISOString(), 80),
    baseUpdatedAt: limitedText(item.baseUpdatedAt, "", 80),
    provider: limitedText(item.provider, "规则检查", 120),
    source: item.source === "ai" ? "ai" : "rules",
    headline: limitedText(item.headline, "当前方案审阅", 200),
    summary: limitedText(item.summary, "已检查当前旅行方案。", 1_500),
    note: limitedText(item.note, "所有建议均需人工确认。", 500),
    suggestions: (Array.isArray(item.suggestions) ? item.suggestions : []).slice(0, 12).map(normalizeReviewSuggestion),
  }));
}

function enforceAiReviewSafety(state, reviews = state.aiReviews || []) {
  for (const review of reviews) {
    for (const suggestion of review.suggestions || []) {
      if (suggestion.applyMode !== "direct") continue;
      const candidate = (state.places || []).find((item) => item.id === suggestion.targetId);
      const locked = candidate && Object.prototype.hasOwnProperty.call(candidate.manualOverrides || {}, suggestion.field);
      const protectedPrice = suggestion.field === "priceLabel" && candidate
        && ((candidate.price !== null && Number.isFinite(Number(candidate.price))) || candidate.verificationStatus === "已核实");
      if (!candidate || locked || protectedPrice) suggestion.applyMode = "advice";
    }
  }
  return reviews;
}

function fallbackTripReview(snapshot, error = "") {
  const suggestions = [];
  if (/待|未/.test(String(snapshot.trip.roundTripFlightPerPerson || ""))) suggestions.push({
    category: "资料完善", targetType: "plan", targetLabel: "往返机票", currentValue: snapshot.trip.roundTripFlightPerPerson,
    proposedValue: "补充六个人的实际含税往返机票价格和行李额度", reason: "机票包含在每人6000元预算内，缺少真实票价时无法可靠判断剩余预算。", confidence: "高", evidenceType: "当前方案", field: "roundTripFlightPerPerson",
  });
  if (/待|未/.test(String(snapshot.stay?.name || ""))) suggestions.push({
    category: "候选补充", targetType: "stay", targetLabel: "住宿", currentValue: snapshot.stay?.name || "未选择",
    proposedValue: "先补充3至5个济州市区六人住宿真实链接", reason: "住宿仍是当前方案里最大的未确定部分。", confidence: "高", evidenceType: "当前方案", field: "name",
  });
  const incomplete = snapshot.candidates.filter((item) => item.type === "place" && item.missingFields.length).slice(0, 3);
  for (const candidate of incomplete) suggestions.push({
    category: "资料完善", targetType: "candidate", targetId: candidate.id, targetLabel: candidate.name, field: "missingFields",
    currentValue: candidate.missingFields.join("；"), proposedValue: candidate.missingFields,
    reason: "先补齐动态价格、开放时间或预约要求，再进入最终比较。", confidence: "高", evidenceType: "字段完整度检查",
  });
  return {
    headline: "当前方案仍需补齐关键动态信息",
    summary: "规则检查已整理出最需要人工确认的项目；正式行程不会被自动修改。",
    suggestions: suggestions.slice(0, 8),
    note: error ? `DeepSeek 本次未完成，已使用规则检查：${limitedText(error, "", 220)}` : "当前为规则检查结果。",
  };
}

async function generateTripReview(state) {
  const snapshot = buildTripReviewSnapshot(state);
  const systemPrompt = `你是六人济州岛旅行的“方案审阅与初步策划助手”。请审阅给出的结构化旅行快照，并返回严格 JSON。你可以：预测价格区间、打推荐分、分析优缺点与资料完整度、提出一到三个建议时段、指出路线冲突、给出待核实地点线索。所有预测必须写明依据类型和置信度。你不可以：判断某项是否最终适合团队、自动入选或淘汰、替团队投票、直接改变正式行程第几天、覆盖人工锁定字段。
输出字段：headline, summary, note, suggestions。suggestions 最多8项，每项必须包含 category, targetType(candidate|itinerary|stay|plan|information), targetId, targetLabel, field, currentValue, proposedValue, reason, confidence(高|中|低), evidenceType。
只有针对已有 candidate 的下列字段可以作为可直接采纳修改：score, priceLabel, duration, summary, pros, cons, missingFields, featureTags, aiCompleteness, aiPriceBasis, aiPriceConfidence, aiScoreReason, planningSuggestions。价格预测写入 priceLabel 时必须以“AI估价：”开头。时段和路线建议只能作为 advice，不能修改 day、selected、decisionStatus 或投票。`;
  let source = "ai";
  let result;
  let note = "";
  try {
    result = await modelJsonResponse(systemPrompt, JSON.stringify(snapshot), { maxTokens: 6_000 });
    if (!result) {
      source = "rules";
      result = fallbackTripReview(snapshot);
    }
  } catch (error) {
    source = "rules";
    note = error instanceof Error ? error.message : "AI审阅暂时不可用";
    result = fallbackTripReview(snapshot, note);
  }
  const review = normalizeAiReviews([{
    ...result,
    id: `review-${randomUUID()}`,
    generatedAt: new Date().toISOString(),
    baseUpdatedAt: state.finalPlan?.updatedAt || "",
    provider: source === "ai" ? providerLabel() : "本地规则检查",
    source,
    note: result?.note || note || "所有建议均需人工确认后才能写回。",
  }])[0];
  return enforceAiReviewSafety(state, [review])[0];
}

function applyAiReviewSuggestion(state, reviewId, suggestionId, action) {
  const review = (state.aiReviews || []).find((item) => item.id === reviewId);
  if (!review) throw new Error("AI审阅记录不存在");
  const suggestion = review.suggestions.find((item) => item.id === suggestionId);
  if (!suggestion) throw new Error("AI建议不存在");
  if (!['accept', 'dismiss'].includes(action)) throw new Error("请选择采纳或忽略");
  suggestion.appliedAt = new Date().toISOString();
  if (action === "dismiss") {
    suggestion.status = "已忽略";
    suggestion.appliedResult = "团队决定暂不采用，正式方案未修改";
    return suggestion;
  }
  suggestion.status = "已采纳";
  if (suggestion.applyMode !== "direct") {
    suggestion.appliedResult = "已保留为人工调整事项，未自动修改正式行程";
    return suggestion;
  }
  const candidate = state.places.find((item) => item.id === suggestion.targetId);
  if (!candidate) throw new Error("建议对应的候选不存在，请重新审阅");
  if (Object.prototype.hasOwnProperty.call(candidate.manualOverrides || {}, suggestion.field)) {
    throw new Error("这个字段已经人工锁定，AI建议不会覆盖");
  }
  const field = suggestion.field;
  if (!reviewPatchFields.has(field)) throw new Error("这项建议只能作为参考，不能自动写回");
  if (["pros", "cons", "missingFields", "featureTags", "planningSuggestions"].includes(field)) {
    candidate[field] = stringList(suggestion.proposedValue).slice(0, field === "planningSuggestions" ? 3 : 12);
  } else if (field === "score") {
    candidate.score = Math.min(5, Math.max(0, Number(suggestion.proposedValue) || candidate.score || 3.5));
  } else if (field === "aiCompleteness") {
    candidate.aiCompleteness = normalizedAiCompleteness(suggestion.proposedValue, candidate.missingFields) ?? candidate.aiCompleteness;
  } else if (field === "aiPriceConfidence") {
    candidate.aiPriceConfidence = normalizeAiConfidence(suggestion.proposedValue);
  } else if (field === "priceLabel") {
    if ((candidate.price !== null && Number.isFinite(Number(candidate.price))) || candidate.verificationStatus === "已核实") {
      throw new Error("已有核实价格，AI估价不会覆盖");
    }
    const estimate = limitedText(suggestion.proposedValue, "", 160);
    candidate.priceLabel = estimate.startsWith("AI估价：") ? estimate : `AI估价：${estimate}`;
  } else {
    candidate[field] = limitedText(suggestion.proposedValue, candidate[field] || "", field === "summary" ? 1_000 : 500);
  }
  suggestion.appliedResult = `已更新候选“${candidate.name}”的${field}字段，并写入 Excel`;
  return suggestion;
}

function parseMoneyRange(value) {
  const source = String(value || "").trim();
  const currency = source.match(/(?:¥|￥|CNY|人民币)\s*([\d,]+(?:\.\d+)?)(?:\s*(?:-|\u2013|—|~|至|到)\s*(?:(?:¥|￥|CNY|人民币)\s*)?([\d,]+(?:\.\d+)?))?/i);
  const contextual = source.match(/(?:预算|总价|费用|价格)[^\d]{0,12}([\d,]+(?:\.\d+)?)(?:\s*(?:-|\u2013|—|~|至|到)\s*([\d,]+(?:\.\d+)?))?/i);
  const match = currency || contextual;
  if (!match) return null;
  const first = Number(String(match[1]).replace(/,/g, ""));
  const second = Number(String(match[2] || match[1]).replace(/,/g, ""));
  if (!Number.isFinite(first) || !Number.isFinite(second)) return null;
  return { min: Math.min(first, second), max: Math.max(first, second) };
}

function budgetItemCategory(item) {
  const context = `${item?.category || ""} ${item?.title || ""}`;
  if (/早餐|午餐|晚餐|餐饮|美食|烤肉|海鲜|小吃|宵夜|咖啡|市场/.test(context)) return "餐饮";
  if (/交通|公交|拼车|出租|机场|返程|抵达/.test(context)) return "岛内交通";
  if (/购物|伴手礼|纪念品/.test(context)) return "购物";
  if (/住宿|民宿|酒店|公寓/.test(context)) return "住宿";
  return "游玩";
}

function formatYuanRange(minimum, maximum = minimum) {
  if (!Number.isFinite(minimum) || !Number.isFinite(maximum)) return "待确认";
  const roundedMinimum = Math.round(minimum);
  const roundedMaximum = Math.round(maximum);
  return roundedMinimum === roundedMaximum
    ? `¥${roundedMinimum.toLocaleString("zh-CN")} / 人`
    : `¥${roundedMinimum.toLocaleString("zh-CN")}–${roundedMaximum.toLocaleString("zh-CN")} / 人`;
}

function buildBudgetSnapshot(state) {
  const finalPlan = state.finalPlan || {};
  const people = boundedInteger(finalPlan.people || state.project?.people, 6, 1, 50);
  const target = parseMoneyRange(finalPlan.perPersonBudget);
  const budgetLabel = String(finalPlan.perPersonBudget || "");
  const excludesFlights = /不含.*机票|机票.*不含/.test(budgetLabel);
  const includesFlights = !excludesFlights && /(?:包含|含).{0,12}机票|机票.{0,12}(?:包含|含)/.test(budgetLabel);
  const flight = includesFlights ? parseMoneyRange(finalPlan.roundTripFlightPerPerson) : null;
  const stayTotal = parseMoneyRange(finalPlan.stay?.twoNightTotal);
  const stay = stayTotal ? { min: stayTotal.min / people, max: stayTotal.max / people } : null;
  const itemTotals = { "餐饮": 0, "岛内交通": 0, "游玩": 0, "购物": 0, "住宿": 0 };
  let zeroCostItems = 0;
  for (const item of Array.isArray(finalPlan.itinerary) ? finalPlan.itinerary : []) {
    const cost = boundedCost(item?.cost, 0);
    const category = budgetItemCategory(item);
    itemTotals[category] = (itemTotals[category] || 0) + cost;
    if (!cost && category !== "住宿") zeroCostItems += 1;
  }
  const itineraryTotal = Object.entries(itemTotals)
    .filter(([name]) => name !== "住宿")
    .reduce((total, [, value]) => total + value, 0);
  const knownMin = itineraryTotal + (stay?.min || 0) + (flight?.min || 0);
  const knownMax = itineraryTotal + (stay?.max || 0) + (flight?.max || 0);
  const missingInputs = [];
  if (!target) missingInputs.push("人均总预算");
  if (includesFlights && !flight) missingInputs.push("往返济州机票人均含税价格（含行李）");
  if (!stay) missingInputs.push("住宿含税两晚总价");
  if (!finalPlan.stay?.sourceUrl) missingInputs.push("住宿真实链接和动态价格");
  if (zeroCostItems) missingInputs.push(`${zeroCostItems} 个行程项目的实际费用`);
  const unpricedPlaces = (state.places || []).filter((place) => normalizeCandidateType(place.candidateType, place) === "place" && !Number.isFinite(Number(place.price))).length;
  if (unpricedPlaces) missingInputs.push(`${unpricedPlaces} 个真实候选的动态价格`);
  const categories = [
    ...(includesFlights ? [{ name: "往返机票", min: flight?.min || 0, max: flight?.max || 0, basis: flight ? finalPlan.roundTripFlightPerPerson : "待填写实际含税票价（含行李）" }] : []),
    { name: "住宿", min: stay?.min || 0, max: stay?.max || 0, basis: stay ? finalPlan.stay.twoNightTotal : "待确认真实住宿" },
    { name: "餐饮", min: itemTotals["餐饮"], max: itemTotals["餐饮"], basis: "按当前三日行程的每人费用合计" },
    { name: "岛内交通", min: itemTotals["岛内交通"], max: itemTotals["岛内交通"], basis: "公交优先，必要时短途拼车" },
    { name: "游玩", min: itemTotals["游玩"], max: itemTotals["游玩"], basis: "按当前景点与活动费用合计" },
    ...(itemTotals["购物"] ? [{ name: "购物", min: itemTotals["购物"], max: itemTotals["购物"], basis: "按当前行程费用合计" }] : []),
  ];
  return {
    destination: finalPlan.destination || state.project?.destination || "待确认目的地",
    dates: finalPlan.dates || "待确认",
    people,
    days: Number(state.project?.days) || 3,
    includesFlights,
    excludesFlights,
    flight,
    target,
    targetLabel: String(finalPlan.perPersonBudget || "待确认"),
    known: { min: knownMin, max: knownMax },
    categories,
    missingInputs: [...new Set(missingInputs)].slice(0, 8),
    strategy: String(finalPlan.summary || "").slice(0, 600),
  };
}

function fallbackBudgetAdvice(snapshot) {
  const targetMiddle = snapshot.target ? (snapshot.target.min + snapshot.target.max) / 2 : 0;
  const assessments = snapshot.categories.map((category) => {
    const middle = (category.min + category.max) / 2;
    let assessment = "合理";
    let suggestion = "先保留这个区间，等真实链接和含税价格补齐后再调整。";
    if (!middle) {
      assessment = "待确认";
      suggestion = "当前还没有可靠金额，先补真实价格，暂不判断高低。";
    } else if (category.name === "往返机票" && targetMiddle && middle > targetMiddle * 0.55) {
      assessment = "偏高";
      suggestion = "机票已占人均总预算的一半以上；先核对是否含税、行李额和退改签，再决定是否压缩其他项目。";
    } else if (category.name === "往返机票") {
      suggestion = "核对含税总价、托运行李额和退改签规则，确认后再分配其余预算。";
    } else if (category.name === "住宿" && targetMiddle && middle > targetMiddle * 0.42) {
      assessment = "偏高";
      suggestion = "住宿已占较大比例，优先比较含税总价、房型和取消政策。";
    } else if (category.name === "住宿" && targetMiddle && middle < targetMiddle * 0.2) {
      assessment = "偏低";
      suggestion = "这个住宿价位可能没有包含清洁费或服务费，建议再核对。";
    } else if (category.name === "餐饮" && middle > 450) {
      assessment = "偏高";
      suggestion = "餐饮超出目前“正餐约 ¥100/人”的策略，可减少一顿正式聚餐。";
    } else if (category.name === "餐饮" && middle < 180) {
      assessment = "偏低";
      suggestion = "餐饮缓冲较少，如果想多吃一顿特色餐，建议多留 ¥100–200/人。";
    } else if (category.name === "岛内交通" && middle > 250) {
      assessment = "偏高";
      suggestion = "交通费接近包车型方案，建议重新核对公交与短途拼车的比例。";
    } else if (category.name === "岛内交通" && middle < 60) {
      assessment = "偏低";
      suggestion = "雨天、行李或公交接驳可能需要打车，建议增加小额交通机动金。";
    } else if (category.name === "游玩" && middle < 80) {
      assessment = "偏低";
      suggestion = "当前是免费海岸和步道优先的省钱方案；如果后期想加付费项目，可再补预算。";
    }
    const planned = category.name === "往返机票" && !snapshot.flight
      ? "待确认"
      : formatYuanRange(category.min, category.max);
    return { name: category.name, planned, assessment, suggestion };
  });
  let overallStatus = "信息不足";
  let headline = "预算框架已建好，还需真实价格才能判断。";
  let reserveAdvice = "先补住宿和候选的真实价格。";
  let nextAction = "优先补住宿含税总价和主要候选价格。";
  if (snapshot.target) {
    const reserveLow = Math.max(0, snapshot.target.min - snapshot.known.max);
    const reserveHigh = Math.max(0, snapshot.target.max - snapshot.known.min);
    reserveAdvice = `按已填金额估算，每人还有约 ¥${Math.round(reserveLow).toLocaleString("zh-CN")}–${Math.round(reserveHigh).toLocaleString("zh-CN")} 可用于未定项和机动支出。`;
    if (snapshot.known.min > snapshot.target.max) {
      overallStatus = "超出预算";
      const suggested = Math.ceil((snapshot.known.max + 200) / 100) * 100;
      headline = "已知支出已超过当前预算上限。";
      nextAction = `先删减偏高项；如果都要保留，建议将人均预算至少调到约 ¥${suggested.toLocaleString("zh-CN")}。`;
    } else if (snapshot.known.max > snapshot.target.max || snapshot.target.min - snapshot.known.max < 150) {
      overallStatus = "偏紧";
      headline = "当前方案可执行，但机动金偏少。";
      nextAction = "可先保持方案；真实价格高于现值时，再增加 ¥100–300/人或替换一个付费项目。";
    } else if (snapshot.known.max < snapshot.target.min * 0.7) {
      overallStatus = "信息不足";
      headline = "已填支出明显低于预算，可能还有项目没有计入。";
    } else {
      overallStatus = "合理";
      headline = "当前已知支出在预算内，仍有机动空间。";
      nextAction = "先不增加总预算，等住宿和候选真实价格补齐后再决定。";
    }
    if (snapshot.includesFlights && !snapshot.flight) {
      overallStatus = "信息不足";
      headline = "人均 ¥6,000 总预算已明确，先补机票价格再分配。";
      reserveAdvice = "往返机票包含在 ¥6,000/人的总预算内，但实际票价尚未填写；现在不能把账面差额全部视为机动金。";
      nextAction = "先填写每人的往返机票含税价和行李费用，再判断住宿、餐饮、交通和游玩可以增加多少。";
    }
  }
  return {
    headline,
    overallStatus,
    summary: `已知项目约 ${formatYuanRange(snapshot.known.min, snapshot.known.max)}，目标为 ${snapshot.targetLabel}。${snapshot.includesFlights ? "本次将往返济州机票计入总预算。" : snapshot.excludesFlights ? "本次不将往返济州机票计入。" : ""}`,
    categories: assessments,
    reserveAdvice,
    nextAction,
    missingInputs: snapshot.missingInputs,
  };
}

function normalizeBudgetAdvice(value, fallback, snapshot) {
  const source = value && typeof value === "object" ? value : {};
  const allowedAssessments = new Set(["偏高", "合理", "偏低", "待确认"]);
  const suppliedCategories = Array.isArray(source.categories) ? source.categories : [];
  const normalized = {
    headline: limitedText(source.headline, fallback.headline, 160),
    overallStatus: ["合理", "偏紧", "超出预算", "信息不足"].includes(source.overallStatus) ? source.overallStatus : fallback.overallStatus,
    summary: limitedText(source.summary, fallback.summary, 600),
    categories: fallback.categories.map((base) => {
      const item = suppliedCategories.find((candidate) => candidate && candidate.name === base.name) || {};
      const suppliedSuggestion = limitedText(item.suggestion, base.suggestion, 300);
      const unsupportedMarketClaim = /旺季|淡季|当地|市场(?:价|水平|通常)|在济州.{0,20}(?:合理|中等|偏高|偏低)/.test(suppliedSuggestion);
      return {
        name: base.name,
        planned: base.planned,
        assessment: !unsupportedMarketClaim && allowedAssessments.has(item.assessment) ? item.assessment : base.assessment,
        suggestion: unsupportedMarketClaim ? base.suggestion : suppliedSuggestion,
      };
    }),
    reserveAdvice: limitedText(source.reserveAdvice, fallback.reserveAdvice, 400),
    nextAction: limitedText(source.nextAction, fallback.nextAction, 400),
    missingInputs: stringList(source.missingInputs, fallback.missingInputs).slice(0, 8),
  };
  if (snapshot?.includesFlights && !snapshot.flight) {
    normalized.overallStatus = "信息不足";
    normalized.reserveAdvice = fallback.reserveAdvice;
    normalized.nextAction = fallback.nextAction;
    normalized.missingInputs = fallback.missingInputs;
  }
  return normalized;
}

async function generateBudgetAdvice(state, { refresh = false } = {}) {
  const snapshot = buildBudgetSnapshot(state);
  const fingerprint = createHash("sha256").update(JSON.stringify({ provider: providerLabel(), snapshot })).digest("hex").slice(0, 20);
  if (refresh) budgetAdviceCache.delete(fingerprint);
  if (budgetAdviceCache.has(fingerprint)) return budgetAdviceCache.get(fingerprint);
  if (budgetAdvicePending.has(fingerprint)) return budgetAdvicePending.get(fingerprint);
  const task = (async () => {
    const fallback = fallbackBudgetAdvice(snapshot);
    const provider = (process.env.AI_PROVIDER || "demo").toLowerCase();
    let normalized = fallback;
    let source = "rules";
    let note = provider === "demo" ? "当前测试环境使用基础预算规则。" : "";
    if (provider !== "demo") {
      const systemPrompt = `你是六人济州岛旅行的预算助手。仅根据用户提供的总预算、机票、住宿和行程费用做分析，不得猜测实时市场价，也不得使用“旺季”“当地中等价位”“市场通常价格”等未在输入中提供的外部判断。未填或未核实的价格必须标为“待确认”。如果目标预算包含往返机票，必须把机票作为独立类别计入；机票实际价格未填时，overallStatus 必须为“信息不足”，不得把未分配差额都称为机动金或断言预算充足。必须区分“当前已知花费”和“团队可用总预算”，说明哪一部分偏高、合理、偏低或待确认，并建议剩余预算适合增加在哪些体验；只有确认后的总支出真的不足时才建议提高总预算。输出严格 JSON：headline, overallStatus(合理|偏紧|超出预算|信息不足), summary, categories(数组，每项 name, planned, assessment(偏高|合理|偏低|待确认), suggestion), reserveAdvice, nextAction, missingInputs(字符串数组)。`;
      try {
        const result = await modelJsonResponse(systemPrompt, JSON.stringify(snapshot), { maxTokens: 2_400 });
        normalized = normalizeBudgetAdvice(result, fallback, snapshot);
        source = "ai";
      } catch (error) {
        note = `AI 暂时未返回，已使用基础预算规则。${error instanceof Error ? ` ${error.message}` : ""}`.slice(0, 220);
      }
    }
    const advice = { ...normalized, source, provider: providerLabel(), note, fingerprint, generatedAt: new Date().toISOString() };
    budgetAdviceCache.set(fingerprint, advice);
    if (budgetAdviceCache.size > 20) budgetAdviceCache.delete(budgetAdviceCache.keys().next().value);
    return advice;
  })();
  budgetAdvicePending.set(fingerprint, task);
  try {
    return await task;
  } finally {
    budgetAdvicePending.delete(fingerprint);
  }
}

function demoAnalysis(title, text, category, url, destination = "") {
  const priceMatch = text.match(/(?:₩|KRW|韩元|¥|￥|人均|价格)[^\d]{0,6}(\d{2,7})/i);
  const price = priceMatch ? Number(priceMatch[1]) : null;
  const hostname = new URL(url).hostname.replace(/^www\./, "");
  return {
    name: title || hostname,
    category,
    subCategory: inferSubCategory(`${title} ${text}`, category),
    featureTags: inferFeatureTags(`${title} ${text}`, category),
    area: destination && text.includes(destination) ? `${destination}（具体区域待核实）` : "地点待核实",
    price,
    priceLabel: price ? `参考 ₩${price}` : "价格待核实",
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

function textSubmissionTitle(text) {
  const compact = String(text || "").replace(/\s+/g, " ").trim();
  return `团队需求｜${compact.slice(0, 42)}${compact.length > 42 ? "…" : ""}`;
}

function demoTextAnalysis(text, category, destination = "") {
  const compact = String(text || "").replace(/\s+/g, " ").trim();
  const groupSize = compact.match(/(?:适合|容纳|支持|我们|团队)?\s*(\d{1,2})\s*人/i)?.[1];
  const priceMatch = compact.match(/(?:人均|预算|不超过|最多|上限|价格|韩元|KRW)[^\d]{0,8}(\d{2,7})/i);
  const price = priceMatch ? Number(priceMatch[1]) : null;
  const details = { ...defaultDetails, evidence: `团队原始诉求：${compact.slice(0, 300)}` };
  if (category === "住宿") {
    if (groupSize) details.capacity = `团队明确要求适合 ${groupSize} 人`;
    if (/烧烤/.test(compact)) details.barbecue = "团队希望可以烧烤，需用真实房源核实";
    if (/整租/.test(compact)) details.entireRental = "团队偏好整租";
  } else if (category === "餐饮") {
    if (groupSize) details.groupSuitability = `团队明确要求适合 ${groupSize} 人`;
    details.usage = /早餐|早茶/.test(compact) ? "早餐 / 早茶" : "团队餐饮候选";
  } else if (category === "密室") {
    if (/微恐/.test(compact)) details.horrorLevel = "微恐";
    else if (/中恐/.test(compact)) details.horrorLevel = "中恐";
    else if (/重恐/.test(compact)) details.horrorLevel = "重恐";
    else if (/无恐/.test(compact)) details.horrorLevel = "无恐";
    if (groupSize) details.capacity = `${groupSize} 人`;
  } else if (category === "休闲娱乐") {
    details.leisureFacilities = inferSubCategory(compact, category);
  }
  return {
    name: textSubmissionTitle(compact).replace(/^团队需求｜/, ""),
    category,
    subCategory: inferSubCategory(compact, category),
    featureTags: inferFeatureTags(compact, category),
    area: destination && compact.includes(destination) ? `${destination}（具体区域待匹配）` : "地点待匹配",
    price,
    priceLabel: price ? `预算参考 ₩${price}` : "预算待补充",
    duration: "时长待匹配",
    summary: `团队成员提出：${compact.slice(0, 260)}`,
    tags: [category, "团队诉求", ...inferFeatureTags(compact, category)].slice(0, 8),
    pros: ["需求已经结构化，可直接与后续真实候选比较"],
    cons: ["这是团队偏好，不代表已找到符合条件的真实商户"],
    score: 3.5,
    dataStatus: "团队文字需求，具体商户待匹配",
    factsFound: ["团队原始诉求", "内容分类", "关键偏好"],
    missingFields: ["具体商户", "详细地址", "真实价格", "营业或预订规则"],
    details,
  };
}

function stringList(value, fallback = []) {
  const normalized = Array.isArray(value) ? value.map((item) => String(item).trim()).filter(Boolean) : [];
  return normalized.length ? normalized : fallback;
}

function filterTextMissingFields(items, originalText, destination = "") {
  const text = String(originalText || "");
  return items.filter((item) => {
    if (/恐怖程度/.test(item) && /(无恐|微恐|中恐|重恐)/.test(text)) return false;
    if (/难度/.test(item) && /(?:难度.{0,4})?(简单|中等|困难|高难)/.test(text)) return false;
    if (/NPC|真人互动/.test(item) && /NPC|真人互动/i.test(text)) return false;
    if (/适合.*人|人数|同时|开场/.test(item) && /\d{1,2}\s*人|[二三四五六七八九十]人/.test(text)) return false;
    if (/价格|预算/.test(item) && /(?:预算|人均|单人|不超过|以内).{0,10}\d+/.test(text)) return false;
    if (/位置|区域/.test(item) && destination && text.includes(destination)) return false;
    if (/预约/.test(item) && /预约/.test(text)) return false;
    return true;
  });
}

function normalizeDetails(value) {
  const supplied = value && typeof value === "object" && !Array.isArray(value) ? value : {};
  return Object.fromEntries(Object.keys(defaultDetails).map((key) => {
    const detail = supplied[key];
    const readable = typeof detail === "boolean" ? (detail ? "是" : "否") : detail;
    return [key, readable === null || readable === undefined || readable === "" ? defaultDetails[key] : String(readable).slice(0, 500)];
  }));
}

function applyManualOverrides(place, overrides) {
  const source = overrides && typeof overrides === "object" && !Array.isArray(overrides) ? overrides : {};
  const next = { ...place, details: { ...place.details }, manualOverrides: { ...source } };
  for (const [key, value] of Object.entries(source)) {
    if (key.startsWith("details.")) next.details[key.slice(8)] = value;
    else next[key] = value;
  }
  return next;
}

function rememberManualOverride(overrides, key, nextValue, currentValue) {
  if (nextValue === "" || nextValue === null || nextValue === undefined) return;
  const same = Array.isArray(nextValue)
    ? JSON.stringify(nextValue) === JSON.stringify(Array.isArray(currentValue) ? currentValue : [])
    : String(nextValue) === String(currentValue ?? "");
  if (!same) overrides[key] = nextValue;
}

function manualOverrideSummary(place) {
  const labels = {
    name: "名称", area: "区域", subCategory: "子分类", featureTags: "特征标签", price: "价格", pros: "优点", cons: "缺点",
    "details.address": "地址", "details.cuisineType": "餐饮类型", "details.signatureDishes": "招牌菜", "details.sixPersonTotal": "团队总价",
    "details.horrorLevel": "恐怖程度", "details.difficulty": "难度", "details.venueSize": "场地规模", "details.leisureFacilities": "休闲设施",
  };
  return Object.keys(place.manualOverrides || {}).map((key) => labels[key] || key.replace(/^details\./, "")).join("；");
}

async function updateLink(id, patch) {
  const state = await readState();
  const index = state.links.findIndex((item) => item.id === id);
  if (index < 0) return null;
  state.links[index] = { ...state.links[index], ...patch, updatedAt: new Date().toISOString() };
  await writeState(state);
  return state.links[index];
}

function candidateIdentity(candidate) {
  return `${candidate.candidateType || ""}|${candidate.category || ""}|${String(candidate.name || "").replace(/\s+/g, "").toLowerCase()}`;
}

function normalizedCandidateName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[“”"'`·•\s\-_（）()\[\]【】，,。.!！?？:：/\\]/g, "")
    .replace(/候选|推荐|攻略|风景区|景区/g, "");
}

function candidateNameSimilarity(left, right) {
  const a = normalizedCandidateName(left);
  const b = normalizedCandidateName(right);
  if (!a || !b) return 0;
  if (a === b) return 1;
  if (Math.min(a.length, b.length) >= 3 && (a.includes(b) || b.includes(a))) return 0.9;
  const grams = (text) => new Set(Array.from({ length: Math.max(0, text.length - 1) }, (_, index) => text.slice(index, index + 2)));
  const aGrams = grams(a);
  const bGrams = grams(b);
  if (!aGrams.size || !bGrams.size) return 0;
  const overlap = [...aGrams].filter((item) => bGrams.has(item)).length;
  return overlap / new Set([...aGrams, ...bGrams]).size;
}

function findExistingCandidate(existingForSource, unusedExisting, proposed) {
  const exact = existingForSource.find((item) => unusedExisting.has(item.id) && candidateIdentity(item) === candidateIdentity(proposed));
  if (exact) return exact;
  const ranked = existingForSource
    .filter((item) => unusedExisting.has(item.id) && item.candidateType === proposed.candidateType && item.category === proposed.category)
    .map((item) => ({
      item,
      similarity: candidateNameSimilarity(item.name, proposed.name),
      categoryBonus: item.category === proposed.category ? 0.12 : 0,
    }))
    .sort((left, right) => (right.similarity + right.categoryBonus) - (left.similarity + left.categoryBonus));
  return ranked[0] && ranked[0].similarity + ranked[0].categoryBonus >= 0.78 ? ranked[0].item : null;
}

function hasProtectedHumanState(place) {
  return Boolean(
    place.selected
    || Object.keys(place.votes || {}).length
    || Object.keys(place.manualOverrides || {}).length
    || String(place.manualNote || "").trim()
    || ["拟定", "备选", "淘汰"].includes(place.decisionStatus),
  );
}

function preserveUnmatchedCandidate(place) {
  const note = "新一轮分析未匹配到此条，因含人工结论已保留，请在主电脑确认";
  return {
    ...place,
    dataStatus: String(place.dataStatus || "").includes("新一轮分析未匹配") ? place.dataStatus : [place.dataStatus, note].filter(Boolean).join("；"),
    missingFields: [...new Set([...(place.missingFields || []), "确认是否被新分析结果替代"])],
  };
}

function generatedCandidateId(sourceId, candidate) {
  const location = normalizedCandidateName(candidate.details?.address || candidate.area || candidate.sourceUrl || "");
  const digest = createHash("sha1").update(`${sourceId}|${candidateIdentity(candidate)}|${location}`).digest("hex").slice(0, 16);
  return `place-${digest}`;
}

function dedupeAnalysisCandidates(candidates, { isText = false, fallbackCategory = "其他", fallbackName = "候选" } = {}) {
  const seen = new Set();
  return candidates.filter((candidate) => {
    const resolvedCategory = normalizeCategory(candidate.category || fallbackCategory, `${candidate.name || fallbackName} ${candidate.summary || ""}`);
    const candidateType = isText ? "requirement" : normalizeCandidateType(candidate.candidateType, {
      sourceType: "链接",
      category: resolvedCategory,
      dataStatus: candidate.dataStatus,
    });
    const effectiveCategory = candidateType === "guide" ? "攻略" : resolvedCategory;
    const name = normalizedCandidateName(limitedText(candidate.name, fallbackName, 160));
    const location = normalizedCandidateName(candidate.details?.address || candidate.area || candidate.sourceUrl || "");
    const key = `${candidateType}|${effectiveCategory}|${name}|${location}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function processLink(id) {
  const initialState = await readState();
  const initialRecord = initialState.links.find((item) => item.id === id);
  const isText = initialRecord?.sourceType === "文字" || (!initialRecord?.url && Boolean(initialRecord?.inputText));
  const record = await serializeMutation(() => updateLink(id, {
    status: isText ? "AI分析中" : "正在读取",
    readStatus: isText ? "文字已接收" : "等待读取",
    organizedStatus: "整理中",
    resultNote: isText ? "正在理解并整理团队诉求" : "正在读取网页",
    model: providerLabel(),
    error: "",
  }));
  if (!record) return;
  try {
    let title;
    let text;
    if (isText) {
      text = String(record.inputText || "").trim();
      if (text.length < 3) throw new Error("文字内容过少，请补充更具体的旅行诉求");
      title = record.title || textSubmissionTitle(text);
    } else {
      const response = await fetchExternalPage(record.url);
      if (!response.ok) throw new Error(`网页返回 ${response.status}`);
      const html = (await response.text()).slice(0, 1_500_000);
      title = extractTitle(html, new URL(record.url).hostname);
      text = pageText(html);
      if (text.length < 40) throw new Error("网页正文过少，可能需要登录或验证码");
    }
    const category = classify(`${title} ${text}`, record.category);
    await serializeMutation(() => updateLink(id, {
      status: "AI分析中",
      title,
      category,
      readStatus: isText ? "文字已接收" : "成功读取",
      organizedStatus: "整理中",
      resultNote: isText ? "文字已接收，正在提取需求条件" : "网页已读取，正在提取事实",
    }));
    const destination = String(initialState.project?.destination || initialState.finalPlan?.destination || "");
    const fallback = isText ? demoTextAnalysis(text, category, destination) : demoAnalysis(title, text, category, record.url, destination);
    const modelInput = isText
      ? `旅行上下文：\n${JSON.stringify(buildTripReviewSnapshot(initialState))}\n\n来源类型：团队文字诉求\n提交人：${record.submitter}\n原始文字：${text}\n补充说明：${record.note || "无"}`
      : `旅行上下文：\n${JSON.stringify(buildTripReviewSnapshot(initialState))}\n\n来源类型：网页链接\n链接：${record.url}\n标题：${title}\n正文：${text}`;
    const analysis = await modelAnalysis(modelInput, fallback, category, isText ? "文字" : "链接", initialState);
    const analysedCandidates = dedupeAnalysisCandidates(
      normalizeAnalysisCandidates(analysis, fallback, { sourceType: isText ? "文字" : "链接" }),
      { isText, fallbackCategory: category, fallbackName: title },
    );
    await serializeMutation(async () => {
    const state = await readState();
    const linkIndex = state.links.findIndex((item) => item.id === id);
    if (linkIndex < 0) return;
    const existingForSource = state.places.filter((item) => item.sourceId === id);
    const unusedExisting = new Set(existingForSource.map((item) => item.id));
    const mappedPlaces = analysedCandidates.map((candidate) => {
      const resolvedCategory = normalizeCategory(candidate.category || category, `${candidate.name || title} ${candidate.summary || ""}`);
      const candidateType = isText ? "requirement" : normalizeCandidateType(candidate.candidateType, {
        sourceType: "链接",
        category: resolvedCategory,
        dataStatus: candidate.dataStatus,
      });
      const effectiveCategory = candidateType === "guide" ? "攻略" : resolvedCategory;
      const inferredSubCategory = inferSubCategory(`${candidate.name || title} ${candidate.summary || ""} ${text}`, effectiveCategory);
      const proposedSubCategory = String(candidate.subCategory || inferredSubCategory).slice(0, 80);
      const subCategory = validSubCategory(effectiveCategory, proposedSubCategory) ? proposedSubCategory : inferredSubCategory;
      const featureTags = [...new Set([
        ...stringList(candidate.featureTags),
        ...inferFeatureTags(`${candidate.name || title} ${candidate.summary || ""} ${subCategory}`, effectiveCategory),
      ])].slice(0, 12);
      const rawMissingFields = stringList(candidate.missingFields, fallback.missingFields);
      const candidateMissingFields = (isText ? filterTextMissingFields(rawMissingFields, text, destination) : rawMissingFields).slice(0, 12);
      const proposedIdentity = { ...candidate, name: limitedText(candidate.name, title, 160), candidateType, category: effectiveCategory };
      let existingPlace = findExistingCandidate(existingForSource, unusedExisting, proposedIdentity);
      if (!existingPlace && analysedCandidates.length === 1 && existingForSource.length === 1
        && existingForSource[0].candidateType === candidateType && existingForSource[0].category === effectiveCategory) existingPlace = existingForSource[0];
      if (existingPlace) unusedExisting.delete(existingPlace.id);
      const dataStatus = candidate.dataStatus || (isText ? "团队文字需求，具体商户待匹配" : "AI总结，等待人工核实");
      const analysedPlace = {
        id: existingPlace?.id || generatedCandidateId(id, { ...candidate, candidateType, category: effectiveCategory }),
        sourceId: id,
        candidateType,
        verificationStatus: existingPlace?.verificationStatus || normalizeVerificationStatus(candidate.verificationStatus, candidateType, dataStatus),
        name: limitedText(candidate.name, title, 160),
        category: effectiveCategory,
        subCategory,
        featureTags,
        area: limitedText(candidate.area, "地点待核实", 200),
        price: Number.isFinite(Number(candidate.price)) && candidate.price !== null && candidate.price !== "" ? Number(candidate.price) : null,
        priceLabel: limitedText(candidate.priceLabel, "价格待核实", 120),
        duration: limitedText(candidate.duration, "时长待核实", 120),
        summary: limitedText(candidate.summary, "", 1_000),
        factsFound: stringList(candidate.factsFound, fallback.factsFound).slice(0, 12),
        missingFields: candidateMissingFields,
        score: Math.max(0, Math.min(5, Number(candidate.score) || 3.5)),
        aiCompleteness: normalizedAiCompleteness(candidate.aiCompleteness, candidateMissingFields),
        aiPriceBasis: limitedText(candidate.aiPriceBasis, "", 500),
        aiPriceConfidence: normalizeAiConfidence(candidate.aiPriceConfidence),
        aiScoreReason: limitedText(candidate.scoreReason || candidate.aiScoreReason, "", 500),
        planningSuggestions: stringList(candidate.planningSuggestions).slice(0, 3),
        candidateLeads: normalizeCandidateLeads(candidate.candidateLeads),
        tags: stringList(candidate.tags, featureTags).slice(0, 8),
        pros: stringList(candidate.pros, fallback.pros || ["等待进一步分析"]).slice(0, 4),
        cons: stringList(candidate.cons, fallback.cons || ["关键信息待核实"]).slice(0, 4),
        selected: existingPlace?.selected || false,
        votes: existingPlace?.votes || {},
        decisionStatus: existingPlace?.decisionStatus || (existingPlace?.selected ? "拟定" : "待比较"),
        manualNote: existingPlace?.manualNote || "",
        manualOverrides: existingPlace?.manualOverrides || {},
        sourceUrl: safeWebUrl(candidate.sourceUrl, record.url),
        dataStatus,
        details: normalizeDetails(candidate.details),
      };
      return applyManualOverrides(analysedPlace, existingPlace?.manualOverrides);
    });
    const nextPlaces = [...new Map(mappedPlaces.map((item) => [item.id, item])).values()];
    const preservedUnmatched = existingForSource
      .filter((item) => unusedExisting.has(item.id) && hasProtectedHumanState(item))
      .map(preserveUnmatchedCandidate);
    const sourcePlaces = [...nextPlaces, ...preservedUnmatched];
    const envelope = analysis && typeof analysis === "object" && !Array.isArray(analysis) ? analysis : {};
    const factsFound = [...new Set([
      ...stringList(envelope.factsFound),
      ...nextPlaces.flatMap((item) => item.factsFound || []),
    ])].slice(0, 20);
    const missingFields = [...new Set([
      ...stringList(envelope.missingFields),
      ...nextPlaces.flatMap((item) => item.missingFields || []),
    ])].slice(0, 20);
    const candidateCategories = [...new Set(nextPlaces.map((item) => item.category))];
    const resolvedLinkCategory = candidateCategories.length === 1 ? candidateCategories[0] : category;
    const linkSubCategory = nextPlaces.length === 1
      ? nextPlaces[0].subCategory
      : inferSubCategory(`${title} ${text}`, resolvedLinkCategory);
    state.links[linkIndex] = {
      ...state.links[linkIndex],
      title,
      category: resolvedLinkCategory,
      subCategory: linkSubCategory,
      summary: limitedText(envelope.summary, nextPlaces.map((item) => item.summary).filter(Boolean).join("；") || fallback.summary, 2_000),
      status: "已写入Excel",
      readStatus: isText ? "文字已接收" : "成功读取",
      organizedStatus: "已整理",
      candidateCount: sourcePlaces.length,
      factsFound,
      missingFields,
      resultNote: preservedUnmatched.length
        ? `已整理出 ${nextPlaces.length} 个新结果，另保留 ${preservedUnmatched.length} 个含人工结论的旧结果待确认`
        : missingFields.length ? `已整理出 ${nextPlaces.length} 个候选；仍有 ${missingFields.length} 项待核实` : `已完整整理出 ${nextPlaces.length} 个候选并写入 Excel`,
      model: providerLabel(),
      updatedAt: new Date().toISOString(),
      error: "",
    };
    state.places = [...sourcePlaces, ...state.places.filter((item) => item.sourceId !== id)];
    await writeState(state);
    await syncToExcel(state);
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "未知错误";
    const blocked = !isText && /登录|验证码|正文过少|403|401/.test(message);
    await serializeMutation(async () => {
    await updateLink(id, {
      status: blocked ? "需要人工补充" : "处理失败",
      readStatus: isText ? "文字解析失败" : blocked ? "读取受限" : "读取失败",
      organizedStatus: "未整理",
      factsFound: [],
      missingFields: isText ? ["可识别的具体诉求", "分类与偏好条件"] : ["网页正文", "名称", "地址", "价格", "特点与预订信息"],
      resultNote: `未整理：${message}`,
      error: message,
      summary: isText ? "文字需求未能自动整理，原始内容仍保留在投递记录中。" : "网页未能自动读取，系统没有生成未经证实的内容。",
    });
    const state = await readState();
    await syncToExcel(state).catch(() => {});
    });
  }
}

function serializeMutation(task) {
  const run = mutationChain.then(task, task);
  mutationChain = run.catch((error) => {
    console.error("State mutation failed:", error);
  });
  return run;
}

function enqueueLink(id) {
  processingChain = processingChain.then(() => processLink(id)).catch((error) => {
    console.error("Link task failed:", error);
  });
}

const headers = {
  links: ["ID", "来源类型", "提交时间", "提交人", "原始内容", "页面标题 / 需求名称", "原始链接", "大类", "子分类", "读取结果", "整理结果", "处理状态", "生成候选数", "已提取信息", "缺失信息", "AI摘要", "分析模型", "错误原因", "备注"],
  report: ["ID", "来源类型", "提交人", "页面标题 / 需求名称", "原始内容", "大类", "子分类", "读取结果", "整理结果", "生成候选数", "结果说明", "已提取信息", "缺失信息", "错误原因", "原始链接", "更新时间"],
  decisions: ["ID", "记录类型", "核实状态", "大类", "子分类", "名称", "区域 / 位置", "参考价格", "核心规格", "特征标签", "主要亮点", "主要风险", "资料完整度", "缺失信息", "AI推荐分", "AI评分说明", "AI建议时段", "AI候选线索", "AI估价依据", "想去票", "可以票", "不考虑票", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接"],
  requirements: ["ID", "需求名称", "大类", "子分类", "偏好标签", "目标区域", "预算", "期望时长", "需求摘要", "已表达条件", "仍需匹配", "提交人", "团队原始诉求", "处理状态", "人工结论", "是否采用", "人工备注"],
  stays: ["ID", "名称", "子分类", "特征标签", "区域", "详细地址", "地段特点", "核心景点距离", "每晚价格", "两晚总价", "额外费用", "押金", "适合人数", "户型", "房间", "床位", "床型", "卫浴", "环境特点", "是否整租", "厨房", "能否烧烤", "烧烤设备/费用", "早餐", "停车", "交通", "入住时间", "退房时间", "取消政策", "预订要求", "预订状态", "优点", "缺点", "推荐分", "资料完整度", "缺失信息", "证据摘要", "想去票", "可以票", "不考虑票", "投票详情", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接", "数据状态"],
  food: ["ID", "名称", "子分类", "特征标签", "适合安排", "区域", "详细地址", "人均价格", "团队预计总价", "招牌菜", "包间", "团队适合度", "排队情况", "营业时间", "预约要求", "取消政策", "停车", "环境特点", "预订状态", "优点", "缺点", "推荐分", "资料完整度", "缺失信息", "证据摘要", "想去票", "可以票", "不考虑票", "投票详情", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接", "数据状态"],
  escapes: ["ID", "名称", "主题名称", "子分类", "特征标签", "区域", "详细地址", "单人价格", "团队预计总价", "恐怖程度", "难度", "玩法类型", "场地规模", "房间数量", "推荐人数", "最少人数", "最多人数", "团队独立开场", "时长", "NPC/真人互动", "体力消耗", "换装", "营业时间", "预约要求", "取消政策", "停车", "预订状态", "优点", "缺点", "推荐分", "资料完整度", "缺失信息", "证据摘要", "想去票", "可以票", "不考虑票", "投票详情", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接", "数据状态"],
  leisure: ["ID", "名称", "子分类", "特征标签", "区域", "详细地址", "人均/套餐价格", "团队预计总价", "包含设施", "套餐内容", "营业时间", "能否过夜", "是否含餐", "休息区域", "独立房间", "男女分区", "适合人数", "使用限制", "环境特点", "停车", "预约要求", "取消政策", "预订状态", "优点", "缺点", "推荐分", "资料完整度", "缺失信息", "证据摘要", "想去票", "可以票", "不考虑票", "投票详情", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接", "数据状态"],
  attractions: ["ID", "名称", "子分类", "特征标签", "区域", "详细地址", "票价", "票价说明", "营业时间", "建议时长", "室内/室外", "天气影响", "预约要求", "取消政策", "停车", "预订状态", "核心看点", "注意事项", "推荐分", "资料完整度", "缺失信息", "证据摘要", "想去票", "可以票", "不考虑票", "投票详情", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接", "数据状态"],
  guides: ["ID", "标题", "子分类", "特征标签", "涉及区域", "AI摘要", "避坑信息", "已提取信息", "缺失信息", "资料完整度", "证据摘要", "想去票", "可以票", "不考虑票", "投票详情", "人工结论", "是否入选", "人工备注", "人工保护字段", "来源类型", "团队原始诉求", "原始链接", "数据状态"],
  itinerary: ["日期", "开始时间", "结束时间", "类型", "地点", "活动安排", "详细地址", "交通", "预计费用/人", "预订状态", "注意事项", "来源ID", "来源链接"],
  reservations: ["ID", "待办事项", "类型", "计划时间", "当前状态", "负责人", "完成期限", "核对说明"],
  aiReviews: ["审阅ID", "建议ID", "生成时间", "审阅摘要", "建议类型", "修改对象", "字段", "当前内容", "AI建议", "修改理由", "置信度", "依据类型", "处理状态", "应用结果", "分析模型"],
  settings: ["设置项", "当前内容", "说明"],
  overview: ["项目指标", "当前值", "说明"],
};

const sheetDescriptions = {
  "投递汇总": "链接和团队文字诉求的完整台账；失败记录也会保留，不会静默丢失。",
  "处理报告": "快速查看哪些投递已成功整理、哪些未整理，以及缺失了什么。",
  "候选决策台": "决策的唯一修改入口：黄色列会同步到网站。真实地点的“是否入选”表示地点入选；团队需求或攻略的“是”表示线索已采纳，不会自动写入行程。",
  "团队需求": "团队成员写下的愿望和约束，只作为匹配条件；标记“采用”表示确认这条需求，不代表已经选定真实商户。",
  "住宿候选": "民宿完整明细：价格、位置、户型、床位、烧烤、费用和取消政策；黄色列可直接修改。",
  "美食餐饮": "按烧烤、火锅、炒菜、早茶等分类；比较人均、团队总价、包间、排队和招牌菜。",
  "密室候选": "按主题和玩法分类；比较微恐/中恐、难度、规模、人数、价格、NPC 和预约规则。",
  "室内休闲": "按汗蒸、桑拿、洗浴、温泉等分类；比较套餐、设施、过夜、餐食和使用限制。",
  "景点户外": "园林、博物馆、历史街区和户外项目；当前优先确认具体地点、票价、开放时间与天气影响。",
  "攻略文章": "攻略完整明细：摘要、避坑、证据和缺失信息；决策状态请统一到“候选决策台”修改。",
  "三日行程": "济州岛 3 天 2 晚初版安排；日期、航班和动态价格仍需团队确认。",
  "预订清单": "所有需要团队确认或下单的事项；网站不会代替你付款。",
  "AI审阅建议": "DeepSeek 对当前 Excel 与行程快照的审阅记录；建议不会直接覆盖投票、入选结论或人工锁定字段。",
  "项目设置": "本次团队出行的需求约束与运行设置。",
  "项目总览": "当前资料完成度、候选数量和下一步重点。",
};

function columnWidth(header) {
  if (["ID", "来源ID"].includes(header)) return 20;
  if (/原始链接|来源链接/.test(header)) return 36;
  if (/原始内容|原始诉求/.test(header)) return 38;
  if (/摘要|优点|缺点|说明|信息|特点|交通|证据|政策|备注|亮点|风险/.test(header)) return 30;
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

function hasUsefulFact(value) {
  if (typeof value === "number") return Number.isFinite(value);
  const text = String(value ?? "").trim();
  return Boolean(text) && !/待核实|待确认|待补充|未知|暂无|未选择|空位|占位|日期.*确认后|具体.*待/.test(text);
}

function candidateQuality(place, link) {
  const d = place.details || defaultDetails;
  let checks;
  if (place.candidateType === "requirement") {
    checks = [["原始诉求", link?.inputText], ["需求分类", place.category], ["关键偏好", place.featureTags?.[0]], ["需求摘要", place.summary || link?.summary]];
  } else if (place.category === "住宿") {
    checks = [["详细地址", d.address], ["价格", place.price ?? d.twoNightTotal], ["团队人数容量", d.capacity], ["户型", d.roomType], ["房间", d.rooms], ["床位 / 床型", hasUsefulFact(d.beds) ? d.beds : d.bedTypes], ["卫浴", d.bathrooms], ["烧烤", d.barbecue], ["取消政策", d.cancellationPolicy]];
  } else if (place.category === "餐饮") {
    checks = [["子分类", place.subCategory], ["详细地址", d.address], ["人均价格", place.price], ["招牌菜", d.signatureDishes], ["团队适合度", d.groupSuitability], ["营业时间", d.openingHours], ["预约要求", d.reservation]];
  } else if (place.category === "密室") {
    checks = [["主题名称", d.themeName], ["详细地址", d.address], ["单人价格", place.price], ["恐怖程度", d.horrorLevel], ["难度", d.difficulty], ["场地规模", d.venueSize], ["适合人数", d.capacity], ["团队独立开场", d.sixPersonSession], ["时长", place.duration], ["预约要求", d.reservation]];
  } else if (place.category === "休闲娱乐") {
    checks = [["子分类", place.subCategory], ["详细地址", d.address], ["价格", place.price], ["包含设施", d.leisureFacilities], ["营业时间", d.openingHours], ["能否过夜", d.overnight], ["适合人数", d.capacity], ["使用限制", d.serviceRestrictions]];
  } else if (place.category === "景点") {
    checks = [["子分类", place.subCategory], ["详细地址", d.address], ["票价", hasUsefulFact(d.ticketInfo) ? d.ticketInfo : place.price], ["开放时间", d.openingHours], ["建议时长", hasUsefulFact(d.recommendedDuration) ? d.recommendedDuration : place.duration], ["预约要求", d.reservation], ["天气影响", d.weatherImpact]];
  } else {
    checks = [["涉及区域", place.area], ["摘要", link?.summary], ["主要亮点", place.pros?.[0]], ["主要风险", place.cons?.[0]]];
  }
  const missing = checks.filter(([, value]) => !hasUsefulFact(value)).map(([label]) => label);
  const ratio = checks.length ? (checks.length - missing.length) / checks.length : 0;
  return { ratio, percent: Math.round(ratio * 100), missing };
}

function candidateCompletenessRatio(place, quality) {
  return place.aiCompleteness !== null && place.aiCompleteness !== undefined && place.aiCompleteness !== "" && Number.isFinite(Number(place.aiCompleteness))
    ? Number(place.aiCompleteness) / 100
    : quality.ratio;
}

function placePriceSummary(place) {
  if (place.category === "住宿" && hasUsefulFact(place.details?.twoNightTotal)) return String(place.details.twoNightTotal);
  if (["餐饮", "密室", "休闲娱乐"].includes(place.category) && hasUsefulFact(place.details?.sixPersonTotal)) return String(place.details.sixPersonTotal);
  if (place.category === "景点" && hasUsefulFact(place.details?.ticketInfo)) return String(place.details.ticketInfo);
  if (Number.isFinite(place.price)) return `₩${place.price}`;
  return place.priceLabel || "待核实";
}

function placeCoreSpec(place) {
  const d = place.details || defaultDetails;
  let parts;
  if (place.category === "住宿") parts = [d.capacity, d.roomType, d.rooms, d.beds, `烧烤：${d.barbecue}`];
  else if (place.category === "餐饮") parts = [place.subCategory, d.signatureDishes, d.privateRoom, d.groupSuitability];
  else if (place.category === "密室") parts = [d.horrorLevel, d.difficulty, d.venueSize, d.sixPersonSession];
  else if (place.category === "休闲娱乐") parts = [place.subCategory, d.leisureFacilities, d.overnight, d.includedMeals];
  else if (place.category === "景点") parts = [place.subCategory, d.recommendedDuration || place.duration, d.indoorOutdoor, d.weatherImpact];
  else parts = place.tags || [];
  const useful = parts.filter(hasUsefulFact).slice(0, 4);
  return useful.length ? useful.join("；") : "关键规格待核实";
}

function ensureSheet(workbook, name, sheetHeaders) {
  let sheet = workbook.getWorksheet(name);
  const created = !sheet;
  if (!sheet) {
    sheet = workbook.addWorksheet(name, { views: [{ showGridLines: false, state: "frozen", xSplit: 3, ySplit: 3 }] });
  }
  if (sheet.getCell(1, 1).isMerged) sheet.unMergeCells(1, 1, 1, 1);
  if (sheet.getCell(2, 1).isMerged) sheet.unMergeCells(2, 1, 2, 1);
  sheet.mergeCells(1, 1, 1, sheetHeaders.length);
  sheet.mergeCells(2, 1, 2, sheetHeaders.length);
  sheet.views = [{ showGridLines: false, state: "frozen", xSplit: 3, ySplit: 3 }];
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
  const editableColumns = ["核实状态", "人工结论", "是否入选", "是否采用", "人工备注"];
  for (let rowNumber = 4; rowNumber < 4 + Math.max(rowCount, 1); rowNumber += 1) {
    for (const header of editableColumns) {
      const column = sheetHeaders.indexOf(header) + 1;
      if (!column) continue;
      const cell = sheet.getCell(rowNumber, column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4D6" } };
      cell.font = { name: "PingFang SC", color: { argb: "8A542C" }, size: 9, bold: header !== "人工备注" };
      cell.alignment = { vertical: "middle", horizontal: header === "人工备注" ? "left" : "center", wrapText: true };
      if (["是否入选", "是否采用"].includes(header)) cell.dataValidation = { type: "list", allowBlank: false, formulae: ['"是,否"'] };
      if (header === "人工结论") cell.dataValidation = { type: "list", allowBlank: false, formulae: ['"待比较,备选,拟定,淘汰"'] };
      if (header === "核实状态") cell.dataValidation = { type: "list", allowBlank: false, formulae: ['"待核实,部分核实,已核实,团队诉求"'] };
    }
  }
}

function styleEditableFields(sheet, sheetHeaders, rowCount, editableHeaders) {
  const columns = editableHeaders.map((header) => sheetHeaders.indexOf(header) + 1).filter(Boolean);
  for (let rowNumber = 4; rowNumber < 4 + rowCount; rowNumber += 1) {
    for (const column of columns) {
      const cell = sheet.getCell(rowNumber, column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9E8" } };
      cell.font = { name: "PingFang SC", color: { argb: "5F553E" }, size: 9 };
    }
  }
}

function formatCandidateColumns(sheet, sheetHeaders, rowCount) {
  for (const header of ["每晚价格", "价格", "人均价格", "单人价格", "人均/套餐价格", "票价"]) {
    const column = sheetHeaders.indexOf(header) + 1;
    if (column) sheet.getColumn(column).numFmt = "₩#,##0";
  }
  const scoreColumn = sheetHeaders.indexOf("推荐分") + 1 || sheetHeaders.indexOf("AI推荐分") + 1;
  if (scoreColumn) sheet.getColumn(scoreColumn).numFmt = "0.0";
  const completenessColumn = sheetHeaders.indexOf("资料完整度") + 1;
  if (completenessColumn) sheet.getColumn(completenessColumn).numFmt = "0%";
  for (let rowNumber = 4; rowNumber < 4 + rowCount; rowNumber += 1) {
    if (completenessColumn) {
      const cell = sheet.getCell(rowNumber, completenessColumn);
      const ratio = Number(cell.value);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ratio >= 0.75 ? "DCECE4" : ratio >= 0.5 ? "FFF4D6" : "F5DFCD" } };
      cell.font = { name: "PingFang SC", color: { argb: ratio >= 0.75 ? "23634C" : "8A542C" }, size: 9, bold: true };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    }
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
  sheet.getCell("A1").value = `${fp.destination}旅行 · 最终方案首页`;
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
    ["行程结构", fp.schedule, "例如：3 天 2 晚 · 国际航班直达济州"],
    ["同行人数", fp.people, "用于住宿、餐饮和活动人数判断"],
    ["住宿晚数", fp.nights, "当前按 3 天 2 晚设计"],
    ["人均预算", fp.perPersonBudget, "可填写数字或预算区间"],
    ["往返机票 / 人", fp.roundTripFlightPerPerson, "填写含税、行李额后的每人实际往返票价"],
    ["方案说明", fp.summary, "网站首页的行程摘要"],
  ];
  const stayFields = [
    ["民宿名称", stay.name, "确定住宿后填写真实名称"],
    ["详细地址", stay.address, "尽量填写完整门牌或平台可见地址"],
    ["适合人数", stay.capacity, `确认房源允许 ${fp.people} 人入住`],
    ["房间 / 床位", stay.roomsBeds, "写清房间数、床型和床数"],
    ["两晚总价", stay.twoNightTotal, "填写含清洁费、服务费后的总价"],
    ["入住 / 退房", stay.checkInOut, "填写具体时间和延迟入住限制"],
    ["能否烧烤", stay.barbecue, "非硬性条件，需要时再核对"],
    ["烧烤设备 / 费用", stay.bbqEquipment, "烤炉、炭火、食材、清洁费与限制"],
    ["早餐安排", stay.breakfast, "第 2、3 天早餐的具体门店"],
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
    sheet.getRow(row).height = label === "方案说明" ? 70 : label === "烧烤设备 / 费用" || label === "早餐安排" ? 46 : 34;
  }

  section(5, "一、基础信息", "网站顶部与旅行概览");
  basics.forEach((item, index) => fieldRow(6 + index, ...item));
  section(15, "二、住宿确认", "民宿未确定前保持橙色“待补充”");
  stayFields.forEach((item, index) => fieldRow(16 + index, ...item));

  section(27, "三、三日行程", "直接增删或修改下面的行程，网站按这里显示");
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
    row.values = [item.day, item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost || null, item.bookingStatus, item.sourceUrl || null, item.note || null];
    row.height = 42;
    for (let column = 1; column <= 12; column += 1) {
      const cell = row.getCell(column);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: index % 2 ? "FFF9EF" : "FFF4D6" } };
      cell.font = { name: "PingFang SC", color: { argb: column === 11 ? "1B5E7A" : "173D3B" }, size: 9 };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "hair", color: { argb: "D7DED8" } } };
    }
    row.getCell(9).numFmt = "#,##0";
    row.getCell(10).dataValidation = { type: "list", allowBlank: true, formulae: ['"待确认,未预订,已预订,已完成,无需预订,无需预约,出发前复核,现场购票"'] };
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
    workbook.creator = "出行共创台";
    workbook.created = new Date();
    writeFinalPlanSheet(workbook, state);

    const decisionPlaces = [...state.places].sort((left, right) => {
      const statusWeight = { 拟定: 4, 备选: 3, 待比较: 2, 淘汰: 0 };
      return Number(right.selected) - Number(left.selected)
        || (statusWeight[right.decisionStatus] || 0) - (statusWeight[left.decisionStatus] || 0)
        || (Number(right.score) || 0) - (Number(left.score) || 0);
    });
    const decisionSheet = ensureSheet(workbook, "候选决策台", headers.decisions);
    decisionSheet.properties.tabColor = { argb: "D9703E" };
    replaceRows(decisionSheet, decisionPlaces.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      const missing = [...new Set([...(link?.missingFields || []), ...quality.missing])];
      const [support, okay, reject] = placeVoteColumns(item);
      const leads = (item.candidateLeads || []).map((lead) => `${lead.name}（${lead.area}，${lead.confidence}置信）`).join("；");
      const priceBasis = [item.aiPriceBasis, item.aiPriceConfidence ? `置信度：${item.aiPriceConfidence}` : ""].filter(Boolean).join("；");
      return [item.id, candidateTypeLabel(item.candidateType), item.verificationStatus, item.category, item.subCategory, item.name, item.area, placePriceSummary(item), placeCoreSpec(item), item.featureTags.join("；"), item.pros[0] || "等待整理", item.cons[0] || "等待核实", candidateCompletenessRatio(item, quality), missing.join("；"), item.score, item.aiScoreReason || "等待说明", (item.planningSuggestions || []).join("；"), leads, priceBasis, support, okay, reject, item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl];
    }), headers.decisions.length);
    styleCandidateDecisions(decisionSheet, headers.decisions, decisionPlaces.length);
    formatCandidateColumns(decisionSheet, headers.decisions, decisionPlaces.length);
    [20, 12, 12, 10, 16, 24, 18, 16, 30, 24, 24, 24, 13, 28, 12, 30, 28, 32, 30, 10, 10, 10, 12, 10, 28, 22, 12, 38, 36]
      .forEach((width, index) => { decisionSheet.getColumn(index + 1).width = width; });
    decisionSheet.views = [{ showGridLines: false, state: "frozen", xSplit: 6, ySplit: 3 }];

    const linkSheet = ensureSheet(workbook, "投递汇总", headers.links);
    replaceRows(linkSheet, state.links.map((item) => [item.id, item.sourceType || "链接", item.createdAt, item.submitter, item.inputText || item.url, item.title, item.url, item.category, item.subCategory || "等待识别", item.readStatus, item.organizedStatus, item.status, item.candidateCount || 0, item.factsFound.join("；"), item.missingFields.join("；"), item.summary, item.model, item.error || "", item.note]), headers.links.length);

    const reportSheet = ensureSheet(workbook, "处理报告", headers.report);
    replaceRows(reportSheet, state.links.map((item) => [item.id, item.sourceType || "链接", item.submitter, item.title || (item.sourceType === "文字" ? "需求名称待生成" : "标题未读取"), item.inputText || "", item.category, item.subCategory || "等待识别", item.readStatus, item.organizedStatus, item.candidateCount || 0, item.resultNote, item.factsFound.join("；"), item.missingFields.join("；"), item.error || "", item.url, item.updatedAt]), headers.report.length);

    const reviewSheet = ensureSheet(workbook, "AI审阅建议", headers.aiReviews);
    reviewSheet.properties.tabColor = { argb: "2F62FF" };
    const reviewRows = (state.aiReviews || []).flatMap((review) => review.suggestions.map((suggestion) => [
      review.id,
      suggestion.id,
      review.generatedAt,
      `${review.headline}｜${review.summary}`,
      suggestion.category,
      suggestion.targetLabel,
      suggestion.field || "仅供参考",
      suggestion.currentValue,
      Array.isArray(suggestion.proposedValue) ? suggestion.proposedValue.join("；") : suggestion.proposedValue,
      suggestion.reason,
      suggestion.confidence,
      suggestion.evidenceType,
      suggestion.status,
      suggestion.appliedResult,
      review.provider,
    ]));
    replaceRows(reviewSheet, reviewRows, headers.aiReviews.length);
    [24, 24, 20, 42, 14, 24, 18, 34, 38, 42, 10, 16, 12, 34, 22]
      .forEach((width, index) => { reviewSheet.getColumn(index + 1).width = width; });
    reviewSheet.views = [{ showGridLines: false, state: "frozen", xSplit: 3, ySplit: 3 }];

    const requirements = state.places.filter((item) => item.candidateType === "requirement");
    const requirementSheet = ensureSheet(workbook, "团队需求", headers.requirements);
    replaceRows(requirementSheet, requirements.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      return [item.id, item.name, item.category, item.subCategory, item.featureTags.join("；"), item.area, placePriceSummary(item), item.duration, item.summary || link?.summary || "", (item.factsFound || link?.factsFound || []).join("；"), (item.missingFields || link?.missingFields || []).join("；"), link?.submitter || "团队成员", link?.inputText || "", link?.resultNote || "", item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || ""];
    }), headers.requirements.length);
    styleEditableFields(requirementSheet, headers.requirements, requirements.length, ["需求名称", "大类", "子分类", "偏好标签", "目标区域", "预算", "期望时长", "需求摘要", "已表达条件", "仍需匹配"]);
    formatCandidateColumns(requirementSheet, headers.requirements, requirements.length);

    const realPlaces = state.places.filter((item) => item.candidateType === "place");
    const stays = realPlaces.filter((item) => item.category === "住宿");
    const staySheet = ensureSheet(workbook, "住宿候选", headers.stays);
    replaceRows(staySheet, stays.map((item) => {
      const d = item.details;
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      const missing = [...new Set([...(link?.missingFields || []), ...quality.missing])].join("；");
      return [item.id, item.name, item.subCategory, item.featureTags.join("；"), item.area, d.address, d.locationHighlights, d.distanceToCore, item.price, d.twoNightTotal, d.extraFees, d.deposit, d.capacity, d.roomType, d.rooms, d.beds, d.bedTypes, d.bathrooms, d.environment, d.entireRental, d.kitchen, d.barbecue, d.bbqEquipment, d.breakfast, d.parking, d.transport, d.checkIn, d.checkOut, d.cancellationPolicy, d.reservation, d.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, candidateCompletenessRatio(item, quality), missing, d.evidence, ...placeVoteColumns(item), item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl, item.dataStatus];
    }), headers.stays.length);
    styleEditableFields(staySheet, headers.stays, stays.length, ["名称", "子分类", "特征标签", "区域", "详细地址", "每晚价格", "两晚总价", "适合人数", "户型", "房间", "床位", "床型", "卫浴", "能否烧烤", "烧烤设备/费用", "取消政策"]);
    formatCandidateColumns(staySheet, headers.stays, stays.length);

    const food = realPlaces.filter((item) => item.category === "餐饮");
    const foodSheet = ensureSheet(workbook, "美食餐饮", headers.food);
    replaceRows(foodSheet, food.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      const missing = [...new Set([...(link?.missingFields || []), ...quality.missing])].join("；");
      const d = item.details;
      return [item.id, item.name, item.subCategory, item.featureTags.join("；"), d.usage, item.area, d.address, item.price, d.sixPersonTotal, d.signatureDishes, d.privateRoom, d.groupSuitability, d.queueInfo, d.openingHours, d.reservation, d.cancellationPolicy, d.parking, d.environment, d.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, candidateCompletenessRatio(item, quality), missing, d.evidence, ...placeVoteColumns(item), item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl, item.dataStatus];
    }), headers.food.length);
    styleEditableFields(foodSheet, headers.food, food.length, ["名称", "子分类", "特征标签", "适合安排", "区域", "详细地址", "人均价格", "团队预计总价", "招牌菜", "包间", "团队适合度", "排队情况", "营业时间", "预约要求"]);
    formatCandidateColumns(foodSheet, headers.food, food.length);

    const escapes = realPlaces.filter((item) => item.category === "密室");
    const escapeSheet = ensureSheet(workbook, "密室候选", headers.escapes);
    replaceRows(escapeSheet, escapes.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      const missing = [...new Set([...(link?.missingFields || []), ...quality.missing])].join("；");
      const d = item.details;
      return [item.id, item.name, d.themeName, item.subCategory, item.featureTags.join("；"), item.area, d.address, item.price, d.sixPersonTotal, d.horrorLevel, d.difficulty, d.escapeStyle, d.venueSize, d.roomCount, d.capacity, d.minPlayers, d.maxPlayers, d.sixPersonSession, item.duration, d.npcInteraction, d.physicalIntensity, d.costume, d.openingHours, d.reservation, d.cancellationPolicy, d.parking, d.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, candidateCompletenessRatio(item, quality), missing, d.evidence, ...placeVoteColumns(item), item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl, item.dataStatus];
    }), headers.escapes.length);
    styleEditableFields(escapeSheet, headers.escapes, escapes.length, ["名称", "主题名称", "子分类", "特征标签", "区域", "详细地址", "单人价格", "团队预计总价", "恐怖程度", "难度", "玩法类型", "场地规模", "房间数量", "推荐人数", "团队独立开场", "时长", "NPC/真人互动"]);
    formatCandidateColumns(escapeSheet, headers.escapes, escapes.length);

    const leisure = realPlaces.filter((item) => item.category === "休闲娱乐");
    const leisureSheet = ensureSheet(workbook, "室内休闲", headers.leisure);
    replaceRows(leisureSheet, leisure.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      const missing = [...new Set([...(link?.missingFields || []), ...quality.missing])].join("；");
      const d = item.details;
      return [item.id, item.name, item.subCategory, item.featureTags.join("；"), item.area, d.address, item.price, d.sixPersonTotal, d.leisureFacilities, d.packageInfo, d.openingHours, d.overnight, d.includedMeals, d.restArea, d.privateRoom, d.genderArrangement, d.capacity, d.serviceRestrictions, d.environment, d.parking, d.reservation, d.cancellationPolicy, d.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, candidateCompletenessRatio(item, quality), missing, d.evidence, ...placeVoteColumns(item), item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl, item.dataStatus];
    }), headers.leisure.length);
    styleEditableFields(leisureSheet, headers.leisure, leisure.length, ["名称", "子分类", "特征标签", "区域", "详细地址", "人均/套餐价格", "团队预计总价", "包含设施", "套餐内容", "营业时间", "能否过夜", "是否含餐", "休息区域", "独立房间", "男女分区", "适合人数", "使用限制"]);
    formatCandidateColumns(leisureSheet, headers.leisure, leisure.length);

    const attractions = realPlaces.filter((item) => item.category === "景点");
    const attractionSheet = ensureSheet(workbook, "景点户外", headers.attractions);
    replaceRows(attractionSheet, attractions.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      const missing = [...new Set([...(link?.missingFields || []), ...quality.missing])].join("；");
      const d = item.details;
      return [item.id, item.name, item.subCategory, item.featureTags.join("；"), item.area, d.address, item.price, d.ticketInfo, d.openingHours, d.recommendedDuration || item.duration, d.indoorOutdoor, d.weatherImpact, d.reservation, d.cancellationPolicy, d.parking, d.bookingStatus, item.pros.join("；"), item.cons.join("；"), item.score, candidateCompletenessRatio(item, quality), missing, d.evidence, ...placeVoteColumns(item), item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl, item.dataStatus];
    }), headers.attractions.length);
    styleEditableFields(attractionSheet, headers.attractions, attractions.length, ["名称", "子分类", "特征标签", "区域", "详细地址", "票价", "票价说明", "营业时间", "建议时长", "室内/室外", "天气影响", "预约要求"]);
    formatCandidateColumns(attractionSheet, headers.attractions, attractions.length);

    const guides = state.places.filter((item) => item.candidateType === "guide");
    const guideSheet = ensureSheet(workbook, "攻略文章", headers.guides);
    replaceRows(guideSheet, guides.map((item) => {
      const link = state.links.find((candidate) => candidate.id === item.sourceId);
      const quality = candidateQuality(item, link);
      return [item.id, item.name, item.subCategory, item.featureTags.join("；"), item.area, item.summary || link?.summary || item.pros.join("；"), item.cons.join("；"), item.factsFound?.join("；") || link?.factsFound?.join("；") || "", item.missingFields?.join("；") || link?.missingFields?.join("；") || quality.missing.join("；"), candidateCompletenessRatio(item, quality), item.details.evidence, ...placeVoteColumns(item), item.decisionStatus || "待比较", item.selected ? "是" : "否", item.manualNote || "", manualOverrideSummary(item), link?.sourceType || (item.sourceId ? "链接" : "示例"), link?.inputText || "", item.sourceUrl, item.dataStatus];
    }), headers.guides.length);
    styleEditableFields(guideSheet, headers.guides, guides.length, ["标题", "子分类", "特征标签", "涉及区域", "AI摘要", "避坑信息", "证据摘要"]);
    styleEditableFields(guideSheet, headers.guides, guides.length, ["标题", "子分类", "特征标签", "涉及区域"]);
    formatCandidateColumns(guideSheet, headers.guides, guides.length);

    const itinerarySheet = ensureSheet(workbook, "三日行程", headers.itinerary);
    const itineraryRows = [
      ...state.itinerary.day0.map((item) => ["第1天", item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost || "", item.bookingStatus, item.note, item.sourceId, state.places.find((place) => place.id === item.sourceId)?.sourceUrl || ""]),
      ...state.itinerary.day1.map((item) => ["第2天", item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost || "", item.bookingStatus, item.note, item.sourceId, state.places.find((place) => place.id === item.sourceId)?.sourceUrl || ""]),
      ...state.itinerary.day2.map((item) => ["第3天", item.time, item.endTime, item.category, item.title, item.subtitle, item.address, item.transport, item.cost || "", item.bookingStatus, item.note, item.sourceId, state.places.find((place) => place.id === item.sourceId)?.sourceUrl || ""]),
    ];
    replaceRows(itinerarySheet, itineraryRows, headers.itinerary.length);

    const reservationSheet = ensureSheet(workbook, "预订清单", headers.reservations);
    replaceRows(reservationSheet, state.reservations.map((item) => [item.id, item.item, item.type, item.targetTime, item.status, item.owner, item.deadline, item.note]), headers.reservations.length);

    state.settings.provider = providerLabel();
    const settingsSheet = recreateSheet(workbook, "项目设置", headers.settings);
    replaceRows(settingsSheet, [
      ["目的地", state.project.destination, `本次主题为${state.project.destination}`],
      ["出行结构", state.tripProfile.schedule, "具体日期待团队确认"],
      ["同行人数", state.tripProfile.groupSize, `按 ${state.tripProfile.groupSize} 人统一比较住宿与活动`],
      ["住宿晚数", state.tripProfile.nights, "3 天行程当前按 2 晚设计"],
      ["住宿偏好", state.tripProfile.stayPreference, "重点检查房间、床位、卫浴与环境"],
      ["住宿预算", state.tripProfile.accommodationBudget, "确认日期后再定预算"],
      ["住宿附加需求", state.tripProfile.barbecue, "如需要烧烤，订房前确认规则"],
      ["早餐", state.tripProfile.breakfasts.join("；"), "两顿早餐待团队补充具体店铺"],
      ["分类体系", "住宿 / 餐饮 / 密室 / 休闲娱乐 / 景点 / 攻略", "每类使用不同的提取字段与完整度标准"],
      ["团队活动", state.tripProfile.activity, "密室、汗蒸、桑拿、洗浴或同类活动"],
      ["AI 分析模型", providerLabel(), "所有 AI 结论仍需人工核实"],
      ["访问方式", "本地后台 + 密码保护的公网网址", "朋友无需账号、无需同一 Wi-Fi"],
    ], headers.settings.length);

    const completedLinks = state.links.filter((item) => item.organizedStatus === "已整理").length;
    const unorganizedLinks = state.links.filter((item) => item.organizedStatus === "未整理").length;
    const overviewSheet = recreateSheet(workbook, "项目总览", headers.overview);
    replaceRows(overviewSheet, [
      ["旅行框架", state.tripProfile.schedule, `${state.tripProfile.groupSize} 人，住宿 ${state.tripProfile.nights} 晚`],
      ["已收集投递", state.links.length, `链接 ${state.links.filter((item) => item.sourceType !== "文字").length} · 文字 ${state.links.filter((item) => item.sourceType === "文字").length}`],
      ["已整理投递", completedLinks, "已形成候选资料并写入分类表"],
      ["未整理投递", unorganizedLinks, "请在处理报告查看原因并补充信息"],
      ["候选资料", state.places.length, `真实候选 ${state.places.filter((item) => item.candidateType === "place").length} · 团队需求 ${state.places.filter((item) => item.candidateType === "requirement").length} · 攻略 ${state.places.filter((item) => item.candidateType === "guide").length}`],
      ["分类明细", supportedCategories.map((category) => `${category} ${state.places.filter((item) => item.category === category && item.candidateType !== "requirement").length}`).join(" · "), "团队诉求请查看“团队需求”工作表"],
      ["已入选地点", state.places.filter((item) => item.candidateType === "place" && item.selected).length, "真实地点，可用于当前行程草案"],
      ["已采纳需求 / 攻略", state.places.filter((item) => item.candidateType !== "place" && item.selected).length, "筛选条件或来源线索，不会自动写入行程"],
      ["待确认预订", state.reservations.filter((item) => !/已完成|已预订/.test(item.status)).length, "付款前由团队最终确认"],
      ["当前重点", "确定日期与直达航班、住宿、公交班次和具体餐厅", "可投递链接，也可直接写文字诉求"],
    ], headers.overview.length);

    for (const sheet of workbook.worksheets) {
      sheet.pageSetup = { ...(sheet.pageSetup || {}), orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };
      sheet.headerFooter = { ...(sheet.headerFooter || {}), oddFooter: `&L${state.project.destination}出行共创&R第 &P / &N 页` };
    }
    state.settings.workbookPath = path.relative(rootDir, workbookPath);
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

function preservedItinerarySourceId(existingItems, importedItem) {
  const items = Array.isArray(existingItems) ? existingItems : [];
  const sourceUrl = String(importedItem?.sourceUrl || "").trim();
  const exact = items.find((item) => normalizePlanDay(item.day) === normalizePlanDay(importedItem?.day)
    && String(item.title || "").trim() === String(importedItem?.title || "").trim()
    && String(item.time || "").trim() === String(importedItem?.time || "").trim());
  if (exact?.sourceId) return exact.sourceId;
  if (sourceUrl) return items.find((item) => String(item.sourceUrl || "").trim() === sourceUrl && item.sourceId)?.sourceId || "";
  return "";
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
      const importedItem = {
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
      };
      itinerary.push({ ...importedItem, sourceId: preservedItinerarySourceId(existing.itinerary, importedItem) });
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
    roundTripFlightPerPerson: field("往返机票 / 人", existing.roundTripFlightPerPerson || "待填写实际含税票价（含托运行李）"),
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
  const importedBudget = parseMoneyRange(state.finalPlan.perPersonBudget);
  if (importedBudget) state.project.budget = importedBudget.max;
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
    { sheetName: "住宿候选", category: "住宿", priceColumn: "每晚价格", detailColumns: { address: "详细地址", locationHighlights: "地段特点", distanceToCore: "核心景点距离", capacity: "适合人数", roomType: "户型", rooms: "房间", beds: "床位", bedTypes: "床型", bathrooms: "卫浴", twoNightTotal: "两晚总价", extraFees: "额外费用", deposit: "押金", environment: "环境特点", entireRental: "是否整租", kitchen: "厨房", barbecue: "能否烧烤", bbqEquipment: "烧烤设备/费用", breakfast: "早餐", parking: "停车", transport: "交通", checkIn: "入住时间", checkOut: "退房时间", cancellationPolicy: "取消政策", reservation: "预订要求", bookingStatus: "预订状态", evidence: "证据摘要" } },
    { sheetName: "美食餐饮", category: "餐饮", priceColumn: "人均价格", detailColumns: { usage: "适合安排", address: "详细地址", sixPersonTotal: "团队预计总价", signatureDishes: "招牌菜", privateRoom: "包间", groupSuitability: "团队适合度", queueInfo: "排队情况", openingHours: "营业时间", reservation: "预约要求", cancellationPolicy: "取消政策", parking: "停车", environment: "环境特点", bookingStatus: "预订状态", evidence: "证据摘要" } },
    { sheetName: "密室候选", category: "密室", priceColumn: "单人价格", durationColumn: "时长", detailColumns: { themeName: "主题名称", address: "详细地址", sixPersonTotal: "团队预计总价", horrorLevel: "恐怖程度", difficulty: "难度", escapeStyle: "玩法类型", venueSize: "场地规模", roomCount: "房间数量", capacity: "推荐人数", minPlayers: "最少人数", maxPlayers: "最多人数", sixPersonSession: "团队独立开场", npcInteraction: "NPC/真人互动", physicalIntensity: "体力消耗", costume: "换装", openingHours: "营业时间", reservation: "预约要求", cancellationPolicy: "取消政策", parking: "停车", bookingStatus: "预订状态", evidence: "证据摘要" } },
    { sheetName: "室内休闲", category: "休闲娱乐", priceColumn: "人均/套餐价格", detailColumns: { address: "详细地址", sixPersonTotal: "团队预计总价", leisureFacilities: "包含设施", packageInfo: "套餐内容", openingHours: "营业时间", overnight: "能否过夜", includedMeals: "是否含餐", restArea: "休息区域", privateRoom: "独立房间", genderArrangement: "男女分区", capacity: "适合人数", serviceRestrictions: "使用限制", environment: "环境特点", parking: "停车", reservation: "预约要求", cancellationPolicy: "取消政策", bookingStatus: "预订状态", evidence: "证据摘要" } },
    { sheetName: "景点户外", category: "景点", priceColumn: "票价", detailColumns: { address: "详细地址", ticketInfo: "票价说明", openingHours: "营业时间", recommendedDuration: "建议时长", indoorOutdoor: "室内/室外", weatherImpact: "天气影响", reservation: "预约要求", cancellationPolicy: "取消政策", parking: "停车", bookingStatus: "预订状态", evidence: "证据摘要" } },
  ];
  for (const { sheetName, category, priceColumn, durationColumn, detailColumns } of configs) {
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
      const rawPrice = value(row, map, priceColumn);
      const numericPrice = rawPrice === "" || rawPrice === null || rawPrice === undefined ? Number.NaN : Number(rawPrice);
      const details = { ...defaultDetails, ...(current.details || {}) };
      const manualOverrides = { ...(current.manualOverrides || {}) };
      for (const [key, columnName] of Object.entries(detailColumns)) {
        const edited = String(value(row, map, columnName) || "").trim();
        if (!edited) continue;
        rememberManualOverride(manualOverrides, `details.${key}`, edited, current.details?.[key]);
        details[key] = edited;
      }
      const editedName = String(value(row, map, "名称") || current.name).trim();
      const editedArea = String(value(row, map, "区域") || current.area).trim();
      const editedSubCategory = String(value(row, map, "子分类") || current.subCategory).trim();
      const editedFeatureTags = String(value(row, map, "特征标签") || current.featureTags.join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedPros = String(value(row, map, "优点") || value(row, map, "核心看点") || current.pros.join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedCons = String(value(row, map, "缺点") || value(row, map, "注意事项") || current.cons.join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedDuration = durationColumn ? String(value(row, map, durationColumn) || current.duration).trim() : current.duration;
      const nextPrice = Number.isFinite(numericPrice) && numericPrice >= 0 ? numericPrice : current.price;
      rememberManualOverride(manualOverrides, "name", editedName, current.name);
      rememberManualOverride(manualOverrides, "area", editedArea, current.area);
      rememberManualOverride(manualOverrides, "subCategory", editedSubCategory, current.subCategory);
      rememberManualOverride(manualOverrides, "featureTags", editedFeatureTags, current.featureTags);
      rememberManualOverride(manualOverrides, "pros", editedPros, current.pros);
      rememberManualOverride(manualOverrides, "cons", editedCons, current.cons);
      rememberManualOverride(manualOverrides, "duration", editedDuration, current.duration);
      rememberManualOverride(manualOverrides, "price", nextPrice, current.price);
      state.places[index] = {
        ...current,
        name: editedName,
        category,
        subCategory: editedSubCategory,
        featureTags: editedFeatureTags,
        area: editedArea,
        price: nextPrice,
        priceLabel: Number.isFinite(nextPrice) ? `参考 ₩${nextPrice}` : current.priceLabel,
        duration: editedDuration,
        score: Number(value(row, map, "推荐分")) || current.score,
        pros: editedPros,
        cons: editedCons,
        dataStatus: String(value(row, map, "数据状态") || current.dataStatus),
        manualOverrides,
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
      const current = state.places[index];
      const manualOverrides = { ...(current.manualOverrides || {}) };
      const editedName = String(value(row, map, "标题") || current.name).trim();
      const editedArea = String(value(row, map, "涉及区域") || current.area).trim();
      const editedSubCategory = String(value(row, map, "子分类") || current.subCategory).trim();
      const editedFeatureTags = String(value(row, map, "特征标签") || current.featureTags.join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedSummary = String(value(row, map, "AI摘要") || current.summary || "").trim();
      const editedCons = String(value(row, map, "避坑信息") || current.cons.join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedEvidence = String(value(row, map, "证据摘要") || current.details?.evidence || "").trim();
      rememberManualOverride(manualOverrides, "name", editedName, current.name);
      rememberManualOverride(manualOverrides, "area", editedArea, current.area);
      rememberManualOverride(manualOverrides, "subCategory", editedSubCategory, current.subCategory);
      rememberManualOverride(manualOverrides, "featureTags", editedFeatureTags, current.featureTags);
      rememberManualOverride(manualOverrides, "summary", editedSummary, current.summary);
      rememberManualOverride(manualOverrides, "cons", editedCons, current.cons);
      rememberManualOverride(manualOverrides, "details.evidence", editedEvidence, current.details?.evidence);
      state.places[index] = {
        ...current,
        name: editedName,
        area: editedArea,
        subCategory: editedSubCategory,
        featureTags: editedFeatureTags,
        summary: editedSummary,
        cons: editedCons,
        details: { ...current.details, evidence: editedEvidence || current.details?.evidence || "待核实" },
        manualOverrides,
      };
    });
  }

  const requirementSheet = workbook.getWorksheet("团队需求");
  if (requirementSheet) {
    const map = headerIndex(requirementSheet);
    requirementSheet.eachRow((row, number) => {
      if (number < 4) return;
      const id = String(value(row, map, "ID") || "").trim();
      if (!id) return;
      const index = state.places.findIndex((item) => item.id === id && item.candidateType === "requirement");
      if (index < 0) return;
      const current = state.places[index];
      const manualOverrides = { ...(current.manualOverrides || {}) };
      const editedName = String(value(row, map, "需求名称") || current.name).trim();
      const editedCategory = normalizeCategory(value(row, map, "大类") || current.category, editedName);
      const editedSubCategory = String(value(row, map, "子分类") || current.subCategory).trim();
      const editedFeatureTags = String(value(row, map, "偏好标签") || current.featureTags.join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedArea = String(value(row, map, "目标区域") || current.area).trim();
      const editedPriceLabel = String(value(row, map, "预算") || current.priceLabel).trim();
      const editedDuration = String(value(row, map, "期望时长") || current.duration).trim();
      const editedSummary = String(value(row, map, "需求摘要") || current.summary || "").trim();
      const editedFacts = String(value(row, map, "已表达条件") || (current.factsFound || []).join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      const editedMissing = String(value(row, map, "仍需匹配") || (current.missingFields || []).join("；")).split(/[；;]/).map((item) => item.trim()).filter(Boolean);
      for (const [key, edited, previous] of [
        ["name", editedName, current.name], ["category", editedCategory, current.category], ["subCategory", editedSubCategory, current.subCategory],
        ["featureTags", editedFeatureTags, current.featureTags], ["area", editedArea, current.area], ["priceLabel", editedPriceLabel, current.priceLabel],
        ["duration", editedDuration, current.duration], ["summary", editedSummary, current.summary], ["factsFound", editedFacts, current.factsFound],
        ["missingFields", editedMissing, current.missingFields],
      ]) rememberManualOverride(manualOverrides, key, edited, previous);
      state.places[index] = {
        ...current,
        name: editedName,
        category: editedCategory,
        subCategory: validSubCategory(editedCategory, editedSubCategory) ? editedSubCategory : inferSubCategory(`${editedName} ${editedSubCategory}`, editedCategory),
        featureTags: editedFeatureTags,
        area: editedArea,
        priceLabel: editedPriceLabel,
        duration: editedDuration,
        summary: editedSummary,
        factsFound: editedFacts,
        missingFields: editedMissing,
        manualOverrides,
      };
    });
  }

  const decisionSheet = workbook.getWorksheet("候选决策台");
  if (decisionSheet) {
    const map = headerIndex(decisionSheet);
    decisionSheet.eachRow((row, number) => {
      if (number < 4) return;
      const id = String(value(row, map, "ID") || "").trim();
      if (!id) return;
      const index = state.places.findIndex((item) => item.id === id);
      if (index < 0) return;
      const current = state.places[index];
      const decisionStatus = map.has("人工结论") ? String(value(row, map, "人工结论") || current.decisionStatus || "待比较").trim() : current.decisionStatus;
      const verificationStatus = map.has("核实状态") ? String(value(row, map, "核实状态") || current.verificationStatus || "待核实").trim() : current.verificationStatus;
      const normalizedVerification = normalizeVerificationStatus(verificationStatus, current.candidateType, current.dataStatus);
      const manualOverrides = { ...(current.manualOverrides || {}) };
      if (map.has("核实状态")) rememberManualOverride(manualOverrides, "verificationStatus", normalizedVerification, current.verificationStatus);
      state.places[index] = {
        ...current,
        decisionStatus: ["待比较", "备选", "拟定", "淘汰"].includes(decisionStatus) ? decisionStatus : current.decisionStatus || "待比较",
        verificationStatus: normalizedVerification,
        selected: map.has("是否入选") ? String(value(row, map, "是否入选")).trim() === "是" : current.selected,
        manualNote: map.has("人工备注") ? String(value(row, map, "人工备注") || "") : current.manualNote,
        manualOverrides,
      };
    });
  }

  const itinerarySheet = workbook.getWorksheet("三日行程") || workbook.getWorksheet("两日行程");
  if (itinerarySheet && !hasFinalPlanHome) {
    const map = headerIndex(itinerarySheet);
    const imported = { day0: [], day1: [], day2: [] };
    itinerarySheet.eachRow((row, number) => {
      if (number < 4) return;
      const dayLabel = String(value(row, map, "日期") || "").trim();
      const title = String(value(row, map, "地点") || "").trim();
      if (!dayLabel || !title) return;
      const dayKey = /第\s*1\s*天|周五|Day\s*1/i.test(dayLabel) ? "day0" : /第\s*2\s*天|周六|Day\s*2/i.test(dayLabel) ? "day1" : "day2";
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

function sanitizeCandidatePatch(input, current) {
  const body = input && typeof input === "object" ? input : {};
  const next = { ...current, details: { ...defaultDetails, ...(current.details || {}) } };
  const manualOverrides = { ...(current.manualOverrides || {}) };
  const textFields = {
    name: 160, area: 200, priceLabel: 120, duration: 120, summary: 1_000,
    dataStatus: 200, decisionStatus: 40, manualNote: 1_000,
  };
  for (const [key, maximum] of Object.entries(textFields)) {
    if (!Object.prototype.hasOwnProperty.call(body, key)) continue;
    const edited = limitedText(body[key], "", maximum);
    if (key === "decisionStatus" && !["待比较", "备选", "拟定", "淘汰"].includes(edited)) continue;
    rememberManualOverride(manualOverrides, key, edited, current[key]);
    next[key] = edited;
  }
  for (const key of ["featureTags", "tags", "pros", "cons", "factsFound", "missingFields"]) {
    if (!Object.prototype.hasOwnProperty.call(body, key)) continue;
    const maximum = ["pros", "cons"].includes(key) ? 4 : 12;
    const edited = stringList(body[key]).map((item) => limitedText(item, "", 160)).slice(0, maximum);
    rememberManualOverride(manualOverrides, key, edited, current[key]);
    next[key] = edited;
  }
  if (Object.prototype.hasOwnProperty.call(body, "category")) {
    const editedCategory = normalizeCategory(body.category, next.name);
    rememberManualOverride(manualOverrides, "category", editedCategory, current.category);
    next.category = editedCategory;
  }
  if (Object.prototype.hasOwnProperty.call(body, "candidateType")) {
    const editedType = current.candidateType === "requirement"
      ? "requirement"
      : normalizeCandidateType(body.candidateType, { category: next.category, dataStatus: next.dataStatus });
    rememberManualOverride(manualOverrides, "candidateType", editedType, current.candidateType);
    next.candidateType = editedType;
  }
  if (next.candidateType !== "requirement" && next.category === "攻略") next.candidateType = "guide";
  if (next.candidateType === "guide") next.category = "攻略";
  if (Object.prototype.hasOwnProperty.call(body, "subCategory")) {
    const proposed = limitedText(body.subCategory, "", 80);
    const editedSubCategory = validSubCategory(next.category, proposed) ? proposed : inferSubCategory(`${next.name} ${proposed}`, next.category);
    rememberManualOverride(manualOverrides, "subCategory", editedSubCategory, current.subCategory);
    next.subCategory = editedSubCategory;
  } else if (!validSubCategory(next.category, next.subCategory)) {
    next.subCategory = inferSubCategory(next.name, next.category);
  }
  if (Object.prototype.hasOwnProperty.call(body, "price")) {
    const numeric = body.price === null || body.price === "" ? null : Number(body.price);
    if (numeric === null || (Number.isFinite(numeric) && numeric >= 0 && numeric <= 1_000_000)) {
      rememberManualOverride(manualOverrides, "price", numeric, current.price);
      next.price = numeric;
    }
  }
  if (Object.prototype.hasOwnProperty.call(body, "score")) {
    const editedScore = Math.max(0, Math.min(5, Number(body.score) || 0));
    rememberManualOverride(manualOverrides, "score", editedScore, current.score);
    next.score = editedScore;
  }
  if (Object.prototype.hasOwnProperty.call(body, "selected")) next.selected = Boolean(body.selected);
  if (Object.prototype.hasOwnProperty.call(body, "sourceUrl")) {
    const editedUrl = safeWebUrl(body.sourceUrl);
    rememberManualOverride(manualOverrides, "sourceUrl", editedUrl, current.sourceUrl);
    next.sourceUrl = editedUrl;
  }
  if (body.details && typeof body.details === "object" && !Array.isArray(body.details)) {
    for (const key of Object.keys(defaultDetails)) {
      if (!Object.prototype.hasOwnProperty.call(body.details, key)) continue;
      const edited = limitedText(body.details[key], "", 500);
      rememberManualOverride(manualOverrides, `details.${key}`, edited, current.details?.[key]);
      next.details[key] = edited || defaultDetails[key];
    }
  }
  if (Object.prototype.hasOwnProperty.call(body, "verificationStatus")) {
    const editedVerification = normalizeVerificationStatus(body.verificationStatus, next.candidateType, next.dataStatus);
    rememberManualOverride(manualOverrides, "verificationStatus", editedVerification, current.verificationStatus);
    next.verificationStatus = editedVerification;
  } else {
    next.verificationStatus = normalizeVerificationStatus(current.verificationStatus, next.candidateType, next.dataStatus);
  }
  next.manualOverrides = manualOverrides;
  return next;
}

function adoptCandidateIntoState(state, candidate, input = {}) {
  if (candidate.candidateType !== "place") throw new Error("只有真实候选可以加入最终方案；团队需求和攻略需先匹配具体地点");
  const details = { ...defaultDetails, ...(candidate.details || {}) };
  candidate.selected = true;
  candidate.decisionStatus = "拟定";
  if (candidate.category === "住宿") {
    const roomParts = [details.roomType, details.rooms, details.beds, details.bedTypes].filter(hasUsefulFact);
    state.finalPlan = sanitizeFinalPlan({
      ...state.finalPlan,
      stay: {
        ...state.finalPlan.stay,
        name: candidate.name,
        address: hasUsefulFact(details.address) ? details.address : candidate.area,
        capacity: details.capacity,
        roomsBeds: roomParts.length ? roomParts.join("；") : "待确认",
        twoNightTotal: hasUsefulFact(details.twoNightTotal) ? details.twoNightTotal : candidate.priceLabel,
        checkInOut: [details.checkIn, details.checkOut].filter(hasUsefulFact).join(" · ") || "待确认",
        barbecue: details.barbecue,
        bbqEquipment: details.bbqEquipment,
        breakfast: details.breakfast,
        sourceUrl: candidate.sourceUrl,
      },
    }, state.finalPlan);
    return { target: "stay" };
  }
  const rawDay = limitedText(input.day, "", 20);
  if (!rawDay || !/第\s*[123]\s*天|Day\s*[123]|周五|星期五|周六|星期六|周日|星期日|周天/i.test(rawDay)) {
    throw new Error("加入行程时请指定 day：第1天、第2天或第3天");
  }
  const day = normalizePlanDay(rawDay);
  const itineraryItem = {
    day,
    time: limitedText(input.time, "待安排", 30),
    endTime: limitedText(input.endTime, "待安排", 30),
    category: candidate.category,
    title: candidate.name,
    subtitle: limitedText(input.subtitle, candidate.summary || candidate.pros?.[0] || "", 800),
    address: hasUsefulFact(details.address) ? details.address : candidate.area,
    transport: limitedText(input.transport, "待补充", 300),
    cost: boundedCost(Object.prototype.hasOwnProperty.call(input, "cost") ? input.cost : candidate.price),
    bookingStatus: details.bookingStatus || "待确认",
    sourceUrl: candidate.sourceUrl,
    note: limitedText(input.note, candidate.cons?.[0] || "", 500),
    sourceId: candidate.id,
  };
  const itinerary = (state.finalPlan.itinerary || []).filter((item) => item.sourceId !== candidate.id);
  const insertionPoint = itinerary.reduce((last, item, index) => normalizePlanDay(item.day) === day ? index + 1 : last, itinerary.length);
  itinerary.splice(insertionPoint, 0, itineraryItem);
  state.finalPlan = sanitizeFinalPlan({ ...state.finalPlan, itinerary }, state.finalPlan);
  return { target: "itinerary", day };
}

async function stateForClient(request) {
  const state = await readState();
  state.settings.provider = providerLabel();
  state.settings.workbookPath = path.relative(rootDir, workbookPath);
  const canManage = isLocalManagerRequest(request);
  state.capabilities = { canManage };
  state.places = state.places.map((place) => {
    const link = state.links.find((candidate) => candidate.id === place.sourceId);
    const quality = candidateQuality(place, link);
    const enriched = {
      ...place,
      completeness: place.aiCompleteness ?? quality.percent,
      keyMissing: [...new Set([...(place.missingFields || link?.missingFields || []), ...quality.missing])],
    };
    if (canManage) return enriched;
    const detailKeys = {
      住宿: ["address", "twoNightTotal", "roomType", "rooms", "beds", "bedTypes", "capacity", "barbecue", "extraFees", "cancellationPolicy", "bookingStatus"],
      餐饮: ["address", "sixPersonTotal", "signatureDishes", "privateRoom", "groupSuitability", "queueInfo", "openingHours", "bookingStatus"],
      密室: ["address", "themeName", "horrorLevel", "difficulty", "venueSize", "roomCount", "sixPersonSession", "npcInteraction", "bookingStatus"],
      休闲娱乐: ["address", "sixPersonTotal", "leisureFacilities", "overnight", "includedMeals", "restArea", "serviceRestrictions", "bookingStatus"],
      景点: ["address", "ticketInfo", "openingHours", "recommendedDuration", "indoorOutdoor", "weatherImpact", "reservation", "bookingStatus"],
      攻略: ["address", "evidence"],
    }[place.category] || ["address", "bookingStatus"];
    return {
      ...enriched,
      factsFound: undefined,
      manualOverrides: undefined,
      details: Object.fromEntries(detailKeys.map((key) => [key, enriched.details?.[key]])),
    };
  });
  if (!canManage) {
    state.links = state.links.map((item) => ({
      id: item.id,
      sourceType: item.sourceType,
      inputText: item.inputText,
      url: item.url,
      title: item.title,
      category: item.category,
      subCategory: item.subCategory,
      status: item.status,
      readStatus: item.readStatus,
      organizedStatus: item.organizedStatus,
      factsFound: item.factsFound,
      missingFields: item.missingFields,
      resultNote: item.resultNote,
      submitter: item.submitter,
      createdAt: item.createdAt,
      updatedAt: item.updatedAt,
      model: item.model,
      candidateCount: item.candidateCount,
    }));
    delete state.aiReviews;
    state.itinerary = { day0: [], day1: [], day2: [] };
    state.reservations = [];
    state.settings.workbookPath = "";
  }
  return state;
}

async function sendStateResponse(request, response) {
  const state = await stateForClient(request);
  const bytes = Buffer.from(JSON.stringify(state));
  const etag = `W/"${createHash("sha1").update(bytes).digest("hex").slice(0, 20)}"`;
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, If-None-Match",
    "Access-Control-Allow-Methods": "GET,POST,PATCH,DELETE,OPTIONS",
    "Cache-Control": "private, no-cache",
    ETag: etag,
  };
  if (String(request.headers["if-none-match"] || "") === etag) return sendBuffer(response, 304, headers, Buffer.alloc(0));
  return sendBuffer(response, 200, headers, bytes);
}

async function proxyToUi(request, response, url) {
  const cacheableAsset = /^(?:\/assets\/|\/favicon\.(?:svg|ico|png)$)/.test(url.pathname) && ["GET", "HEAD"].includes(request.method || "GET");
  const assetCacheKey = `${url.pathname}${url.search}`;
  const cachedAsset = cacheableAsset ? uiAssetCache.get(assetCacheKey) : null;
  if (cachedAsset) {
    const etag = cachedAsset.headers.etag || cachedAsset.headers.ETag;
    if (etag && String(request.headers["if-none-match"] || "") === etag) {
      return sendBuffer(response, 304, cachedAsset.headers, Buffer.alloc(0));
    }
    return sendBuffer(response, cachedAsset.status, cachedAsset.headers, cachedAsset.bytes);
  }
  const upstreamUrl = new URL(`${url.pathname}${url.search}`, `http://127.0.0.1:${uiPort}`);
  const upstreamHeaders = new Headers();
  for (const [name, value] of Object.entries(request.headers)) {
    if (value == null || ["host", "connection", "content-length", "accept-encoding"].includes(name.toLowerCase())) continue;
    upstreamHeaders.set(name, Array.isArray(value) ? value.join(", ") : value);
  }
  upstreamHeaders.set("x-forwarded-host", request.headers.host || `localhost:${localPort}`);
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
  if (cacheableAsset && upstream.ok && bytes.length) uiAssetCache.set(assetCacheKey, { status: upstream.status, headers: responseHeaders, bytes });
  return sendBuffer(response, upstream.status, responseHeaders, bytes);
}

function createTravelServer(accessMode) {
  return http.createServer((request, response) => {
    Object.defineProperty(request, accessModeSymbol, { value: accessMode });
    Object.defineProperty(response, responseRequestSymbol, { value: request });
    const handle = async () => {
      const url = new URL(request.url || "/", `http://${request.headers.host || "localhost"}`);
      if (!isAllowedBrowserOrigin(request)) return sendJson(response, 403, { error: "请从当前出行共创页面完成操作" });
      if (request.method === "OPTIONS") return sendJson(response, 204, {});
      try {
    if (!(await authorizePublicRequest(request, response, url))) return;
    if (request.method === "GET" && url.pathname === "/api/health") {
      return sendJson(response, 200, { ok: true, provider: providerLabel(), workbook: path.relative(rootDir, workbookPath), now: new Date().toISOString() });
    }
    if (request.method === "GET" && url.pathname === "/api/state") {
      return sendStateResponse(request, response);
    }
    if (request.method === "GET" && url.pathname === "/api/budget-advice") {
      const state = await readState();
      return sendJson(response, 200, await generateBudgetAdvice(state, { refresh: url.searchParams.get("refresh") === "1" }));
    }
    if (request.method === "POST" && url.pathname === "/api/ai-review") {
      if (!requireManagement(request, response)) return;
      await serializeMutation(() => importFromExcel());
      const reviewBase = await readState();
      const review = await generateTripReview(reviewBase);
      await serializeMutation(async () => {
        const state = await readState();
        if (review.baseUpdatedAt && review.baseUpdatedAt !== state.finalPlan?.updatedAt) {
          review.note = `${review.note} 审阅期间方案发生了变化，请逐条核对后再采纳。`;
        }
        state.aiReviews = normalizeAiReviews([review, ...(state.aiReviews || [])]);
        await syncToExcel(state);
      });
      return sendJson(response, 200, { ok: true, review, state: await stateForClient(request) });
    }
    const reviewSuggestionMatch = url.pathname.match(/^\/api\/ai-reviews\/([^/]+)\/suggestions\/([^/]+)$/);
    if (request.method === "POST" && reviewSuggestionMatch) {
      if (!requireManagement(request, response)) return;
      const body = await readBody(request);
      const state = await readState();
      const suggestion = applyAiReviewSuggestion(
        state,
        decodeURIComponent(reviewSuggestionMatch[1]),
        decodeURIComponent(reviewSuggestionMatch[2]),
        limitedText(body.action, "", 20),
      );
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, suggestion, state: await stateForClient(request) });
    }
    if (request.method === "POST" && url.pathname === "/api/final-plan") {
      if (!requireManagement(request, response)) return;
      const body = await readBody(request);
      const state = await readState();
      const baseUpdatedAt = limitedText(body.baseUpdatedAt, "", 80);
      if (baseUpdatedAt && baseUpdatedAt !== state.finalPlan.updatedAt) {
        return sendJson(response, 409, {
          error: "这份方案刚刚被其他人或 Excel 更新过。请刷新后再编辑，避免覆盖最新内容。",
          state: await stateForClient(request),
        });
      }
      state.finalPlan = sanitizeFinalPlan(body.finalPlan, state.finalPlan);
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, state: await stateForClient(request) });
    }
    const adoptMatch = url.pathname.match(/^\/api\/places\/([^/]+)\/adopt$/);
    if (request.method === "POST" && adoptMatch) {
      if (!requireManagement(request, response)) return;
      const id = decodeURIComponent(adoptMatch[1]);
      const body = await readBody(request);
      const state = await readState();
      const candidate = state.places.find((item) => item.id === id);
      if (!candidate) return sendJson(response, 404, { error: "候选项目不存在" });
      const adopted = adoptCandidateIntoState(state, candidate, body);
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, ...adopted, state: await stateForClient(request) });
    }
    const candidateMatch = url.pathname.match(/^\/api\/places\/([^/]+)$/);
    if (request.method === "PATCH" && candidateMatch) {
      if (!requireManagement(request, response)) return;
      const id = decodeURIComponent(candidateMatch[1]);
      const body = await readBody(request);
      const state = await readState();
      const index = state.places.findIndex((item) => item.id === id);
      if (index < 0) return sendJson(response, 404, { error: "候选项目不存在" });
      state.places[index] = sanitizeCandidatePatch(body, state.places[index]);
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, state: await stateForClient(request) });
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
      return sendJson(response, 200, { ok: true, removed, state: await stateForClient(request) });
    }
    if (request.method === "GET" && url.pathname === "/api/download/excel") {
      if (!requireManagement(request, response)) return;
      const bytes = await fs.readFile(workbookPath);
      response.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(workbookFileName)}`,
        "Access-Control-Allow-Origin": "*",
        "Content-Length": bytes.length,
      });
      return response.end(bytes);
    }
    if (request.method === "POST" && ["/api/links", "/api/submissions"].includes(url.pathname)) {
      const body = await readBody(request);
      const incomingUrls = Array.isArray(body.urls) ? body.urls : [];
      const incomingTexts = [
        ...(Array.isArray(body.texts) ? body.texts : []),
        ...(typeof body.text === "string" ? [body.text] : []),
      ].map((item) => String(item || "").trim()).filter(Boolean).slice(0, 10);
      if (!incomingUrls.length && !incomingTexts.length) return sendJson(response, 400, { error: "请提交至少一个链接，或写下一段旅行诉求" });
      const submitter = String(body.submitter || "").trim().slice(0, 40);
      if (!submitter || submitter === "团队成员") return sendJson(response, 400, { error: "请先填写团队昵称" });
      const state = await readState();
      const existingUrls = new Set(state.links.map((item) => item.url).filter(Boolean));
      const existingTexts = new Set(state.links
        .filter((item) => item.sourceType === "文字" && item.inputText)
        .map((item) => `${item.submitter}:${String(item.inputText).replace(/\s+/g, " ").trim()}`));
      const created = [];
      let linkCreated = 0;
      let textCreated = 0;
      let duplicates = 0;
      let invalid = 0;
      for (const raw of incomingUrls.slice(0, 30)) {
        let normalized;
        try { normalized = cleanUrl(raw); } catch { invalid += 1; continue; }
        if (existingUrls.has(normalized)) { duplicates += 1; continue; }
        existingUrls.add(normalized);
        const now = new Date().toISOString();
        const id = `link-${createHash("sha1").update(`${normalized}-${now}`).digest("hex").slice(0, 12)}`;
        state.links.unshift({
          id,
          sourceType: "链接",
          inputText: "",
          url: normalized,
          title: "",
          category: body.category || "自动识别",
          status: "等待处理",
          readStatus: "等待读取",
          organizedStatus: "整理中",
          candidateCount: 0,
          factsFound: [],
          missingFields: [],
          resultNote: "已进入后台处理队列",
          submitter,
          note: String(body.note || "").slice(0, 300),
          createdAt: now,
          updatedAt: now,
          summary: "",
          model: "",
          error: "",
        });
        created.push(id);
        linkCreated += 1;
      }
      for (const rawText of incomingTexts) {
        const inputText = rawText.slice(0, 4_000);
        if (inputText.length < 3) { invalid += 1; continue; }
        const textKey = `${submitter}:${inputText.replace(/\s+/g, " ").trim()}`;
        if (existingTexts.has(textKey)) { duplicates += 1; continue; }
        existingTexts.add(textKey);
        const now = new Date().toISOString();
        const id = `text-${createHash("sha1").update(`${textKey}-${now}`).digest("hex").slice(0, 12)}`;
        state.links.unshift({
          id,
          sourceType: "文字",
          inputText,
          url: "",
          title: textSubmissionTitle(inputText),
          category: body.category || "自动识别",
          status: "等待处理",
          readStatus: "文字已接收",
          organizedStatus: "整理中",
          candidateCount: 0,
          factsFound: [],
          missingFields: [],
          resultNote: "已进入 DeepSeek 整理队列",
          submitter,
          note: String(body.note || "").slice(0, 300),
          createdAt: now,
          updatedAt: now,
          summary: "",
          model: "",
          error: "",
        });
        created.push(id);
        textCreated += 1;
      }
      if (!created.length) {
        const error = invalid
          ? "没有可接收的公网链接或有效文字；本机、局域网和非 http(s) 地址不会读取"
          : "这些内容已经投递过了，不需要重复提交";
        return sendJson(response, invalid ? 400 : 409, { error, created: 0, linkCreated, textCreated, duplicates, invalid, ids: [] });
      }
      await writeState(state);
      await syncToExcel(state);
      created.forEach((id, index) => setTimeout(() => enqueueLink(id), 250 + index * 80));
      return sendJson(response, 202, { created: created.length, linkCreated, textCreated, duplicates, invalid, ids: created });
    }
    const editLinkMatch = url.pathname.match(/^\/api\/links\/([^/]+)$/);
    if (request.method === "PATCH" && editLinkMatch) {
      if (!requireManagement(request, response)) return;
      const id = decodeURIComponent(editLinkMatch[1]);
      const body = await readBody(request);
      const state = await readState();
      const index = state.links.findIndex((item) => item.id === id);
      if (index < 0) return sendJson(response, 404, { error: "投递记录不存在" });
      const current = state.links[index];
      if (["等待处理", "正在读取", "AI分析中"].includes(current.status)) {
        return sendJson(response, 409, { error: "这条投递正在处理，完成后再修改，避免覆盖新结果" });
      }
      const patch = {};
      for (const [key, maximum] of Object.entries({ title: 160, note: 300, submitter: 40 })) {
        if (Object.prototype.hasOwnProperty.call(body, key)) patch[key] = limitedText(body[key], "", maximum);
      }
      if (Object.prototype.hasOwnProperty.call(body, "category")) {
        patch.category = body.category === "自动识别" ? "自动识别" : normalizeCategory(body.category, current.title);
      }
      if (Object.prototype.hasOwnProperty.call(body, "subCategory")) {
        const effectiveCategory = patch.category || current.category;
        const proposed = limitedText(body.subCategory, "", 80);
        patch.subCategory = effectiveCategory === "自动识别"
          ? proposed || "等待识别"
          : validSubCategory(effectiveCategory, proposed) ? proposed : inferSubCategory(`${current.title} ${proposed}`, effectiveCategory);
      }
      if (Object.prototype.hasOwnProperty.call(body, "inputText") && current.sourceType === "文字") {
        patch.inputText = limitedText(body.inputText, "", 4_000);
        if (patch.inputText.length < 3) return sendJson(response, 400, { error: "文字内容至少需要 3 个字" });
      }
      if (Object.prototype.hasOwnProperty.call(body, "url") && current.sourceType !== "文字") {
        try { patch.url = cleanUrl(body.url); } catch { return sendJson(response, 400, { error: "链接格式不正确" }); }
      }
      if (body.reprocess === true) Object.assign(patch, { status: "等待处理", organizedStatus: "整理中", resultNote: "已重新加入后台处理队列", error: "" });
      state.links[index] = { ...current, ...patch, updatedAt: new Date().toISOString() };
      await syncToExcel(state);
      if (body.reprocess === true) setTimeout(() => enqueueLink(id), 200);
      return sendJson(response, body.reprocess === true ? 202 : 200, { ok: true, state: await stateForClient(request) });
    }
    const retryMatch = url.pathname.match(/^\/api\/links\/([^/]+)\/retry$/);
    if (request.method === "POST" && retryMatch) {
      if (!requireManagement(request, response)) return;
      const id = decodeURIComponent(retryMatch[1]);
      const state = await readState();
      const existing = state.links.find((item) => item.id === id);
      if (!existing) return sendJson(response, 404, { error: "投递记录不存在，无法重新处理" });
      if (["正在读取", "AI分析中"].includes(existing.status)) return sendJson(response, 409, { error: "这条投递正在处理中，请稍后再试" });
      const isText = existing.sourceType === "文字";
      await updateLink(id, { status: "等待处理", readStatus: isText ? "文字已接收" : "等待读取", organizedStatus: "整理中", resultNote: "已重新加入后台处理队列", error: "" });
      await syncToExcel(await readState());
      setTimeout(() => enqueueLink(id), 200);
      return sendJson(response, 202, { ok: true });
    }
    const deleteMatch = url.pathname.match(/^\/api\/links\/([^/]+)$/);
    if (request.method === "DELETE" && deleteMatch) {
      if (!requireManagement(request, response)) return;
      const id = decodeURIComponent(deleteMatch[1]);
      const state = await readState();
      const link = state.links.find((item) => item.id === id);
      if (!link) return sendJson(response, 404, { error: "投递记录不存在" });
      if (["等待处理", "正在读取", "AI分析中"].includes(link.status)) {
        return sendJson(response, 409, { error: "这条投递正在处理，完成前不能删除" });
      }
      const linkedCandidates = state.places.filter((item) => item.sourceId === id);
      const itinerarySourceIds = new Set((state.finalPlan?.itinerary || []).map((item) => item.sourceId).filter(Boolean));
      const protectedCandidates = linkedCandidates.filter((item) => hasProtectedHumanState(item) || itinerarySourceIds.has(item.id));
      if (protectedCandidates.length) {
        return sendJson(response, 409, {
          error: `这条投递还有 ${protectedCandidates.length} 个已入选、已投票或人工修改的结果。请先在候选与最终方案中处理后再删除。`,
        });
      }
      state.links = state.links.filter((item) => item.id !== id);
      state.places = state.places.filter((item) => item.sourceId !== id);
      await writeState(state);
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, state: await stateForClient(request) });
    }
    if (request.method === "POST" && url.pathname === "/api/sync/from-excel") {
      if (!requireManagement(request, response)) return;
      await importFromExcel();
      return sendJson(response, 200, { ok: true, state: await stateForClient(request) });
    }
    if (request.method === "POST" && url.pathname === "/api/sync/to-excel") {
      if (!requireManagement(request, response)) return;
      const state = await readState();
      await syncToExcel(state);
      return sendJson(response, 200, { ok: true, state: await stateForClient(request) });
    }
    if (!url.pathname.startsWith("/api/")) return proxyToUi(request, response, url);
    return sendJson(response, 404, { error: "接口不存在" });
      } catch (error) {
        return sendJson(response, 500, { error: error instanceof Error ? error.message : "服务器错误" });
      }
    };
    const isLongAiReview = request.method === "POST" && /^\/api\/ai-review(?:\?|$)/.test(request.url || "");
    const isMutation = !isLongAiReview && ["POST", "PATCH", "DELETE"].includes(request.method || "") && /^\/api\//.test(request.url || "");
    if (isMutation) void serializeMutation(handle).catch((error) => {
      if (!response.headersSent) sendJson(response, 500, { error: error instanceof Error ? error.message : "服务器错误" });
    });
    else void handle();
  });
}

const localServer = createTravelServer("local");
const publicServer = createTravelServer("public");

if (!/^(1|true|yes)$/i.test(process.env.TEST_SKIP_LISTEN || "")) {
  localServer.listen(localPort, localHost, () => {
    console.log(`Travel Co-creation local manager: http://localhost:${localPort}`);
  });
  publicServer.listen(publicPort, publicHost, () => {
    console.log(`Travel Co-creation public gateway: http://${publicHost}:${publicPort}`);
  });

  try { lastKnownWorkbookMtime = (await fs.stat(workbookPath)).mtimeMs; } catch { lastKnownWorkbookMtime = 0; }
  setInterval(async () => {
    if (writeInProgress) return;
    try {
      const mtime = (await fs.stat(workbookPath)).mtimeMs;
      if (lastKnownWorkbookMtime && mtime > lastKnownWorkbookMtime + 100) await serializeMutation(() => importFromExcel());
      else if (!lastKnownWorkbookMtime) lastKnownWorkbookMtime = mtime;
    } catch { /* The workbook is generated during setup. */ }
  }, 4_000).unref();

  for (const signal of ["SIGINT", "SIGTERM"]) {
    process.on(signal, () => {
      let pending = 2;
      const done = () => {
        pending -= 1;
        if (pending === 0) process.exit(0);
      };
      localServer.close(done);
      publicServer.close(done);
    });
  }
}

export {
  adoptCandidateIntoState,
  applyAiReviewSuggestion,
  buildBudgetSnapshot,
  buildTripReviewSnapshot,
  cleanUrl,
  dedupeAnalysisCandidates,
  enforceAiReviewSafety,
  fallbackBudgetAdvice,
  fallbackTripReview,
  isBlockedNetworkAddress,
  isLocalManagerRequest,
  normalizeAnalysisCandidates,
  normalizeAiReviews,
  normalizeCandidateType,
  preservedItinerarySourceId,
  sanitizeCandidatePatch,
};
