import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server renders the Jeju travel co-creation workspace", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>济州岛旅行共创台 · 六个人一起完成三日计划<\/title>/);
  assert.match(html, /JEJU,/);
  assert.match(html, /TOGETHER\./);
  assert.match(html, /OUR JEJU NOTE/);
  assert.match(html, /分享一个济州想法/);
  assert.match(html, /先把住哪里定下来/);
  assert.match(html, /三天，把海岸和小城慢慢走完/);
  assert.match(html, /已确定/);
  assert.match(html, /待补充/);
  assert.match(html, /投递 ADD/);
  assert.match(html, /候选 PICK/);
  assert.doesNotMatch(html, /主电脑管理|下载 Excel 副本|读取主电脑原文件/);
  assert.doesNotMatch(html, /Your site is taking shape|react-loading-skeleton/);
});

test("团队昵称默认留空，投递和投票都明确标记必填", async () => {
  const source = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  assert.doesNotMatch(source, /useState\("团队成员"\)/);
  assert.match(source, /你的昵称（必填）/);
  assert.match(source, /我的昵称（必填）/);
  assert.match(source, /方便大家区分是谁提交的/);
});

test("团队流程先展示真实候选，并提供投递后引导和 AI 预算建议", async () => {
  const source = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  assert.match(source, /下一步：去看真实候选并投票/);
  assert.match(source, /AI 预算助手/);
  assert.match(source, /团队参与/);
  assert.match(source, /先让 AI 做初步审阅，再由你逐条决定/);
  assert.match(source, /采纳并写入 Excel/);
  assert.ok(source.indexOf('id="library-results"') < source.indexOf('className="reference-library"'));
});

test("local data and the generated workbook are present", async () => {
  const [stateText, templateText, workbook] = await Promise.all([
    readFile(new URL("../data/store.json", import.meta.url), "utf8"),
    readFile(new URL("../data/trip-template.json", import.meta.url), "utf8"),
    readFile(
      new URL(
        "../outputs/019fa660-42f1-74a1-ab98-31c99abb6300/韩国济州岛旅行共创.xlsx",
        import.meta.url,
      ),
    ),
  ]);
  const state = JSON.parse(stateText);
  const template = JSON.parse(templateText);

  assert.equal(state.project.destination, "韩国济州岛");
  assert.equal(state.project.days, 3);
  assert.equal(state.project.people, 6);
  assert.equal(state.finalPlan.dates, "2026年8月21日–23日（周五–周日）");
  assert.equal(state.project.budget, 6000);
  assert.equal(state.finalPlan.perPersonBudget, "¥6,000 / 人（包含往返济州机票）");
  assert.equal(state.finalPlan.roundTripFlightPerPerson, "待填写实际含税票价（含托运行李）");
  assert.match(state.finalPlan.summary, /¥36,000/);
  assert.match(state.finalPlan.summary, /正餐约 ¥100\/人/);
  assert.match(state.finalPlan.summary, /公交优先/);
  assert.ok(state.finalPlan.itinerary.some((item) => item.title === "济州市区出发 · 东线公交"));
  assert.ok(state.finalPlan.itinerary.every((item) => item.title !== "济州市区出发 · 东线包车"));
  assert.ok(state.places.length >= 1);
  assert.ok(state.places.every((place) => ["住宿", "餐饮", "密室", "休闲娱乐", "景点", "攻略"].includes(place.category)));
  assert.ok(state.places.every((place) => typeof place.subCategory === "string" && place.subCategory.length > 0));
  assert.ok(state.places.every((place) => Array.isArray(place.featureTags)));
  assert.ok(state.links.every((item) => ["链接", "文字"].includes(item.sourceType)));
  assert.ok(state.links.every((item) => typeof item.inputText === "string"));
  assert.equal(template.project.destination, "待确定目的地");
  assert.deepEqual(template.links, []);
  assert.deepEqual(template.places, []);
  const subCategories = {
    住宿: ["民宿", "酒店", "公寓", "客栈", "度假村", "其他住宿"],
    餐饮: ["烧烤", "火锅", "早茶早餐", "炒菜正餐", "自助餐", "夜宵", "甜品饮品", "咖啡", "酒吧", "其他餐饮"],
    密室: ["推理密室", "机关密室", "剧情密室", "恐怖密室", "未分类密室"],
    休闲娱乐: ["汗蒸", "桑拿", "洗浴中心", "温泉", "SPA", "足疗按摩", "KTV", "桌游", "电竞", "电玩城", "沉浸式剧场", "其他休闲"],
    景点: ["博物馆", "园林景区", "历史街区", "公园", "古镇", "露营", "户外运动", "演出展览", "其他景点"],
    攻略: ["美食攻略", "住宿攻略", "行程攻略", "综合攻略"],
  };
  assert.ok(state.places.every((place) => subCategories[place.category].includes(place.subCategory)));
  assert.equal(workbook.subarray(0, 2).toString(), "PK");
});
