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

test("server renders the reusable travel co-creation workspace", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>出行共创台 · 把大家的想法整理成一份行程<\/title>/);
  assert.match(html, /团队最终行程/);
  assert.match(html, /投递链接或想法/);
  assert.match(html, /住宿与待确认事项/);
  assert.match(html, /最终内容只来自 Excel 第一张/);
  assert.match(html, /已确定/);
  assert.match(html, /待补充/);
  assert.match(html, /投递想法/);
  assert.match(html, /重点候选/);
  assert.match(html, /Excel 管理/);
  assert.doesNotMatch(html, /Your site is taking shape|react-loading-skeleton/);
});

test("local data and the generated workbook are present", async () => {
  const [stateText, templateText, workbook] = await Promise.all([
    readFile(new URL("../data/store.json", import.meta.url), "utf8"),
    readFile(new URL("../data/trip-template.json", import.meta.url), "utf8"),
    readFile(
      new URL(
        "../outputs/019f7eda-a998-7660-a73d-e43b3af67965/出行共创项目.xlsx",
        import.meta.url,
      ),
    ),
  ]);
  const state = JSON.parse(stateText);
  const template = JSON.parse(templateText);

  assert.equal(state.project.destination, "扬州");
  assert.equal(state.project.days, 2);
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
