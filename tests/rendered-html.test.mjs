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

test("server renders the Yangzhou trip workspace", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>下扬州 · 团队旅行共创台<\/title>/);
  assert.match(html, /网页与 Excel「行程首页」双向同步/);
  assert.match(html, /编辑最终方案/);
  assert.match(html, /最终方案信息/);
  assert.match(html, /最终方案和候选资料已经分开/);
  assert.match(html, /已确定/);
  assert.match(html, /待补充/);
  assert.match(html, /链接收集/);
  assert.match(html, /Excel 工作台/);
  assert.doesNotMatch(html, /Your site is taking shape|react-loading-skeleton/);
});

test("local data and the generated workbook are present", async () => {
  const [stateText, workbook] = await Promise.all([
    readFile(new URL("../data/store.json", import.meta.url), "utf8"),
    readFile(
      new URL(
        "../outputs/019f7eda-a998-7660-a73d-e43b3af67965/扬州团队旅行攻略.xlsx",
        import.meta.url,
      ),
    ),
  ]);
  const state = JSON.parse(stateText);

  assert.equal(state.project.destination, "扬州");
  assert.equal(state.project.days, 2);
  assert.ok(state.places.length >= 1);
  assert.equal(workbook.subarray(0, 2).toString(), "PK");
});
