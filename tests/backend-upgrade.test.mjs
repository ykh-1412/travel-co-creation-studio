import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import http from "node:http";
import net from "node:net";
import path from "node:path";
import { setTimeout as delay } from "node:timers/promises";
import test, { after } from "node:test";
import { fileURLToPath } from "node:url";

const repositoryRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const unitFixtureRoot = await mkdtemp(path.join(repositoryRoot, ".test-backend-unit-"));

process.env.TEST_SKIP_LISTEN = "1";
process.env.TRIP_DATA_FILE = path.relative(repositoryRoot, path.join(unitFixtureRoot, "store.json"));
process.env.TRIP_OUTPUT_DIR = path.relative(repositoryRoot, path.join(unitFixtureRoot, "outputs"));
process.env.WORKBOOK_FILE_NAME = "unit-test.xlsx";
process.env.AI_PROVIDER = "demo";

const {
  buildBudgetSnapshot,
  cleanUrl,
  dedupeAnalysisCandidates,
  fallbackBudgetAdvice,
  isBlockedNetworkAddress,
  isLocalManagerRequest,
  normalizeAnalysisCandidates,
  normalizeCandidateType,
  preservedItinerarySourceId,
  sanitizeCandidatePatch,
} = await import(`../server/index.mjs?backend-unit=${process.pid}-${Date.now()}`);

after(async () => {
  await rm(unitFixtureRoot, { recursive: true, force: true });
});

const fallbackCandidate = {
  candidateType: "guide",
  name: "整篇攻略",
  category: "攻略",
  subCategory: "综合攻略",
  area: "济州",
  summary: "原始攻略摘要",
  factsFound: ["页面标题"],
  missingFields: ["动态价格"],
  details: {
    address: "待核实",
    evidence: "来源页面",
  },
};

test("normalizeCandidateType keeps new values and upgrades legacy records", () => {
  assert.equal(normalizeCandidateType("PLACE"), "place");
  assert.equal(normalizeCandidateType("guide"), "guide");
  assert.equal(normalizeCandidateType("place", { sourceType: "文字", category: "餐饮" }), "requirement");
  assert.equal(normalizeCandidateType("", { sourceType: "文字", category: "餐饮" }), "requirement");
  assert.equal(normalizeCandidateType(undefined, { category: "攻略" }), "guide");
  assert.equal(normalizeCandidateType(undefined, { category: "住宿" }), "place");
  assert.equal(normalizeCandidateType(undefined, { category: "餐饮", dataStatus: "团队诉求，待匹配" }), "requirement");
});

test("normalizeAnalysisCandidates accepts a multi-candidate envelope", () => {
  const result = normalizeAnalysisCandidates({
    summary: "济州美食合集",
    candidates: [
      {
        candidateType: "place",
        name: "甲餐厅",
        category: "餐饮",
        factsFound: ["招牌菜"],
        missingFields: ["预约方式"],
        details: { address: "济州市甲路 1 号" },
      },
      {
        name: "乙密室",
        category: "密室",
        factsFound: ["中恐"],
        details: { horrorLevel: "中恐" },
      },
    ],
  }, fallbackCandidate, { sourceType: "链接" });

  assert.equal(result.length, 2);
  assert.deepEqual(result.map((item) => item.name), ["甲餐厅", "乙密室"]);
  assert.deepEqual(result.map((item) => item.candidateType), ["place", "place"]);
  assert.deepEqual(result.map((item) => item.category), ["餐饮", "密室"]);
  assert.equal(result[0].details.address, "济州市甲路 1 号");
  assert.equal(result[0].details.evidence, "来源页面");
  assert.equal(result[1].details.horrorLevel, "中恐");
  assert.deepEqual(result[1].factsFound, ["中恐"]);
});

test("normalizeAnalysisCandidates remains compatible with legacy single-object and array output", () => {
  const legacy = normalizeAnalysisCandidates({
    name: "旧版民宿结果",
    category: "住宿",
    details: { capacity: "6 人" },
  }, fallbackCandidate, { sourceType: "链接" });
  assert.equal(legacy.length, 1);
  assert.equal(legacy[0].candidateType, "place");
  assert.equal(legacy[0].name, "旧版民宿结果");
  assert.equal(legacy[0].details.capacity, "6 人");

  const directArray = normalizeAnalysisCandidates([
    { name: "景点甲", category: "景点" },
    { name: "攻略原文", category: "攻略" },
  ], fallbackCandidate, { sourceType: "链接" });
  assert.deepEqual(directArray.map((item) => item.candidateType), ["place", "guide"]);

  const textRequirements = normalizeAnalysisCandidates({
    candidates: [
      { name: "想住可烧烤民宿", category: "住宿" },
      { name: "不应生成第二条", category: "餐饮" },
    ],
  }, fallbackCandidate, { sourceType: "文字" });
  assert.equal(textRequirements.length, 1);
  assert.equal(textRequirements[0].candidateType, "requirement");
});

test("duplicate AI candidates are removed before matching so human state cannot be overwritten", () => {
  const duplicate = {
    candidateType: "place",
    name: "济州海鲜（东门店）",
    category: "餐饮",
    summary: "第一条结果会继续匹配并保留旧投票",
    details: { address: "济州市中央路 38 号" },
  };
  const result = dedupeAnalysisCandidates([
    duplicate,
    { ...duplicate, summary: "模型重复返回的第二条" },
    { ...duplicate, details: { address: "济州市莲洞路 2 号" } },
  ], { fallbackCategory: "餐饮", fallbackName: "来源页面" });

  assert.equal(result.length, 2);
  assert.equal(result[0].summary, "第一条结果会继续匹配并保留旧投票");
  assert.equal(result[1].details.address, "济州市莲洞路 2 号");
});

test("submitted links accept public web pages but reject local and private network targets", () => {
  assert.equal(cleanUrl("https://example.com/path?utm_source=test&keep=1#section"), "https://example.com/path?keep=1");
  for (const target of [
    "file:///etc/passwd",
    "ftp://example.com/file",
    "http://localhost:8787/api/state",
    "http://127.0.0.1:8787/api/state",
    "http://10.0.0.1/",
    "http://172.16.1.1/",
    "http://192.168.1.1/",
    "http://169.254.169.254/latest/meta-data/",
    "http://[::1]/",
  ]) assert.throws(() => cleanUrl(target));
  assert.equal(isBlockedNetworkAddress("198.18.0.1"), true);
  assert.equal(isBlockedNetworkAddress("198.18.0.1", { allowProxySynthetic: true }), false);
  assert.equal(isBlockedNetworkAddress("8.8.8.8"), false);
});

test("management capability is decided by listener access mode, never by Host", () => {
  assert.equal(isLocalManagerRequest({ accessMode: "local", headers: { host: "public.example" } }), true);
  assert.equal(isLocalManagerRequest({ accessMode: "public", headers: { host: "localhost" } }), false);
  assert.equal(isLocalManagerRequest({ headers: { host: "localhost" } }), false);
});

test("sanitizeCandidatePatch preserves verification unless explicitly edited and protects classification edits", () => {
  const current = {
    id: "place-edit",
    candidateType: "place",
    name: "济州黑猪烤肉",
    category: "餐饮",
    subCategory: "烧烤",
    summary: "AI 摘要",
    dataStatus: "AI 总结，等待人工核实",
    verificationStatus: "已核实",
    manualOverrides: {},
    details: { address: "济州市测试路 1 号" },
  };

  const summaryOnly = sanitizeCandidatePatch({ summary: "人工补充摘要" }, current);
  assert.equal(summaryOnly.verificationStatus, "已核实");
  assert.equal(summaryOnly.manualOverrides.verificationStatus, undefined);
  assert.equal(summaryOnly.manualOverrides.summary, "人工补充摘要");

  const reclassified = sanitizeCandidatePatch({
    category: "密室",
    subCategory: "恐怖密室",
    verificationStatus: "部分核实",
  }, current);
  assert.equal(reclassified.category, "密室");
  assert.equal(reclassified.subCategory, "恐怖密室");
  assert.equal(reclassified.verificationStatus, "部分核实");
  assert.equal(reclassified.manualOverrides.category, "密室");
  assert.equal(reclassified.manualOverrides.subCategory, "恐怖密室");
  assert.equal(reclassified.manualOverrides.verificationStatus, "部分核实");
  assert.deepEqual(current.manualOverrides, {});
});

test("budget advice treats the 6000 yuan target as including round-trip flights", () => {
  const snapshot = buildBudgetSnapshot({
    project: { people: 6, days: 3 },
    finalPlan: {
      destination: "韩国济州岛",
      dates: "2026年8月21日–23日",
      people: 6,
      perPersonBudget: "¥6,000 / 人（包含往返济州机票）",
      roundTripFlightPerPerson: "待填写实际含税票价（含托运行李）",
      summary: "公交优先",
      stay: { twoNightTotal: "六人两晚约 ¥1,800–3,000", sourceUrl: "" },
      itinerary: [
        { category: "晚餐", title: "黑猪烤肉", cost: 100 },
        { category: "交通", title: "东线公交", cost: 50 },
        { category: "景点", title: "城山日出峰", cost: 20 },
      ],
    },
    places: [],
  });
  assert.deepEqual(snapshot.target, { min: 6000, max: 6000 });
  assert.deepEqual(snapshot.known, { min: 470, max: 670 });
  assert.equal(snapshot.includesFlights, true);
  assert.equal(snapshot.excludesFlights, false);
  assert.equal(snapshot.flight, null);
  assert.equal(snapshot.categories[0].name, "往返机票");
  assert.equal(snapshot.categories.find((item) => item.name === "住宿").min, 300);
  assert.equal(snapshot.categories.find((item) => item.name === "餐饮").max, 100);
  const advice = fallbackBudgetAdvice(snapshot);
  assert.equal(advice.categories.length, 5);
  assert.equal(advice.categories[0].planned, "待确认");
  assert.equal(advice.overallStatus, "信息不足");
  assert.match(advice.summary, /将往返济州机票计入总预算/);
  assert.match(advice.reserveAdvice, /不能把账面差额全部视为机动金/);
  assert.ok(advice.missingInputs.includes("往返济州机票人均含税价格（含行李）"));
  assert.ok(advice.missingInputs.includes("住宿真实链接和动态价格"));

  const withFlight = buildBudgetSnapshot({
    project: { people: 6, days: 3 },
    finalPlan: {
      destination: "韩国济州岛",
      people: 6,
      perPersonBudget: "¥6,000 / 人（包含往返济州机票）",
      roundTripFlightPerPerson: "¥2,800 / 人（含税及托运行李）",
      stay: { twoNightTotal: "六人两晚约 ¥1,800–3,000", sourceUrl: "" },
      itinerary: [
        { category: "晚餐", title: "黑猪烤肉", cost: 100 },
        { category: "交通", title: "东线公交", cost: 50 },
        { category: "景点", title: "城山日出峰", cost: 20 },
      ],
    },
    places: [],
  });
  assert.deepEqual(withFlight.flight, { min: 2800, max: 2800 });
  assert.deepEqual(withFlight.known, { min: 3270, max: 3470 });
});

test("Excel itinerary imports preserve their candidate source links", () => {
  const existing = [
    { day: "第2天", time: "10:00", title: "城山日出峰", sourceUrl: "https://example.com/seongsan", sourceId: "place-seongsan" },
    { day: "第3天", time: "14:30", title: "国际航班返程", sourceUrl: "https://example.com/entry", sourceId: "guide-entry" },
  ];
  assert.equal(preservedItinerarySourceId(existing, {
    day: "第2天", time: "10:00", title: "城山日出峰", sourceUrl: "https://example.com/seongsan",
  }), "place-seongsan");
  assert.equal(preservedItinerarySourceId(existing, {
    day: "第3天", time: "15:00", title: "返程时间已调整", sourceUrl: "https://example.com/entry",
  }), "guide-entry");
  assert.equal(preservedItinerarySourceId(existing, {
    day: "第1天", time: "18:00", title: "新增自由活动", sourceUrl: "",
  }), "");
});

async function reservePorts(count) {
  const servers = [];
  try {
    for (let index = 0; index < count; index += 1) {
      const server = net.createServer();
      servers.push(server);
      await new Promise((resolve, reject) => {
        server.once("error", reject);
        server.listen(0, "127.0.0.1", resolve);
      });
    }
    return servers.map((server) => server.address().port);
  } finally {
    await Promise.all(servers.map((server) => new Promise((resolve) => server.close(resolve))));
  }
}

function requestServer(port, pathname, { method = "GET", headers = {}, body = "" } = {}) {
  return new Promise((resolve, reject) => {
    const payload = typeof body === "string" ? body : JSON.stringify(body);
    const request = http.request({
      hostname: "127.0.0.1",
      port,
      path: pathname,
      method,
      headers: {
        Host: "localhost",
        ...(payload ? { "Content-Length": Buffer.byteLength(payload) } : {}),
        ...headers,
      },
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => resolve({
        status: response.statusCode,
        headers: response.headers,
        text: Buffer.concat(chunks).toString("utf8"),
      }));
    });
    request.once("error", reject);
    if (payload) request.write(payload);
    request.end();
  });
}

async function waitForServer(port, child, logs) {
  const deadline = Date.now() + 15_000;
  while (Date.now() < deadline) {
    if (child.exitCode !== null) {
      throw new Error(`测试服务提前退出：\n${logs.join("")}`);
    }
    try {
      const response = await requestServer(port, "/api/health");
      if (response.status === 200) return;
    } catch {
      // The child process has not started listening yet.
    }
    await delay(50);
  }
  throw new Error(`等待测试服务超时：\n${logs.join("")}`);
}

async function stopChild(child) {
  if (child.exitCode !== null) return;
  const exited = new Promise((resolve) => child.once("exit", resolve));
  child.kill("SIGTERM");
  await Promise.race([exited, delay(3_000)]);
  if (child.exitCode === null) {
    child.kill("SIGKILL");
    await exited;
  }
}

test("public listener is read-only even when Host is forged as localhost", { timeout: 30_000 }, async () => {
  const fixtureRoot = await mkdtemp(path.join(repositoryRoot, ".test-backend-http-"));
  const outputDirectory = path.join(fixtureRoot, "outputs");
  const [localPort, publicPort] = await reservePorts(2);
  const logs = [];
  const child = spawn(process.execPath, [path.join(repositoryRoot, "server", "index.mjs")], {
    cwd: repositoryRoot,
    env: {
      ...process.env,
      TEST_SKIP_LISTEN: "0",
      API_PORT: String(localPort),
      PUBLIC_API_PORT: String(publicPort),
      API_HOST: "127.0.0.1",
      PUBLIC_API_HOST: "127.0.0.1",
      TRIP_DATA_FILE: path.relative(repositoryRoot, path.join(fixtureRoot, "store.json")),
      TRIP_OUTPUT_DIR: path.relative(repositoryRoot, outputDirectory),
      WORKBOOK_FILE_NAME: "http-test.xlsx",
      AI_PROVIDER: "demo",
      PUBLIC_REQUIRE_PASSWORD: "true",
      PUBLIC_ACCESS_PASSWORD: "integration-password",
    },
    stdio: ["ignore", "pipe", "pipe"],
  });
  child.stdout.on("data", (chunk) => logs.push(chunk.toString()));
  child.stderr.on("data", (chunk) => logs.push(chunk.toString()));

  try {
    await waitForServer(localPort, child, logs);
    await waitForServer(publicPort, child, logs);

    const localStateResponse = await requestServer(localPort, "/api/state");
    assert.equal(localStateResponse.status, 200);
    assert.equal(JSON.parse(localStateResponse.text).capabilities.canManage, true);

    const loginBody = new URLSearchParams({ password: "integration-password" }).toString();
    const loginResponse = await requestServer(publicPort, "/__team-login", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: loginBody,
    });
    assert.equal(loginResponse.status, 303, logs.join(""));
    const cookie = loginResponse.headers["set-cookie"]?.[0]?.split(";", 1)[0];
    assert.ok(cookie, "公网登录应返回访问 cookie");

    const publicHeaders = {
      Cookie: cookie,
      Host: "localhost",
      "Content-Type": "application/json",
    };
    const publicStateResponse = await requestServer(publicPort, "/api/state", { headers: publicHeaders });
    assert.equal(publicStateResponse.status, 200);
    assert.equal(JSON.parse(publicStateResponse.text).capabilities.canManage, false);

    const budgetAdviceResponse = await requestServer(publicPort, "/api/budget-advice", { headers: publicHeaders });
    assert.equal(budgetAdviceResponse.status, 200, budgetAdviceResponse.text);
    const budgetAdvice = JSON.parse(budgetAdviceResponse.text);
    assert.equal(budgetAdvice.source, "rules");
    assert.ok(Array.isArray(budgetAdvice.categories));
    assert.ok(budgetAdvice.categories.some((item) => item.name === "住宿"));

    const pristineState = await readFile(path.join(fixtureRoot, "store.json"), "utf8");
    const unnamedSubmission = await requestServer(publicPort, "/api/submissions", {
      method: "POST",
      headers: publicHeaders,
      body: { text: "想吃济州黑猪烤肉", submitter: "" },
    });
    assert.equal(unnamedSubmission.status, 400, unnamedSubmission.text);
    assert.match(unnamedSubmission.text, /请先填写团队昵称/);

    const privateLinkSubmission = await requestServer(publicPort, "/api/submissions", {
      method: "POST",
      headers: publicHeaders,
      body: { urls: ["http://127.0.0.1:8787/api/state"], submitter: "公网测试" },
    });
    assert.equal(privateLinkSubmission.status, 400, privateLinkSubmission.text);
    assert.match(privateLinkSubmission.text, /本机、局域网/);
    const managementRequests = [
      ["POST", "/api/final-plan", { finalPlan: {} }],
      ["POST", "/api/places/not-found/adopt", { day: "周六" }],
      ["PATCH", "/api/places/not-found", { name: "不应写入" }],
      ["PATCH", "/api/links/not-found", { title: "不应写入" }],
      ["POST", "/api/links/not-found/retry", {}],
      ["DELETE", "/api/links/not-found", ""],
      ["POST", "/api/sync/from-excel", {}],
      ["POST", "/api/sync/to-excel", {}],
    ];

    for (const [method, pathname, body] of managementRequests) {
      const response = await requestServer(publicPort, pathname, {
        method,
        headers: publicHeaders,
        body,
      });
      assert.equal(response.status, 403, `${method} ${pathname} 应在公网端口被拒绝，实际响应：${response.status} ${response.text}`);
      assert.match(response.text, /仅能在发起人的本机管理界面完成/);
    }

    const downloadResponse = await requestServer(publicPort, "/api/download/excel", { headers: publicHeaders });
    assert.equal(downloadResponse.status, 403);
    assert.equal(await readFile(path.join(fixtureRoot, "store.json"), "utf8"), pristineState);
    assert.deepEqual(await readdir(outputDirectory), []);
  } finally {
    await stopChild(child);
    await rm(fixtureRoot, { recursive: true, force: true });
  }
});

function fixtureLink(id, title) {
  const now = "2026-07-26T00:00:00.000Z";
  return {
    id,
    sourceType: "链接",
    inputText: "",
    url: `https://example.com/${id}`,
    title,
    category: "餐饮",
    subCategory: "烧烤",
    status: "已写入Excel",
    readStatus: "成功读取",
    organizedStatus: "已整理",
    candidateCount: 1,
    factsFound: ["页面标题"],
    missingFields: ["动态价格"],
    resultNote: "测试记录",
    submitter: "测试成员",
    note: "",
    createdAt: now,
    updatedAt: now,
    summary: "测试摘要",
    model: "测试",
    error: "",
  };
}

function fixturePlace(id, sourceId, overrides = {}) {
  return {
    id,
    sourceId,
    candidateType: "place",
    verificationStatus: "待核实",
    name: `测试候选 ${id}`,
    category: "餐饮",
    subCategory: "烧烤",
    featureTags: ["餐饮", "烧烤"],
    area: "济州",
    price: 100,
    priceLabel: "人均 ¥100",
    duration: "2 小时",
    summary: "临时回归测试候选",
    factsFound: ["名称"],
    missingFields: ["预约方式"],
    score: 3.5,
    tags: ["烧烤"],
    pros: ["适合团队"],
    cons: ["待核实"],
    selected: false,
    votes: {},
    decisionStatus: "待比较",
    manualNote: "",
    manualOverrides: {},
    sourceUrl: `https://example.com/${sourceId}`,
    dataStatus: "测试数据",
    details: { address: "济州市测试路" },
    ...overrides,
  };
}

test("local mutations preserve concurrent votes and block deletion of human decisions", { timeout: 30_000 }, async () => {
  const fixtureRoot = await mkdtemp(path.join(repositoryRoot, ".test-backend-mutations-"));
  const outputDirectory = path.join(fixtureRoot, "outputs");
  const dataPath = path.join(fixtureRoot, "store.json");
  const [localPort, publicPort] = await reservePorts(2);
  const logs = [];
  let child;

  try {
    const state = JSON.parse(await readFile(path.join(repositoryRoot, "data", "trip-template.json"), "utf8"));
    state.links = [
      fixtureLink("link-vote", "并发投票来源"),
      fixtureLink("link-protected", "有人工结果来源"),
      fixtureLink("link-plain", "普通可删除来源"),
    ];
    state.places = [
      fixturePlace("place-vote", "link-vote"),
      fixturePlace("place-protected", "link-protected", {
        selected: true,
        votes: { "已投票成员": "想去" },
        decisionStatus: "拟定",
        manualNote: "人工确认保留",
        manualOverrides: { summary: "人工编辑过的摘要" },
      }),
      fixturePlace("place-plain", "link-plain"),
    ];
    await writeFile(dataPath, `${JSON.stringify(state, null, 2)}\n`, "utf8");

    child = spawn(process.execPath, [path.join(repositoryRoot, "server", "index.mjs")], {
      cwd: repositoryRoot,
      env: {
        ...process.env,
        TEST_SKIP_LISTEN: "0",
        API_PORT: String(localPort),
        PUBLIC_API_PORT: String(publicPort),
        API_HOST: "127.0.0.1",
        PUBLIC_API_HOST: "127.0.0.1",
        TRIP_DATA_FILE: path.relative(repositoryRoot, dataPath),
        TRIP_OUTPUT_DIR: path.relative(repositoryRoot, outputDirectory),
        WORKBOOK_FILE_NAME: "mutation-test.xlsx",
        AI_PROVIDER: "demo",
        PUBLIC_REQUIRE_PASSWORD: "true",
        PUBLIC_ACCESS_PASSWORD: "integration-password",
      },
      stdio: ["ignore", "pipe", "pipe"],
    });
    child.stdout.on("data", (chunk) => logs.push(chunk.toString()));
    child.stderr.on("data", (chunk) => logs.push(chunk.toString()));
    await waitForServer(localPort, child, logs);
    await waitForServer(publicPort, child, logs);

    const jsonHeaders = { "Content-Type": "application/json" };
    const [firstVote, secondVote] = await Promise.all([
      requestServer(localPort, "/api/places/place-vote/vote", {
        method: "POST",
        headers: jsonHeaders,
        body: { nickname: "小王", choice: "想去" },
      }),
      requestServer(localPort, "/api/places/place-vote/vote", {
        method: "POST",
        headers: jsonHeaders,
        body: { nickname: "小李", choice: "可以" },
      }),
    ]);
    assert.equal(firstVote.status, 200, firstVote.text);
    assert.equal(secondVote.status, 200, secondVote.text);

    const afterVotesResponse = await requestServer(localPort, "/api/state");
    assert.equal(afterVotesResponse.status, 200);
    const afterVotes = JSON.parse(afterVotesResponse.text);
    assert.deepEqual(afterVotes.places.find((item) => item.id === "place-vote").votes, {
      "小王": "想去",
      "小李": "可以",
    });

    const protectedDelete = await requestServer(localPort, "/api/links/link-protected", {
      method: "DELETE",
      headers: jsonHeaders,
    });
    assert.equal(protectedDelete.status, 409, protectedDelete.text);
    assert.match(protectedDelete.text, /已入选、已投票或人工修改/);

    const plainDelete = await requestServer(localPort, "/api/links/link-plain", {
      method: "DELETE",
      headers: jsonHeaders,
    });
    assert.equal(plainDelete.status, 200, plainDelete.text);

    const finalResponse = await requestServer(localPort, "/api/state");
    assert.equal(finalResponse.status, 200);
    const finalState = JSON.parse(finalResponse.text);
    assert.ok(finalState.links.some((item) => item.id === "link-protected"));
    assert.ok(finalState.places.some((item) => item.id === "place-protected"));
    assert.ok(!finalState.links.some((item) => item.id === "link-plain"));
    assert.ok(!finalState.places.some((item) => item.id === "place-plain"));

    const persistedState = JSON.parse(await readFile(dataPath, "utf8"));
    assert.deepEqual(persistedState.places.find((item) => item.id === "place-vote").votes, {
      "小王": "想去",
      "小李": "可以",
    });
    assert.ok(persistedState.links.some((item) => item.id === "link-protected"));
    assert.ok(!persistedState.links.some((item) => item.id === "link-plain"));
  } finally {
    if (child) await stopChild(child);
    await rm(fixtureRoot, { recursive: true, force: true });
  }
});
