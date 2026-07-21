"use client";

import { useEffect, useMemo, useState } from "react";
import type { FormEvent } from "react";

type LinkStatus = "等待处理" | "正在读取" | "AI分析中" | "已写入Excel" | "需要人工补充" | "处理失败";

type LinkRecord = {
  id: string;
  url: string;
  title: string;
  category: string;
  status: LinkStatus;
  readStatus: string;
  organizedStatus: string;
  factsFound: string[];
  missingFields: string[];
  resultNote: string;
  submitter: string;
  note: string;
  createdAt: string;
  updatedAt: string;
  summary: string;
  model: string;
  error?: string;
};

type PlaceDetails = {
  address: string;
  capacity: string;
  rooms: string;
  beds: string;
  bathrooms: string;
  twoNightTotal: string;
  environment: string;
  entireRental: string;
  kitchen: string;
  barbecue: string;
  bbqEquipment: string;
  breakfast: string;
  parking: string;
  transport: string;
  checkIn: string;
  checkOut: string;
  bookingStatus: string;
  reservation: string;
  openingHours: string;
  difficulty: string;
  horrorLevel: string;
  usage: string;
};

type Place = {
  id: string;
  sourceId?: string;
  name: string;
  category: string;
  area: string;
  price: number | null;
  priceLabel: string;
  duration: string;
  score: number;
  tags: string[];
  pros: string[];
  cons: string[];
  selected: boolean;
  sourceUrl: string;
  dataStatus: string;
  details: PlaceDetails;
};

type ItineraryItem = {
  time: string;
  endTime: string;
  title: string;
  subtitle: string;
  transport: string;
  cost: number;
  category: string;
  note: string;
  address: string;
  bookingStatus: string;
  sourceId: string;
};

type Reservation = {
  id: string;
  item: string;
  type: string;
  targetTime: string;
  status: string;
  owner: string;
  deadline: string;
  note: string;
};

type FinalPlan = {
  version: string;
  title: string;
  destination: string;
  dates: string;
  schedule: string;
  people: number;
  nights: number;
  perPersonBudget: string;
  summary: string;
  stay: {
    name: string;
    address: string;
    capacity: string;
    roomsBeds: string;
    twoNightTotal: string;
    checkInOut: string;
    barbecue: string;
    bbqEquipment: string;
    breakfast: string;
    sourceUrl: string;
  };
  itinerary: Array<ItineraryItem & { day: string; sourceUrl: string }>;
  reservations: Reservation[];
  updatedAt: string;
};

type AppState = {
  project: { name: string; destination: string; days: number; people: number; budget: number; status: string; tagline: string };
  tripProfile: { dates: string; schedule: string; groupSize: number; nights: number; stayPreference: string; barbecue: string; breakfasts: string[]; activity: string; accommodationBudget: string };
  links: LinkRecord[];
  places: Place[];
  itinerary: { day0: ItineraryItem[]; day1: ItineraryItem[]; day2: ItineraryItem[] };
  reservations: Reservation[];
  finalPlan: FinalPlan;
  settings: { provider: string; workbookPath: string; lastExcelSync: string };
};

const EMPTY_STATE: AppState = {
  project: { name: "扬州周末共创攻略", destination: "扬州", days: 2, people: 6, budget: 6000, status: "方案共创中", tagline: "周五晚集合，周末一起住、一起吃、一起玩。" },
  tripProfile: { dates: "待团队确认", schedule: "周五晚抵达 · 周日傍晚返程", groupSize: 6, nights: 2, stayPreference: "环境好、整租优先、至少 3 个独立睡眠空间", barbecue: "周六晚在民宿烧烤", breakfasts: ["周六早茶", "周日早餐"], activity: "6 人密室 / 团队活动", accommodationBudget: "待团队确认" },
  links: [],
  places: [],
  itinerary: { day0: [], day1: [], day2: [] },
  reservations: [],
  finalPlan: {
    version: "excel-home-v1",
    title: "6 人扬州周末旅行",
    destination: "扬州",
    dates: "待团队确认",
    schedule: "周五晚抵达 · 周日傍晚返程",
    people: 6,
    nights: 2,
    perPersonBudget: "待团队确认",
    summary: "周五晚集合，周末一起住、一起吃、一起玩。",
    stay: { name: "未选择", address: "待补充", capacity: "6 人", roomsBeds: "待补充", twoNightTotal: "待补充", checkInOut: "待补充", barbecue: "待确认", bbqEquipment: "待确认", breakfast: "待确认", sourceUrl: "" },
    itinerary: [],
    reservations: [],
    updatedAt: "",
  },
  settings: { provider: "演示分析", workbookPath: "", lastExcelSync: "" },
};

const tabs = [["plan", "行程首页"], ["collect", "链接收集"], ["library", "候选资料库"], ["excel", "Excel 工作台"], ["settings", "运行设置"]] as const;
const categories = ["自动识别", "攻略文章", "住宿", "密室/活动", "景点", "餐饮"];

function formatTime(value: string) {
  if (!value) return "尚未同步";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString("zh-CN", { hour12: false });
}

function statusTone(status: string) {
  if (["已写入Excel", "已整理", "成功读取", "已预订", "已完成", "已确定"].includes(status)) return "success";
  if (["处理失败", "需要人工补充", "未整理", "读取失败", "读取受限", "待补充", "待确认", "待预订"].includes(status)) return "warning";
  if (["正在读取", "AI分析中", "整理中", "等待读取", "等待处理"].includes(status)) return "active";
  return "muted";
}

function isPending(value: unknown) {
  const text = String(value ?? "").trim();
  return !text || /待|未选择|未确定|未知|尚未|占位|空位|暂无/.test(text);
}

function apiBase() {
  if (typeof window === "undefined") return "http://localhost:8787";
  if (!window.location.port || window.location.port === "8787") return window.location.origin;
  return `${window.location.protocol}//${window.location.hostname}:8787`;
}

function sourceName(url: string) {
  try { return new URL(url).hostname.replace(/^www\./, ""); } catch { return "链接"; }
}

function importantDetails(place: Place) {
  const d = place.details;
  if (place.category === "住宿") return [["详细地址", d.address], ["6 人容量", d.capacity], ["房间 / 床位", `${d.rooms} / ${d.beds}`], ["两晚总价", d.twoNightTotal], ["环境", d.environment], ["烧烤", d.barbecue], ["早餐", d.breakfast], ["预订", d.bookingStatus]];
  if (["活动", "景点"].includes(place.category)) return [["详细地址", d.address], ["适合人数", d.capacity], ["难度 / 恐怖", `${d.difficulty} / ${d.horrorLevel}`], ["营业时间", d.openingHours], ["预约", d.reservation], ["预订", d.bookingStatus]];
  if (place.category === "餐饮") return [["适合安排", d.usage], ["详细地址", d.address], ["营业时间", d.openingHours], ["预约", d.reservation], ["预订", d.bookingStatus]];
  return [["涉及区域", place.area], ["数据状态", place.dataStatus]];
}

function clonePlan(plan: FinalPlan): FinalPlan {
  return JSON.parse(JSON.stringify(plan)) as FinalPlan;
}

export default function Home() {
  const [state, setState] = useState<AppState>(EMPTY_STATE);
  const [activeTab, setActiveTab] = useState<(typeof tabs)[number][0]>("plan");
  const [urls, setUrls] = useState("");
  const [category, setCategory] = useState("自动识别");
  const [submitter, setSubmitter] = useState("团队成员");
  const [note, setNote] = useState("");
  const [sending, setSending] = useState(false);
  const [connected, setConnected] = useState(false);
  const [message, setMessage] = useState("");
  const [libraryFilter, setLibraryFilter] = useState("全部");
  const [backendBase, setBackendBase] = useState("http://localhost:8787");
  const [draftPlan, setDraftPlan] = useState<FinalPlan | null>(null);
  const [savingPlan, setSavingPlan] = useState(false);

  async function refresh(silent = false) {
    try {
      const response = await fetch(`${apiBase()}/api/state`, { cache: "no-store" });
      if (!response.ok) throw new Error("后台暂时不可用");
      setState(await response.json());
      setConnected(true);
    } catch {
      setConnected(false);
      if (!silent) setMessage("正在等待本地后台连接，页面暂时展示行程结构。");
    }
  }

  useEffect(() => {
    const initial = window.setTimeout(() => { setBackendBase(apiBase()); refresh(); }, 0);
    const timer = window.setInterval(() => refresh(true), 3500);
    return () => { window.clearTimeout(initial); window.clearInterval(timer); };
  }, []);

  const stats = useMemo(() => ({
    links: state.links.length,
    completed: state.links.filter((item) => item.organizedStatus === "已整理").length,
    unorganized: state.links.filter((item) => item.organizedStatus === "未整理").length,
    processing: state.links.filter((item) => item.organizedStatus === "整理中").length,
    selected: state.places.filter((item) => item.selected).length,
  }), [state]);

  const finalPlan = state.finalPlan || EMPTY_STATE.finalPlan;
  const finalFields = useMemo(() => [
    finalPlan.title,
    finalPlan.destination,
    finalPlan.dates,
    finalPlan.schedule,
    finalPlan.people,
    finalPlan.nights,
    finalPlan.perPersonBudget,
    finalPlan.summary,
    finalPlan.stay.name,
    finalPlan.stay.address,
    finalPlan.stay.capacity,
    finalPlan.stay.roomsBeds,
    finalPlan.stay.twoNightTotal,
    finalPlan.stay.checkInOut,
    finalPlan.stay.barbecue,
    finalPlan.stay.bbqEquipment,
    finalPlan.stay.breakfast,
    finalPlan.stay.sourceUrl,
  ], [finalPlan]);
  const finalProgress = useMemo(() => ({
    confirmed: finalFields.filter((item) => !isPending(item)).length,
    pending: finalFields.filter(isPending).length,
  }), [finalFields]);
  const groupedFinalItinerary = useMemo(() => ["周五晚上", "周六", "周日"].map((day) => ({
    day,
    items: finalPlan.itinerary.filter((item) => item.day === day),
  })), [finalPlan.itinerary]);

  const filteredPlaces = useMemo(() => state.places.filter((place) => libraryFilter === "全部" || place.category === libraryFilter), [state.places, libraryFilter]);

  async function submitLinks(event: FormEvent) {
    event.preventDefault();
    const list = urls.split(/\n|\s+/).map((item) => item.trim()).filter(Boolean);
    if (!list.length) return setMessage("请先粘贴至少一个链接。");
    setSending(true);
    setMessage("");
    try {
      const response = await fetch(`${apiBase()}/api/links`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ urls: list, category, submitter, note }) });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "提交失败");
      setUrls("");
      setNote("");
      setMessage(`已接收 ${result.created} 个新链接；打不开的链接也会保留在处理报告中。`);
      await refresh(true);
      setActiveTab("collect");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "提交失败，请检查后台服务。");
    } finally { setSending(false); }
  }

  async function retryLink(id: string) {
    await fetch(`${apiBase()}/api/links/${id}/retry`, { method: "POST" });
    setMessage("已重新加入分析队列。");
    refresh(true);
  }

  async function syncExcel() {
    setMessage("正在读取 Excel 的人工修改……");
    try {
      const response = await fetch(`${apiBase()}/api/sync/from-excel`, { method: "POST" });
      if (!response.ok) throw new Error("同步失败");
      setMessage("Excel“行程首页”的最新内容已同步到网站。");
      refresh(true);
    } catch { setMessage("暂时无法读取 Excel，请确认文件没有被移动或占用。"); }
  }

  function startPlanEditing() {
    setDraftPlan(clonePlan(finalPlan));
    setMessage("");
  }

  async function savePlan() {
    if (!draftPlan) return;
    setSavingPlan(true);
    setMessage("");
    try {
      const response = await fetch(`${apiBase()}/api/final-plan`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ finalPlan: draftPlan, baseUpdatedAt: draftPlan.updatedAt }),
      });
      const result = await response.json();
      if (!response.ok) {
        if (result.state) setState(result.state);
        throw new Error(result.error || "保存失败");
      }
      setState(result.state);
      setDraftPlan(null);
      setMessage("已保存：网页和本地 Excel 已同步更新。");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "保存失败，请稍后重试。");
    } finally {
      setSavingPlan(false);
    }
  }

  return (
    <main>
      <header className="topbar">
        <button className="brand" onClick={() => setActiveTab("plan")} aria-label="返回行程首页"><span className="brand-seal">扬</span><span><strong>下扬州</strong><small>团队旅行共创台</small></span></button>
        <nav aria-label="网站主导航">{tabs.map(([key, label]) => <button key={key} className={activeTab === key ? "nav-active" : ""} onClick={() => setActiveTab(key)}>{label}</button>)}</nav>
        <div className={`connection ${connected ? "online" : "offline"}`}><span aria-hidden="true" />{connected ? "本地后台运行中" : "等待后台"}</div>
      </header>

      {message && <div className="toast" role="status"><span>{message}</span><button onClick={() => setMessage("")} aria-label="关闭提示">×</button></div>}

      {draftPlan && <PlanEditor plan={draftPlan} saving={savingPlan} onChange={setDraftPlan} onCancel={() => setDraftPlan(null)} onSave={savePlan} />}

      {activeTab === "plan" && <>
        <section className="excel-source-banner">
          <div className="shell">
            <div><span className="excel-source-icon">↔</span><p><strong>网页与 Excel「行程首页」双向同步</strong><small>可以直接编辑网页；保存后会立即写入本地 Excel。改 Excel 后也会自动更新网页。</small></p></div>
            <div className="excel-source-actions"><button className="primary edit-plan-button" onClick={startPlanEditing}>编辑最终方案</button><a className="small-button" href={`${backendBase}/api/download/excel`}>打开 Excel</a><button className="small-button" onClick={syncExcel}>同步 Excel</button></div>
          </div>
        </section>

        <section className="final-hero shell">
          <div className="final-hero-copy">
            <p className="eyebrow">FINAL PLAN · EXCEL DRIVEN</p>
            <span className="final-kicker">{finalPlan.destination} · {finalPlan.schedule}</span>
            <h1>{finalPlan.title}</h1>
            <p>{finalPlan.summary}</p>
            <div className="final-quick-facts">
              <div><span>目的地</span><strong>{finalPlan.destination}</strong></div>
              <div><span>日期</span><strong className={isPending(finalPlan.dates) ? "pending-value" : ""}>{finalPlan.dates}</strong></div>
              <div><span>人数 / 住宿</span><strong>{finalPlan.people} 人 · {finalPlan.nights} 晚</strong></div>
              <div><span>人均预算</span><strong className={isPending(finalPlan.perPersonBudget) ? "pending-value" : ""}>{finalPlan.perPersonBudget}</strong></div>
            </div>
          </div>
          <aside className="completion-card">
            <span className="paper-tag">最终方案完成度</span>
            <div className="completion-number"><strong>{finalProgress.confirmed}</strong><span>/ {finalFields.length}</span></div>
            <div className="completion-bar"><i style={{ width: `${Math.round(finalProgress.confirmed / Math.max(finalFields.length, 1) * 100)}%` }} /></div>
            <div className="completion-legend"><span className="done-dot">已确定 {finalProgress.confirmed}</span><span className="pending-dot">待补充 {finalProgress.pending}</span></div>
            <p>这里统计基础信息和住宿信息。橙色内容还需要在 Excel 首页补充。</p>
          </aside>
        </section>

        <section className="shell final-overview-section">
          <div className="final-section-heading"><div><p className="eyebrow">01 · OVERVIEW</p><h2>最终方案信息</h2></div><p>只有「行程首页」里的内容会进入这里；后面的工作表只是候选资料和整理记录。</p></div>
          <div className="final-overview-grid">
            <article className="final-panel">
              <header><div><span>基础信息</span><h3>这次旅行</h3></div><StatusPill value={finalPlan.dates} /></header>
              <div className="final-field-list">
                <FinalField label="目的地" value={finalPlan.destination} />
                <FinalField label="出行日期" value={finalPlan.dates} />
                <FinalField label="时间安排" value={finalPlan.schedule} />
                <FinalField label="团队人数" value={`${finalPlan.people} 人`} />
                <FinalField label="住宿晚数" value={`${finalPlan.nights} 晚`} />
                <FinalField label="人均预算" value={finalPlan.perPersonBudget} />
              </div>
            </article>
            <article className="final-panel stay-panel">
              <header><div><span>住宿确认</span><h3>{finalPlan.stay.name}</h3></div><StatusPill value={finalPlan.stay.name} /></header>
              <div className="final-field-list two-column">
                <FinalField label="详细地址" value={finalPlan.stay.address} />
                <FinalField label="容纳人数" value={finalPlan.stay.capacity} />
                <FinalField label="房间 / 床位" value={finalPlan.stay.roomsBeds} />
                <FinalField label="两晚总价" value={finalPlan.stay.twoNightTotal} />
                <FinalField label="入住 / 退房" value={finalPlan.stay.checkInOut} />
                <FinalField label="周六烧烤" value={finalPlan.stay.barbecue} />
                <FinalField label="烧烤设备" value={finalPlan.stay.bbqEquipment} />
                <FinalField label="早餐条件" value={finalPlan.stay.breakfast} />
              </div>
              {finalPlan.stay.sourceUrl && <a className="source-link" href={finalPlan.stay.sourceUrl} target="_blank" rel="noreferrer">查看已选住宿原链接 →</a>}
            </article>
          </div>
        </section>

        <section className="final-itinerary-section">
          <div className="shell">
            <div className="final-section-heading"><div><p className="eyebrow">02 · ITINERARY</p><h2>周五晚到周日</h2></div><p>时间、地点和预订状态均来自 Excel 首页；直接改表格即可调整网页流程。</p></div>
            <div className="final-day-grid">{groupedFinalItinerary.map(({ day, items }) => <FinalDay key={day} day={day} items={items} />)}</div>
          </div>
        </section>

        <section className="shell final-reservation-section">
          <div className="final-section-heading"><div><p className="eyebrow">03 · CHECKLIST</p><h2>确认与预订</h2></div><p>网站只记录进度，不会自动付款或下单。</p></div>
          <div className="reservation-list">{finalPlan.reservations.length ? finalPlan.reservations.map((item) => <article key={item.id}><span className="reservation-type">{item.type}</span><div><h3>{item.item}</h3><p>{item.targetTime} · {item.note}</p></div><div className="reservation-owner"><small>{item.owner}</small><span>截至 {item.deadline}</span></div><b className={`status-badge ${statusTone(item.status)}`}><i />{item.status}</b></article>) : <div className="empty-state">还没有预订事项，请在 Excel 首页底部添加。</div>}</div>
        </section>

        <section className="source-separation"><div className="shell"><div><p className="eyebrow">SOURCE SEPARATION</p><h2>最终方案和候选资料已经分开</h2><p>朋友投递的链接会先进入候选资料库，不会自动挤进最终方案。确认采用后，再把内容填写到 Excel「行程首页」。</p></div><div className="source-separation-actions"><button className="primary" onClick={() => setActiveTab("collect")}>继续投递链接</button><button className="small-button" onClick={() => setActiveTab("library")}>查看 {state.places.length} 条候选资料</button><span>{stats.completed} 条已整理 · {stats.unorganized} 条未整理 · {stats.processing} 条处理中</span></div></div></section>
      </>}

      {activeTab === "collect" && <section className="shell workspace-page">
        <div className="page-title"><p className="eyebrow">LINK INBOX</p><h1>把大家觉得好的链接都放进来。</h1><p>民宿、酒店、攻略、餐厅和密室都可以。每条链接都会留下读取和整理结果，打不开也不会被悄悄忽略。</p></div>
        <div className="report-summary"><div><span>全部链接</span><strong>{stats.links}</strong></div><div className="good"><span>已成功整理</span><strong>{stats.completed}</strong></div><div className="warn"><span>未整理</span><strong>{stats.unorganized}</strong></div><div><span>处理中</span><strong>{stats.processing}</strong></div></div>
        <div className="collect-layout">
          <form className="collect-form" onSubmit={submitLinks}>
            <label htmlFor="urls">链接列表</label><textarea id="urls" value={urls} onChange={(event) => setUrls(event.target.value)} placeholder={"粘贴攻略、民宿、密室或餐厅链接……\n每行一个，也可以一次粘贴多个"} />
            <div className="form-row"><label>内容类型<select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((item) => <option key={item}>{item}</option>)}</select></label><label>提交人<input value={submitter} onChange={(event) => setSubmitter(event.target.value)} /></label></div>
            <label>重点关注<input value={note} onChange={(event) => setNote(event.target.value)} placeholder="例如：重点看 6 人能否住、周六能否烧烤" /></label>
            <button className="primary wide" disabled={sending}>{sending ? "正在提交……" : "开始后台整理"}</button><p className="form-hint">需要登录、验证码或限制抓取的平台会标记为“读取受限 / 未整理”，系统不会编造内容。</p>
          </form>
          <aside className="pipeline-card"><span className="paper-tag">自动流水线</span><h3>每个链接的 6 个步骤</h3>{["保留原始链接与提交人", "尝试读取公开网页", "DeepSeek 提取事实", "列出已提取与缺失字段", "写入 Excel 对应工作表", "网站实时展示最新结果"].map((item, index) => <div className="pipeline-step" key={item}><b>{String(index + 1).padStart(2, "0")}</b><span>{item}</span></div>)}</aside>
        </div>
        <div className="task-panel"><div className="panel-heading"><div><h2>链接处理报告</h2><p>成功和失败都在这里；完整报告也已写入 Excel。</p></div><button className="small-button" onClick={() => refresh()}>刷新状态</button></div><div className="task-list">
          {state.links.length === 0 && <div className="empty-state">还没有团队链接。提交第一个链接后，处理过程会显示在这里。</div>}
          {state.links.map((item) => <article className="task-report" key={item.id}>
            <div className="task-report-main"><span className="source-icon">链</span><div><div className="task-title-line"><strong>{item.title || sourceName(item.url)}</strong><span>{item.category}</span></div><a href={item.url} target="_blank" rel="noreferrer">{item.url}</a><p>{item.resultNote}</p></div></div>
            <div className="task-statuses"><span className={`status-badge ${statusTone(item.readStatus)}`}><i />{item.readStatus}</span><span className={`status-badge ${statusTone(item.organizedStatus)}`}><i />{item.organizedStatus}</span><small>{formatTime(item.updatedAt)}</small></div>
            <div className="fact-columns"><div><b>已提取</b><p>{item.factsFound.length ? item.factsFound.join(" · ") : "暂无"}</p></div><div><b>还缺少</b><p>{item.missingFields.length ? item.missingFields.join(" · ") : "无明显缺失"}</p></div></div>
            <div className="task-report-foot"><span>{item.submitter} · {item.model || "等待分配模型"}</span>{(item.organizedStatus === "未整理" || item.status === "处理失败") && <button className="small-button" onClick={() => retryLink(item.id)}>重新尝试</button>}</div>
          </article>)}
        </div></div>
      </section>}

      {activeTab === "library" && <section className="shell workspace-page">
        <div className="page-title split"><div><p className="eyebrow">CANDIDATE LIBRARY</p><h1>候选资料库</h1><p>这里是 AI 整理出的备选内容，不会直接出现在最终方案。团队确认采用后，请把关键信息填写到 Excel「行程首页」。</p></div><div className="library-count"><strong>{state.places.length}</strong><span>个候选项</span></div></div>
        <div className="candidate-notice"><strong>候选 ≠ 已确定</strong><span>卡片显示的是参考信息；最终以 Excel 首页和网站「行程首页」为准。</span></div>
        <div className="filter-bar">{["全部", "住宿", "活动", "景点", "餐饮", "攻略"].map((item) => <button key={item} className={libraryFilter === item ? "selected" : ""} onClick={() => setLibraryFilter(item)}>{item}</button>)}</div>
        <div className="place-grid">{filteredPlaces.map((place) => <article className={`place-card ${place.category === "住宿" ? "stay-card" : ""}`} key={place.id}>
          <div className="place-top"><span className="place-category">{place.category}</span><span className={place.selected ? "selected-mark" : "candidate-mark"}>{place.selected ? "已入选" : "候选"}</span></div>
          <h3>{place.name}</h3><p className="place-meta">{place.area} · {place.duration || "时长待核实"}</p>
          <div className="detail-grid">{importantDetails(place).map(([label, detail]) => <div key={label}><span>{label}</span><strong>{detail}</strong></div>)}</div>
          <div className="place-tags">{place.tags.map((tag) => <span key={tag}>{tag}</span>)}</div>
          <div className="pros-cons"><div><b>适合的理由</b><p>{place.pros[0] || "等待分析"}</p></div><div><b>需要留意</b><p>{place.cons[0] || "等待核实"}</p></div></div>
          <div className="place-footer"><div><span>参考预算</span><strong>{place.priceLabel}</strong></div><div className="score"><span>推荐度</span><strong>{place.score.toFixed(1)}</strong></div></div>
          <div className="data-state">数据状态：{place.dataStatus}</div>{place.sourceUrl && <a className="source-link" href={place.sourceUrl} target="_blank" rel="noreferrer">查看原始链接 →</a>}
        </article>)}</div>
      </section>}

      {activeTab === "excel" && <section className="shell workspace-page">
        <div className="page-title"><p className="eyebrow">EXCEL WORKSPACE</p><h1>先改 Excel 首页，网站跟着更新。</h1><p>第一张「行程首页」是最终方案的唯一来源；其余工作表保存候选、链接处理结果和参考信息。</p></div>
        <div className="excel-hero"><div className="excel-file-icon">X</div><div className="excel-file-info"><span>当前工作簿</span><h2>扬州团队旅行攻略.xlsx</h2><p>最近同步：{formatTime(state.settings.lastExcelSync)}</p></div><div className="excel-actions"><a className="primary" href={`${backendBase}/api/download/excel`}>下载 / 打开 Excel</a><button className="small-button" onClick={syncExcel}>读取最新修改</button></div></div>
        <div className="sheet-grid">{[["行程首页","唯一最终方案；黄色单元格可修改"],["项目总览","完成度、候选和下一步重点"],["链接汇总","所有原始链接与完整字段"],["处理报告","成功、失败和缺失信息"],["住宿候选","6 人两晚、烧烤与房型比较"],["活动候选","密室、景点和预约条件"],["餐饮候选","两顿早餐与正餐备选"],["攻略文章","文章摘要、避坑和缺失项"],["两日行程","旧版明细与参考计划"],["预订清单","旧版负责人、期限与状态明细"],["项目设置","人数、住宿偏好与模型设置"]].map(([name, desc], index) => <div className={`sheet-card ${index === 0 ? "primary-sheet" : ""}`} key={name}><b>{String(index + 1).padStart(2, "0")}</b><div><strong>{name}</strong><p>{desc}</p></div></div>)}</div>
        <div className="sync-explainer"><div><span>团队</span><b>提交链接</b></div><i>→</i><div><span>候选区</span><b>DeepSeek 整理事实</b></div><i>→</i><div className="highlight"><span>Excel 第一张表</span><b>人工确认最终方案</b></div><i>→</i><div><span>网站首页</span><b>按表格自动生成</b></div></div>
      </section>}

      {activeTab === "settings" && <section className="shell workspace-page narrow-page">
        <div className="page-title"><p className="eyebrow">LOCAL SERVICE</p><h1>运行与模型设置</h1><p>API 密钥只保存在这台电脑，不会出现在网页或 Excel 中。</p></div>
        <div className="settings-card"><div className="setting-row"><div><span>当前分析方式</span><strong>{state.settings.provider}</strong></div><span className="setting-state good">已配置</span></div><div className="setting-row"><div><span>本地后台</span><strong>{connected ? "正在运行" : "未连接"}</strong></div><span className={`setting-state ${connected ? "good" : ""}`}>{connected ? "可接收链接" : "请启动服务"}</span></div><div className="setting-row"><div><span>Excel 自动同步</span><strong>每 4 秒检查一次修改</strong></div><span className="setting-state good">已开启</span></div></div>
        <div className="scope-note"><strong>访问范围</strong><p>后台和 Excel 仍保存在你的电脑；外地朋友通过密码保护的 HTTPS 公网网址访问，无需注册账号，也不需要处于同一个 Wi-Fi。</p></div>
        <div className="scope-note safety"><strong>预订边界</strong><p>网站负责整理、比较和列出待办，不会自动付款。住宿、早餐、密室和门票都需要团队确认真实价格与取消政策后再下单。</p></div>
      </section>}

      <footer><div className="shell"><span>下扬州 · 团队旅行共创台</span><span>本地数据 · Excel 可编辑 · DeepSeek 整理</span></div></footer>
    </main>
  );
}

function PlanEditor({ plan, saving, onChange, onCancel, onSave }: {
  plan: FinalPlan;
  saving: boolean;
  onChange: (plan: FinalPlan) => void;
  onCancel: () => void;
  onSave: () => void;
}) {
  function updatePlan(field: keyof FinalPlan, value: string | number) {
    onChange({ ...plan, [field]: value } as FinalPlan);
  }

  function updateStay(field: keyof FinalPlan["stay"], value: string) {
    onChange({ ...plan, stay: { ...plan.stay, [field]: value } });
  }

  function updateItinerary(index: number, field: keyof FinalPlan["itinerary"][number], value: string | number) {
    const itinerary = plan.itinerary.map((item, itemIndex) => itemIndex === index ? { ...item, [field]: value } : item);
    onChange({ ...plan, itinerary });
  }

  function updateReservation(index: number, field: keyof Reservation, value: string) {
    const reservations = plan.reservations.map((item, itemIndex) => itemIndex === index ? { ...item, [field]: value } : item);
    onChange({ ...plan, reservations });
  }

  function addItinerary() {
    onChange({
      ...plan,
      itinerary: [...plan.itinerary, { day: "周六", time: "", endTime: "", category: "安排", title: "新增安排", subtitle: "", address: "待补充", transport: "待补充", cost: 0, bookingStatus: "待确认", sourceUrl: "", sourceId: "", note: "" }],
    });
  }

  function addReservation() {
    onChange({
      ...plan,
      reservations: [...plan.reservations, { id: `reserve-web-${Date.now()}`, item: "新增待办", type: "待分类", targetTime: "待确认", status: "待确认", owner: "待认领", deadline: "待确认", note: "" }],
    });
  }

  return <div className="plan-editor-overlay" role="dialog" aria-modal="true" aria-labelledby="plan-editor-title">
    <form className="plan-editor" onSubmit={(event) => { event.preventDefault(); onSave(); }}>
      <header className="plan-editor-header">
        <div><p className="eyebrow">WEB EDITOR · EXCEL SYNC</p><h1 id="plan-editor-title">编辑最终方案</h1><span>保存后会立即更新网站，并写入本地 Excel 第一张表「行程首页」。</span></div>
        <div className="plan-editor-actions"><button type="button" className="small-button" onClick={onCancel} disabled={saving}>取消</button><button type="submit" className="primary" disabled={saving}>{saving ? "正在同步……" : "保存并同步 Excel"}</button></div>
      </header>

      <div className="plan-editor-body">
        <section className="editor-section">
          <div className="editor-section-title"><b>01</b><div><h2>基础信息</h2><p>控制网站标题、日期、人数和顶部介绍。</p></div></div>
          <div className="editor-grid basic-grid">
            <EditorField label="方案标题" value={plan.title} onChange={(value) => updatePlan("title", value)} wide />
            <EditorField label="目的地" value={plan.destination} onChange={(value) => updatePlan("destination", value)} />
            <EditorField label="出行日期" value={plan.dates} onChange={(value) => updatePlan("dates", value)} />
            <EditorField label="行程结构" value={plan.schedule} onChange={(value) => updatePlan("schedule", value)} wide />
            <EditorField label="同行人数" value={plan.people} inputType="number" onChange={(value) => updatePlan("people", Number(value) || 1)} />
            <EditorField label="住宿晚数" value={plan.nights} inputType="number" onChange={(value) => updatePlan("nights", Number(value) || 1)} />
            <EditorField label="人均预算" value={plan.perPersonBudget} onChange={(value) => updatePlan("perPersonBudget", value)} />
            <EditorField label="方案说明" value={plan.summary} onChange={(value) => updatePlan("summary", value)} multiline wide />
          </div>
        </section>

        <section className="editor-section">
          <div className="editor-section-title"><b>02</b><div><h2>住宿信息</h2><p>民宿名称、地址、烧烤和早餐都会同步到 Excel。</p></div></div>
          <div className="editor-grid">
            <EditorField label="民宿名称" value={plan.stay.name} onChange={(value) => updateStay("name", value)} wide />
            <EditorField label="详细地址" value={plan.stay.address} onChange={(value) => updateStay("address", value)} wide />
            <EditorField label="适合人数" value={plan.stay.capacity} onChange={(value) => updateStay("capacity", value)} />
            <EditorField label="房间 / 床位" value={plan.stay.roomsBeds} onChange={(value) => updateStay("roomsBeds", value)} />
            <EditorField label="两晚总价" value={plan.stay.twoNightTotal} onChange={(value) => updateStay("twoNightTotal", value)} />
            <EditorField label="入住 / 退房" value={plan.stay.checkInOut} onChange={(value) => updateStay("checkInOut", value)} />
            <EditorField label="能否烧烤" value={plan.stay.barbecue} onChange={(value) => updateStay("barbecue", value)} multiline />
            <EditorField label="烧烤设备 / 费用" value={plan.stay.bbqEquipment} onChange={(value) => updateStay("bbqEquipment", value)} multiline />
            <EditorField label="早餐安排" value={plan.stay.breakfast} onChange={(value) => updateStay("breakfast", value)} multiline wide />
            <EditorField label="民宿原始链接" value={plan.stay.sourceUrl} inputType="url" onChange={(value) => updateStay("sourceUrl", value)} wide />
          </div>
        </section>

        <section className="editor-section">
          <div className="editor-section-title with-action"><b>03</b><div><h2>周末行程</h2><p>可以增加、删除和调整每一项安排。</p></div><button type="button" className="small-button" onClick={addItinerary}>＋ 添加行程</button></div>
          <div className="editor-card-list">
            {plan.itinerary.map((item, index) => <article className="itinerary-editor-card" key={`${item.day}-${index}`}>
              <header><strong>{String(index + 1).padStart(2, "0")} · {item.title || "未命名安排"}</strong><button type="button" className="remove-button" onClick={() => onChange({ ...plan, itinerary: plan.itinerary.filter((_, itemIndex) => itemIndex !== index) })}>删除</button></header>
              <div className="editor-grid compact-grid">
                <label className="editor-field"><span>日期</span><select value={item.day} onChange={(event) => updateItinerary(index, "day", event.target.value)}><option>周五晚上</option><option>周六</option><option>周日</option></select></label>
                <EditorField label="开始时间" value={item.time} inputType="time" onChange={(value) => updateItinerary(index, "time", value)} />
                <EditorField label="结束时间" value={item.endTime} inputType="time" onChange={(value) => updateItinerary(index, "endTime", value)} />
                <EditorField label="类型" value={item.category} onChange={(value) => updateItinerary(index, "category", value)} />
                <EditorField label="安排名称" value={item.title} onChange={(value) => updateItinerary(index, "title", value)} wide />
                <EditorField label="详细说明" value={item.subtitle} onChange={(value) => updateItinerary(index, "subtitle", value)} multiline wide />
                <EditorField label="地址" value={item.address} onChange={(value) => updateItinerary(index, "address", value)} wide />
                <EditorField label="交通" value={item.transport} onChange={(value) => updateItinerary(index, "transport", value)} />
                <EditorField label="费用 / 人" value={item.cost} inputType="number" onChange={(value) => updateItinerary(index, "cost", Number(value) || 0)} />
                <EditorField label="预订状态" value={item.bookingStatus} onChange={(value) => updateItinerary(index, "bookingStatus", value)} />
                <EditorField label="来源链接" value={item.sourceUrl} inputType="url" onChange={(value) => updateItinerary(index, "sourceUrl", value)} wide />
                <EditorField label="备注" value={item.note} onChange={(value) => updateItinerary(index, "note", value)} multiline wide />
              </div>
            </article>)}
            {!plan.itinerary.length && <div className="editor-empty">还没有行程，点击“添加行程”开始安排。</div>}
          </div>
        </section>

        <section className="editor-section">
          <div className="editor-section-title with-action"><b>04</b><div><h2>确认与预订</h2><p>网站只记录进度，不会自动付款或下单。</p></div><button type="button" className="small-button" onClick={addReservation}>＋ 添加待办</button></div>
          <div className="editor-card-list">
            {plan.reservations.map((item, index) => <article className="reservation-editor-card" key={item.id || index}>
              <header><strong>{item.item || "未命名待办"}</strong><button type="button" className="remove-button" onClick={() => onChange({ ...plan, reservations: plan.reservations.filter((_, itemIndex) => itemIndex !== index) })}>删除</button></header>
              <div className="editor-grid compact-grid">
                <EditorField label="待办事项" value={item.item} onChange={(value) => updateReservation(index, "item", value)} wide />
                <EditorField label="类型" value={item.type} onChange={(value) => updateReservation(index, "type", value)} />
                <EditorField label="计划时间" value={item.targetTime} onChange={(value) => updateReservation(index, "targetTime", value)} />
                <EditorField label="当前状态" value={item.status} onChange={(value) => updateReservation(index, "status", value)} />
                <EditorField label="负责人" value={item.owner} onChange={(value) => updateReservation(index, "owner", value)} />
                <EditorField label="完成期限" value={item.deadline} onChange={(value) => updateReservation(index, "deadline", value)} />
                <EditorField label="核对说明" value={item.note} onChange={(value) => updateReservation(index, "note", value)} multiline wide />
              </div>
            </article>)}
            {!plan.reservations.length && <div className="editor-empty">还没有预订待办，可以从住宿、早餐或密室开始添加。</div>}
          </div>
        </section>
      </div>

      <footer className="plan-editor-footer"><span>本次保存会同时覆盖网站最终方案和 Excel「行程首页」。</span><div className="plan-editor-actions"><button type="button" className="small-button" onClick={onCancel} disabled={saving}>取消</button><button type="submit" className="primary" disabled={saving}>{saving ? "正在同步……" : "保存并同步 Excel"}</button></div></footer>
    </form>
  </div>;
}

function EditorField({ label, value, onChange, inputType = "text", multiline = false, wide = false }: {
  label: string;
  value: string | number;
  onChange: (value: string) => void;
  inputType?: string;
  multiline?: boolean;
  wide?: boolean;
}) {
  return <label className={`editor-field ${wide ? "wide" : ""}`}><span>{label}</span>{multiline
    ? <textarea value={value} onChange={(event) => onChange(event.target.value)} />
    : <input type={inputType} min={inputType === "number" ? 0 : undefined} value={value} onChange={(event) => onChange(event.target.value)} />}</label>;
}

function StatusPill({ value }: { value: unknown }) {
  const pending = isPending(value);
  return <span className={`final-state ${pending ? "pending" : "confirmed"}`}><i />{pending ? "待补充" : "已确定"}</span>;
}

function FinalField({ label, value }: { label: string; value: unknown }) {
  return <div className="final-field"><span>{label}</span><strong className={isPending(value) ? "pending-value" : ""}>{String(value || "待补充")}</strong><StatusPill value={value} /></div>;
}

function FinalDay({ day, items }: { day: string; items: FinalPlan["itinerary"] }) {
  return <article className="final-day-card">
    <header><span>{day === "周五晚上" ? "FRI" : day === "周六" ? "SAT" : "SUN"}</span><h3>{day}</h3><b>{items.length} 项安排</b></header>
    <div className="final-day-items">{items.length ? items.map((item, index) => <div className="final-day-item" key={`${day}-${item.time}-${item.title}-${index}`}>
      <time>{item.time}{item.endTime ? `–${item.endTime}` : ""}</time><i /><div><span>{item.category}</span><h4>{item.title}</h4><p>{item.subtitle || item.note}</p><small>{[item.address, item.transport].filter(Boolean).join(" · ") || "地点待补充"}</small><b className={`status-badge ${statusTone(item.bookingStatus)}`}><i />{item.bookingStatus || "待确认"}</b></div>
    </div>) : <div className="day-empty">请在 Excel 首页添加当天安排</div>}</div>
  </article>;
}
