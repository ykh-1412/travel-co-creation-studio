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

type AppState = {
  project: { name: string; destination: string; days: number; people: number; budget: number; status: string; tagline: string };
  tripProfile: { dates: string; schedule: string; groupSize: number; nights: number; stayPreference: string; barbecue: string; breakfasts: string[]; activity: string; accommodationBudget: string };
  links: LinkRecord[];
  places: Place[];
  itinerary: { day0: ItineraryItem[]; day1: ItineraryItem[]; day2: ItineraryItem[] };
  reservations: Reservation[];
  settings: { provider: string; workbookPath: string; lastExcelSync: string };
};

const EMPTY_STATE: AppState = {
  project: { name: "扬州周末共创攻略", destination: "扬州", days: 2, people: 6, budget: 6000, status: "方案共创中", tagline: "周五晚集合，周末一起住、一起吃、一起玩。" },
  tripProfile: { dates: "待团队确认", schedule: "周五晚抵达 · 周日傍晚返程", groupSize: 6, nights: 2, stayPreference: "环境好、整租优先、至少 3 个独立睡眠空间", barbecue: "周六晚在民宿烧烤", breakfasts: ["周六早茶", "周日早餐"], activity: "6 人密室 / 团队活动", accommodationBudget: "待团队确认" },
  links: [],
  places: [],
  itinerary: { day0: [], day1: [], day2: [] },
  reservations: [],
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
  if (["已写入Excel", "已整理", "成功读取", "已预订", "已完成"].includes(status)) return "success";
  if (["处理失败", "需要人工补充", "未整理", "读取失败", "读取受限"].includes(status)) return "warning";
  if (["正在读取", "AI分析中", "整理中", "等待读取", "等待处理"].includes(status)) return "active";
  return "muted";
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
    const initial = window.setTimeout(() => refresh(), 0);
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
      setMessage("Excel 中的候选、行程和预订状态已同步到网站。");
      refresh(true);
    } catch { setMessage("暂时无法读取 Excel，请确认文件没有被移动或占用。"); }
  }

  return (
    <main>
      <header className="topbar">
        <button className="brand" onClick={() => setActiveTab("plan")} aria-label="返回行程首页"><span className="brand-seal">扬</span><span><strong>下扬州</strong><small>团队旅行共创台</small></span></button>
        <nav aria-label="网站主导航">{tabs.map(([key, label]) => <button key={key} className={activeTab === key ? "nav-active" : ""} onClick={() => setActiveTab(key)}>{label}</button>)}</nav>
        <div className={`connection ${connected ? "online" : "offline"}`}><span aria-hidden="true" />{connected ? "本地后台运行中" : "等待后台"}</div>
      </header>

      {message && <div className="toast" role="status"><span>{message}</span><button onClick={() => setMessage("")} aria-label="关闭提示">×</button></div>}

      {activeTab === "plan" && <>
        <section className="hero shell">
          <div className="hero-copy">
            <p className="eyebrow">FRIDAY NIGHT → SUNDAY · YANGZHOU</p>
            <h1>六个人，两晚住在一起，<br />把<em>扬州</em>过成一个周末。</h1>
            <p className="hero-intro">{state.project.tagline} 日期：{state.tripProfile.dates}。</p>
            <div className="trip-facts"><div><strong>{state.tripProfile.groupSize}</strong><span>人同行</span></div><div><strong>{state.tripProfile.nights}</strong><span>晚住宿</span></div><div><strong>1</strong><span>场民宿烧烤</span></div></div>
            <div className="hero-actions"><button className="primary" onClick={() => setActiveTab("collect")}>投递新链接</button><button className="text-button" onClick={() => setActiveTab("library")}>查看 {state.places.length} 条候选资料 →</button></div>
          </div>
          <div className="route-card" aria-label="周末路线摘要">
            <div className="route-card-head"><div><span>当前框架</span><strong>住在一起的扬州周末</strong></div><span className="paper-tag">{state.project.status}</span></div>
            <div className="arrival-line"><b>FRI</b><div><strong>周五晚上抵达 · 入住两晚</strong><small>{state.itinerary.day0.map((item) => item.title).join(" → ")}</small></div></div>
            <div className="route-line"><div className="route-day"><b>六</b><span>SAT</span></div><div className="route-stops">{state.itinerary.day1.slice(0, 4).map((item, index) => <div key={`${item.time}-${item.title}`}><i>{index + 1}</i><span><strong>{item.title}</strong><small>{item.time} · {item.bookingStatus}</small></span></div>)}</div></div>
            <div className="route-divider" />
            <div className="route-line compact"><div className="route-day"><b>日</b><span>SUN</span></div><div className="route-stops horizontal">{state.itinerary.day2.slice(0, 4).map((item) => <span key={`${item.time}-${item.title}`}>{item.title}</span>)}</div></div>
            <p className="route-note">这是可继续修改的第一版。动态价格、营业时间和预订结果以最终核实为准。</p>
          </div>
        </section>

        <section className="status-strip"><div className="shell status-grid"><div><span>团队链接</span><strong>{stats.links}</strong><small>全部保留记录</small></div><div><span>已整理</span><strong>{stats.completed}</strong><small>已写入 Excel</small></div><div><span>未整理</span><strong>{stats.unorganized}</strong><small>原因清楚可见</small></div><div className="status-accent"><span>正在处理</span><strong>{stats.processing}</strong><small>DeepSeek 后台任务</small></div></div></section>

        <section className="shell requirements-section">
          <div className="section-heading"><div><p className="eyebrow">TRIP BRIEF</p><h2>这次要找什么</h2></div><p>后续投递的链接都会按这些条件提取和比较，不满足或没写清楚的地方会明确标成“待核实”。</p></div>
          <div className="requirement-grid">
            <article><span>01 · 住宿</span><h3>6 人住两晚</h3><p>{state.tripProfile.stayPreference}</p><small>预算：{state.tripProfile.accommodationBudget}</small></article>
            <article><span>02 · 周六晚上</span><h3>回民宿烧烤</h3><p>{state.tripProfile.barbecue}</p><small>订房前确认设备、费用和邻里限制</small></article>
            <article><span>03 · 早餐</span><h3>两顿都安排</h3><p>{state.tripProfile.breakfasts.join(" · ")}</p><small>优先可预约、6 人同桌、交通顺路</small></article>
            <article><span>04 · 团队活动</span><h3>密室或同类活动</h3><p>{state.tripProfile.activity}</p><small>核对主题、难度、恐怖程度与时长</small></article>
          </div>
        </section>

        <section className="flow-section"><div className="shell"><div className="section-heading"><div><p className="eyebrow">HOW IT FLOWS</p><h2>从链接，到可以出发</h2></div><p>链接读不到也不会消失；处理报告会告诉你缺了什么，团队可以换链接或补充说明。</p></div><div className="flow-track">{[["01","投递链接"],["02","读取网页"],["03","AI 提取"],["04","写入 Excel"],["05","团队入选"],["06","形成行程"]].map(([num,label], index) => <div key={num}><b>{num}</b><span>{label}</span>{index < 5 && <i>→</i>}</div>)}</div></div></section>

        <section className="shell section itinerary-section">
          <div className="section-heading"><div><p className="eyebrow">WEEKEND PLAN</p><h2>周五晚到周日的初版方案</h2></div><p>路线先把住宿、烧烤、两顿早餐和密室固定下来；真实链接到位后，再替换地址和价格。</p></div>
          <DayPlan day="周五晚上" theme="抵达 · 入住 · 六人碰头" items={state.itinerary.day0} compact />
          <div className="day-columns"><DayPlan day="周六" theme="早茶、园林与民宿烧烤" items={state.itinerary.day1} /><DayPlan day="周日" theme="早餐、密室与弹性返程" items={state.itinerary.day2} /></div>
        </section>

        <section className="reservation-section"><div className="shell"><div className="section-heading"><div><p className="eyebrow">BOOKING CHECKLIST</p><h2>需要确认和预订的事</h2></div><p>这里只管理准备工作，不会自动付款或替你下单。最终价格与取消政策请在预订页面再次核对。</p></div><div className="reservation-list">{state.reservations.map((item) => <article key={item.id}><span className="reservation-type">{item.type}</span><div><h3>{item.item}</h3><p>{item.targetTime} · {item.note}</p></div><div className="reservation-owner"><small>{item.owner}</small><span>截至 {item.deadline}</span></div><b className={`status-badge ${statusTone(item.status)}`}><i />{item.status}</b></article>)}</div></div></section>
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
        <div className="page-title split"><div><p className="eyebrow">CANDIDATE LIBRARY</p><h1>候选资料库</h1><p>住宿卡重点显示 6 人两晚、地址、烧烤和房型；“待核实”表示原链接没有给出可靠事实。</p></div><div className="library-count"><strong>{state.places.length}</strong><span>个候选项</span></div></div>
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
        <div className="page-title"><p className="eyebrow">EXCEL WORKSPACE</p><h1>Excel 是最终可核对的旅行台账。</h1><p>除了候选资料，现在还新增了“处理报告”和“预订清单”；行程与预订状态在 Excel 修改后可以同步回网站。</p></div>
        <div className="excel-hero"><div className="excel-file-icon">X</div><div className="excel-file-info"><span>当前工作簿</span><h2>扬州团队旅行攻略.xlsx</h2><p>最近同步：{formatTime(state.settings.lastExcelSync)}</p></div><div className="excel-actions"><a className="primary" href={`${apiBase()}/api/download/excel`}>下载 / 打开 Excel</a><button className="small-button" onClick={syncExcel}>读取最新修改</button></div></div>
        <div className="sheet-grid">{[["项目总览","完成度、候选和下一步重点"],["链接汇总","所有原始链接与完整字段"],["处理报告","成功、失败和缺失信息"],["住宿候选","6 人两晚、烧烤与房型比较"],["活动候选","密室、景点和预约条件"],["餐饮候选","两顿早餐与正餐备选"],["攻略文章","文章摘要、避坑和缺失项"],["两日行程","周五晚 + 周六周日计划"],["预订清单","负责人、期限与当前状态"],["项目设置","人数、住宿偏好与模型设置"]].map(([name, desc], index) => <div className="sheet-card" key={name}><b>{String(index + 1).padStart(2, "0")}</b><div><strong>{name}</strong><p>{desc}</p></div></div>)}</div>
        <div className="sync-explainer"><div><span>团队</span><b>提交链接</b></div><i>→</i><div><span>本地后台</span><b>读取 + DeepSeek 分析</b></div><i>→</i><div className="highlight"><span>Excel</span><b>完整归档 / 人工修改</b></div><i>→</i><div><span>网站</span><b>更新候选与行程</b></div></div>
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

function DayPlan({ day, theme, items, compact = false }: { day: string; theme: string; items: ItineraryItem[]; compact?: boolean }) {
  const total = items.reduce((sum, item) => sum + (item.cost || 0), 0);
  return <article className={`day-plan ${compact ? "arrival-plan" : ""}`}><header><div><span>{day}</span><h3>{theme}</h3></div><p>预计 ¥{total} / 人</p></header><div className="timeline">{items.map((item) => <div className="timeline-item" key={`${item.time}-${item.title}`}><time>{item.time}</time><i /><div><span className="timeline-category">{item.category}</span><h4>{item.title}</h4><p>{item.subtitle}</p><small>{item.address} · {item.transport}</small><span className={`mini-booking ${statusTone(item.bookingStatus)}`}>{item.bookingStatus}</span>{item.note && <em>{item.note}</em>}</div></div>)}</div></article>;
}
