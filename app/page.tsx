"use client";

import { useEffect, useMemo, useState } from "react";
import type { FormEvent } from "react";

type LinkStatus = "等待处理" | "正在读取" | "AI分析中" | "已写入Excel" | "需要人工补充" | "处理失败";

type LinkRecord = {
  id: string;
  sourceType: "链接" | "文字";
  inputText: string;
  url: string;
  title: string;
  category: string;
  subCategory: string;
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
  locationHighlights: string;
  distanceToCore: string;
  capacity: string;
  roomType: string;
  rooms: string;
  beds: string;
  bedTypes: string;
  bathrooms: string;
  twoNightTotal: string;
  extraFees: string;
  deposit: string;
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
  cancellationPolicy: string;
  evidence: string;
  cuisineType: string;
  signatureDishes: string;
  sixPersonTotal: string;
  privateRoom: string;
  queueInfo: string;
  groupSuitability: string;
  themeName: string;
  escapeStyle: string;
  venueSize: string;
  roomCount: string;
  minPlayers: string;
  maxPlayers: string;
  sixPersonSession: string;
  npcInteraction: string;
  physicalIntensity: string;
  costume: string;
  leisureFacilities: string;
  packageInfo: string;
  overnight: string;
  includedMeals: string;
  restArea: string;
  genderArrangement: string;
  serviceRestrictions: string;
  attractionType: string;
  indoorOutdoor: string;
  weatherImpact: string;
  ticketInfo: string;
  recommendedDuration: string;
};

type Place = {
  id: string;
  sourceId?: string;
  name: string;
  category: string;
  subCategory: string;
  featureTags: string[];
  area: string;
  price: number | null;
  priceLabel: string;
  duration: string;
  score: number;
  tags: string[];
  pros: string[];
  cons: string[];
  selected: boolean;
  votes?: Record<string, VoteChoice>;
  decisionStatus?: "待比较" | "备选" | "拟定" | "淘汰";
  manualNote?: string;
  manualOverrides?: Record<string, unknown>;
  completeness?: number;
  keyMissing?: string[];
  sourceUrl: string;
  dataStatus: string;
  details: PlaceDetails;
};

type VoteChoice = "想去" | "可以" | "不考虑";

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
  project: { name: "我的出行共创项目", destination: "待确定目的地", days: 2, people: 6, budget: 6000, status: "方案共创中", tagline: "把分散的链接和想法，整理成大家都看得懂的出行方案。" },
  tripProfile: { dates: "待团队确认", schedule: "周五晚抵达 · 周日傍晚返程", groupSize: 6, nights: 2, stayPreference: "环境好、整租优先、至少 3 个独立睡眠空间", barbecue: "周六晚在民宿烧烤", breakfasts: ["周六早茶", "周日早餐"], activity: "6 人密室 / 团队活动", accommodationBudget: "待团队确认" },
  links: [],
  places: [],
  itinerary: { day0: [], day1: [], day2: [] },
  reservations: [],
  finalPlan: {
    version: "excel-home-v1",
    title: "我的出行共创项目",
    destination: "待确定目的地",
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

const tabs = [["plan", "看行程"], ["collect", "投递想法"], ["library", "重点候选"], ["manage", "Excel 管理"]] as const;
const categories = ["自动识别", "攻略文章", "住宿", "餐饮", "密室", "室内休闲", "景点户外"];
const mainFilters = ["全部", "住宿", "餐饮", "密室", "休闲娱乐", "景点", "攻略"];
const voteChoices: VoteChoice[] = ["想去", "可以", "不考虑"];

function formatTime(value: string) {
  if (!value) return "尚未同步";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString("zh-CN", { hour12: false });
}

function statusTone(status: string) {
  if (["已写入Excel", "已整理", "成功读取", "文字已接收", "已预订", "已完成", "已确定"].includes(status)) return "success";
  if (["处理失败", "需要人工补充", "未整理", "读取失败", "读取受限", "文字解析失败", "待补充", "待确认", "待预订"].includes(status)) return "warning";
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

function mapSearchUrl(keyword: string, city: string) {
  return `https://uri.amap.com/search?keyword=${encodeURIComponent(keyword)}&city=${encodeURIComponent(city)}&src=travel-co-creation`;
}

function voteSummary(place: Place) {
  const values = Object.values(place.votes || {});
  return {
    support: values.filter((item) => item === "想去").length,
    okay: values.filter((item) => item === "可以").length,
    reject: values.filter((item) => item === "不考虑").length,
  };
}

function importantDetails(place: Place, teamPeople: number) {
  const d = place.details;
  if (place.category === "住宿") return [["位置", d.address], ["两晚价格", !isPending(d.twoNightTotal) ? d.twoNightTotal : place.priceLabel], ["户型 / 房间", `${d.roomType} / ${d.rooms}`], ["床位 / 床型", `${d.beds} / ${d.bedTypes}`], [`${teamPeople} 人容量`, d.capacity], ["烧烤", d.barbecue], ["额外费用", d.extraFees], ["取消政策", d.cancellationPolicy]];
  if (place.category === "餐饮") return [["餐饮分类", place.subCategory], ["位置", d.address], ["参考人均", place.priceLabel], ["团队总价", d.sixPersonTotal], ["招牌菜", d.signatureDishes], ["包间 / 团队", `${d.privateRoom} / ${d.groupSuitability}`], ["排队", d.queueInfo], ["营业时间", d.openingHours]];
  if (place.category === "密室") return [["主题", d.themeName], ["位置", d.address], ["恐怖 / 难度", `${d.horrorLevel} / ${d.difficulty}`], ["规模", `${d.venueSize} / ${d.roomCount}`], ["六人开场", d.sixPersonSession], ["价格", !isPending(d.sixPersonTotal) ? d.sixPersonTotal : place.priceLabel], ["时长", place.duration], ["NPC", d.npcInteraction]];
  if (place.category === "休闲娱乐") return [["休闲类型", place.subCategory], ["位置", d.address], ["参考价格", place.priceLabel], ["包含设施", d.leisureFacilities], ["能否过夜", d.overnight], ["是否含餐", d.includedMeals], ["休息区域", d.restArea], ["使用限制", d.serviceRestrictions]];
  if (place.category === "景点") return [["游玩类型", place.subCategory], ["具体地点", d.address], ["票价", !isPending(d.ticketInfo) ? d.ticketInfo : place.priceLabel], ["开放时间", d.openingHours], ["建议时长", !isPending(d.recommendedDuration) ? d.recommendedDuration : place.duration], ["室内 / 室外", d.indoorOutdoor], ["天气影响", d.weatherImpact], ["预约", d.reservation]];
  return [["涉及区域", place.area], ["数据状态", place.dataStatus]];
}

function candidateRank(place: Place) {
  const decisionWeight = { 拟定: 80, 备选: 45, 待比较: 20, 淘汰: -100 }[place.decisionStatus || "待比较"];
  const votes = voteSummary(place);
  return Number(place.selected) * 120 + decisionWeight + place.score * 5 + (place.completeness || 0) / 5 + votes.support * 4 + votes.okay;
}

function clonePlan(plan: FinalPlan): FinalPlan {
  return JSON.parse(JSON.stringify(plan)) as FinalPlan;
}

export default function Home() {
  const [state, setState] = useState<AppState>(EMPTY_STATE);
  const [activeTab, setActiveTab] = useState<(typeof tabs)[number][0]>("plan");
  const [submissionMode, setSubmissionMode] = useState<"link" | "text">("link");
  const [urls, setUrls] = useState("");
  const [ideaText, setIdeaText] = useState("");
  const [category, setCategory] = useState("自动识别");
  const [submitter, setSubmitter] = useState("团队成员");
  const [nickname, setNickname] = useState("团队成员");
  const [note, setNote] = useState("");
  const [sending, setSending] = useState(false);
  const [connected, setConnected] = useState(false);
  const [message, setMessage] = useState("");
  const [libraryFilter, setLibraryFilter] = useState("全部");
  const [subCategoryFilter, setSubCategoryFilter] = useState("全部子类");
  const [showAllCandidates, setShowAllCandidates] = useState(false);
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
    const initial = window.setTimeout(() => {
      const savedNickname = window.localStorage.getItem("travel-team-nickname") || window.localStorage.getItem("yangzhou-team-nickname") || "团队成员";
      setNickname(savedNickname);
      setSubmitter(savedNickname);
      setBackendBase(apiBase());
      refresh();
    }, 0);
    const timer = window.setInterval(() => refresh(true), 3500);
    return () => { window.clearTimeout(initial); window.clearInterval(timer); };
  }, []);

  function rememberNickname(value: string) {
    const next = value.slice(0, 20);
    setNickname(next);
    setSubmitter(next || "团队成员");
    if (next.trim()) window.localStorage.setItem("travel-team-nickname", next.trim());
  }

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

  const filteredPlaces = useMemo(() => {
    const ranked = [...state.places]
      .filter((place) => libraryFilter === "全部" || place.category === libraryFilter)
      .filter((place) => subCategoryFilter === "全部子类" || place.subCategory === subCategoryFilter || place.featureTags.includes(subCategoryFilter))
      .sort((left, right) => candidateRank(right) - candidateRank(left));
    if (showAllCandidates) return ranked;
    const visible = ranked.filter((place) => place.decisionStatus !== "淘汰");
    if (libraryFilter !== "全部") return visible.slice(0, 5);
    const perCategory = new Map<string, number>();
    return visible.filter((place) => {
      const count = perCategory.get(place.category) || 0;
      if (count >= 3 && !place.selected && place.decisionStatus !== "拟定") return false;
      perCategory.set(place.category, count + 1);
      return true;
    }).slice(0, 10);
  }, [state.places, libraryFilter, subCategoryFilter, showAllCandidates]);

  const availableSubCategories = useMemo(() => {
    const relevant = state.places.filter((place) => libraryFilter === "全部" || place.category === libraryFilter);
    return [...new Set(relevant.flatMap((place) => [place.subCategory, ...place.featureTags].filter(Boolean)))].sort((left, right) => left.localeCompare(right, "zh-CN")).slice(0, 18);
  }, [state.places, libraryFilter]);

  const decisionStats = useMemo(() => ({
    total: state.places.length,
    shortlist: state.places.filter((place) => place.selected || ["拟定", "备选"].includes(place.decisionStatus || "")).length,
    complete: state.places.filter((place) => (place.completeness || 0) >= 75).length,
    rejected: state.places.filter((place) => place.decisionStatus === "淘汰").length,
  }), [state.places]);

  async function submitIdeas(event: FormEvent) {
    event.preventDefault();
    const list = urls.split(/\n|\s+/).map((item) => item.trim()).filter(Boolean);
    const text = ideaText.trim();
    if (submissionMode === "link" && !list.length) return setMessage("请先粘贴至少一个链接。");
    if (submissionMode === "text" && text.length < 3) return setMessage("请写下更具体的旅行想法。");
    setSending(true);
    setMessage("");
    try {
      const response = await fetch(`${apiBase()}/api/submissions`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ urls: submissionMode === "link" ? list : [], text: submissionMode === "text" ? text : "", category, submitter, note }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "提交失败");
      if (submissionMode === "link") setUrls("");
      else setIdeaText("");
      setNote("");
      setMessage(submissionMode === "text"
        ? `已接收 ${result.textCreated || result.created} 条文字想法，DeepSeek 正在整理。`
        : `已接收 ${result.linkCreated || result.created} 个新链接；打不开的链接也会保留。`);
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

  async function voteForPlace(id: string, choice: VoteChoice) {
    const voter = nickname.trim();
    if (!voter || voter === "团队成员") {
      setMessage("请先填写你的昵称，再参与选择。");
      return;
    }
    window.localStorage.setItem("travel-team-nickname", voter);
    try {
      const response = await fetch(`${apiBase()}/api/places/${encodeURIComponent(id)}/vote`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ nickname: voter, choice }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "选择保存失败");
      setState(result.state);
      setMessage(result.removed ? `已取消 ${voter} 的“${choice}”选择，并同步到 Excel。` : `已记录：${voter}选择“${choice}”，并同步到 Excel。`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "选择保存失败，请稍后重试。");
    }
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

  const destination = finalPlan.destination || state.project.destination || "待确定目的地";
  const brandName = isPending(destination) ? "出行共创" : `去${destination}`;
  const brandSeal = isPending(destination) ? "行" : destination.slice(0, 1);
  const workbookName = state.settings.workbookPath.split(/[\\/]/).pop() || "出行共创项目.xlsx";

  return (
    <main>
      <header className="topbar">
        <button className="brand" onClick={() => setActiveTab("plan")} aria-label="返回行程"><span className="brand-seal">{brandSeal}</span><span><strong>{brandName}</strong><small>{finalPlan.people} 人 · {state.project.days} 天共创</small></span></button>
        <nav aria-label="网站主导航">{tabs.map(([key, label]) => <button key={key} className={activeTab === key ? "nav-active" : ""} onClick={() => setActiveTab(key)}>{label}</button>)}</nav>
        <div className={`connection ${connected ? "online" : "offline"}`}><span aria-hidden="true" />{connected ? "已连接并自动同步" : "正在连接"}</div>
      </header>

      {message && <div className="toast" role="status"><span>{message}</span><button onClick={() => setMessage("")} aria-label="关闭提示">×</button></div>}

      {draftPlan && <PlanEditor plan={draftPlan} saving={savingPlan} onChange={setDraftPlan} onCancel={() => setDraftPlan(null)} onSave={savePlan} />}

      {activeTab === "plan" && <>
        <section className="simple-hero">
          <div className="shell simple-hero-grid">
            <div className="simple-hero-copy">
              <div className="sync-line"><span className={connected ? "online" : ""} />{connected ? `已和 Excel 同步 · ${formatTime(state.settings.lastExcelSync)}` : "正在连接本地资料"}</div>
              <p className="eyebrow">团队最终行程</p>
              <h1>{finalPlan.title}</h1>
              <p className="simple-summary">{finalPlan.summary}</p>
              <div className="simple-actions"><button className="primary" onClick={() => setActiveTab("collect")}>＋ 投递链接或想法</button><button className="small-button" onClick={() => setActiveTab("library")}>一起选候选</button></div>
            </div>
            <aside className="trip-at-a-glance">
              <div><span>什么时候</span><strong className={isPending(finalPlan.dates) ? "pending-value" : ""}>{finalPlan.dates}</strong></div>
              <div><span>怎么安排</span><strong>{finalPlan.schedule}</strong></div>
              <div><span>几个人</span><strong>{finalPlan.people} 人 · {finalPlan.nights} 晚</strong></div>
              <div><span>预算</span><strong className={isPending(finalPlan.perPersonBudget) ? "pending-value" : ""}>{finalPlan.perPersonBudget}</strong></div>
            </aside>
          </div>
        </section>

        <section className="decision-strip"><div className="shell"><div><strong>{finalProgress.confirmed}</strong><span>项已确定</span></div><div className="attention"><strong>{finalProgress.pending}</strong><span>项还要确认</span></div><p>最终内容只来自 Excel 第一张「行程首页」，候选资料不会自动进入行程。</p><button className="small-button" onClick={() => setActiveTab("manage")}>去管理</button></div></section>

        <section className="shell essentials-section">
          <div className="simple-section-title"><div><p className="eyebrow">先看重点</p><h2>住宿与待确认事项</h2></div><span>橙色内容表示还没有最终确定</span></div>
          <div className="essentials-grid">
            <article className="stay-summary-card">
              <header><div><span>两晚住宿</span><h3>{finalPlan.stay.name}</h3></div><StatusPill value={finalPlan.stay.name} /></header>
              <div className="stay-summary-fields"><FinalField label="地址" value={finalPlan.stay.address} /><FinalField label="房间 / 床位" value={finalPlan.stay.roomsBeds} /><FinalField label="两晚总价" value={finalPlan.stay.twoNightTotal} /><FinalField label="周六烧烤" value={finalPlan.stay.barbecue} /></div>
              <div className="card-link-row">{!isPending(finalPlan.stay.address) && <a href={mapSearchUrl(finalPlan.stay.address, destination)} target="_blank" rel="noreferrer">在高德地图查看</a>}{finalPlan.stay.sourceUrl && <a href={finalPlan.stay.sourceUrl} target="_blank" rel="noreferrer">查看民宿原链接</a>}</div>
            </article>
            <article className="pending-card">
              <header><span>下一步</span><h3>优先确认这几项</h3></header>
              {[{ label: "出行日期", value: finalPlan.dates }, { label: "人均预算", value: finalPlan.perPersonBudget }, { label: "民宿", value: finalPlan.stay.name }, { label: "烧烤条件", value: finalPlan.stay.barbecue }].map((item) => <div key={item.label}><StatusPill value={item.value} /><span>{item.label}</span><strong>{String(item.value)}</strong></div>)}
              <button className="text-button" onClick={() => setActiveTab("manage")}>在 Excel 或网页中修改 →</button>
            </article>
          </div>
        </section>

        <section className="final-itinerary-section simple-itinerary-section"><div className="shell"><div className="simple-section-title"><div><p className="eyebrow">两天怎么玩</p><h2>周五晚到周日</h2></div><span>修改 Excel 后，这里会自动刷新</span></div><div className="final-day-grid">{groupedFinalItinerary.map(({ day, items }) => <FinalDay key={day} day={day} items={items} />)}</div></div></section>

        <section className="shell final-reservation-section simple-reservation-section"><div className="simple-section-title"><div><p className="eyebrow">出发前清单</p><h2>谁来确认、什么时候完成</h2></div><span>网站只记录进度，不会自动下单</span></div><div className="reservation-list">{finalPlan.reservations.length ? finalPlan.reservations.map((item) => <article key={item.id}><span className="reservation-type">{item.type}</span><div><h3>{item.item}</h3><p>{item.targetTime} · {item.note}</p></div><div className="reservation-owner"><small>{item.owner}</small><span>截至 {item.deadline}</span></div><b className={`status-badge ${statusTone(item.status)}`}><i />{item.status}</b></article>) : <div className="empty-state">还没有预订事项，请在 Excel 首页底部添加。</div>}</div></section>
      </>}

      {activeTab === "collect" && <section className="shell workspace-page">
        <div className="page-title"><p className="eyebrow">投递灵感</p><h1>有链接就粘贴，没有链接就直接说想法。</h1><p>DeepSeek 会把“想住能烧烤的六人民宿”“想玩中恐密室”这类文字，和网页链接一样整理成分类、条件、缺失项与候选资料，再写进 Excel。</p></div>
        <div className="report-summary"><div><span>全部投递</span><strong>{stats.links}</strong></div><div className="good"><span>已成功整理</span><strong>{stats.completed}</strong></div><div className="warn"><span>未整理</span><strong>{stats.unorganized}</strong></div><div><span>处理中</span><strong>{stats.processing}</strong></div></div>
        <div className="collect-layout">
          <form className="collect-form" onSubmit={submitIdeas}>
            <div className="submission-switch" role="group" aria-label="选择投递方式"><button type="button" className={submissionMode === "link" ? "selected" : ""} aria-pressed={submissionMode === "link"} onClick={() => setSubmissionMode("link")}><b>粘贴链接</b><span>攻略、民宿、餐厅或活动页面</span></button><button type="button" className={submissionMode === "text" ? "selected" : ""} aria-pressed={submissionMode === "text"} onClick={() => setSubmissionMode("text")}><b>直接写想法</b><span>没有链接，也能表达自己的诉求</span></button></div>
            {submissionMode === "link"
              ? <><label htmlFor="urls">链接列表</label><textarea id="urls" value={urls} onChange={(event) => setUrls(event.target.value)} placeholder={"粘贴攻略、民宿、密室或餐厅链接……\n每行一个，也可以一次粘贴多个"} /></>
              : <><label htmlFor="ideaText">你想要什么</label><textarea id="ideaText" value={ideaText} onChange={(event) => setIdeaText(event.target.value.slice(0, 4000))} placeholder={"例如：我想住环境安静的整租民宿，6 个人能住，周六晚上可以烧烤，最好靠近东关街，两晚总价不要太高。"} /><div className="text-counter">{ideaText.length} / 4000</div></>}
            <div className="form-row"><label>大概是什么<select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((item) => <option key={item}>{item}</option>)}</select></label><label>你的昵称<input value={submitter} onChange={(event) => rememberNickname(event.target.value)} placeholder="例如：小王" /></label></div>
            <label>补充说明（可选）<input value={note} onChange={(event) => setNote(event.target.value)} placeholder="例如：这是我最在意的条件，优先级比较高" /></label>
            <button className="primary wide" disabled={sending}>{sending ? "正在提交……" : submissionMode === "text" ? "交给 DeepSeek 整理" : "开始读取并整理"}</button><p className="form-hint">文字会按“团队偏好”保存，不会冒充真实商户信息；链接若需要登录或验证码，会明确标记为读取受限。</p>
          </form>
          <aside className="pipeline-card"><span className="paper-tag">提交后会发生什么</span><h3>原话保留，条件拆开</h3>{["链接或原始文字先进入投递台账", "DeepSeek 识别分类、预算和关键偏好", "结构化结果进入候选表，等待团队筛选"].map((item, index) => <div className="pipeline-step" key={item}><b>{String(index + 1).padStart(2, "0")}</b><span>{item}</span></div>)}<p className="pipeline-note">Excel 同时保留原始诉求与 AI 整理结果，方便以后核对和继续补充。</p></aside>
        </div>
        <div className="task-panel"><div className="panel-heading"><div><h2>投递处理报告</h2><p>链接和文字都会显示处理进度；完整记录也会写入 Excel。</p></div><button className="small-button" onClick={() => refresh()}>刷新状态</button></div><div className="task-list">
          {state.links.length === 0 && <div className="empty-state">还没有团队投递。粘贴链接或写下第一个想法后，处理过程会显示在这里。</div>}
          {state.links.map((item) => <article className="task-report" key={item.id}>
            <div className="task-report-main"><span className="source-icon">{item.sourceType === "文字" ? "文" : "链"}</span><div><div className="task-title-line"><strong>{item.title || (item.sourceType === "文字" ? "团队文字需求" : sourceName(item.url))}</strong><span>{item.category}{item.subCategory ? ` · ${item.subCategory}` : ""}</span></div>{item.url ? <a href={item.url} target="_blank" rel="noreferrer">{item.url}</a> : <p className="task-source-text">“{item.inputText}”</p>}<p>{item.resultNote}</p></div></div>
            <div className="task-statuses"><span className={`status-badge ${statusTone(item.readStatus)}`}><i />{item.readStatus}</span><span className={`status-badge ${statusTone(item.organizedStatus)}`}><i />{item.organizedStatus}</span><small>{formatTime(item.updatedAt)}</small></div>
            <details className="task-details"><summary>查看整理详情</summary><div className="fact-columns"><div><b>已提取</b><p>{item.factsFound.length ? item.factsFound.join(" · ") : "暂无"}</p></div><div><b>还缺少</b><p>{item.missingFields.length ? item.missingFields.join(" · ") : "无明显缺失"}</p></div></div></details>
            <div className="task-report-foot"><span>{item.submitter} · {item.sourceType || "链接"} · {item.model || "等待分配模型"}</span>{(item.organizedStatus === "未整理" || item.status === "处理失败") && <button className="small-button" onClick={() => retryLink(item.id)}>重新尝试</button>}</div>
          </article>)}
        </div></div>
      </section>}

      {activeTab === "library" && <section className="shell workspace-page">
        <div className="page-title split"><div><p className="eyebrow">重点候选</p><h1>先选大类，再按子类快速缩小范围。</h1><p>可以直接筛选烧烤、火锅、微恐、中恐、汗蒸、桑拿等标签；完整字段和人工修改记录都保留在 Excel。</p></div><div className="library-count"><strong>{filteredPlaces.length}</strong><span>当前展示 / 共 {state.places.length} 个</span></div></div>
        <div className="nickname-bar"><div><strong>我的昵称</strong><span>不用注册，只用于区分是谁做的选择</span></div><input value={nickname} onChange={(event) => rememberNickname(event.target.value)} placeholder="例如：小王" maxLength={20} /></div>
        <div className="candidate-notice candidate-mode"><div><strong>{showAllCandidates ? "正在查看全部候选" : "智能收起已开启"}</strong><span>{showAllCandidates ? "包括待比较和已淘汰项目，完整数据仍以 Excel 为准。" : "每类最多显示 3 个重点项目；已淘汰候选默认隐藏。"}</span></div><button className="small-button" onClick={() => setShowAllCandidates((value) => !value)}>{showAllCandidates ? "收起，只看重点" : `查看全部 ${state.places.length} 个`}</button></div>
        <div className="filter-stack"><div className="filter-label"><strong>大类</strong><span>先选住宿、餐饮或游玩方向</span></div><div className="filter-bar">{mainFilters.map((item) => <button key={item} className={libraryFilter === item ? "selected" : ""} onClick={() => { setLibraryFilter(item); setSubCategoryFilter("全部子类"); setShowAllCandidates(false); }}>{item}</button>)}</div>{availableSubCategories.length > 0 && <><div className="filter-label secondary"><strong>子类 / 标签</strong><span>一个候选可以同时拥有多个标签</span></div><div className="filter-bar sub-filter"><button className={subCategoryFilter === "全部子类" ? "selected" : ""} onClick={() => setSubCategoryFilter("全部子类")}>全部子类</button>{availableSubCategories.map((item) => <button key={item} className={subCategoryFilter === item ? "selected" : ""} onClick={() => setSubCategoryFilter(item)}>{item}</button>)}</div></>}</div>
        <div className="place-grid simple-place-grid">{filteredPlaces.map((place) => {
          const votes = voteSummary(place);
          const myVote = place.votes?.[nickname.trim()];
          const details = importantDetails(place, finalPlan.people).slice(0, 4);
          return <article className={`place-card simple-place-card ${place.selected ? "chosen-card" : ""}`} key={place.id}>
          <div className="place-top"><span className="place-category">{place.category} · {place.subCategory}</span><span className={place.selected ? "selected-mark" : `candidate-mark decision-${place.decisionStatus || "待比较"}`}>{place.selected ? "已入选" : place.decisionStatus || "待比较"}</span></div>
          <h3>{place.name}</h3><p className="place-meta">{place.area} · {place.category === "住宿" && !isPending(place.details.twoNightTotal) ? place.details.twoNightTotal : place.priceLabel}</p>
          <div className="candidate-metrics"><div><span>资料完整度</span><strong>{place.completeness || 0}%</strong></div><div><span>AI 推荐</span><strong>{place.score.toFixed(1)}</strong></div><div><span>还缺</span><strong>{place.keyMissing?.length || 0} 项</strong></div></div>
          <div className="candidate-highlight"><div><span>为什么值得看</span><strong>{place.pros[0] || "等待分析"}</strong></div><div><span>还要核实</span><strong>{place.cons[0] || "等待核实"}</strong></div></div>
          <div className="detail-grid compact-details">{details.map(([label, detail]) => <div key={label}><span>{label}</span><strong>{detail}</strong></div>)}</div>
          <div className="vote-panel"><div className="vote-summary"><strong>{votes.support}</strong><span>人想去</span><small>{votes.okay} 人可以 · {votes.reject} 人不考虑</small></div><div className="vote-buttons">{voteChoices.map((choice) => <button key={choice} className={myVote === choice ? "selected" : ""} onClick={() => voteForPlace(place.id, choice)}>{choice}</button>)}</div></div>
          <details className="candidate-details"><summary>查看更多资料</summary><div className="place-tags">{[...new Set([place.subCategory, ...place.featureTags, ...place.tags])].map((tag) => <span key={tag}>{tag}</span>)}</div>{place.keyMissing?.length ? <p className="candidate-missing"><b>缺失项：</b>{place.keyMissing.join(" · ")}</p> : null}{place.manualNote ? <p className="candidate-manual-note"><b>人工备注：</b>{place.manualNote}</p> : null}{Object.keys(place.manualOverrides || {}).length ? <p className="candidate-manual-note"><b>人工保护：</b>{Object.keys(place.manualOverrides || {}).length} 个字段不会被重新分析覆盖</p> : null}<div className="place-footer"><div><span>参考预算</span><strong>{place.priceLabel}</strong></div><div className="score"><span>推荐度</span><strong>{place.score.toFixed(1)}</strong></div></div><p className="data-state">数据状态：{place.dataStatus}</p></details>
          <div className="card-link-row">{place.details.address && !isPending(place.details.address) && <a href={mapSearchUrl(place.details.address, destination)} target="_blank" rel="noreferrer">地图</a>}{place.sourceUrl && <a href={place.sourceUrl} target="_blank" rel="noreferrer">原始链接</a>}</div>
        </article>})}</div>
        {!filteredPlaces.length && <div className="empty-state">当前筛选下没有重点候选，可以查看全部或继续投递链接与想法。</div>}
      </section>}

      {activeTab === "manage" && <section className="shell workspace-page manage-page">
        <div className="page-title"><p className="eyebrow">Excel 管理</p><h1>你在 Excel 里分类、比较和定稿。</h1><p>先在「候选决策台」看全局，再到美食、密室、休闲或住宿分表补充黄色字段，最后把确定结果写进「行程首页」。</p></div>
        <div className="decision-overview"><div><span>全部候选</span><strong>{decisionStats.total}</strong></div><div><span>拟定 / 备选</span><strong>{decisionStats.shortlist}</strong></div><div><span>资料完整 ≥ 75%</span><strong>{decisionStats.complete}</strong></div><div><span>已淘汰</span><strong>{decisionStats.rejected}</strong></div></div>
        <div className="excel-hero"><div className="excel-file-icon">X</div><div className="excel-file-info"><span>主操作文件</span><h2>{workbookName}</h2><p>打开后先看第二张「候选决策台」 · 最近同步：{formatTime(state.settings.lastExcelSync)}</p></div><div className="excel-actions"><a className="primary" href={`${backendBase}/api/download/excel`}>打开 Excel 决策台</a><button className="small-button" onClick={syncExcel}>立即读取修改</button></div></div>

        <div className="manage-choice-grid">
          <article className="recommended-choice"><span>推荐流程</span><h2>先收集想法，再补全，最后定行程</h2><p>链接和文字都会进入分类表；黄色单元格是你可以人工确认和修正的内容。</p><ol><li>在「候选决策台」按大类和子分类筛选</li><li>查看团队原始诉求，再补充真实商户信息</li><li>黄色列改过的内容会被人工保护</li><li>最后把确定内容写进「行程首页」</li></ol></article>
          <article><span>备用</span><h2>在网页中快速修改</h2><p>适合临时改一个时间、负责人或民宿信息，保存后同样会写回 Excel。</p><button className="small-button" onClick={startPlanEditing}>打开网页编辑器</button></article>
        </div>

        <div className="simple-section-title manage-sheet-title"><div><p className="eyebrow">四组工作表</p><h2>全局决策、分类细节、最终行程和处理记录</h2></div></div>
        <div className="sheet-grid simple-sheet-grid">{[["候选决策台","统一比较大类、子分类、价格、位置和团队诉求"],["六张分类明细","住宿、美食、密室、室内休闲、景点与攻略"],["行程首页","最终网页的唯一内容来源"],["投递汇总 / 处理报告","查看链接和文字是否成功整理、还缺什么"]].map(([name, desc], index) => <div className={`sheet-card ${index === 0 ? "primary-sheet" : ""}`} key={name}><b>{String(index + 1).padStart(2, "0")}</b><div><strong>{name}</strong><p>{desc}</p></div></div>)}</div>

        <details className="advanced-settings"><summary>查看运行状态</summary><div className="settings-card"><div className="setting-row"><div><span>当前分析方式</span><strong>{state.settings.provider}</strong></div><span className="setting-state good">已配置</span></div><div className="setting-row"><div><span>本地后台</span><strong>{connected ? "正在运行" : "未连接"}</strong></div><span className={`setting-state ${connected ? "good" : ""}`}>{connected ? "可用" : "请启动服务"}</span></div><div className="setting-row"><div><span>Excel 自动同步</span><strong>每 4 秒检查一次</strong></div><span className="setting-state good">已开启</span></div></div></details>
        <div className="scope-note safety"><strong>不会自动下单</strong><p>网站只负责整理和展示。住宿、餐厅、密室与门票都需要大家确认真实价格和取消政策后再预订。</p></div>
      </section>}

      <footer><div className="shell"><span>{destination} · 出行共创台</span><span>本地数据 · Excel 可编辑 · DeepSeek 整理</span></div></footer>
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
      <time>{item.time}{item.endTime ? `–${item.endTime}` : ""}</time><i /><div><span>{item.category}</span><h4>{item.title}</h4><p>{item.subtitle || item.note}</p><small>{[item.address, item.transport].filter(Boolean).join(" · ") || "地点待补充"}</small><div className="itinerary-item-actions"><b className={`status-badge ${statusTone(item.bookingStatus)}`}><i />{item.bookingStatus || "待确认"}</b>{item.address && !isPending(item.address) && <a href={mapSearchUrl(item.address)} target="_blank" rel="noreferrer">地图</a>}{item.sourceUrl && <a href={item.sourceUrl} target="_blank" rel="noreferrer">来源</a>}</div></div>
    </div>) : <div className="day-empty">请在 Excel 首页添加当天安排</div>}</div>
  </article>;
}
