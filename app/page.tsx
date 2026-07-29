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
  candidateCount?: number;
  error?: string;
};

type CandidateKind = "place" | "requirement" | "guide";

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
  candidateType?: CandidateKind | "真实地点" | "团队诉求" | "攻略资料" | string;
  name: string;
  category: string;
  subCategory: string;
  featureTags: string[];
  area: string;
  price: number | null;
  priceLabel: string;
  duration: string;
  summary?: string;
  factsFound?: string[];
  missingFields?: string[];
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
  verificationStatus: string;
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
  capabilities?: { canManage: boolean };
};

const EMPTY_STATE: AppState = {
  project: { name: "济州岛三日旅行共创", destination: "韩国济州岛", days: 3, people: 6, budget: 0, status: "方案共创中", tagline: "六个人一起把济州岛的链接和想法整理成可执行计划。" },
  tripProfile: { dates: "待团队确认", schedule: "3 天 2 晚 · 国际航班直达济州", groupSize: 6, nights: 2, stayPreference: "济州市区交通方便、环境舒适，适合 6 人入住并可简单做饭", barbecue: "待团队确认", breakfasts: ["民宿简餐或附近早餐", "民宿简餐或附近早餐"], activity: "东线自然景观、海岸步道与济州美食", accommodationBudget: "六人两晚约 ¥1,800–3,000" },
  links: [],
  places: [],
  itinerary: { day0: [], day1: [], day2: [] },
  reservations: [],
  finalPlan: {
    version: "excel-home-v1",
    title: "我的出行共创项目",
    destination: "待确定目的地",
    dates: "待团队确认",
    schedule: "3 天 2 晚 · 国际航班直达济州",
    people: 6,
    nights: 2,
    perPersonBudget: "¥1,000–1,500 / 人（不含往返济州机票）",
    summary: "餐饮按正餐约 ¥100/人并穿插民宿做饭；岛内公交优先、必要时短途拼车，贵景点可替换为免费海岸与步道。",
    stay: { name: "未选择", address: "待补充", capacity: "6 人", roomsBeds: "优先整租并确认厨房可用", twoNightTotal: "六人两晚约 ¥1,800–3,000", checkInOut: "待补充", barbecue: "非硬性条件", bbqEquipment: "按需确认", breakfast: "民宿简餐或附近早餐", sourceUrl: "" },
    itinerary: [],
    reservations: [],
    updatedAt: "",
  },
  settings: { provider: "演示分析", workbookPath: "", lastExcelSync: "" },
  capabilities: { canManage: false },
};

const teamTabs = [["plan", "行程 PLAN"], ["collect", "投递 ADD"], ["library", "候选 PICK"]] as const;
const manageTab = ["manage", "管理 EDIT"] as const;
type TabKey = (typeof teamTabs)[number][0] | typeof manageTab[0];
const categories = ["自动识别", "攻略文章", "住宿", "餐饮", "密室", "室内休闲", "景点户外"];
const mainFilters = ["全部", "住宿", "餐饮", "密室", "休闲娱乐", "景点"];
const voteChoices: VoteChoice[] = ["想去", "可以", "不考虑"];

function formatTime(value: string) {
  if (!value) return "尚未同步";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString("zh-CN", { hour12: false });
}

function statusTone(status: string) {
  if (["已写入Excel", "已整理", "成功读取", "文字已接收", "已预订", "已完成", "已确定", "已核实"].includes(status)) return "success";
  if (["处理失败", "需要人工补充", "未整理", "读取失败", "读取受限", "文字解析失败", "待补充", "待确认", "待预订", "待核实"].includes(status)) return "warning";
  if (["正在读取", "AI分析中", "整理中", "等待读取", "等待处理", "部分核实"].includes(status)) return "active";
  return "muted";
}

function isPending(value: unknown) {
  const text = String(value ?? "").trim();
  return !text || /待|未选择|未确定|未知|尚未|占位|空位|暂无/.test(text);
}

function apiBase() {
  if (typeof window === "undefined") return "http://localhost:8887";
  if (window.location.port === "3100") return `${window.location.protocol}//${window.location.hostname}:8887`;
  return window.location.origin;
}

function sourceName(url: string) {
  try { return new URL(url).hostname.replace(/^www\./, ""); } catch { return "链接"; }
}

function mapSearchUrl(keyword: string, city?: string) {
  const query = [keyword, city && !isPending(city) ? city : ""].filter(Boolean).join(" ");
  return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(query)}`;
}

function candidateKind(place: Place): CandidateKind {
  const value = String(place.candidateType || "").toLowerCase();
  if (/requirement|需求|诉求|偏好/.test(value)) return "requirement";
  if (/guide|攻略|article|资料/.test(value) || place.category === "攻略" || place.category === "攻略文章") return "guide";
  return "place";
}

function candidateKindLabel(place: Place) {
  return candidateKind(place) === "requirement" ? "团队诉求" : candidateKind(place) === "guide" ? "攻略资料" : "真实地点";
}

function linkBucket(item: LinkRecord): "processing" | "needs" | "complete" {
  if (["等待处理", "正在读取", "AI分析中"].includes(item.status) || item.organizedStatus === "整理中") return "processing";
  if (item.status === "处理失败" || item.status === "需要人工补充" || item.organizedStatus === "未整理" || /失败|受限|补充/.test(item.readStatus)) return "needs";
  return "complete";
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
  if (place.category === "密室") return [["主题", d.themeName], ["位置", d.address], ["恐怖 / 难度", `${d.horrorLevel} / ${d.difficulty}`], ["规模", `${d.venueSize} / ${d.roomCount}`], [`${teamPeople}人开场`, d.sixPersonSession], ["价格", !isPending(d.sixPersonTotal) ? d.sixPersonTotal : place.priceLabel], ["时长", place.duration], ["NPC", d.npcInteraction]];
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
  const [activeTab, setActiveTab] = useState<TabKey>("plan");
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
  const [candidateSort, setCandidateSort] = useState<"recommended" | "votes" | "complete" | "price">("recommended");
  const [compareIds, setCompareIds] = useState<string[]>([]);
  const [sourceFocus, setSourceFocus] = useState("");
  const [editingLinkId, setEditingLinkId] = useState("");
  const [linkDraft, setLinkDraft] = useState({ title: "", category: "", subCategory: "", note: "" });
  const [editingPlaceId, setEditingPlaceId] = useState("");
  const [adoptDay, setAdoptDay] = useState<"第1天" | "第2天" | "第3天">("第2天");
  const [backendBase, setBackendBase] = useState("http://localhost:8887");
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

  const canManage = Boolean(state.capabilities?.canManage);
  const visibleTabs = useMemo(() => canManage ? [...teamTabs, manageTab] : [...teamTabs], [canManage]);
  const displayTab: TabKey = !canManage && activeTab === "manage" ? "plan" : activeTab;

  function rememberNickname(value: string) {
    const next = value.slice(0, 20);
    setNickname(next);
    setSubmitter(next || "团队成员");
    if (next.trim()) window.localStorage.setItem("travel-team-nickname", next.trim());
  }

  const groupedLinks = useMemo(() => ({
    processing: state.links.filter((item) => linkBucket(item) === "processing"),
    needs: state.links.filter((item) => linkBucket(item) === "needs"),
    complete: state.links.filter((item) => linkBucket(item) === "complete"),
  }), [state.links]);

  const stats = useMemo(() => ({
    links: state.links.length,
    completed: groupedLinks.complete.length,
    unorganized: groupedLinks.needs.length,
    processing: groupedLinks.processing.length,
    selected: state.places.filter((item) => candidateKind(item) === "place" && item.selected).length,
  }), [groupedLinks, state.links.length, state.places]);

  const finalPlan = state.finalPlan || EMPTY_STATE.finalPlan;
  const finalProgress = useMemo(() => {
    const mealItems = finalPlan.itinerary.filter((item) => /餐|早茶|早餐|午餐|晚餐|烧烤|美食/.test(`${item.category}${item.title}`));
    const activityItems = finalPlan.itinerary.filter((item) => /密室|活动|游玩|景点|景观|海岸|市场|休闲|洗浴|桑拿|汗蒸/.test(`${item.category}${item.title}`));
    const isResolvedItem = (item: FinalPlan["itinerary"][number]) => !isPending(item.title)
      && !isPending(item.address)
      && !isPending(item.time)
      && /已预订|已确认|已完成|无需预订|无需预约|现场购票/.test(item.bookingStatus);
    const resolvedMeals = mealItems.filter(isResolvedItem);
    const resolvedActivities = activityItems.filter(isResolvedItem);
    const reservationsReady = finalPlan.reservations.length > 0 && finalPlan.reservations.every((item) => /已预订|已完成|已确认/.test(item.status));
    const groups = [
      { key: "date", label: "日期", ready: !isPending(finalPlan.dates), note: finalPlan.dates },
      { key: "stay", label: "住宿", ready: !isPending(finalPlan.stay.name) && !isPending(finalPlan.stay.address) && !isPending(finalPlan.stay.twoNightTotal), note: finalPlan.stay.name },
      { key: "meal", label: "餐饮", ready: resolvedMeals.length >= 2, note: mealItems.length ? `已落实 ${resolvedMeals.length}/${mealItems.length} 项` : "还没有餐饮安排" },
      { key: "activity", label: "活动", ready: resolvedActivities.length > 0, note: activityItems.length ? `已落实 ${resolvedActivities.length}/${activityItems.length} 项` : "还没有游玩活动" },
      { key: "budget", label: "预算", ready: !isPending(finalPlan.perPersonBudget), note: finalPlan.perPersonBudget },
      { key: "booking", label: "预订", ready: reservationsReady, note: finalPlan.reservations.length ? `${finalPlan.reservations.filter((item) => /已预订|已完成|已确认/.test(item.status)).length}/${finalPlan.reservations.length} 已确认` : "还没有预订清单" },
    ];
    return { groups, confirmed: groups.filter((item) => item.ready).length, pending: groups.filter((item) => !item.ready).length };
  }, [finalPlan]);
  const groupedFinalItinerary = useMemo(() => ["第1天", "第2天", "第3天"].map((day) => ({
    day,
    items: finalPlan.itinerary.filter((item) => item.day === day),
  })), [finalPlan.itinerary]);

  const realPlaces = useMemo(() => state.places.filter((place) => candidateKind(place) === "place"), [state.places]);
  const focusedRealPlaces = useMemo(() => realPlaces.filter((place) => !sourceFocus || place.sourceId === sourceFocus), [realPlaces, sourceFocus]);
  const requirements = useMemo(() => state.places.filter((place) => candidateKind(place) === "requirement" && (!sourceFocus || place.sourceId === sourceFocus)), [sourceFocus, state.places]);
  const guides = useMemo(() => state.places.filter((place) => candidateKind(place) === "guide" && (!sourceFocus || place.sourceId === sourceFocus)), [sourceFocus, state.places]);

  const filteredPlaces = useMemo(() => {
    const ranked = [...realPlaces]
      .filter((place) => !sourceFocus || place.sourceId === sourceFocus)
      .filter((place) => libraryFilter === "全部" || place.category === libraryFilter)
      .filter((place) => subCategoryFilter === "全部子类" || place.subCategory === subCategoryFilter || place.featureTags.includes(subCategoryFilter))
      .sort((left, right) => {
        if (candidateSort === "votes") return voteSummary(right).support - voteSummary(left).support || candidateRank(right) - candidateRank(left);
        if (candidateSort === "complete") return (right.completeness || 0) - (left.completeness || 0) || candidateRank(right) - candidateRank(left);
        if (candidateSort === "price") return (left.price ?? Number.MAX_SAFE_INTEGER) - (right.price ?? Number.MAX_SAFE_INTEGER);
        return candidateRank(right) - candidateRank(left);
      });
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
  }, [realPlaces, sourceFocus, libraryFilter, subCategoryFilter, showAllCandidates, candidateSort]);

  const availableSubCategories = useMemo(() => {
    const relevant = focusedRealPlaces.filter((place) => libraryFilter === "全部" || place.category === libraryFilter);
    return [...new Set(relevant.flatMap((place) => [place.subCategory, ...place.featureTags].filter(Boolean)))].sort((left, right) => left.localeCompare(right, "zh-CN")).slice(0, 18);
  }, [focusedRealPlaces, libraryFilter]);

  const decisionStats = useMemo(() => ({
    total: realPlaces.length,
    shortlist: realPlaces.filter((place) => place.selected || ["拟定", "备选"].includes(place.decisionStatus || "")).length,
    complete: realPlaces.filter((place) => (place.completeness || 0) >= 75).length,
    rejected: realPlaces.filter((place) => place.decisionStatus === "淘汰").length,
  }), [realPlaces]);
  const comparedPlaces = useMemo(() => compareIds.map((id) => realPlaces.find((place) => place.id === id)).filter((place): place is Place => Boolean(place)), [compareIds, realPlaces]);

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
      const skipped = [result.duplicates ? `${result.duplicates} 条重复` : "", result.invalid ? `${result.invalid} 条无效` : ""].filter(Boolean).join("、");
      setMessage(submissionMode === "text"
        ? `已接收 ${result.textCreated || result.created} 条文字想法，DeepSeek 正在整理。${skipped ? `另跳过 ${skipped}。` : ""}`
        : `已接收 ${result.linkCreated || result.created} 个新链接；读取失败也会留在处理报告中。${skipped ? `另跳过 ${skipped}。` : ""}`);
      await refresh(true);
      setActiveTab("collect");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "提交失败，请检查后台服务。");
    } finally { setSending(false); }
  }

  async function retryLink(id: string) {
    try {
      const response = await fetch(`${apiBase()}/api/links/${encodeURIComponent(id)}/retry`, { method: "POST" });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "重新处理失败");
      setMessage("已重新加入分析队列。");
      await refresh(true);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "重新处理失败");
    }
  }

  function startLinkEditing(item: LinkRecord) {
    setEditingLinkId(item.id);
    setLinkDraft({ title: item.title || "", category: item.category || "", subCategory: item.subCategory || "", note: item.note || "" });
  }

  async function saveLinkEdit(id: string) {
    try {
      const response = await fetch(`${apiBase()}/api/links/${encodeURIComponent(id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(linkDraft),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "投递记录保存失败");
      if (result.state) setState(result.state); else await refresh(true);
      setEditingLinkId("");
      setMessage("投递记录已修改，并保存到主电脑 Excel。");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "投递记录保存失败");
    }
  }

  async function deleteLink(id: string) {
    if (!window.confirm("确认删除这条投递及其全部未采用结果吗？如果其中有已入选、已投票或人工修改的内容，系统会阻止删除。")) return;
    try {
      const response = await fetch(`${apiBase()}/api/links/${encodeURIComponent(id)}`, { method: "DELETE" });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "删除失败");
      if (result.state) setState(result.state); else await refresh(true);
      setEditingLinkId("");
      setMessage("投递记录已删除。");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "删除失败");
    }
  }

  function viewSourceCandidates(id: string) {
    setSourceFocus(id);
    setLibraryFilter("全部");
    setSubCategoryFilter("全部子类");
    setShowAllCandidates(true);
    setActiveTab("library");
    window.setTimeout(() => document.getElementById("source-results")?.scrollIntoView({ behavior: "smooth", block: "start" }), 80);
  }

  function toggleCompare(id: string) {
    setCompareIds((current) => {
      const validIds = new Set(realPlaces.map((place) => place.id));
      const normalized = current.filter((item) => validIds.has(item));
      if (normalized.includes(id)) return normalized.filter((item) => item !== id);
      if (normalized.length >= 4) {
        setMessage("最多同时比较 4 个真实地点，请先移除一个。");
        return normalized;
      }
      return [...normalized, id];
    });
  }

  async function savePlaceEdit(id: string, changes: Partial<Place>) {
    try {
      const response = await fetch(`${apiBase()}/api/places/${encodeURIComponent(id)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(changes),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "候选保存失败");
      if (result.state) setState(result.state); else await refresh(true);
      setEditingPlaceId("");
      setMessage("候选资料已修改，并保存到主电脑 Excel。");
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "候选保存失败");
    }
  }

  async function adoptPlace(place: Place, day = adoptDay) {
    const target = place.category === "住宿" ? "stay" : "itinerary";
    try {
      const response = await fetch(`${apiBase()}/api/places/${encodeURIComponent(place.id)}/adopt`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ target, day }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "采用失败");
      if (result.state) setState(result.state); else await refresh(true);
      setMessage(target === "stay" ? `已将“${place.name}”采用为住宿。` : `已将“${place.name}”加入${day}行程，可在最终方案中调整时间。`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "采用失败");
    }
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
  const isJeju = destination.includes("济州");
  const brandName = isJeju ? "JEJU 같이" : isPending(destination) ? "出行共创" : `去${destination}`;
  const brandSeal = isJeju ? "ㅈ" : isPending(destination) ? "行" : destination.slice(0, 1);
  const workbookName = state.settings.workbookPath.split(/[\\/]/).pop() || "出行共创项目.xlsx";

  return (
    <main>
      <header className="topbar">
        <button className="brand" onClick={() => setActiveTab("plan")} aria-label="返回行程"><span className="brand-seal">{brandSeal}</span><span><strong>{brandName}</strong><small>{isJeju ? "제주 여행을 같이 만들어요" : `${finalPlan.people} 人 · ${state.project.days} 天共创`}</small></span></button>
        <nav aria-label="网站主导航">{visibleTabs.map(([key, label]) => <button key={key} className={displayTab === key ? "nav-active" : ""} onClick={() => setActiveTab(key)}>{label}</button>)}</nav>
        <div className={`connection ${connected ? "online" : "offline"}`}><span aria-hidden="true" />{connected ? "已连接并自动同步" : "正在连接"}</div>
      </header>

      {message && <div className="toast" role="status"><span>{message}</span><button onClick={() => setMessage("")} aria-label="关闭提示">×</button></div>}

      {draftPlan && <PlanEditor plan={draftPlan} saving={savingPlan} onChange={setDraftPlan} onCancel={() => setDraftPlan(null)} onSave={savePlan} />}

      {displayTab === "plan" && <>
        <section className="simple-hero">
          <div className="shell simple-hero-grid">
            <div className="simple-hero-copy">
              <div className="sync-line"><span className={connected ? "online" : ""} />{connected ? `已保存到主电脑 Excel · ${formatTime(state.settings.lastExcelSync)}（不等于已备份 GitHub）` : "正在连接主电脑资料"}</div>
              <p className="eyebrow"><span>제주</span> OUR JEJU NOTE · 03 DAYS</p>
              <h1><span className="hero-english">JEJU,<br />TOGETHER.</span><span className="hero-chinese">{finalPlan.title}</span></h1>
              <p className="simple-summary">{finalPlan.summary}</p>
              <div className="simple-actions"><button className="primary" onClick={() => setActiveTab("collect")}>＋ 分享一个济州想法</button><button className="small-button" onClick={() => setActiveTab("library")}>看看大家的候选</button></div>
            </div>
            <div className="hero-board"><div className="jeju-ticket"><span>JEJU ISLAND</span><strong>제주</strong><em>06 FRIENDS · 03 DAYS</em></div><aside className="trip-at-a-glance">
              <div><span>WHEN · 日期</span><strong className={isPending(finalPlan.dates) ? "pending-value" : ""}>{finalPlan.dates}</strong></div>
              <div><span>ROUTE · 路线</span><strong>{finalPlan.schedule}</strong></div>
              <div><span>CREW · 成员</span><strong>{finalPlan.people} 人 · {finalPlan.nights} 晚</strong></div>
              <div><span>BUDGET · 预算</span><strong className={isPending(finalPlan.perPersonBudget) ? "pending-value" : ""}>{finalPlan.perPersonBudget}</strong></div>
            </aside></div>
          </div>
        </section>

        <section className="decision-strip"><div className="shell progress-shell"><div className="progress-heading"><strong>{finalProgress.confirmed}/{finalProgress.groups.length}</strong><span>关键环节已确定</span></div><div className="progress-groups">{finalProgress.groups.map((item) => <div className={item.ready ? "ready" : "pending"} key={item.key}><i /> <span>{item.label}</span><small>{item.note}</small></div>)}</div><p>这里按日期、住宿、餐饮、活动、预算和预订的真实完成情况计算，不再按文字字段凑数。</p>{canManage ? <button className="small-button" onClick={() => setActiveTab("manage")}>主电脑管理</button> : <button className="small-button" onClick={() => setActiveTab("collect")}>继续补充</button>}</div></section>

        <section className="shell essentials-section">
          <div className="simple-section-title"><div><p className="eyebrow">STAY · CHECK</p><h2>先把住哪里定下来</h2></div><span>济州橘标记的内容还等大家确认</span></div>
          <div className="essentials-grid">
            <article className="stay-summary-card">
              <header><div><span>{finalPlan.nights} 晚住宿</span><h3>{finalPlan.stay.name}</h3></div><StatusPill value={finalPlan.stay.name} /></header>
              <div className="stay-summary-fields"><FinalField label="地址" value={finalPlan.stay.address} /><FinalField label="房间 / 床位" value={finalPlan.stay.roomsBeds} /><FinalField label={`${finalPlan.nights}晚总价`} value={finalPlan.stay.twoNightTotal} /><FinalField label="烧烤条件" value={finalPlan.stay.barbecue} /></div>
              <div className="card-link-row">{!isPending(finalPlan.stay.address) && <a href={mapSearchUrl(finalPlan.stay.address, destination)} target="_blank" rel="noreferrer">在地图查看</a>}{finalPlan.stay.sourceUrl && <a href={finalPlan.stay.sourceUrl} target="_blank" rel="noreferrer">查看住宿原链接</a>}</div>
            </article>
            <article className="pending-card">
              <header><span>下一步</span><h3>优先确认这几项</h3></header>
              {[{ label: "出行日期", value: finalPlan.dates }, { label: "人均预算", value: finalPlan.perPersonBudget }, { label: "民宿", value: finalPlan.stay.name }, { label: "烧烤条件", value: finalPlan.stay.barbecue }].map((item) => <div key={item.label}><StatusPill value={item.value} /><span>{item.label}</span><strong>{String(item.value)}</strong></div>)}
              {canManage ? <button className="text-button" onClick={() => setActiveTab("manage")}>在主电脑 Excel 或网页中修改 →</button> : <button className="text-button" onClick={() => setActiveTab("collect")}>投递你知道的信息 →</button>}
            </article>
          </div>
        </section>

        <section className="final-itinerary-section simple-itinerary-section"><div className="shell"><div className="simple-section-title"><div><p className="eyebrow">ROUTE · 제주</p><h2>三天，把海岸和小城慢慢走完</h2></div><span>济州市区 · 东线 · 西北海岸</span></div><div className="final-day-grid">{groupedFinalItinerary.map(({ day, items }) => <FinalDay key={day} day={day} items={items} destination={destination} />)}</div></div></section>

        <section className="shell final-reservation-section simple-reservation-section"><div className="simple-section-title"><div><p className="eyebrow">READY · GO</p><h2>出发前，六个人各有分工</h2></div><span>网站只记录进度，不会自动下单</span></div><div className="reservation-list">{finalPlan.reservations.length ? finalPlan.reservations.map((item) => <article key={item.id}><span className="reservation-type">{item.type}</span><div><h3>{item.item}</h3><p>{item.targetTime} · {item.note}</p></div><div className="reservation-owner"><small>{item.owner}</small><span>截至 {item.deadline}</span></div><b className={`status-badge ${statusTone(item.status)}`}><i />{item.status}</b></article>) : <div className="empty-state">还没有预订事项，请在 Excel 首页底部添加。</div>}</div></section>
      </>}

      {displayTab === "collect" && <section className="shell workspace-page">
        <div className="page-title"><p className="eyebrow">ADD TO JEJU</p><h1>把你种草的济州岛，都丢进来。</h1><p>有链接就粘贴，没有链接就直接说想法。DeepSeek 会把民宿、黑猪烤肉、海边咖啡和团队偏好分开整理，保留原话后写入 Excel。</p></div>
        <div className="report-summary"><div><span>全部投递</span><strong>{stats.links}</strong></div><div className="good"><span>已成功整理</span><strong>{stats.completed}</strong></div><div className="warn"><span>未整理</span><strong>{stats.unorganized}</strong></div><div><span>处理中</span><strong>{stats.processing}</strong></div></div>
        <div className="collect-layout">
          <form className="collect-form" onSubmit={submitIdeas}>
            <div className="submission-switch" role="group" aria-label="选择投递方式"><button type="button" className={submissionMode === "link" ? "selected" : ""} aria-pressed={submissionMode === "link"} onClick={() => setSubmissionMode("link")}><b>粘贴链接</b><span>攻略、民宿、餐厅或活动页面</span></button><button type="button" className={submissionMode === "text" ? "selected" : ""} aria-pressed={submissionMode === "text"} onClick={() => setSubmissionMode("text")}><b>直接写想法</b><span>没有链接，也能表达自己的诉求</span></button></div>
            {submissionMode === "link"
              ? <><label htmlFor="urls">链接列表</label><textarea id="urls" value={urls} onChange={(event) => setUrls(event.target.value)} placeholder={"粘贴攻略、民宿、密室或餐厅链接……\n每行一个，也可以一次粘贴多个"} /></>
              : <><label htmlFor="ideaText">你想要什么</label><textarea id="ideaText" value={ideaText} onChange={(event) => setIdeaText(event.target.value.slice(0, 4000))} placeholder={"例如：我想住济州市区交通方便的酒店，6 个人入住，两晚总价不要太高，附近最好有黑猪烤肉和早餐。"} /><div className="text-counter">{ideaText.length} / 4000</div></>}
            <div className="form-row"><label>大概是什么<select value={category} onChange={(event) => setCategory(event.target.value)}>{categories.map((item) => <option key={item}>{item}</option>)}</select></label><label>你的昵称<input value={submitter} onChange={(event) => rememberNickname(event.target.value)} placeholder="例如：小王" /></label></div>
            <label>补充说明（可选）<input value={note} onChange={(event) => setNote(event.target.value)} placeholder="例如：这是我最在意的条件，优先级比较高" /></label>
            <button className="primary wide" disabled={sending}>{sending ? "正在提交……" : submissionMode === "text" ? "交给 DeepSeek 整理" : "开始读取并整理"}</button><p className="form-hint">文字会按“团队偏好”保存，不会冒充真实商户信息；链接若需要登录或验证码，会明确标记为读取受限。</p>
          </form>
          <aside className="pipeline-card"><span className="paper-tag">提交后会发生什么</span><h3>原话保留，条件拆开</h3>{["链接或原始文字先进入投递台账", "DeepSeek 识别分类、预算和关键偏好", "结构化结果进入候选表，等待团队筛选"].map((item, index) => <div className="pipeline-step" key={item}><b>{String(index + 1).padStart(2, "0")}</b><span>{item}</span></div>)}<p className="pipeline-note">Excel 同时保留原始诉求与 AI 整理结果，方便以后核对和继续补充。</p></aside>
        </div>
        <div className="task-panel"><div className="panel-heading"><div><h2>投递处理中心</h2><p>一眼区分“正在处理、需要补充、已经完成”，并能看到每条投递生成了多少个结果。</p></div><button className="small-button" onClick={() => refresh()}>刷新状态</button></div><div className="task-groups">
          {state.links.length === 0 && <div className="empty-state">还没有团队投递。粘贴链接或写下第一个想法后，处理过程会显示在这里。</div>}
          {([[
            "processing", "正在处理", "DeepSeek 正在读取或整理，请稍后刷新"
          ], ["needs", "需要补充", "链接受限、字段不足或处理失败，需要人工协助"], ["complete", "已经完成", "已拆成团队诉求、攻略资料或真实地点"]] as const).map(([groupKey, groupTitle, groupDescription]) => {
            const records = groupedLinks[groupKey];
            return <section className={`task-group task-group-${groupKey}`} key={groupKey}>
              <header><div><h3>{groupTitle}</h3><p>{groupDescription}</p></div><b>{records.length}</b></header>
              <div className="task-list">{records.length ? records.map((item) => {
                const candidateCount = item.candidateCount ?? state.places.filter((place) => place.sourceId === item.id).length;
                return <article className="task-report" key={item.id}>
                  <div className="task-report-main"><span className="source-icon">{item.sourceType === "文字" ? "文" : "链"}</span><div><div className="task-title-line"><strong>{item.title || (item.sourceType === "文字" ? "团队文字需求" : sourceName(item.url))}</strong><span>{item.category}{item.subCategory ? ` · ${item.subCategory}` : ""}</span></div>{item.url ? <a href={item.url} target="_blank" rel="noreferrer">{item.url}</a> : <p className="task-source-text">“{item.inputText}”</p>}<p>{item.resultNote}</p></div></div>
                  <div className="task-statuses"><span className={`status-badge ${statusTone(item.readStatus)}`}><i />{item.readStatus}</span><span className={`status-badge ${statusTone(item.organizedStatus)}`}><i />{item.organizedStatus}</span><span className="candidate-count-badge">生成 {candidateCount} 个整理结果</span><small>{formatTime(item.updatedAt)}</small></div>
                  <details className="task-details"><summary>查看提取与缺失信息</summary><div className="fact-columns"><div><b>已提取</b><p>{item.factsFound.length ? item.factsFound.join(" · ") : "暂无"}</p></div><div><b>还缺少</b><p>{item.missingFields.length ? item.missingFields.join(" · ") : "无明显缺失"}</p></div></div></details>
                  {editingLinkId === item.id && <div className="inline-editor link-inline-editor"><label>标题<input value={linkDraft.title} onChange={(event) => setLinkDraft({ ...linkDraft, title: event.target.value })} /></label><label>大类<input value={linkDraft.category} onChange={(event) => setLinkDraft({ ...linkDraft, category: event.target.value })} /></label><label>子类<input value={linkDraft.subCategory} onChange={(event) => setLinkDraft({ ...linkDraft, subCategory: event.target.value })} /></label><label className="wide">管理员备注<textarea value={linkDraft.note} onChange={(event) => setLinkDraft({ ...linkDraft, note: event.target.value })} /></label><div className="inline-editor-actions"><button type="button" className="small-button" onClick={() => setEditingLinkId("")}>取消</button><button type="button" className="primary" onClick={() => saveLinkEdit(item.id)}>保存修改</button></div></div>}
                  <div className="task-report-foot"><span>{item.submitter} · {item.sourceType || "链接"} · {item.model || "等待分配模型"}</span><div className="task-actions">{candidateCount > 0 && <button className="small-button" onClick={() => viewSourceCandidates(item.id)}>查看 {candidateCount} 个结果</button>}{canManage && groupKey === "needs" && <button className="small-button" onClick={() => retryLink(item.id)}>重新尝试</button>}{canManage && groupKey === "complete" && <button className="small-button" onClick={() => window.confirm("重新分析会再次读取原始内容。投票、入选和人工修改会优先保留，继续吗？") && retryLink(item.id)}>重新分析</button>}{canManage && groupKey !== "processing" && <><button className="small-button" onClick={() => startLinkEditing(item)}>修改</button><button className="danger-button" onClick={() => deleteLink(item.id)}>删除</button></>}</div></div>
                </article>;
              }) : <div className="task-group-empty">当前没有内容</div>}</div>
            </section>;
          })}
        </div></div>
      </section>}

      {displayTab === "library" && <section id="source-results" className="shell workspace-page">
        <div className="page-title split"><div><p className="eyebrow">候选与需求</p><h1>先分清“想要什么”，再比较“具体去哪”。</h1><p>团队诉求、攻略资料和真实地点分别展示。只有有明确商户或地点的真实候选才能投票和加入行程。</p></div><div className="library-count"><strong>{focusedRealPlaces.length}</strong><span>真实地点 · {requirements.length} 条诉求 · {guides.length} 份攻略</span></div></div>

        <div className="candidate-type-summary"><div className="type-place"><b>真实地点</b><strong>{focusedRealPlaces.length}</strong><span>可比较、投票和加入行程</span></div><div className="type-requirement"><b>团队诉求</b><strong>{requirements.length}</strong><span>可标记采纳，用来约束筛选</span></div><div className="type-guide"><b>攻略资料</b><strong>{guides.length}</strong><span>作为线索，继续拆成具体地点</span></div></div>

        {sourceFocus && <div className="source-focus-banner"><div><strong>正在查看同一条投递生成的结果</strong><span>下面只显示这条来源拆出的团队诉求、攻略资料和真实地点。</span></div><button className="small-button" onClick={() => setSourceFocus("")}>清除来源筛选</button></div>}

        {(requirements.length > 0 || guides.length > 0) && <div className="reference-sections">
          {requirements.length > 0 && <details className="reference-group requirement-group" open><summary><span>团队诉求</span><b>{requirements.length}</b><small>这些是“想要什么”，不能投票当作真实商户</small></summary><div className="reference-grid">{requirements.map((place) => <ReferenceCard key={place.id} place={place} canManage={canManage} editing={editingPlaceId === place.id} onEdit={() => setEditingPlaceId(place.id)} onCancel={() => setEditingPlaceId("")} onSave={savePlaceEdit} />)}</div></details>}
          {guides.length > 0 && <details className="reference-group guide-group" open><summary><span>攻略资料</span><b>{guides.length}</b><small>这是来源线索，不等于已经核实过的店或地点</small></summary><div className="reference-grid">{guides.map((place) => <ReferenceCard key={place.id} place={place} canManage={canManage} editing={editingPlaceId === place.id} onEdit={() => setEditingPlaceId(place.id)} onCancel={() => setEditingPlaceId("")} onSave={savePlaceEdit} />)}</div></details>}
        </div>}

        <div id="library-results" className="real-candidate-heading"><div><p className="eyebrow">真实地点</p><h2>比较具体民宿、餐厅和活动</h2></div><div className="nickname-inline"><label htmlFor="voter-name">我的昵称</label><input id="voter-name" value={nickname} onChange={(event) => rememberNickname(event.target.value)} placeholder="例如：小王" maxLength={20} /></div></div>

        <div className="candidate-notice candidate-mode"><div><strong>{showAllCandidates ? "正在查看全部真实候选" : "智能收起已开启"}</strong><span>{showAllCandidates ? "包括待比较和已淘汰地点。每张卡片都会明确显示核实状态。" : "每类最多显示 3 个重点地点；已淘汰内容默认隐藏。"}</span></div><button className="small-button" onClick={() => setShowAllCandidates((value) => !value)}>{showAllCandidates ? "收起，只看重点" : `查看全部 ${focusedRealPlaces.length} 个地点`}</button></div>

        <div className="filter-stack"><div className="candidate-toolbar"><div><div className="filter-label"><strong>大类</strong><span>先选住宿、餐饮或游玩方向</span></div><div className="filter-bar">{mainFilters.map((item) => <button key={item} className={libraryFilter === item ? "selected" : ""} onClick={() => { setLibraryFilter(item); setSubCategoryFilter("全部子类"); setShowAllCandidates(false); }}>{item}</button>)}</div></div><label className="sort-control">排序方式<select value={candidateSort} onChange={(event) => setCandidateSort(event.target.value as typeof candidateSort)}><option value="recommended">综合推荐</option><option value="votes">想去人数</option><option value="complete">资料完整度</option><option value="price">价格从低到高</option></select></label></div>{availableSubCategories.length > 0 && <><div className="filter-label secondary"><strong>子类 / 标签</strong><span>一个候选可以同时拥有多个标签</span></div><div className="filter-bar sub-filter"><button className={subCategoryFilter === "全部子类" ? "selected" : ""} onClick={() => setSubCategoryFilter("全部子类")}>全部子类</button>{availableSubCategories.map((item) => <button key={item} className={subCategoryFilter === item ? "selected" : ""} onClick={() => setSubCategoryFilter(item)}>{item}</button>)}</div></>}</div>

        {comparedPlaces.length > 0 && <ComparisonTable places={comparedPlaces} onRemove={(id) => toggleCompare(id)} />}

        <div className="place-grid simple-place-grid">{filteredPlaces.map((place) => {
          const votes = voteSummary(place);
          const myVote = place.votes?.[nickname.trim()];
          const details = importantDetails(place, finalPlan.people).slice(0, 4);
          const isCompared = comparedPlaces.some((item) => item.id === place.id);
          return <article className={`place-card simple-place-card ${place.selected ? "chosen-card" : ""} ${isCompared ? "compare-selected" : ""}`} key={place.id}>
          <div className="candidate-identity"><span className="candidate-type type-place">真实地点</span><span className={`verification-state ${statusTone(place.verificationStatus)}`}><i />{place.verificationStatus || "待核实"}</span></div>
          <div className="place-top"><span className="place-category">{place.category} · {place.subCategory}</span><span className={place.selected ? "selected-mark" : `candidate-mark decision-${place.decisionStatus || "待比较"}`}>{place.selected ? "已入选" : place.decisionStatus || "待比较"}</span></div>
          <h3>{place.name}</h3><p className="place-meta">{place.area} · {place.category === "住宿" && !isPending(place.details.twoNightTotal) ? place.details.twoNightTotal : place.priceLabel}</p>
          <button className={`compare-toggle ${isCompared ? "selected" : ""}`} onClick={() => toggleCompare(place.id)}>{isCompared ? "✓ 已加入比较" : `＋ 加入比较（${comparedPlaces.length}/4）`}</button>
          <div className="candidate-metrics"><div><span>资料完整度</span><strong>{place.completeness || 0}%</strong></div><div><span>AI 推荐</span><strong>{place.score.toFixed(1)}</strong></div><div><span>还缺</span><strong>{place.keyMissing?.length || 0} 项</strong></div></div>
          <div className="candidate-highlight"><div><span>为什么值得看</span><strong>{place.pros[0] || "等待分析"}</strong></div><div><span>还要核实</span><strong>{place.cons[0] || "等待核实"}</strong></div></div>
          <div className="detail-grid compact-details">{details.map(([label, detail]) => <div key={label}><span>{label}</span><strong>{detail}</strong></div>)}</div>
          <div className="vote-panel"><div className="vote-summary"><strong>{votes.support}</strong><span>人想去</span><small>{votes.okay} 人可以 · {votes.reject} 人不考虑</small></div><div className="vote-buttons">{voteChoices.map((choice) => <button key={choice} className={myVote === choice ? "selected" : ""} onClick={() => voteForPlace(place.id, choice)}>{choice}</button>)}</div></div>
          <details className="candidate-details"><summary>查看更多资料</summary><div className="place-tags">{[...new Set([place.subCategory, ...place.featureTags, ...place.tags])].map((tag) => <span key={tag}>{tag}</span>)}</div>{place.keyMissing?.length ? <p className="candidate-missing"><b>缺失项：</b>{place.keyMissing.join(" · ")}</p> : null}{place.manualNote ? <p className="candidate-manual-note"><b>人工备注：</b>{place.manualNote}</p> : null}{Object.keys(place.manualOverrides || {}).length ? <p className="candidate-manual-note"><b>人工保护：</b>{Object.keys(place.manualOverrides || {}).length} 个字段不会被重新分析覆盖</p> : null}<div className="place-footer"><div><span>参考预算</span><strong>{place.priceLabel}</strong></div><div className="score"><span>推荐度</span><strong>{place.score.toFixed(1)}</strong></div></div></details>
          {editingPlaceId === place.id && <CandidateEditor place={place} onCancel={() => setEditingPlaceId("")} onSave={savePlaceEdit} />}
          {canManage && editingPlaceId !== place.id && <div className="manager-card-actions"><button className="small-button" onClick={() => setEditingPlaceId(place.id)}>修改候选</button>{place.category !== "住宿" && <select aria-label="选择加入哪天行程" value={adoptDay} onChange={(event) => setAdoptDay(event.target.value as typeof adoptDay)}><option>第1天</option><option>第2天</option><option>第3天</option></select>}<button className="primary" onClick={() => adoptPlace(place)}>{place.category === "住宿" ? "采用为住宿" : `加入${adoptDay}行程`}</button></div>}
          <div className="card-link-row">{place.details.address && !isPending(place.details.address) && <a href={mapSearchUrl(place.details.address, destination)} target="_blank" rel="noreferrer">地图</a>}{place.sourceUrl && <a href={place.sourceUrl} target="_blank" rel="noreferrer">原始链接</a>}</div>
        </article>})}</div>
        {!filteredPlaces.length && <div className="empty-state">当前筛选下没有真实地点。团队诉求和攻略资料不会混进这里。</div>}
      </section>}

      {displayTab === "manage" && canManage && <section className="shell workspace-page manage-page">
        <div className="page-title"><p className="eyebrow">主电脑管理</p><h1>在主电脑完成分类、比较和定稿。</h1><p>这个页面只在主电脑开放。先在「候选决策台」看全局，再到分类分表补充黄色字段，最后把确定结果写进「行程首页」。</p></div>
        <div className="decision-overview"><div><span>真实地点</span><strong>{decisionStats.total}</strong></div><div><span>拟定 / 备选</span><strong>{decisionStats.shortlist}</strong></div><div><span>资料完整 ≥ 75%</span><strong>{decisionStats.complete}</strong></div><div><span>已淘汰</span><strong>{decisionStats.rejected}</strong></div></div>
        <div className="excel-hero"><div className="excel-file-icon">X</div><div className="excel-file-info"><span>主电脑原文件</span><h2>{workbookName}</h2><p>最近保存：{formatTime(state.settings.lastExcelSync)} · 已保存到 Excel 不等于已备份到 GitHub</p></div><div className="excel-actions"><a className="primary" href={`${backendBase}/api/download/excel`}>下载 Excel 副本</a><button className="small-button" onClick={syncExcel}>读取主电脑原文件</button></div></div>
        <div className="excel-copy-warning"><strong>请注意：下载得到的是副本</strong><p>朋友下载后自行修改的文件不会自动传回。只有这台主电脑上的原始 Excel 会被网站自动读取和同步；GitHub 备份仍需单独执行。</p></div>

        <div className="manage-choice-grid">
          <article className="recommended-choice"><span>推荐流程</span><h2>先收集想法，再补全，最后定行程</h2><p>链接和文字都会进入分类表；黄色单元格是你可以人工确认和修正的内容。</p><ol><li>在「候选决策台」按大类和子分类筛选</li><li>查看团队原始诉求，再补充真实商户信息</li><li>黄色列改过的内容会被人工保护</li><li>最后把确定内容写进「行程首页」</li></ol></article>
          <article><span>快捷编辑</span><h2>在网页中分段修改</h2><p>适合临时改时间、负责人或民宿信息。编辑器会提示未保存内容，保存后写回主电脑 Excel。</p><button className="small-button" onClick={startPlanEditing}>打开网页编辑器</button></article>
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

function ReferenceCard({ place, canManage, editing, onEdit, onCancel, onSave }: {
  place: Place;
  canManage: boolean;
  editing: boolean;
  onEdit: () => void;
  onCancel: () => void;
  onSave: (id: string, changes: Partial<Place>) => void;
}) {
  const kind = candidateKind(place);
  return <article className={`reference-card type-${kind}`}>
    <div className="candidate-identity"><span className={`candidate-type type-${kind}`}>{candidateKindLabel(place)}</span><span className={`verification-state ${statusTone(place.verificationStatus)}`}><i />{place.verificationStatus || (kind === "requirement" ? "团队诉求" : "待核实")}</span></div>
    <span className="place-category">{place.category}{place.subCategory ? ` · ${place.subCategory}` : ""}</span>
    {place.selected && <span className="selected-mark">{kind === "requirement" ? "已采纳需求" : "已采纳线索"}</span>}
    <h3>{place.name}</h3>
    <p>{place.manualNote || place.summary || place.pros[0] || (kind === "requirement" ? "这是一条团队偏好，用来筛选真实地点。" : "这是一份攻略线索，需要继续拆解和核实具体地点。")}</p>
    <div className="place-tags">{[...new Set([place.subCategory, ...place.featureTags, ...place.tags].filter(Boolean))].slice(0, 8).map((tag) => <span key={tag}>{tag}</span>)}</div>
    {place.keyMissing?.length ? <p className="candidate-missing"><b>还需补充：</b>{place.keyMissing.join(" · ")}</p> : null}
    {editing && <CandidateEditor place={place} onCancel={onCancel} onSave={onSave} />}
    <div className="reference-actions">{place.sourceUrl && <a href={place.sourceUrl} target="_blank" rel="noreferrer">查看原始来源</a>}{canManage && !editing && <button className="small-button" onClick={onEdit}>修改整理结果</button>}</div>
  </article>;
}

function CandidateEditor({ place, onCancel, onSave }: {
  place: Place;
  onCancel: () => void;
  onSave: (id: string, changes: Partial<Place>) => void;
}) {
  const [draft, setDraft] = useState({
    name: place.name,
    category: place.category,
    subCategory: place.subCategory,
    area: place.area,
    priceLabel: place.priceLabel,
    address: place.details.address,
    decisionStatus: place.decisionStatus || "待比较",
    dataStatus: place.dataStatus,
    verificationStatus: place.verificationStatus || (candidateKind(place) === "requirement" ? "团队诉求" : "待核实"),
    manualNote: place.manualNote || "",
  });
  return <div className="inline-editor candidate-inline-editor">
    <label className="wide">名称<input value={draft.name} onChange={(event) => setDraft({ ...draft, name: event.target.value })} /></label>
    <label>大类<input value={draft.category} onChange={(event) => setDraft({ ...draft, category: event.target.value })} /></label>
    <label>子类<input value={draft.subCategory} onChange={(event) => setDraft({ ...draft, subCategory: event.target.value })} /></label>
    <label>区域<input value={draft.area} onChange={(event) => setDraft({ ...draft, area: event.target.value })} /></label>
    <label>参考价格<input value={draft.priceLabel} onChange={(event) => setDraft({ ...draft, priceLabel: event.target.value })} /></label>
    <label className="wide">详细地址<input value={draft.address} onChange={(event) => setDraft({ ...draft, address: event.target.value })} /></label>
    <label>决策状态<select value={draft.decisionStatus} onChange={(event) => setDraft({ ...draft, decisionStatus: event.target.value as NonNullable<Place["decisionStatus"]> })}><option>待比较</option><option>备选</option><option>拟定</option><option>淘汰</option></select></label>
    <label>核实状态<select value={draft.verificationStatus} disabled={candidateKind(place) === "requirement"} onChange={(event) => setDraft({ ...draft, verificationStatus: event.target.value })}>{candidateKind(place) === "requirement" && <option>团队诉求</option>}<option>待核实</option><option>部分核实</option><option>已核实</option></select></label>
    <label className="wide">数据来源说明<input value={draft.dataStatus} onChange={(event) => setDraft({ ...draft, dataStatus: event.target.value })} /></label>
    <label className="wide">人工备注<textarea value={draft.manualNote} onChange={(event) => setDraft({ ...draft, manualNote: event.target.value })} /></label>
    <div className="inline-editor-actions"><button type="button" className="small-button" onClick={onCancel}>取消</button><button type="button" className="primary" onClick={() => onSave(place.id, { name: draft.name, category: draft.category, subCategory: draft.subCategory, area: draft.area, priceLabel: draft.priceLabel, decisionStatus: draft.decisionStatus, dataStatus: draft.dataStatus, verificationStatus: draft.verificationStatus, manualNote: draft.manualNote, details: { ...place.details, address: draft.address } })}>保存候选</button></div>
  </div>;
}

function ComparisonTable({ places, onRemove }: { places: Place[]; onRemove: (id: string) => void }) {
  const rows: Array<[string, (place: Place) => string]> = [
    ["类别", (place) => `${place.category} · ${place.subCategory}`],
    ["区域", (place) => place.area || "待补充"],
    ["价格", (place) => place.category === "住宿" && !isPending(place.details.twoNightTotal) ? place.details.twoNightTotal : place.priceLabel],
    ["位置", (place) => place.details.address || "待补充"],
    ["想去人数", (place) => `${voteSummary(place).support} 人`],
    ["资料完整", (place) => `${place.completeness || 0}%`],
    ["推荐理由", (place) => place.pros[0] || "等待分析"],
    ["注意事项", (place) => place.cons[0] || "等待核实"],
    ["核实状态", (place) => place.verificationStatus || "待核实"],
  ];
  return <section className="comparison-panel"><header><div><span>并排比较</span><h2>已选择 {places.length}/4 个真实地点</h2></div><small>左右滑动查看全部列</small></header><div className="comparison-scroll"><table><thead><tr><th>比较项目</th>{places.map((place) => <th key={place.id}><strong>{place.name}</strong><button onClick={() => onRemove(place.id)} aria-label={`移除 ${place.name}`}>移除</button></th>)}</tr></thead><tbody>{rows.map(([label, value]) => <tr key={label}><th>{label}</th>{places.map((place) => <td key={place.id}>{value(place)}</td>)}</tr>)}</tbody></table></div></section>;
}

function PlanEditor({ plan, saving, onChange, onCancel, onSave }: {
  plan: FinalPlan;
  saving: boolean;
  onChange: (plan: FinalPlan) => void;
  onCancel: () => void;
  onSave: () => void;
}) {
  const [activeSection, setActiveSection] = useState<"basic" | "stay" | "itinerary" | "reservations">("basic");
  const [initialSnapshot] = useState(() => JSON.stringify(plan));
  const dirty = JSON.stringify(plan) !== initialSnapshot;

  useEffect(() => {
    function warnBeforeLeaving(event: BeforeUnloadEvent) {
      if (!dirty) return;
      event.preventDefault();
      event.returnValue = "";
    }
    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") requestCancel();
    }
    window.addEventListener("beforeunload", warnBeforeLeaving);
    window.addEventListener("keydown", closeOnEscape);
    return () => {
      window.removeEventListener("beforeunload", warnBeforeLeaving);
      window.removeEventListener("keydown", closeOnEscape);
    };
  });

  function requestCancel() {
    if (!dirty || window.confirm("还有未保存的修改，确定放弃并关闭吗？")) onCancel();
  }

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
      itinerary: [...plan.itinerary, { day: "第2天", time: "", endTime: "", category: "安排", title: "新增安排", subtitle: "", address: "待补充", transport: "待补充", cost: 0, bookingStatus: "待确认", sourceUrl: "", sourceId: "", note: "" }],
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
        <div><p className="eyebrow">WEB EDITOR · EXCEL SYNC</p><h1 id="plan-editor-title">分段编辑最终方案</h1><span>{dirty ? "● 有未保存修改" : "尚未修改"} · 保存后写入主电脑 Excel「行程首页」。</span></div>
        <div className="plan-editor-actions"><button type="button" className="small-button" onClick={requestCancel} disabled={saving}>关闭</button><button type="submit" className="primary" disabled={saving || !dirty}>{saving ? "正在同步……" : dirty ? "保存到主电脑 Excel" : "没有新修改"}</button></div>
      </header>

      <nav className="editor-tabs" aria-label="编辑内容分段">{([ ["basic", "基础信息"], ["stay", "住宿"], ["itinerary", `行程 ${plan.itinerary.length}`], ["reservations", `待办 ${plan.reservations.length}`] ] as const).map(([key, label], index) => <button type="button" className={activeSection === key ? "selected" : ""} onClick={() => setActiveSection(key)} key={key}><b>{String(index + 1).padStart(2, "0")}</b><span>{label}</span></button>)}</nav>

      <div className="plan-editor-body">
        {activeSection === "basic" && <section className="editor-section">
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
        </section>}

        {activeSection === "stay" && <section className="editor-section">
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
        </section>}

        {activeSection === "itinerary" && <section className="editor-section">
          <div className="editor-section-title with-action"><b>03</b><div><h2>三日行程</h2><p>可以增加、删除和调整每一项安排。</p></div><button type="button" className="small-button" onClick={addItinerary}>＋ 添加行程</button></div>
          <div className="editor-card-list">
            {plan.itinerary.map((item, index) => <article className="itinerary-editor-card" key={`${item.day}-${index}`}>
              <header><strong>{String(index + 1).padStart(2, "0")} · {item.title || "未命名安排"}</strong><button type="button" className="remove-button" onClick={() => onChange({ ...plan, itinerary: plan.itinerary.filter((_, itemIndex) => itemIndex !== index) })}>删除</button></header>
              <div className="editor-grid compact-grid">
                <label className="editor-field"><span>日期</span><select value={item.day} onChange={(event) => updateItinerary(index, "day", event.target.value)}><option>第1天</option><option>第2天</option><option>第3天</option></select></label>
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
        </section>}

        {activeSection === "reservations" && <section className="editor-section">
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
        </section>}
      </div>

      <footer className="plan-editor-footer"><span>{dirty ? "有未保存内容；关闭页面前会再次确认。" : "修改任意内容后即可保存。"}</span><div className="plan-editor-actions"><button type="button" className="small-button" onClick={requestCancel} disabled={saving}>关闭</button><button type="submit" className="primary" disabled={saving || !dirty}>{saving ? "正在同步……" : dirty ? "保存到主电脑 Excel" : "没有新修改"}</button></div></footer>
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

function FinalDay({ day, items, destination }: { day: string; items: FinalPlan["itinerary"]; destination: string }) {
  const dayNumber = day.match(/\d+/)?.[0] || "·";
  return <article className="final-day-card">
    <header><span>D{dayNumber}</span><h3>{day}</h3><b>{items.length} 项安排</b></header>
    <div className="final-day-items">{items.length ? items.map((item, index) => <div className="final-day-item" key={`${day}-${item.time}-${item.title}-${index}`}>
      <time>{item.time}{item.endTime ? `–${item.endTime}` : ""}</time><i /><div><span>{item.category}</span><h4>{item.title}</h4><p>{item.subtitle || item.note}</p><small>{[item.address, item.transport].filter(Boolean).join(" · ") || "地点待补充"}</small><div className="itinerary-item-actions"><b className={`status-badge ${statusTone(item.bookingStatus)}`}><i />{item.bookingStatus || "待确认"}</b>{item.address && !isPending(item.address) && <a href={mapSearchUrl(item.address, destination)} target="_blank" rel="noreferrer">地图</a>}{item.sourceUrl && <a href={item.sourceUrl} target="_blank" rel="noreferrer">来源</a>}</div></div>
    </div>) : <div className="day-empty">请在 Excel 首页添加当天安排</div>}</div>
  </article>;
}
