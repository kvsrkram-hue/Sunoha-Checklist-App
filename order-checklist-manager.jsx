import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ─── API Helper (Google Apps Script backend) ──────────────────
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzOJoOFNvYtmYFZJkIyJCoaKiH05t9iqhk8bNkdxisI4FLjHzlpoeK09oofLZW2rF0b/exec";

// Track in-flight API calls for the loading indicator
let _apiCallCount = 0;
let _apiListeners = [];
function notifyApiListeners() { _apiListeners.forEach(fn => fn(_apiCallCount > 0)); }

// Actions that must NOT auto-redirect to login on AUTH_EXPIRED. For these,
// losing the current context (e.g. an in-progress edit) is worse than a re-auth prompt —
// surface an inline error so the user can refresh deliberately.
const AUTH_SOFT_FAIL_ACTIONS = new Set(["getUntaggedResponse", "editUntaggedResponse"]);

class AuthExpiredError extends Error {
  constructor(action) { super("Session expired. Please refresh the page."); this.name = "AuthExpiredError"; this.action = action; }
}

const API = {
  _token: localStorage.getItem("ocm_token"),
  setToken(t) { this._token = t; },
  _startCall() { _apiCallCount++; notifyApiListeners(); },
  _endCall() { _apiCallCount = Math.max(0, _apiCallCount - 1); notifyApiListeners(); },
  _handleAuthExpired(action) {
    if (AUTH_SOFT_FAIL_ACTIONS.has(action)) {
      throw new AuthExpiredError(action);
    }
    localStorage.removeItem("ocm_token");
    localStorage.removeItem("ocm_user");
    window.location.reload();
  },
  get: async (action, params = {}) => {
    API._startCall();
    try {
      const url = new URL(APPS_SCRIPT_URL);
      url.searchParams.set("action", action);
      if (API._token) url.searchParams.set("token", API._token);
      Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
      const r = await fetch(url.toString());
      if (!r.ok) throw new Error(`GET ${action}: ${r.status}`);
      const json = await r.json();
      if (json?.error === "AUTH_EXPIRED") { API._handleAuthExpired(action); return; }
      if (json?.error) throw new Error(json.error);
      return json;
    } finally { API._endCall(); }
  },
  post: async (action, data = {}) => {
    API._startCall();
    try {
      const r = await fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({ action, token: API._token, ...data }),
      });
      if (!r.ok) throw new Error(`POST ${action}: ${r.status}`);
      const json = await r.json();
      if (json?.error === "AUTH_EXPIRED") { API._handleAuthExpired(action); return; }
      if (json?.error) throw new Error(json.error);
      return json;
    } finally { API._endCall(); }
  },
};

// ─── Icons ────────────────────────────────────────────────────

const Icon = ({ name, size = 20, color = "currentColor", style: extraStyle }) => {
  const icons = {
    plus: <path d="M12 5v14M5 12h14" strokeWidth="2" strokeLinecap="round" />,
    check: <path d="M20 6L9 17l-5-5" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />,
    chevron: <path d="M9 18l6-6-6-6" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />,
    settings: <><path d="M12.22 2h-.44a2 2 0 00-2 2v.18a2 2 0 01-1 1.73l-.43.25a2 2 0 01-2 0l-.15-.08a2 2 0 00-2.73.73l-.22.38a2 2 0 00.73 2.73l.15.1a2 2 0 011 1.72v.51a2 2 0 01-1 1.74l-.15.09a2 2 0 00-.73 2.73l.22.38a2 2 0 002.73.73l.15-.08a2 2 0 012 0l.43.25a2 2 0 011 1.73V20a2 2 0 002 2h.44a2 2 0 002-2v-.18a2 2 0 011-1.73l.43-.25a2 2 0 012 0l.15.08a2 2 0 002.73-.73l.22-.39a2 2 0 00-.73-2.73l-.15-.08a2 2 0 01-1-1.74v-.5a2 2 0 011-1.74l.15-.09a2 2 0 00.73-2.73l-.22-.38a2 2 0 00-2.73-.73l-.15.08a2 2 0 01-2 0l-.43-.25a2 2 0 01-1-1.73V4a2 2 0 00-2-2z" strokeWidth="1.5" /><circle cx="12" cy="12" r="3" strokeWidth="1.5" /></>,
    clipboard: <><path d="M16 4h2a2 2 0 012 2v14a2 2 0 01-2 2H6a2 2 0 01-2-2V6a2 2 0 012-2h2" strokeWidth="1.5" /><rect x="8" y="2" width="8" height="4" rx="1" ry="1" strokeWidth="1.5" /></>,
    package: <><path d="M16.5 9.4l-9-5.19M21 16V8a2 2 0 00-1-1.73l-7-4a2 2 0 00-2 0l-7 4A2 2 0 003 8v8a2 2 0 001 1.73l7 4a2 2 0 002 0l7-4A2 2 0 0021 16z" strokeWidth="1.5" /><path d="M3.27 6.96L12 12.01l8.73-5.05M12 22.08V12" strokeWidth="1.5" /></>,
    externalLink: <><path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6M15 3h6v6M10 14L21 3" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" /></>,
    trash: <><path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2" strokeWidth="1.5" strokeLinecap="round" /></>,
    edit: <><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" strokeWidth="1.5" /><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" strokeWidth="1.5" /></>,
    back: <path d="M19 12H5M12 19l-7-7 7-7" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />,
    link: <><path d="M10 13a5 5 0 007.54.54l3-3a5 5 0 00-7.07-7.07l-1.72 1.71" strokeWidth="1.5" strokeLinecap="round" /><path d="M14 11a5 5 0 00-7.54-.54l-3 3a5 5 0 007.07 7.07l1.71-1.71" strokeWidth="1.5" strokeLinecap="round" /></>,
    coffee: <><path d="M17 8h1a4 4 0 010 8h-1M3 8h14v9a4 4 0 01-4 4H7a4 4 0 01-4-4V8zM6 1v3M10 1v3M14 1v3" strokeWidth="1.5" strokeLinecap="round" /></>,
    user: <><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" strokeWidth="1.5" /><circle cx="12" cy="7" r="4" strokeWidth="1.5" /></>,
    users: <><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" strokeWidth="1.5" /><circle cx="9" cy="7" r="4" strokeWidth="1.5" /><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75" strokeWidth="1.5" /></>,
    clock: <><circle cx="12" cy="12" r="10" strokeWidth="1.5" /><path d="M12 6v6l4 2" strokeWidth="1.5" strokeLinecap="round" /></>,
    checkCircle: <><path d="M22 11.08V12a10 10 0 11-5.93-9.14" strokeWidth="1.5" /><path d="M22 4L12 14.01l-3-3" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></>,
    layers: <><path d="M12 2L2 7l10 5 10-5-10-5z" strokeWidth="1.5" /><path d="M2 17l10 5 10-5" strokeWidth="1.5" /><path d="M2 12l10 5 10-5" strokeWidth="1.5" /></>,
    logOut: <><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4" strokeWidth="1.5" strokeLinecap="round" /><path d="M16 17l5-5-5-5M21 12H9" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></>,
    archive: <><path d="M21 8v13H3V8M1 3h22v5H1zM10 12h4" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></>,
    undo: <><path d="M1 4v6h6" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /><path d="M3.51 15a9 9 0 105.67-8.51L1 10" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></>,
    key: <><path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 11-7.778 7.778 5.5 5.5 0 017.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></>,
    lock: <><rect x="3" y="11" width="18" height="11" rx="2" ry="2" strokeWidth="1.5" /><path d="M7 11V7a5 5 0 0110 0v4" strokeWidth="1.5" strokeLinecap="round" /></>,
    refresh: <><path d="M23 4v6h-6" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" /></>,
    x: <path d="M18 6L6 18M6 6l12 12" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />,
    "alert-triangle": <><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/><path d="M12 9v4M12 17h.01" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></>,
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} xmlns="http://www.w3.org/2000/svg" style={extraStyle}>{icons[name]}</svg>;
};

// ─── Theme ────────────────────────────────────────────────────

const T = {
  bg:"#0F1117",surface:"#1A1D27",surfaceHover:"#222636",card:"#1E2230",border:"#2A2E3F",borderLight:"#353A50",
  accent:"#D4A574",accentBg:"rgba(212,165,116,0.08)",accentBorder:"rgba(212,165,116,0.2)",
  success:"#6BCB77",successBg:"rgba(107,203,119,0.08)",successBorder:"rgba(107,203,119,0.2)",
  danger:"#E85D5D",dangerBg:"rgba(232,93,93,0.08)",
  info:"#5B9CF6",infoBg:"rgba(91,156,246,0.08)",infoBorder:"rgba(91,156,246,0.2)",
  text:"#E8E4DF",textSec:"#9B96A0",textMut:"#6B6777",
  font:"'Outfit',sans-serif",mono:"'JetBrains Mono',monospace",rad:"12px",radSm:"8px",
  warning:"#F0AD4E",warningBg:"rgba(240,173,78,0.08)",warningBorder:"rgba(240,173,78,0.25)",
};

// ─── Date Formatting Helpers ─────────────────────────────────

const formatDate = (dateStr) => {
  if (!dateStr) return "—";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return String(dateStr);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
};

const formatDateTime = (dateStr) => {
  if (!dateStr) return "—";
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return String(dateStr);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${dd}-${mm}-${yyyy} ${hh}:${mi}`;
};

// ─── Question Utilities ───────────────────────────────────────

const normalizeQuestions = (questions) => {
  if (!Array.isArray(questions)) return [];
  const DEFAULTS = { text: "", type: "text", formula: null, ideal: null, remarkCondition: null, isApprovalGate: false, linkedSource: null, inventoryLink: null, isMasterQuantity: false, inventoryCategory: "", idealLabel: "", idealUnit: "", remarksTargetIdx: null, autoFillMapping: null, dateComparison: null };
  return questions.map(q => {
    if (typeof q === "string") return { ...DEFAULTS, text: q };
    // Start with defaults, then overlay ALL existing fields from q to preserve any extra metadata
    const result = { ...DEFAULTS };
    for (const k in q) { if (q.hasOwnProperty(k) && q[k] !== undefined) result[k] = q[k]; }
    // Normalize specific fields
    if (!result.text) result.text = "";
    if (!result.type) result.type = "text";
    result.remarksTargetIdx = (result.remarksTargetIdx === null || result.remarksTargetIdx === undefined || result.remarksTargetIdx === "") ? null : Number(result.remarksTargetIdx);
    return result;
  });
};

// Format a date as dd-mm-yyyy for display. Accepts Date / yyyy-mm-dd / ISO / freeform.
const formatDateDisplay = (value) => {
  if (!value) return "";
  const d = value instanceof Date ? value : new Date(String(value));
  if (isNaN(d.getTime())) return String(value);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = d.getFullYear();
  return `${dd}-${mm}-${yy}`;
};

// Convert a stored response value to a human-friendly display string based on question type.
// inventory_item: ID → name (or already-name passthrough); date: dd-mm-yyyy; other: as-is.
const displayResponseValue = (q, raw, inventoryItems) => {
  if (raw === undefined || raw === null || raw === "") return "—";
  if (!q) return String(raw);
  if (q.type === "inventory_item") {
    if (Array.isArray(inventoryItems)) {
      const it = inventoryItems.find(i => i.id === raw || i.name === raw);
      if (it) return it.name;
    }
    return String(raw);
  }
  if (q.type === "date") return formatDateDisplay(raw);
  return String(raw);
};

// Format the time elapsed since a given date for display next to date inputs
const formatAgeFromDate = (value) => {
  if (!value) return "";
  const d = value instanceof Date ? value : new Date(String(value));
  if (isNaN(d.getTime())) return "";
  const now = new Date();
  const diffMs = now.getTime() - d.setHours(0,0,0,0);
  if (diffMs < 0) {
    const days = Math.ceil(Math.abs(diffMs) / 86400000);
    return `in ${days} day${days===1?"":"s"}`;
  }
  const days = Math.floor(diffMs / 86400000);
  if (days === 0) return "today";
  if (days === 1) return "yesterday";
  if (days < 30) return `${days} days ago`;
  const months = Math.floor(days / 30);
  if (months < 12) return `${months} month${months===1?"":"s"} ago`;
  const years = Math.floor(months / 12);
  const remMonths = months - years * 12;
  if (remMonths === 0) return `${years} year${years===1?"":"s"} ago`;
  return `${years} year${years===1?"":"s"}, ${remMonths} month${remMonths===1?"":"s"} ago`;
};

// Source has a trackable IN-linked quantity (or legacy master quantity flag)
const sourceHasInTrackedQty = (srcCk) => {
  if (!srcCk?.questions) return false;
  return srcCk.questions.some(qq => (qq.inventoryLink?.enabled && qq.inventoryLink.txType === "IN") || qq.isMasterQuantity);
};

// ─── Auto ID Helpers (Frontend Preview) ───────────────────────

const DEFAULT_AUTO_ID_PREFIXES = {
  "Green Bean QC Sample Check": "GBS",
  "Green Beans Quality Check": "GB",
  "Roasted Beans Quality Check": "RB",
  "Grinding & Packing Checklist": "RG",
  "Tagging Roasted Beans": "TG",
  "Sample Retention Checklist": "SR",
  "Coffee with Chicory Mix": "CC",
};
const getDefaultPrefixForChecklist = (name) => DEFAULT_AUTO_ID_PREFIXES[name] || "";

const formatDateDDMMYY = (value) => {
  if (!value) return "";
  const d = value instanceof Date ? value : new Date(String(value));
  if (isNaN(d.getTime())) return "";
  return String(d.getDate()).padStart(2, "0") + String(d.getMonth() + 1).padStart(2, "0") + String(d.getFullYear()).slice(-2);
};

const sanitizeItemCodeToken = (value) => {
  if (!value) return "";
  return String(value).replace(/[^A-Za-z0-9]/g, "").slice(0, 6).toUpperCase();
};

// Build live preview "PREFIX-DDMMYY-CODE-###". Sequence is shown as "###" placeholder.
// inventoryItems lets us resolve item dropdown selections to abbreviations.
const buildAutoIdPreview = (checklist, responses, fallbackDate, inventoryItems) => {
  if (!checklist || !checklist.autoIdConfig || !checklist.autoIdConfig.enabled) return "";
  const cfg = checklist.autoIdConfig;
  const prefix = (cfg.prefix || getDefaultPrefixForChecklist(checklist.name) || "AUTO").toUpperCase();
  let dateStr = "";
  if (cfg.dateFieldIdx !== null && cfg.dateFieldIdx !== undefined && cfg.dateFieldIdx !== "") {
    dateStr = formatDateDDMMYY(responses?.[cfg.dateFieldIdx]);
  }
  if (!dateStr) dateStr = formatDateDDMMYY(fallbackDate || new Date());
  let code = "";
  if (cfg.itemCodeFieldIdx !== null && cfg.itemCodeFieldIdx !== undefined && cfg.itemCodeFieldIdx !== "") {
    const raw = responses?.[cfg.itemCodeFieldIdx];
    if (raw && Array.isArray(inventoryItems)) {
      const match = inventoryItems.find(it => it.id === raw || it.name === raw);
      if (match && match.abbreviation) code = String(match.abbreviation).toUpperCase();
    }
    if (!code) code = sanitizeItemCodeToken(raw);
  }
  if (!code) code = "X";
  return `${prefix}-${dateStr || "DDMMYY"}-${code}-###`;
};

// Sum quantities across all batchAllocations entries
const sumBatchAllocations = (batchAllocations) => {
  if (!batchAllocations || typeof batchAllocations !== "object") return 0;
  let total = 0;
  Object.values(batchAllocations).forEach(arr => {
    if (Array.isArray(arr)) arr.forEach(a => { total += parseFloat(a?.quantity) || 0; });
  });
  return total;
};

const getQText = (q) => typeof q === "string" ? q : (q?.text || "");

const evaluateFormula = (formula, getFieldValue) => {
  if (!formula?.fields?.length) return null;
  let result = null, operator = null;
  for (const item of formula.fields) {
    if (item.type === "operator") { operator = item.value; continue; }
    let value = item.type === "constant" ? parseFloat(item.value) : parseFloat(getFieldValue(item.checklist, item.question));
    if (isNaN(value)) return null;
    if (result === null) { result = value; }
    else if (operator) {
      if (operator === "+") result += value;
      else if (operator === "-") result -= value;
      else if (operator === "×") result *= value;
      else if (operator === "÷") result = value !== 0 ? result / value : null;
      else if (operator === "%") result = result * value / 100;
      operator = null;
    }
  }
  return result !== null ? Math.round(result * 100) / 100 : null;
};

const checkRemarkCondition = (responseValue, idealValue, condition) => {
  if (!condition || idealValue == null || responseValue == null || responseValue === "") return false;
  const resp = parseFloat(responseValue), ideal = parseFloat(idealValue);
  if (isNaN(resp) || isNaN(ideal)) return false;
  switch (condition.type) {
    case "gt_ideal": return resp > ideal;
    case "lt_ideal": return resp < ideal;
    case "ne_ideal": return Math.abs(resp - ideal) > 0.001;
    case "differs_by_percent": return ideal !== 0 && Math.abs((resp - ideal) / ideal * 100) > (condition.value || 0);
    case "differs_by_units": return Math.abs(resp - ideal) > (condition.value || 0);
    default: return false;
  }
};

const formulaPreview = (fields, allChecklists, selfQuestions) => {
  if (!fields?.length) return "";
  return fields.map(f => {
    if (f.type === "operator") return ` ${f.value} `;
    if (f.type === "constant") return String(f.value);
    if (f.checklist === "self" && selfQuestions) return selfQuestions[f.question]?.text || selfQuestions[f.question] || `Q${f.question+1}`;
    const ck = allChecklists?.find(c => c.id === f.checklist);
    const qText = ck ? getQText(ck.questions[f.question]) : `Q${f.question+1}`;
    return ck ? `[${ck.name}] ${qText}` : qText;
  }).join("");
};

const globalCss = `
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
html{font-size:16px}
body{font-family:${T.font};background:${T.bg};color:${T.text};min-height:100vh;-webkit-font-smoothing:antialiased}
input,textarea,select,button{font-family:inherit}
::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:${T.border};border-radius:3px}
@keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
@keyframes slideIn{from{opacity:0;transform:translateX(-16px)}to{opacity:1;transform:translateX(0)}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
@keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
@keyframes spinnerBar{0%{transform:scaleX(0);transform-origin:left}50%{transform:scaleX(1);transform-origin:left}50.1%{transform-origin:right}100%{transform:scaleX(0);transform-origin:right}}
@keyframes toastIn{from{opacity:0;transform:translateY(16px) scale(.95)}to{opacity:1;transform:translateY(0) scale(1)}}
@keyframes toastOut{from{opacity:1;transform:translateY(0) scale(1)}to{opacity:0;transform:translateY(16px) scale(.95)}}
.fade-up{animation:fadeUp .4s ease forwards}
.slide-in{animation:slideIn .35s ease forwards}
`;

// ─── UI Primitives ────────────────────────────────────────────

const Badge = ({ children, variant="default", style={} }) => {
  const v = { default:{bg:T.accentBg,color:T.accent,border:T.accentBorder}, success:{bg:T.successBg,color:T.success,border:T.successBorder}, muted:{bg:"rgba(107,103,119,0.1)",color:T.textSec,border:"rgba(107,103,119,0.15)"}, danger:{bg:T.dangerBg,color:T.danger,border:"rgba(232,93,93,0.2)"}, info:{bg:T.infoBg,color:T.info,border:T.infoBorder} }[variant];
  return <span style={{ display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:500,letterSpacing:".02em",background:v.bg,color:v.color,border:`1px solid ${v.border}`,whiteSpace:"nowrap",...style }}>{children}</span>;
};

const Btn = ({ children, variant="primary", onClick, style={}, disabled, small }) => {
  const base = { display:"inline-flex",alignItems:"center",justifyContent:"center",gap:8,padding:small?"8px 14px":"11px 20px",borderRadius:T.radSm,fontWeight:500,fontSize:small?13:14,cursor:disabled?"not-allowed":"pointer",border:"none",transition:"all .2s",opacity:disabled?.5:1,letterSpacing:".01em" };
  const vs = { primary:{...base,background:T.accent,color:T.bg}, secondary:{...base,background:T.surfaceHover,color:T.text,border:`1px solid ${T.border}`}, ghost:{...base,background:"transparent",color:T.textSec}, danger:{...base,background:T.dangerBg,color:T.danger,border:"1px solid rgba(232,93,93,0.2)"}, success:{...base,background:T.successBg,color:T.success,border:`1px solid ${T.successBorder}`} };
  return <button style={{...vs[variant],...style}} onClick={onClick} disabled={disabled}>{children}</button>;
};

// `min` / `max` pass through to the DOM input. For type="number", `min` defaults to "0"
// unless `allowNegative` is set. When a negative value is entered and `allowNegative` is
// false, we clamp to 0 on change/blur AND invoke `onNegativeAttempt` so callers can surface
// an inline error message.
const Input = ({ value, onChange, placeholder, style={}, type="text", readOnly=false, min, max, allowNegative=false, onNegativeAttempt, onBlur }) => {
  const effectiveMin = min !== undefined ? min : (type === "number" && !allowNegative ? "0" : undefined);
  const handleChange = (e) => {
    const v = e.target.value;
    if (type === "number" && !allowNegative && v !== "" && v !== "-") {
      const n = parseFloat(v);
      if (!isNaN(n) && n < 0) {
        if (typeof onNegativeAttempt === "function") onNegativeAttempt(n);
        onChange("0");
        return;
      }
    }
    onChange(v);
  };
  const handleBlur = (e) => {
    if (type === "number" && !allowNegative) {
      const n = parseFloat(e.target.value);
      if (!isNaN(n) && n < 0) {
        if (typeof onNegativeAttempt === "function") onNegativeAttempt(n);
        onChange("0");
      }
    }
    if (typeof onBlur === "function") onBlur(e);
  };
  return (
    <input type={type} value={value} readOnly={readOnly} min={effectiveMin} max={max}
      onChange={handleChange} onBlur={handleBlur} placeholder={placeholder}
      style={{ width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:readOnly?T.surfaceHover:T.bg,border:`1px solid ${T.border}`,color:readOnly?T.textSec:T.text,fontSize:14,outline:"none",transition:"border .2s",cursor:readOnly?"not-allowed":"text",...style }}/>
  );
};

const Field = ({ label, children, style={} }) => (
  <div style={style}><label style={{ display:"block",fontSize:13,fontWeight:500,color:T.textSec,marginBottom:8,letterSpacing:".02em" }}>{label}</label>{children}</div>
);

const Section = ({ children, count, icon, action }) => (
  <div style={{ display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16 }}>
    <div style={{ display:"flex",alignItems:"center",gap:10 }}>
      {icon && <Icon name={icon} size={18} color={T.accent}/>}
      <h3 style={{ fontSize:15,fontWeight:600,letterSpacing:".02em",color:T.text }}>{children}</h3>
      {count!==undefined && <Badge variant="muted">{count}</Badge>}
    </div>
    {action}
  </div>
);

const Empty = ({ icon, text, sub }) => (
  <div style={{ textAlign:"center",padding:"28px 20px",marginBottom:24,background:T.surface,borderRadius:T.rad,border:`1px dashed ${T.border}` }}>
    <Icon name={icon} size={28} color={T.textMut}/><p style={{ fontSize:14,color:T.textSec,marginTop:8 }}>{text}</p>
    {sub && <p style={{ fontSize:12,color:T.textMut,marginTop:4 }}>{sub}</p>}
  </div>
);

const Chip = ({ label, active, onClick }) => (
  <button onClick={onClick} style={{ padding:"7px 14px",borderRadius:20,border:`1px solid ${active?T.accent:T.border}`,background:active?T.accentBg:"transparent",color:active?T.accent:T.textSec,fontSize:13,fontWeight:500,cursor:"pointer",transition:"all .2s" }}>{label}</button>
);

const YesNoButtons = ({ value, onChange }) => (
  <div style={{display:"flex",gap:10}}>
    {["Yes","No"].map(opt=>(
      <button key={opt} onClick={()=>onChange(opt)} style={{
        flex:1,padding:"14px 0",borderRadius:T.radSm,fontSize:16,fontWeight:600,cursor:"pointer",transition:"all .2s",
        border:`2px solid ${value===opt?(opt==="Yes"?T.success:T.danger):T.border}`,
        background:value===opt?(opt==="Yes"?T.successBg:T.dangerBg):"transparent",
        color:value===opt?(opt==="Yes"?T.success:T.danger):T.textSec,
      }}>{opt}</button>
    ))}
  </div>
);

const Toggle = ({ active, onClick, label, sub }) => (
  <button onClick={onClick} style={{ display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderRadius:T.radSm,border:`1px solid ${active?T.accentBorder:T.border}`,background:active?T.accentBg:"transparent",cursor:"pointer",textAlign:"left",transition:"all .2s",width:"100%" }}>
    <div style={{ width:18,height:18,borderRadius:4,border:`2px solid ${active?T.accent:T.borderLight}`,background:active?T.accent:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0 }}>
      {active && <Icon name="check" size={12} color={T.bg}/>}
    </div>
    <div>
      <span style={{ fontSize:13,fontWeight:500,color:T.text }}>{label}</span>
      {sub && <span style={{ display:"block",fontSize:11,color:T.textMut }}>{sub}</span>}
    </div>
  </button>
);

// ─── Question Input Renderer ─────────────────────────────────

function QuestionInputField({ q, qi, currentVal, idealVal, needsRemark, formData, setFormData, approvedEntries, checklists, getFieldValue, orders, customers, inventoryItems, onBatchAllocChange, allQuestions, onInventoryAutoFill }) {
  const [negativeError, setNegativeError] = useState(false);
  useEffect(() => { if (negativeError) { const t = setTimeout(() => setNegativeError(false), 3500); return () => clearTimeout(t); } }, [negativeError]);
  const isAutoFilled = !!(formData.autoFilled && formData.autoFilled[qi]);
  const autoFilledReadOnly = !!(formData.autoFilledReadOnly && formData.autoFilledReadOnly[qi]);
  const autoFilledSource = (formData.autoFilledSource || {})[qi] || null;

  // Apply auto-fill from a selected linked source entry using per-question autoFillMapping
  const applyLinkedAutoFill = (sourceEntry) => {
    if (!sourceEntry || !Array.isArray(sourceEntry.responses) || !Array.isArray(allQuestions)) return;
    const srcCk = checklists?.find(c => c.id === q.linkedSource?.checklistId);
    const srcName = srcCk?.name || "source";
    const srcQ = srcCk ? normalizeQuestions(srcCk.questions) : [];

    setFormData(p => {
      const nextResponses = { ...p.responses };
      const nextAutoFilled = { ...(p.autoFilled || {}) };
      const nextAutoFilledReadOnly = { ...(p.autoFilledReadOnly || {}) };
      const nextAutoFilledSource = { ...(p.autoFilledSource || {}) };

      let hasAnyMapping = false;
      // Scan all questions for per-question autoFillMapping
      allQuestions.forEach((tq, tgtIdx) => {
        if (tgtIdx === qi || tq.linkedSource) return; // skip the linked dropdown itself
        const m = tq.autoFillMapping;
        if (!m || m.sourceFieldIdx === "" || m.sourceFieldIdx === undefined) return;
        hasAnyMapping = true;
        const srcIdx = Number(m.sourceFieldIdx);
        const srcField = srcQ[srcIdx];
        if (!srcField) return;
        const srcResp = sourceEntry.responses.find(r => r.question === srcField.text);
        if (!srcResp || srcResp.response === undefined || srcResp.response === "") return;
        let val = srcResp.response;
        // Inventory item responses are stored as display name — convert back to id for dropdown
        if (tq.type === "inventory_item" && Array.isArray(inventoryItems)) {
          const match = inventoryItems.find(it => it.name === val || it.id === val);
          if (match) val = match.id;
        }
        nextResponses[tgtIdx] = String(val);
        nextAutoFilled[tgtIdx] = true;
        nextAutoFilledReadOnly[tgtIdx] = m.readOnly !== false;
        nextAutoFilledSource[tgtIdx] = srcName;
      });

      if (!hasAnyMapping) {
        // Fallback: match by question text (legacy behavior for checklists without per-question mappings)
        sourceEntry.responses.forEach(r => {
          if (!r || !r.question) return;
          const targetIdx = allQuestions.findIndex((tq, ti) => ti !== qi && tq.text === r.question && !tq.linkedSource);
          if (targetIdx >= 0 && r.response !== undefined && r.response !== "") {
            let val = r.response;
            if (allQuestions[targetIdx].type === "inventory_item" && Array.isArray(inventoryItems)) {
              const match = inventoryItems.find(it => it.name === val || it.id === val);
              if (match) val = match.id;
            }
            nextResponses[targetIdx] = String(val);
            nextAutoFilled[targetIdx] = true;
            nextAutoFilledReadOnly[targetIdx] = true;
            nextAutoFilledSource[targetIdx] = srcName;
          }
        });
      }
      return { ...p, responses: nextResponses, autoFilled: nextAutoFilled, autoFilledReadOnly: nextAutoFilledReadOnly, autoFilledSource: nextAutoFilledSource };
    });

    // Auto-fill inventory tracking section if applicable
    if (typeof onInventoryAutoFill === "function") {
      for (let si = 0; si < srcQ.length; si++) {
        if (srcQ[si].type === "inventory_item") {
          const srcResp = sourceEntry.responses.find(r => r.question === srcQ[si].text);
          if (srcResp && srcResp.response) {
            const itemMatch = (inventoryItems || []).find(it => it.name === srcResp.response || it.id === srcResp.response);
            if (itemMatch) onInventoryAutoFill(itemMatch);
          }
          break;
        }
      }
    }
  };

  // Clear auto-filled values when linked dropdown is cleared
  const clearLinkedAutoFill = () => {
    setFormData(p => {
      const nextResponses = { ...p.responses };
      const nextAutoFilled = { ...(p.autoFilled || {}) };
      const nextAutoFilledReadOnly = { ...(p.autoFilledReadOnly || {}) };
      const nextAutoFilledSource = { ...(p.autoFilledSource || {}) };
      Object.keys(nextAutoFilled).forEach(k => {
        if (nextAutoFilled[k]) {
          delete nextResponses[k];
          delete nextAutoFilled[k];
          delete nextAutoFilledReadOnly[k];
          delete nextAutoFilledSource[k];
        }
      });
      return { ...p, responses: nextResponses, autoFilled: nextAutoFilled, autoFilledReadOnly: nextAutoFilledReadOnly, autoFilledSource: nextAutoFilledSource };
    });
    if (typeof onInventoryAutoFill === "function") onInventoryAutoFill(null);
  };
  // Linked source dropdown — branches into BatchSelector when source has IN-tracked quantity
  if(q.linkedSource&&q.linkedSource.checklistId){
    const entries=approvedEntries?.[q.linkedSource.checklistId]||[];
    const srcCk=checklists?.find(c=>c.id===q.linkedSource.checklistId);
    const useBatch=sourceHasInTrackedQty(srcCk)&&typeof onBatchAllocChange==="function";
    return (
      <div key={qi}>
        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
          <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0,marginTop:1}}>{String(qi+1).padStart(2,"0")}</span>
          <span style={{fontSize:13,color:T.textSec,lineHeight:1.4,flex:1}}>{q.text}</span>
        </div>
        {useBatch ? (
          <BatchSelector entries={entries} allocations={formData.batchAllocations?.[qi]||[]}
            onChange={allocs=>{
              onBatchAllocChange(qi,allocs);
              // Auto-fill matching fields from the FIRST selected batch
              clearLinkedAutoFill();
              const firstId=(allocs||[]).map(a=>a.sourceAutoId).find(Boolean);
              if(firstId){
                const picked=entries.find(en=>(en.autoId&&en.autoId===firstId)||en.linkedId===firstId);
                if(picked) applyLinkedAutoFill(picked);
              }
            }}
            checklistName={srcCk?.name||"Source"}
            emptyMessage={`No approved batches found in ${srcCk?.name||"source checklist"}`}/>
        ) : (
          <LinkedDropdown entries={entries} value={currentVal} onChange={v=>{
            // Clear previous auto-fills first, then apply new
            clearLinkedAutoFill();
            setFormData(p=>({...p,responses:{...p.responses,[qi]:v}}));
            if(v){
              const picked=entries.find(en=>(en.autoId&&en.autoId===v)||en.linkedId===v);
              if(picked) applyLinkedAutoFill(picked);
            }
          }}
            checklistName={srcCk?.name||"Source"} sourceChecklistId={q.linkedSource.checklistId} checklists={checklists} placeholder="Select approved entry..." emptyMessage={`No approved entries found in ${srcCk?.name||"source checklist"}`}/>
        )}
        {(()=>{
          // Derive the linked autoId from either the direct response or the first valid batch allocation
          const directId = currentVal ? (typeof currentVal==="string"?currentVal.split(",")[0].trim():currentVal) : "";
          const batchId = !directId && Array.isArray(formData.batchAllocations?.[qi])
            ? (formData.batchAllocations[qi].find(a=>a.sourceAutoId)?.sourceAutoId || "")
            : "";
          const chainAutoId = directId || batchId;
          return chainAutoId && q.linkedSource?.checklistId
            ? <SourceChainDisplay checklistId={q.linkedSource.checklistId} autoId={chainAutoId} checklists={checklists}/>
            : null;
        })()}
      </div>
    );
  }

  // Inventory item dropdown
  if(q.type === "inventory_item"){
    const items=(inventoryItems||[]).filter(it=>it.isActive&&(!q.inventoryCategory||it.category===q.inventoryCategory));
    return (
      <div key={qi}>
        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
          <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0,marginTop:1}}>{String(qi+1).padStart(2,"0")}</span>
          <span style={{fontSize:13,color:T.textSec,lineHeight:1.4,flex:1}}>{q.text}</span>
          {isAutoFilled&&<Badge variant="info" style={{fontSize:9}}>Auto-filled from {autoFilledSource||"source"}</Badge>}
        </div>
        <SearchableDropdown
          options={items.map(it=>({label:it.name+(it.abbreviation?` (${it.abbreviation})`:""),value:it.id}))}
          value={currentVal} onChange={v=>setFormData(p=>({...p,responses:{...p.responses,[qi]:v}}))}
          disabled={isAutoFilled&&autoFilledReadOnly} placeholder="— Select item —"/>
      </div>
    );
  }

  // Approval gate or yesno type
  if(q.isApprovalGate || q.type === "yesno"){
    return (
      <div key={qi}>
        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
          <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0,marginTop:1}}>{String(qi+1).padStart(2,"0")}</span>
          <span style={{fontSize:13,color:T.textSec,lineHeight:1.4,flex:1}}>{q.text}</span>
          {q.isApprovalGate&&<Badge variant="success" style={{fontSize:10}}>Approval Gate</Badge>}
        </div>
        <YesNoButtons value={currentVal} onChange={v=>setFormData(p=>({...p,responses:{...p.responses,[qi]:v}}))}/>
      </div>
    );
  }

  // Date type
  if(q.type === "date"){
    const age = formatAgeFromDate(currentVal);
    // Date comparison validation
    let dateError = null;
    if (q.dateComparison && q.dateComparison.compareToFieldIdx !== "" && q.dateComparison.compareToFieldIdx !== undefined && currentVal) {
      const cmpIdx = Number(q.dateComparison.compareToFieldIdx);
      const cmpVal = formData.responses[cmpIdx] || "";
      if (cmpVal) {
        const d1 = new Date(currentVal), d2 = new Date(cmpVal);
        if (!isNaN(d1.getTime()) && !isNaN(d2.getTime())) {
          const t1 = d1.setHours(0,0,0,0), t2 = d2.setHours(0,0,0,0);
          const op = q.dateComparison.operator;
          const cmpField = allQuestions?.[cmpIdx]?.text || "the other date";
          if (op === "gte" && t1 < t2) dateError = q.dateComparison.errorMessage || `${q.text} cannot be before ${cmpField}`;
          else if (op === "lte" && t1 > t2) dateError = q.dateComparison.errorMessage || `${q.text} cannot be after ${cmpField}`;
          else if (op === "eq" && t1 !== t2) dateError = q.dateComparison.errorMessage || `${q.text} must be the same as ${cmpField}`;
        }
      }
    }
    return (
      <div key={qi}>
        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
          <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0,marginTop:1}}>{String(qi+1).padStart(2,"0")}</span>
          <span style={{fontSize:13,color:T.textSec,lineHeight:1.4,flex:1}}>{q.text}</span>
          {isAutoFilled&&<Badge variant="info" style={{fontSize:9}}>Auto-filled from {autoFilledSource||"source"}</Badge>}
        </div>
        <input type="date" value={currentVal} readOnly={isAutoFilled&&autoFilledReadOnly} onChange={e=>{if(isAutoFilled&&autoFilledReadOnly)return;setFormData(p=>({...p,responses:{...p.responses,[qi]:e.target.value}}))}}
          style={{width:"100%",padding:"12px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${dateError?T.danger:isAutoFilled?T.infoBorder:T.border}`,color:T.text,fontSize:15,outline:"none",colorScheme:"dark",opacity:(isAutoFilled&&autoFilledReadOnly)?0.7:1}}/>
        {dateError && <div style={{marginTop:4,padding:"6px 10px",background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm,fontSize:12,color:T.danger}}>{dateError}</div>}
        {age && !dateError && <span style={{fontSize:11,color:T.textMut,marginTop:4,display:"block"}}>{age}</span>}
      </div>
    );
  }

  // Invoice/SO searchable dropdown for Grinding & Packing checklist
  if(q.text === "Invoice/SO" && orders) {
    const orderOptions = orders.filter(o => o.canTag !== false && o.status !== "delivered" && o.status !== "cancelled").map(o => {
      const cust = customers?.find(c => c.id === o.customerId);
      return { label: `${o.id} — ${cust?.label||""} — ${o.invoiceSo||"N/A"}`, value: o.invoiceSo || o.id };
    });
    const selectedOrder = orders.find(o => (o.invoiceSo || o.id) === currentVal);
    const selectedCust = selectedOrder ? customers?.find(c => c.id === selectedOrder.customerId) : null;
    return (
      <div key={qi}>
        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
          <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0,marginTop:1}}>{String(qi+1).padStart(2,"0")}</span>
          <span style={{fontSize:13,color:T.textSec,lineHeight:1.4,flex:1}}>{q.text}</span>
        </div>
        <SearchableDropdown options={orderOptions} value={currentVal} onChange={v=>{
          setFormData(p=>({...p,responses:{...p.responses,[qi]:v}}));
          // Auto-populate client name if next question is "Client name"
          const matchOrder = orders.find(o => (o.invoiceSo || o.id) === v);
          if(matchOrder) {
            const matchCust = customers?.find(c => c.id === matchOrder.customerId);
            if(matchCust) setFormData(p=>({...p,responses:{...p.responses,[qi]:v,[qi+1]:matchCust.label}}));
          }
        }} placeholder="Search orders..." emptyMessage="No orders found"/>
        {selectedOrder && selectedOrder.orderLines?.length > 0 && (
          <div style={{marginTop:8,padding:10,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
            <span style={{fontSize:11,color:T.textMut,display:"block",marginBottom:6}}>Order Blend Lines:</span>
            {selectedOrder.orderLines.map((line,li)=>{
              const tq = parseFloat(line.taggedQuantity)||0;
              const rq = parseFloat(line.quantity)||0;
              return <div key={li} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"4px 0",borderBottom:li<selectedOrder.orderLines.length-1?`1px solid ${T.border}`:"none"}}>
                <span style={{fontSize:13,color:T.text,fontWeight:500}}>{line.blend||"Blend "+(li+1)}</span>
                <div style={{display:"flex",gap:8,fontSize:11}}>
                  <span style={{color:T.textSec}}>Req: {rq}</span>
                  <span style={{color:T.warning}}>Tagged: {tq}</span>
                  <span style={{color:rq-tq>0?T.success:T.danger}}>Rem: {rq-tq}</span>
                </div>
              </div>;
            })}
          </div>
        )}
      </div>
    );
  }

  // Forced remark target — another question's remarkCondition points at this question via remarksTargetIdx
  let forcedRemark = null;
  if (Array.isArray(allQuestions) && typeof getFieldValue === "function") {
    for (let oi = 0; oi < allQuestions.length; oi++) {
      const oq = allQuestions[oi];
      if (!oq?.remarkCondition || oq.remarksTargetIdx !== qi || !oq.formula || !oq.ideal) continue;
      const actualVal = evaluateFormula(oq.formula, getFieldValue);
      const idealVal2 = evaluateFormula(oq.ideal, getFieldValue);
      if (actualVal != null && idealVal2 != null && checkRemarkCondition(actualVal, idealVal2, oq.remarkCondition)) {
        const diff = Math.round((actualVal - idealVal2) * 100) / 100;
        const pct = idealVal2 !== 0 ? Math.round((diff / idealVal2) * 100) : 0;
        forcedRemark = { sourceText: oq.text, message: oq.remarkCondition.message || "Value differs from ideal — remarks required", actual: actualVal, ideal: idealVal2, diff, pct };
        break;
      }
    }
  }

  // Default text/number input
  const isNum=q.type==="number"||q.type==="text_number";
  return (
    <div key={qi}>
      <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:6}}>
        <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0,marginTop:1}}>{String(qi+1).padStart(2,"0")}</span>
        <span style={{fontSize:13,color:T.textSec,lineHeight:1.4,flex:1}}>{q.text}</span>
        {idealVal!==null&&<span style={{fontSize:11,color:T.accent,whiteSpace:"nowrap",padding:"2px 8px",background:T.accentBg,borderRadius:12,border:`1px solid ${T.accentBorder}`}}>Ideal: {idealVal}{(q.idealUnit||q.ideal?.suffix)?" "+(q.idealUnit||q.ideal.suffix):""}{q.idealLabel?` — ${q.idealLabel}`:""}</span>}
        {isAutoFilled&&<Badge variant="info" style={{fontSize:9}}>Auto-filled from {autoFilledSource||"source"}</Badge>}
      </div>
      {forcedRemark && (
        <div style={{marginBottom:8,padding:10,background:T.warningBg,border:`1px solid ${T.warningBorder}`,borderRadius:T.radSm}}>
          <p style={{fontSize:12,color:T.warning,fontWeight:600,marginBottom:4}}>⚠ {forcedRemark.message}</p>
          <p style={{fontSize:11,color:T.warning,fontFamily:T.mono}}>Actual: {forcedRemark.actual}{q.idealUnit?` ${q.idealUnit}`:""} · Ideal: {forcedRemark.ideal}{q.idealUnit?` ${q.idealUnit}`:""} · Diff: {forcedRemark.diff>=0?"+":""}{forcedRemark.diff} ({forcedRemark.pct>=0?"+":""}{forcedRemark.pct}%)</p>
        </div>
      )}
      <Input value={currentVal} readOnly={!!q.formula||(isAutoFilled&&autoFilledReadOnly)} onChange={v=>{ if((isAutoFilled&&autoFilledReadOnly)||q.formula) return; setFormData(p=>({...p,responses:{...p.responses,[qi]:v}})); }} placeholder={forcedRemark?"Required: explain the deviation...":(q.formula?"Auto-calculated...":"Enter response...")} type={isNum?"number":"text"} min={isNum?"0":undefined} onNegativeAttempt={isNum?(()=>setNegativeError(true)):undefined} style={{fontSize:15,padding:"12px 14px",opacity:(isAutoFilled&&autoFilledReadOnly)?0.7:1,border:(negativeError?`1px solid ${T.danger}`:forcedRemark?`1px solid ${T.warning}`:undefined)}}/>
      {negativeError && (
        <div style={{marginTop:6,padding:"6px 10px",background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm,fontSize:12,color:T.danger}}>Value cannot be negative</div>
      )}
      {needsRemark&&(
        <div style={{marginTop:8,padding:10,background:T.warningBg,border:`1px solid ${T.warningBorder}`,borderRadius:T.radSm}}>
          <p style={{fontSize:12,color:T.warning,marginBottom:6}}>{q.remarkCondition.message||"Value differs from ideal. Please provide reason."}</p>
          <textarea value={formData.remarks[qi]||""} onChange={e=>setFormData(p=>({...p,remarks:{...p.remarks,[qi]:e.target.value}}))} placeholder="Enter remarks (required)..." rows={2}
            style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.warningBorder}`,color:T.text,fontSize:13,outline:"none",resize:"vertical",fontFamily:T.font}}/>
        </div>
      )}
    </div>
  );
}

// ─── Skeleton Components ──────────────────────────────────────

const shimmerStyle = {
  background:`linear-gradient(90deg, ${T.card} 25%, ${T.surfaceHover} 50%, ${T.card} 75%)`,
  backgroundSize:"200% 100%",
  animation:"shimmer 1.5s infinite",
  borderRadius:T.radSm,
};

const SkeletonCard = ({ delay = 0 }) => (
  <div style={{ background:T.card, borderRadius:T.rad, padding:"16px 18px", border:`1px solid ${T.border}`, marginBottom:10, animationDelay:`${delay}ms` }}>
    <div style={{ display:"flex", justifyContent:"space-between", marginBottom:12 }}>
      <div>
        <div style={{ ...shimmerStyle, width:140, height:16, marginBottom:8 }} />
        <div style={{ display:"flex", gap:8 }}>
          <div style={{ ...shimmerStyle, width:90, height:20, borderRadius:20 }} />
          <div style={{ ...shimmerStyle, width:80, height:20, borderRadius:20 }} />
        </div>
      </div>
      <div style={{ ...shimmerStyle, width:18, height:18 }} />
    </div>
    <div style={{ display:"flex", alignItems:"center", gap:10 }}>
      <div style={{ ...shimmerStyle, flex:1, height:4 }} />
      <div style={{ ...shimmerStyle, width:30, height:14 }} />
    </div>
  </div>
);

const SkeletonList = () => (
  <div className="fade-up">
    <div style={{ ...shimmerStyle, width:"100%", height:44, marginBottom:24 }} />
    <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:16 }}>
      <div style={{ ...shimmerStyle, width:18, height:18 }} />
      <div style={{ ...shimmerStyle, width:120, height:16 }} />
    </div>
    {[0,1,2].map(i => <SkeletonCard key={i} delay={i * 100} />)}
  </div>
);

// ─── Toast Notification System ────────────────────────────────

function ToastContainer({ toasts, onDismiss, onRetry }) {
  if (toasts.length === 0) return null;
  return (
    <div style={{ position:"fixed", bottom:90, left:"50%", transform:"translateX(-50%)", zIndex:100, display:"flex", flexDirection:"column", gap:8, width:"calc(100% - 32px)", maxWidth:560, pointerEvents:"none" }}>
      {toasts.map(t => (
        <div key={t.id} style={{
          background: t.type === "error" ? "rgba(232,93,93,0.95)" : t.type === "success" ? "rgba(107,203,119,0.95)" : "rgba(212,165,116,0.95)",
          color: "#fff", borderRadius:T.rad, padding:"12px 16px", display:"flex", alignItems:"center", gap:10, justifyContent:"space-between",
          animation: t.leaving ? "toastOut .25s ease forwards" : "toastIn .25s ease forwards",
          backdropFilter:"blur(8px)", pointerEvents:"auto", boxShadow:"0 4px 24px rgba(0,0,0,0.4)",
        }}>
          <span style={{ fontSize:13, fontWeight:500, flex:1 }}>{t.message}</span>
          <div style={{ display:"flex", gap:6, flexShrink:0 }}>
            {t.retryFn && <button onClick={() => onRetry(t)} style={{ background:"rgba(255,255,255,0.2)", border:"none", borderRadius:6, padding:"4px 10px", fontSize:12, fontWeight:600, color:"#fff", cursor:"pointer" }}>Retry</button>}
            <button onClick={() => onDismiss(t.id)} style={{ background:"none", border:"none", cursor:"pointer", padding:2, display:"flex" }}><Icon name="x" size={16} color="rgba(255,255,255,0.7)" /></button>
          </div>
        </div>
      ))}
    </div>
  );
}

// ─── Loading Bar ──────────────────────────────────────────────

function LoadingBar({ visible }) {
  if (!visible) return null;
  return <div style={{ position:"absolute", bottom:0, left:0, right:0, height:2, overflow:"hidden" }}>
    <div style={{ width:"100%", height:"100%", background:T.accent, animation:"spinnerBar 1.2s ease-in-out infinite" }} />
  </div>;
}

// ─── Login View ───────────────────────────────────────────────

function LoginView({ onLogin }) {
  const [username,setUsername]=useState("");
  const [password,setPassword]=useState("");
  const [loading,setLoading]=useState(false);
  const [error,setError]=useState(null);

  const handleSubmit = async (e) => {
    e?.preventDefault();
    if (!username.trim()||!password) return;
    setLoading(true); setError(null);
    try {
      const result = await API.post("login", { username: username.trim(), password });
      API.setToken(result.token);
      localStorage.setItem("ocm_token", result.token);
      localStorage.setItem("ocm_user", JSON.stringify(result.user));
      onLogin(result.token, result.user);
    } catch (err) { setError(err.message); }
    setLoading(false);
  };

  return (
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:"100vh",background:T.bg,padding:20}}>
      <style>{globalCss}</style>
      <div className="fade-up" style={{width:"100%",maxWidth:360,textAlign:"center"}}>
        <Icon name="coffee" size={48} color={T.accent}/>
        <h1 style={{fontSize:22,fontWeight:700,color:T.text,marginTop:12,marginBottom:4}}>Sunoha</h1>
        <p style={{fontSize:14,color:T.textSec,marginBottom:32}}>Order Checklists — Sign in to continue</p>
        {error && <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm,padding:"10px 14px",marginBottom:16,textAlign:"left"}}>
          <span style={{fontSize:13,color:T.danger}}>{error}</span>
        </div>}
        <form onSubmit={handleSubmit} style={{display:"flex",flexDirection:"column",gap:14,textAlign:"left"}}>
          <Field label="Username"><Input value={username} onChange={setUsername} placeholder="Enter username"/></Field>
          <Field label="Password"><Input value={password} onChange={setPassword} placeholder="Enter password" type="password"/></Field>
          <Btn onClick={handleSubmit} disabled={loading||!username.trim()||!password} style={{width:"100%",marginTop:8}}>
            <Icon name={loading?"clock":"lock"} size={16} color={T.bg}/> {loading ? "Signing in..." : "Sign In"}
          </Btn>
        </form>
      </div>
    </div>
  );
}

// ─── App ───────────────────────────────────────────────────────

export default function App() {
  const [authToken,setAuthToken]=useState(()=>localStorage.getItem("ocm_token"));
  const [currentUser,setCurrentUser]=useState(()=>{try{return JSON.parse(localStorage.getItem("ocm_user"))}catch{return null}});
  const isAdmin = currentUser?.role === "admin";

  const [view,setView]=useState("orders");
  const [subView,setSubView]=useState(null);
  const [checklists,setChecklists]=useState([]);
  const [orderTypes,setOrderTypes]=useState([]);
  const [customers,setCustomers]=useState([]);
  const [orders,setOrders]=useState([]);
  const [rules,setRules]=useState([]);
  const [untaggedChecklists,setUntaggedChecklists]=useState([]);
  const [approvedEntries,setApprovedEntries]=useState({});
  const [inventoryItems,setInventoryItems]=useState([]);
  const [inventoryCategories,setInventoryCategories]=useState([]);
  const [inventorySummary,setInventorySummary]=useState({greenBeans:0,roastedBeans:0,packedGoods:0,lowStockCount:0});
  const [blends,setBlends]=useState([]);
  const [drafts,setDrafts]=useState([]);
  const [orderStageTemplates,setOrderStageTemplates]=useState({});
  const [resumeDraft,setResumeDraft]=useState(null);
  const [selected,setSelected]=useState(null);
  const [detailOrder,setDetailOrder]=useState(null);
  const [loaded,setLoaded]=useState(false);
  const [error,setError]=useState(null);
  const [busy,setBusy]=useState(null);
  const [showAccount,setShowAccount]=useState(false);

  // API loading indicator state
  const [apiLoading,setApiLoading]=useState(false);
  useEffect(() => {
    const listener = (loading) => setApiLoading(loading);
    _apiListeners.push(listener);
    return () => { _apiListeners = _apiListeners.filter(l => l !== listener); };
  }, []);

  // Toast notification state
  const [toasts,setToasts]=useState([]);
  const toastIdRef = useRef(0);

  const addToast = useCallback((message, type = "error", retryFn = null) => {
    const id = ++toastIdRef.current;
    setToasts(prev => [...prev.slice(-4), { id, message, type, retryFn }]);
    if (!retryFn) setTimeout(() => dismissToast(id), 4000);
    else setTimeout(() => dismissToast(id), 8000);
  }, []);

  const dismissToast = useCallback((id) => {
    setToasts(prev => prev.map(t => t.id === id ? { ...t, leaving: true } : t));
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 300);
  }, []);

  const handleRetryToast = useCallback((toast) => {
    dismissToast(toast.id);
    if (toast.retryFn) toast.retryFn();
  }, [dismissToast]);

  const handleLogin = (token, user) => { setAuthToken(token); setCurrentUser(user); };

  const handleLogout = async () => {
    try { await API.post("logout"); } catch {}
    localStorage.removeItem("ocm_token"); localStorage.removeItem("ocm_user");
    API.setToken(null); setAuthToken(null); setCurrentUser(null);
  };

  // Single batch load using the init endpoint
  const loadAll = useCallback(async () => {
    try {
      const data = await API.get("init");
      setChecklists(data.checklists);
      setOrderTypes(data.orderTypes);
      setOrders(data.orders);
      setRules(data.rules);
      setCustomers(data.customers);
      if (data.untaggedChecklists) setUntaggedChecklists(data.untaggedChecklists);
      if (data.approvedEntries) setApprovedEntries(data.approvedEntries);
      if (data.inventoryItems) setInventoryItems(data.inventoryItems);
      if (data.inventoryCategories) setInventoryCategories(data.inventoryCategories);
      if (data.inventorySummary) setInventorySummary(data.inventorySummary);
      if (data.blends) setBlends(data.blends);
      if (data.drafts) setDrafts(data.drafts);
      if (data.orderStageTemplates) setOrderStageTemplates(data.orderStageTemplates);
      setError(null);
    } catch(e) {
      console.error("Load error:", e);
      setError("Failed to load data. Check the Apps Script URL and deployment.");
    }
    setLoaded(true);
  }, []);

  useEffect(() => { if(authToken) loadAll(); }, [authToken, loadAll]);

  // Individual refresh functions — only used as fallback after specific writes
  const refreshOrders = async () => { try { setOrders(await API.get("getOrders")); } catch(e) { addToast("Failed to refresh orders", "error", refreshOrders); } };
  const refreshChecklists = async () => { try { setChecklists(await API.get("getChecklists")); } catch(e) { addToast("Failed to refresh checklists", "error", refreshChecklists); } };
  const refreshRules = async () => { try { setRules(await API.get("getRules")); } catch(e) { addToast("Failed to refresh rules", "error", refreshRules); } };
  const refreshOrderTypes = async () => { try { setOrderTypes(await API.get("getOrderTypes")); } catch(e) { addToast("Failed to refresh order types", "error", refreshOrderTypes); } };
  const refreshCustomers = async () => { try { setCustomers(await API.get("getCustomers")); } catch(e) { addToast("Failed to refresh customers", "error", refreshCustomers); } };
  const refreshBlends = async () => { try { setBlends(await API.get("getBlends")); } catch(e) { addToast("Failed to refresh blends", "error", refreshBlends); } };
  const refreshDrafts = async () => { try { setDrafts(await API.get("getDrafts")); } catch(e) { addToast("Failed to refresh drafts", "error", refreshDrafts); } };

  const goBack = () => {
    if (subView==="editResponses" && detailOrder) {
      setSubView("orderDetail"); setSelected(detailOrder); setDetailOrder(null);
    } else if (subView) {
      setSubView(null); setSelected(null);
    } else {
      setView("orders"); setSelected(null);
    }
  };

  const switchTab = (tab) => { setView(tab); setSubView(null); setSelected(null); setDetailOrder(null); };
  const currentView = subView || view;
  const titles = {orders:"Sunoha Checklists",responses:"Responses Log",admin:"Settings",users:"User Management",inventory:"Inventory",inventoryLedger:"Item Ledger",orderDetail:"Order Details",newOrder:"New Order",editChecklist:selected?"Edit Checklist":"New Checklist",editRules:"Assignment Rules",addRule:"New Rule",editRule:"Edit Rule",editResponses:"Edit Responses",quickFill:"Fill Checklist",blends:"Blends",editBlend:selected?"Edit Blend":"New Blend"};
  const isTabView = ["orders","responses","admin","users","inventory","blends"].includes(currentView);

  if (!authToken || !currentUser) return <LoginView onLogin={handleLogin}/>;

  // Show skeleton loading instead of blank screen
  if (!loaded) return (
    <><style>{globalCss}</style>
    <div style={{minHeight:"100vh",background:T.bg,paddingBottom:80}}>
      <header style={{position:"sticky",top:0,zIndex:50,background:"rgba(15,17,23,0.85)",backdropFilter:"blur(20px)",borderBottom:`1px solid ${T.border}`,padding:"14px 20px",overflow:"hidden"}}>
        <div style={{maxWidth:600,margin:"0 auto",display:"flex",alignItems:"center",gap:10}}>
          <Icon name="coffee" size={22} color={T.accent}/>
          <h1 style={{fontSize:17,fontWeight:600,color:T.text,letterSpacing:"-.01em"}}>Sunoha Checklists</h1>
        </div>
        <LoadingBar visible={true} />
      </header>
      <div style={{maxWidth:600,margin:"0 auto",padding:"20px 16px"}}>
        <SkeletonList />
      </div>
    </div></>
  );

  return (
    <><style>{globalCss}</style>
    <div style={{minHeight:"100vh",background:T.bg,paddingBottom:isTabView?80:20}}>
      {/* ── Header with loading indicator ── */}
      <header style={{position:"sticky",top:0,zIndex:50,background:"rgba(15,17,23,0.85)",backdropFilter:"blur(20px)",borderBottom:`1px solid ${T.border}`,padding:"14px 20px",overflow:"hidden"}}>
        <div style={{maxWidth:600,margin:"0 auto",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            {!isTabView && <button onClick={goBack} style={{background:"none",border:"none",cursor:"pointer",padding:4,display:"flex"}}><Icon name="back" size={20} color={T.textSec}/></button>}
            <Icon name="coffee" size={22} color={T.accent}/>
            <h1 style={{fontSize:17,fontWeight:600,color:T.text,letterSpacing:"-.01em"}}>{titles[currentView]||"Settings"}</h1>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <button onClick={()=>setShowAccount(true)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:T.textSec,padding:0,fontFamily:T.font}} title="My Account">{currentUser.displayName}</button>
            <button onClick={handleLogout} style={{background:"none",border:"none",cursor:"pointer",padding:6,borderRadius:8,display:"flex"}} title="Sign out"><Icon name="logOut" size={18} color={T.textSec}/></button>
          </div>
        </div>
        <LoadingBar visible={apiLoading} />
      </header>

      {/* ── Content ── */}
      <div style={{maxWidth:600,margin:"0 auto",padding:"20px 16px"}}>
        {busy && <div style={{background:T.accentBg,border:`1px solid ${T.accentBorder}`,borderRadius:T.radSm,padding:"12px 16px",marginBottom:16,display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:13,color:T.accent,animation:"pulse 1.5s infinite"}}>{busy}</span>
        </div>}
        {error && <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm,padding:"12px 16px",marginBottom:16,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontSize:13,color:T.danger}}>{error}</span>
          <Btn variant="ghost" small onClick={()=>{setError(null);loadAll()}}><Icon name="refresh" size={14} color={T.textSec}/> Retry</Btn>
        </div>}

        {currentView==="orders" && <OrdersView orders={orders} checklists={checklists} orderTypes={orderTypes} customers={customers} isAdmin={isAdmin} busy={busy}
          untaggedChecklists={untaggedChecklists} currentUser={currentUser}
          inventoryItems={inventoryItems}
          drafts={drafts} addToast={addToast}
          onEditUntaggedResponse={(ut)=>{setSelected({orderChecklistId:ut.id,checklistId:ut.checklistId,orderId:null,isUntagged:true}); setSubView("editResponses")}}
          onResumeDraft={(d)=>{setResumeDraft(d);setSubView("quickFill")}}
          onDeleteDraft={async(id)=>{
            const old=drafts; setDrafts(prev=>prev.filter(d=>d.id!==id));
            try{await API.post("deleteDraft",{id});addToast("Draft deleted","success")}catch(e){setDrafts(old);addToast(e.message,"error")}
          }}
          onTagAllocation={async(payload)=>{
            try{await API.post("createAllocation",payload);await refreshDrafts();await loadAll();addToast("Tagged successfully","success")}catch(e){addToast(e.message,"error");throw e}
          }}
          onSelect={o=>{setSelected(o);setSubView("orderDetail")}} onNew={()=>setSubView("newOrder")}
          onQuickFill={()=>{setResumeDraft(null);setSubView("quickFill")}}
          onTagUntagged={async(utId,orderId,tagQuantity)=>{
            try{const result=await API.post("tagUntagged",{id:utId,orderId,tagQuantity});
              if(result.fullyTagged){setUntaggedChecklists(prev=>prev.filter(u=>u.id!==utId))}
              else{setUntaggedChecklists(prev=>prev.map(u=>u.id===utId?{...u,taggedQuantity:(u.taggedQuantity||0)+(tagQuantity||u.totalQuantity||0),remainingQuantity:Math.max(0,(u.totalQuantity||0)-(u.taggedQuantity||0)-(tagQuantity||u.totalQuantity||0))}:u))}
              await refreshOrders();addToast("Checklist tagged to order","success")}catch(e){addToast(e.message,"error")}
          }}
          onDeleteOrder={async(id)=>{
            // Optimistic delete
            setOrders(prev=>prev.filter(o=>o.id!==id));
            setBusy("Deleting order...");
            try{await API.post("deleteOrder",{id})}catch(e){addToast(e.message,"error",()=>refreshOrders());await refreshOrders()}finally{setBusy(null)}
          }}/>}

        {currentView==="newOrder" && <NewOrderView orderTypes={orderTypes} customers={customers} checklists={checklists} currentUser={currentUser} blends={blends} orderStageTemplates={orderStageTemplates}
          onAddCustomer={async(name)=>{
            const optimisticId = "cust_"+Date.now();
            const optimisticCust = {id:optimisticId,label:name};
            setCustomers(prev=>[...prev,optimisticCust]);
            try{const c=await API.post("createCustomer",{id:optimisticId,label:name});return c}catch(e){setCustomers(prev=>prev.filter(c=>c.id!==optimisticId));addToast(e.message,"error");throw e}
          }}
          onCreate={async(order)=>{
            try{
              const created=await API.post("createOrder",order);
              // Optimistic: add to orders list immediately
              setOrders(prev=>[created,...prev]);
              setSubView(null);
            }catch(e){addToast(e.message,"error")}
          }}/>}

        {currentView==="quickFill" && <QuickFillView checklists={checklists} orders={orders} customers={customers} currentUser={currentUser} approvedEntries={approvedEntries} inventoryItems={inventoryItems} inventoryCategories={inventoryCategories}
          resumeDraft={resumeDraft} addToast={addToast}
          onSaveDraft={async(payload)=>{
            try{
              const r=await API.post("saveDraft",payload);
              setDrafts(prev=>{const without=prev.filter(d=>d.id!==r.id);return [r,...without];});
              addToast("Draft saved","success");
              return r;
            }catch(e){addToast(e.message,"error");throw e}
          }}
          onSubmit={async(data)=>{
            try{
              const result=await API.post("submitUntagged",data);
              if(!data.orderId){setUntaggedChecklists(prev=>[result,...prev])}else{await refreshOrders()}
              // If we resumed from a draft, delete it now
              if(resumeDraft?.id){try{await API.post("deleteDraft",{id:resumeDraft.id});setDrafts(prev=>prev.filter(d=>d.id!==resumeDraft.id))}catch{}}
              // Refresh approved entries cache after submission
              try{const fresh=await API.get("init");if(fresh.approvedEntries)setApprovedEntries(fresh.approvedEntries)}catch{}
              addToast("Checklist submitted","success");setResumeDraft(null);setSubView(null);
            }catch(e){addToast(e.message,"error")}
          }}/>}

        {currentView==="orderDetail" && selected && <OrderDetailView order={selected} checklists={checklists} customers={customers} isAdmin={isAdmin} currentUser={currentUser} approvedEntries={approvedEntries} inventoryItems={inventoryItems} inventoryCategories={inventoryCategories} untaggedChecklists={untaggedChecklists} blends={blends}
          onUpdate={updated=>{setOrders(prev=>prev.map(o=>o.id===updated.id?updated:o));setSelected(updated)}}
          onEditOrder={async(data)=>{
            // Optimistic update
            const optimistic={...selected,name:data.name||selected.name,customerId:data.customerId||selected.customerId,assignedTo:data.assignedTo!==undefined?data.assignedTo:selected.assignedTo,invoiceSo:data.invoiceSo!==undefined?data.invoiceSo:selected.invoiceSo,orderTypeDetail:data.orderTypeDetail!==undefined?data.orderTypeDetail:selected.orderTypeDetail,productType:data.productType!==undefined?data.productType:selected.productType,missingChecklistReasons:data.missingChecklistReasons!==undefined?data.missingChecklistReasons:selected.missingChecklistReasons};
            setSelected(optimistic);setOrders(prev=>prev.map(o=>o.id===optimistic.id?optimistic:o));
            try{await API.post("editOrder",data)}catch(e){addToast(e.message,"error");const fresh=await API.get("getOrder",{id:data.id});setSelected(fresh);setOrders(prev=>prev.map(o=>o.id===fresh.id?fresh:o))}
          }}
          onDeleteOrder={async(id)=>{
            setOrders(prev=>prev.filter(o=>o.id!==id));
            setBusy("Deleting order...");
            try{await API.post("deleteOrder",{id});setSubView(null);setSelected(null)}catch(e){addToast(e.message,"error");await refreshOrders()}finally{setBusy(null)}
          }}
          onRevertChecklist={async(ocId)=>{
            setBusy("Reverting checklist...");
            try{await API.post("revertChecklist",{id:ocId});const updated=await API.get("getOrder",{id:selected.id});setSelected(updated);setOrders(prev=>prev.map(o=>o.id===updated.id?updated:o))}catch(e){addToast(e.message,"error")}finally{setBusy(null)}
          }}
          onUpdateStatus={async(orderId,status)=>{
            try{const updated=await API.post("updateOrderStatus",{id:orderId,status});setSelected(updated);setOrders(prev=>prev.map(o=>o.id===updated.id?updated:o));addToast("Status updated","success")}catch(e){addToast(e.message,"error")}
          }}
          onTagStage={async(orderId,stageId,autoId,sourceCkId,qty,quantityFieldValue,blendExtras)=>{
            try{
              const body={orderId,stageId,responseId:autoId,sourceChecklistId:sourceCkId,quantity:qty,quantityFieldValue:quantityFieldValue||""};
              if(blendExtras && typeof blendExtras === "object"){
                if(blendExtras.blendLineIndex !== undefined) body.blendLineIndex = blendExtras.blendLineIndex;
                if(blendExtras.componentItemId !== undefined) body.componentItemId = blendExtras.componentItemId;
                if(blendExtras.componentItemName !== undefined) body.componentItemName = blendExtras.componentItemName;
              }
              const r=await API.post("tagChecklistToStage",body);
              if(r.stages){const updated={...selected,stages:r.stages};setSelected(updated);setOrders(prev=>prev.map(o=>o.id===updated.id?updated:o));}
              try{const fresh=await API.get("init");if(fresh.approvedEntries)setApprovedEntries(fresh.approvedEntries);if(fresh.untaggedChecklists)setUntaggedChecklists(fresh.untaggedChecklists);}catch{}
              addToast("Checklist tagged to stage","success");
            }catch(e){addToast(e.message,"error");throw e}
          }}
          onTagMixedStage={async(orderId,payload)=>{
            try{
              const body={orderId,stageId:payload.stageId,isMixedBlend:true,blendLineIndex:payload.blendLineIndex,mixedInventoryItemId:payload.mixedInventoryItemId,mixedInventoryItemName:payload.mixedInventoryItemName,mixedBlendId:payload.mixedBlendId,quantity:payload.quantity};
              const r=await API.post("tagChecklistToStage",body);
              if(r.stages){const updated={...selected,stages:r.stages};setSelected(updated);setOrders(prev=>prev.map(o=>o.id===updated.id?updated:o));}
              addToast("Mixed blend tagged to stage","success");
            }catch(e){addToast(e.message,"error");throw e}
          }}
          onUntagStage={async(orderId,stageId,autoId,spec)=>{
            try{
              const body={orderId,stageId};
              if(spec && typeof spec === "object"){
                if(spec.isMixedBlend){
                  body.isMixedBlend=true;
                  if(spec.mixedItemId) body.mixedItemId=spec.mixedItemId;
                } else if(spec.responseId){
                  body.responseId=spec.responseId;
                }
                if(spec.blendLineIndex !== undefined) body.blendLineIndex=spec.blendLineIndex;
                if(spec.componentItemId !== undefined) body.componentItemId=spec.componentItemId;
              } else {
                body.responseId=autoId;
              }
              const r=await API.post("untagChecklistFromStage",body);
              if(r.stages){const updated={...selected,stages:r.stages};setSelected(updated);setOrders(prev=>prev.map(o=>o.id===updated.id?updated:o));}
              try{const fresh=await API.get("init");if(fresh.approvedEntries)setApprovedEntries(fresh.approvedEntries);if(fresh.untaggedChecklists)setUntaggedChecklists(fresh.untaggedChecklists);}catch{}
              addToast("Untagged from stage","success");
            }catch(e){addToast(e.message,"error");throw e}
          }}
          onDeliver={async(orderId,confirmed)=>{
            try{
              const r=await API.post("deliverOrder",{id:orderId,confirmed});
              if(r.preview) return r;
              if(r.order){setSelected(r.order);setOrders(prev=>prev.map(o=>o.id===r.order.id?r.order:o));}
              try{const s=await API.get("getInventorySummary");setInventorySummary(s)}catch{}
              addToast("Order delivered — inventory updated","success");
              return r;
            }catch(e){addToast(e.message,"error");throw e}
          }}
          onEditResponses={(ocId,ckId)=>{setDetailOrder(selected);setSelected({orderChecklistId:ocId,checklistId:ckId,orderId:selected.id});setSubView("editResponses")}}/>}

        {currentView==="editResponses" && selected && <EditResponseView orderChecklistId={selected.orderChecklistId} checklistId={selected.checklistId} checklists={checklists} approvedEntries={approvedEntries} inventoryItems={inventoryItems} customers={customers} isUntagged={!!selected.isUntagged}
          onSave={async(data)=>{try{
            const action = data.isUntagged ? "editUntaggedResponse" : "editResponse";
            const r=await API.post(action, data);
            addToast("Responses saved","success");
            if(r && r.warning) addToast("Saved. Note: "+r.warning+" — linked data may be affected.","info");
            // Refresh caches so dashboard / linked lists reflect edits
            try{const fresh=await API.get("init");
              if(fresh.approvedEntries)setApprovedEntries(fresh.approvedEntries);
              if(fresh.untaggedChecklists)setUntaggedChecklists(fresh.untaggedChecklists);
              if(fresh.inventorySummary)setInventorySummary(fresh.inventorySummary);
            }catch{}
            goBack();
          }catch(e){addToast(e.message,"error")}}}
          onCancel={goBack}/>}

        {currentView==="responses" && <ResponsesLogView checklists={checklists} inventoryItems={inventoryItems} isAdmin={isAdmin} addToast={addToast}
          onEditResponses={(ocId,ckName)=>{
            const ck=checklists.find(c=>c.name===ckName);
            setSelected({orderChecklistId:ocId,checklistId:ck?.id,orderId:null});setSubView("editResponses");
          }}
          onRevertChecklist={async(ocId)=>{try{await API.post("revertChecklist",{id:ocId});await refreshOrders()}catch(e){addToast(e.message,"error")}}}/>}

        {currentView==="admin" && <AdminView checklists={checklists} orderTypes={orderTypes} customers={customers} rules={rules} isAdmin={isAdmin} addToast={addToast} orderStageTemplates={orderStageTemplates}
          onEditChecklist={ck=>{setSelected(ck);setSubView("editChecklist")}}
          onNewChecklist={()=>{setSelected(null);setSubView("editChecklist")}}
          onEditRules={()=>setSubView("editRules")}
          onDeleteChecklist={async(id)=>{
            setChecklists(prev=>prev.filter(c=>c.id!==id));
            try{await API.post("deleteChecklist",{id});await refreshRules()}catch(e){addToast(e.message,"error");await refreshChecklists()}
          }}
          onAddOrderType={async(label)=>{
            const optimistic={id:"ot_"+Date.now(),label};
            setOrderTypes(prev=>[...prev,optimistic]);
            try{await API.post("createOrderType",optimistic)}catch(e){setOrderTypes(prev=>prev.filter(t=>t.id!==optimistic.id));addToast(e.message,"error")}
          }}
          onDeleteOrderType={async(id)=>{
            setOrderTypes(prev=>prev.filter(t=>t.id!==id));
            try{await API.post("deleteOrderType",{id});await refreshRules()}catch(e){addToast(e.message,"error");await refreshOrderTypes()}
          }}
          onAddCustomer={async(label)=>{
            const optimistic={id:"cust_"+Date.now(),label};
            setCustomers(prev=>[...prev,optimistic]);
            try{await API.post("createCustomer",optimistic)}catch(e){setCustomers(prev=>prev.filter(c=>c.id!==optimistic.id));addToast(e.message,"error")}
          }}
          onDeleteCustomer={async(id)=>{
            setCustomers(prev=>prev.filter(c=>c.id!==id));
            try{await API.post("deleteCustomer",{id});await refreshRules()}catch(e){addToast(e.message,"error");await refreshCustomers()}
          }}
          onArchive={async(days)=>{try{const r=await API.post("archiveOrders",{daysOld:parseInt(days)||30});await refreshOrders();return r}catch(e){addToast(e.message,"error");throw e}}}
          onSaveOrderStageTemplates={async(tpl)=>{const r=await API.post("saveOrderStageTemplates",{templates:tpl});if(r.templates)setOrderStageTemplates(r.templates);return r}}/>}

        {currentView==="editChecklist" && <EditChecklistView checklist={selected} allChecklists={checklists} inventoryItems={inventoryItems} inventoryCategories={inventoryCategories} onSave={async(ck)=>{
          try{
            let saved;
            if(selected) { saved=await API.post("updateChecklist",{id:ck.id,...ck}); setChecklists(prev=>prev.map(c=>c.id===saved.id?saved:c)); }
            else { saved=await API.post("createChecklist",ck); setChecklists(prev=>[...prev,saved]); }
            if(saved && saved.warning) addToast("Saved. Note: "+saved.warning+" — linked data may be affected.","info");
            setSubView(null);
          }catch(e){addToast(e.message,"error")}
        }}/>}

        {currentView==="editRules" && <RulesView rules={rules} orderTypes={orderTypes} customers={customers} checklists={checklists}
          onAddRule={()=>{setSelected(null);setSubView("addRule")}}
          onEditRule={r=>{setSelected(r);setSubView("editRule")}}
          onDeleteRule={async(id)=>{
            setRules(prev=>prev.filter(r=>r.id!==id));
            try{await API.post("deleteRule",{id})}catch(e){addToast(e.message,"error");await refreshRules()}
          }}/>}

        {(currentView==="addRule"||currentView==="editRule") && <EditRuleView rule={currentView==="editRule"?selected:null} orderTypes={orderTypes} customers={customers} checklists={checklists} onSave={async(rule)=>{
          try{
            let saved;
            if(currentView==="editRule") { saved=await API.post("updateRule",{id:rule.id,...rule}); setRules(prev=>prev.map(r=>r.id===saved.id?saved:r)); }
            else { saved=await API.post("createRule",rule); setRules(prev=>[...prev,saved]); }
            setSubView("editRules"); setSelected(null);
          }catch(e){addToast(e.message,"error")}
        }}/>}

        {currentView==="inventory" && <InventoryView items={inventoryItems} categories={inventoryCategories} summary={inventorySummary} isAdmin={isAdmin} addToast={addToast}
          onViewLedger={item=>{setSelected(item);setSubView("inventoryLedger")}}
          onCreateItem={async(item)=>{
            try{const r=await API.post("createInventoryItem",item);setInventoryItems(prev=>[...prev,r]);
              try{const s=await API.get("getInventorySummary");setInventorySummary(s)}catch{}
              addToast("Item created","success")}catch(e){addToast(e.message,"error")}
          }}
          onUpdateItem={async(item)=>{
            try{const result=await API.post("updateInventoryItem",item);if(result&&result.allItems){setInventoryItems(result.allItems)}else{setInventoryItems(prev=>prev.map(i=>i.id===item.id?(result&&result.id?result:{...i,...item}):i))}addToast("Item updated","success")}catch(e){addToast(e.message,"error")}
          }}
          onCreateCategory={async(name)=>{
            try{const r=await API.post("createInventoryCategory",{name});setInventoryCategories(prev=>[...prev,r]);addToast("Category added","success")}catch(e){addToast(e.message,"error")}
          }}
        />}

        {currentView==="inventoryLedger" && selected && <InventoryLedgerView item={selected} isAdmin={isAdmin} addToast={addToast}
          onAdjust={async(data)=>{
            try{const r=await API.post("addInventoryAdjustment",data);
              setInventoryItems(prev=>prev.map(i=>i.id===data.itemId?{...i,currentStock:r.newStock}:i));
              try{const s=await API.get("getInventorySummary");setInventorySummary(s)}catch{}
              addToast("Adjustment saved","success");return r}catch(e){addToast(e.message,"error");throw e}
          }}
        />}

        {currentView==="users" && <UsersView addToast={addToast}/>}

        {currentView==="blends" && <BlendsPage blends={blends} customers={customers} isAdmin={isAdmin} inventoryItems={inventoryItems} addToast={addToast}
          onCreate={()=>{setSelected(null);setSubView("editBlend")}}
          onEdit={b=>{setSelected(b);setSubView("editBlend")}}
          onDelete={async(id)=>{
            const old=blends;
            setBlends(prev=>prev.filter(b=>b.id!==id));
            try{await API.post("deleteBlend",{id});addToast("Blend deleted","success")}catch(e){setBlends(old);addToast(e.message,"error")}
          }}
          onImport={async(selections)=>{
            let created=0, updated=0, skipped=0;
            const errors=[];
            for(const it of selections){
              if(!it.accept){skipped++;continue;}
              const payload={
                name: it.blend.name,
                customer: it.blend.customer,
                description: it.blend.description,
                components: it.blend.components,
              };
              try{
                if(it.isNew){
                  const r=await API.post("createBlend",payload);
                  setBlends(prev=>[...prev,r]);
                  created++;
                } else {
                  // Read-merge-write: reuse existing id + preserve isActive if the import didn't explicitly set one
                  const existing=it.existing;
                  const body={id:existing.id,...payload,isActive:it.blend.isActive!==undefined?it.blend.isActive:existing.isActive};
                  const r=await API.post("updateBlend",body);
                  setBlends(prev=>prev.map(x=>x.id===r.id?r:x));
                  updated++;
                }
              } catch(err){
                errors.push(`${it.blend.name}: ${err.message}`);
                skipped++;
              }
            }
            if(errors.length>0) addToast("Some imports failed:\n"+errors.join("\n"),"error");
            return {created, updated, skipped};
          }}/>}

        {currentView==="editBlend" && <CreateEditBlendForm blend={selected} customers={customers} inventoryItems={inventoryItems} inventoryCategories={inventoryCategories}
          onSave={async(b)=>{
            try{
              if(selected){
                const r=await API.post("updateBlend",{id:selected.id,...b});
                setBlends(prev=>prev.map(x=>x.id===r.id?r:x));
                addToast("Blend updated","success");
              } else {
                const r=await API.post("createBlend",b);
                setBlends(prev=>[...prev,r]);
                addToast("Blend created","success");
              }
              setSubView(null);setSelected(null);
            }catch(e){addToast(e.message,"error")}
          }}/>}
      </div>

      {/* ── Bottom Nav ── */}
      {isTabView && (
        <nav style={{position:"fixed",bottom:0,left:0,right:0,background:"rgba(15,17,23,0.92)",backdropFilter:"blur(20px)",borderTop:`1px solid ${T.border}`,padding:"8px 0",paddingBottom:"max(8px,env(safe-area-inset-bottom))"}}>
          <div style={{maxWidth:600,margin:"0 auto",display:"flex"}}>
            {[{id:"orders",icon:"package",label:"Orders"},{id:"responses",icon:"clipboard",label:"Responses"},{id:"inventory",icon:"layers",label:"Inventory"},{id:"blends",icon:"coffee",label:"Blends"},
              ...(isAdmin?[{id:"admin",icon:"settings",label:"Settings"},{id:"users",icon:"users",label:"Users"}]:[])
            ].map(tab=>(
              <button key={tab.id} onClick={()=>switchTab(tab.id)} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:3,background:"none",border:"none",cursor:"pointer",padding:"6px 0",color:view===tab.id?T.accent:T.textMut,transition:"color .2s"}}>
                <Icon name={tab.icon} size={20} color={view===tab.id?T.accent:T.textMut}/><span style={{fontSize:11,fontWeight:500}}>{tab.label}</span>
              </button>
            ))}
          </div>
        </nav>
      )}

      {/* ── Toasts ── */}
      <ToastContainer toasts={toasts} onDismiss={dismissToast} onRetry={handleRetryToast} />

      {showAccount && <MyAccountModal user={currentUser} onClose={()=>setShowAccount(false)} addToast={addToast}/>}
    </div></>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Orders View ──────────────────────────────────────────────

// TagPicker — tabbed picker for tagging an untagged submission to either an existing checklist or an invoice/order.
// "Tag to Checklist" mode uses the generic createAllocation endpoint via onTagToChecklist.
// "Tag to Invoice" mode delegates to the existing onTagToOrder flow (handleTagUntagged).
function TagPicker({ ut, checklists, orders, onTagToOrder, onTagToChecklist, onCancel }) {
  const [mode, setMode] = useState("checklist");
  const [pickedCkId, setPickedCkId] = useState("");
  const [search, setSearch] = useState("");
  const [qty, setQty] = useState("");
  const [busy, setBusy] = useState(false);

  // Restrict destinations based on the source's canTagTo whitelist (if set)
  const sourceCk = checklists.find(c => c.id === ut.checklistId);
  const allowList = Array.isArray(sourceCk?.canTagTo) ? sourceCk.canTagTo : [];
  const allowAll = allowList.length === 0;
  const canTagOrder = allowAll || allowList.indexOf("order") >= 0;
  const candidateChecklists = (checklists || []).filter(c => c.id !== ut.checklistId && (allowAll || allowList.indexOf(c.id) >= 0));

  // Pull approved/submitted entries of the picked checklist type to choose a specific destination submission
  const [destEntries, setDestEntries] = useState([]);
  const [loadingEntries, setLoadingEntries] = useState(false);
  useEffect(() => {
    if (!pickedCkId) { setDestEntries([]); return; }
    let cancelled = false;
    setLoadingEntries(true);
    API.get("getLinkedEntries", { checklist_id: pickedCkId }).then(arr => {
      if (cancelled) return;
      setDestEntries(Array.isArray(arr) ? arr : []);
    }).catch(() => {}).finally(() => { if (!cancelled) setLoadingEntries(false); });
    return () => { cancelled = true; };
  }, [pickedCkId]);

  const filteredEntries = destEntries.filter(e => {
    const s = search.toLowerCase();
    return (e.autoId || e.linkedId || "").toLowerCase().includes(s);
  });
  const filteredOrders = (orders || []).filter(o => {
    if (o.canTag === false || o.status === "delivered" || o.status === "cancelled") return false;
    const s = search.toLowerCase();
    return o.id.toLowerCase().includes(s) || (o.name || "").toLowerCase().includes(s);
  });

  return (
    <div style={{display:"flex",flexDirection:"column",gap:8,minWidth:260,background:T.bg,padding:10,borderRadius:T.radSm,border:`1px solid ${T.border}`}} onClick={e=>e.stopPropagation()}>
      <div style={{display:"flex",gap:6}}>
        <Chip label="To Checklist" active={mode==="checklist"} onClick={()=>setMode("checklist")}/>
        {canTagOrder && <Chip label="To Invoice" active={mode==="order"} onClick={()=>setMode("order")}/>}
      </div>

      {ut.totalQuantity > 0 && (
        <input value={qty} onChange={e=>{ const v=e.target.value; const n=parseFloat(v); if(!isNaN(n)&&n<0){ setQty("0"); return; } setQty(v); }} onBlur={e=>{ const n=parseFloat(e.target.value); if(!isNaN(n)&&n<0) setQty("0"); }} placeholder={`Quantity (max ${ut.remainingQuantity || ut.totalQuantity})`} type="number" min="0"
          style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
      )}

      {mode === "checklist" && (
        <>
          <select value={pickedCkId} onChange={e=>{setPickedCkId(e.target.value);setSearch("")}}
            style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13}}>
            <option value="">— Pick checklist type —</option>
            {candidateChecklists.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
          {pickedCkId && (
            <>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search by Auto ID..."
                style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
              <div style={{maxHeight:160,overflowY:"auto",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                {loadingEntries ? <p style={{padding:8,fontSize:12,color:T.textMut}}>Loading...</p> :
                 filteredEntries.length === 0 ? <p style={{padding:8,fontSize:12,color:T.textMut}}>No entries found</p> :
                 filteredEntries.map((e, i) => {
                   const id = e.autoId || e.linkedId;
                   return (
                     <button key={i} disabled={busy} onClick={async()=>{
                       const amount = parseFloat(qty) || 0;
                       if (ut.totalQuantity > 0 && amount <= 0) { alert("Enter a quantity"); return; }
                       setBusy(true);
                       await onTagToChecklist({
                         destinationType: "checklist",
                         destinationId: id,
                         destinationAutoId: id,
                         quantity: amount > 0 ? amount : (ut.remainingQuantity || ut.totalQuantity || 0),
                       });
                       setBusy(false);
                     }}
                     style={{display:"block",width:"100%",padding:"8px 10px",background:"none",border:"none",borderBottom:`1px solid ${T.border}`,color:T.text,fontSize:13,cursor:busy?"wait":"pointer",textAlign:"left"}}>
                       <div style={{fontFamily:T.mono}}>{id}</div>
                       {e.remainingQuantity !== undefined && <div style={{fontSize:11,color:T.textMut}}>{e.remainingQuantity} available</div>}
                     </button>
                   );
                 })
                }
              </div>
            </>
          )}
        </>
      )}

      {mode === "order" && (
        <>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search orders..."
            style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
          <div style={{maxHeight:160,overflowY:"auto",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
            {filteredOrders.length === 0 ? <p style={{padding:8,fontSize:12,color:T.textMut}}>No orders found</p> :
              filteredOrders.map(o => (
                <button key={o.id} onClick={()=>onTagToOrder(o.id, parseFloat(qty) || 0)}
                  style={{display:"block",width:"100%",padding:"8px 10px",background:"none",border:"none",borderBottom:`1px solid ${T.border}`,color:T.text,fontSize:13,cursor:"pointer",textAlign:"left"}}>
                  {o.id} — {o.name}
                </button>
              ))
            }
          </div>
        </>
      )}

      <Btn variant="ghost" small onClick={onCancel}>Cancel</Btn>
    </div>
  );
}

function OrdersView({orders,checklists,orderTypes,customers,isAdmin,busy,untaggedChecklists,currentUser,drafts,inventoryItems,onResumeDraft,onDeleteDraft,onTagAllocation,onSelect,onNew,onQuickFill,onTagUntagged,onDeleteOrder,onEditUntaggedResponse,addToast}){
  const active=orders.filter(o=>o.status!=="delivered");
  const delivered=orders.filter(o=>o.status==="delivered");
  const [tagDropdown,setTagDropdown]=useState(null); // utId
  const [tagSearch,setTagSearch]=useState("");
  const [tagQty,setTagQty]=useState("");
  const [expandedGroup,setExpandedGroup]=useState(null);
  const [previewUt,setPreviewUt]=useState(null);
  const [showDelivered,setShowDelivered]=useState(false);
  const [deleteTarget,setDeleteTarget]=useState(null); // {id,entityType,label}
  const filteredOrders=orders.filter(o=>{const s=tagSearch.toLowerCase();return o.id.toLowerCase().includes(s)||o.name.toLowerCase().includes(s)});

  // Group untagged by checklist name
  const untaggedGroups = {};
  (untaggedChecklists||[]).forEach(ut => {
    if (!untaggedGroups[ut.checklistName]) untaggedGroups[ut.checklistName] = [];
    untaggedGroups[ut.checklistName].push(ut);
  });
  const untaggedGroupNames = Object.keys(untaggedGroups);

  return (
    <div className="fade-up">
      <div style={{display:"flex",gap:10,marginBottom:24}}>
        <Btn onClick={onNew} style={{flex:1}}><Icon name="plus" size={18} color={T.bg}/> Create New Order</Btn>
        <Btn variant="secondary" onClick={onQuickFill} style={{flex:1}}><Icon name="clipboard" size={18} color={T.text}/> Fill Checklist</Btn>
      </div>

      {/* ── Untagged Checklists (Grouped) ── */}
      {untaggedGroupNames.length>0 && <>
        <Section icon="clipboard" count={(untaggedChecklists||[]).length}>Untagged Checklists</Section>
        <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:32}}>
          {untaggedGroupNames.map(groupName=>{
            const items=untaggedGroups[groupName];
            const isExpanded=expandedGroup===groupName;
            return <div key={groupName} style={{background:T.card,borderRadius:T.rad,border:`1px solid ${T.warningBorder}`,overflow:"hidden"}}>
              <div onClick={()=>setExpandedGroup(isExpanded?null:groupName)} style={{padding:"14px 16px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <span style={{fontSize:14,fontWeight:600,color:T.text}}>{groupName}</span>
                  <Badge variant="muted">{items.length}</Badge>
                </div>
                <Icon name="chevron" size={16} color={T.textMut} style={{transform:isExpanded?"rotate(90deg)":"rotate(0)",transition:"transform .2s"}}/>
              </div>
              {isExpanded&&<div style={{borderTop:`1px solid ${T.border}`,padding:"0 16px 14px"}}>
                {items.map(ut=>(
                  <div key={ut.id} style={{padding:"12px 0",borderBottom:`1px solid ${T.border}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",gap:8,marginBottom:4,flexWrap:"wrap",alignItems:"center"}}>
                          {ut.autoId&&<button onClick={()=>setPreviewUt(ut)} style={{background:"none",border:`1px solid ${T.accentBorder}`,padding:"2px 8px",borderRadius:12,cursor:"pointer",fontSize:12,fontFamily:T.mono,color:T.accent,fontWeight:600}} title="View responses">{ut.autoId}</button>}
                          {ut.taggedOrderId&&ut.taggedOrderId!=="UNTAGGED"&&<OrderBadge orderId={ut.taggedOrderId} orders={orders} orderTypes={orderTypes} customers={customers}/>}
                          {ut.person&&<span style={{fontSize:12,color:T.textSec}}>by {ut.person}</span>}
                          {ut.date&&<span style={{fontSize:12,color:T.textMut}}>{ut.date}</span>}
                        </div>
                        {ut.totalQuantity>0&&<>
                          <div style={{display:"flex",gap:8,fontSize:11,marginBottom:4}}>
                            <span style={{color:T.textSec}}>Total: {ut.totalQuantity}</span>
                            {ut.taggedQuantity>0&&<span style={{color:T.warning}}>Tagged: {ut.taggedQuantity}</span>}
                            <span style={{color:ut.remainingQuantity>0?T.success:T.danger}}>Remaining: {ut.remainingQuantity||ut.totalQuantity}</span>
                          </div>
                          <div style={{height:4,borderRadius:2,background:T.surfaceHover,overflow:"hidden",maxWidth:240}}>
                            <div style={{width:`${Math.min(100,(ut.taggedQuantity/ut.totalQuantity)*100)}%`,height:"100%",borderRadius:2,background:ut.taggedQuantity>=ut.totalQuantity?T.success:T.warning,transition:"width .5s ease"}}/>
                          </div>
                        </>}
                      </div>
                      {(isAdmin||ut.submittedByUserId===currentUser?.id)&&(
                        <div style={{display:"flex",gap:4,alignItems:"flex-start",flexShrink:0}}>
                        {tagDropdown===ut.id?
                          <TagPicker
                            ut={ut}
                            checklists={checklists}
                            orders={orders.filter(o=>o.canTag!==false&&o.status!=="delivered"&&o.status!=="cancelled")}
                            onTagToOrder={(orderId,qty)=>{onTagUntagged(ut.id,orderId,qty);setTagDropdown(null)}}
                            onTagToChecklist={async(payload)=>{
                              try{
                                if(typeof onTagAllocation==="function") await onTagAllocation({sourceAutoId:ut.autoId,sourceChecklistId:ut.checklistId,...payload});
                                setTagDropdown(null);
                              }catch{}
                            }}
                            onCancel={()=>setTagDropdown(null)}
                          />
                        :<Btn small variant="secondary" onClick={()=>setTagDropdown(ut.id)}>
                          <Icon name="link" size={14} color={T.text}/> Tag
                        </Btn>}
                        {isAdmin && <button onClick={()=>setDeleteTarget({id:ut.id,entityType:"untagged",label:ut.autoId||ut.id})} title="Delete entry" style={{background:"none",border:"none",cursor:"pointer",padding:6,borderRadius:6,marginTop:2}}><Icon name="trash" size={14} color={T.danger}/></button>}
                        </div>
                      )}
                    </div>
                  </div>
                ))}
              </div>}
            </div>;
          })}
        </div>
      </>}

      {Array.isArray(drafts)&&drafts.length>0&&<>
        <Section icon="edit" count={drafts.length}>Drafts</Section>
        <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:32}}>
          {drafts.map(d=>(
            <div key={d.id} style={{background:T.card,borderRadius:T.rad,padding:"12px 14px",border:`1px solid ${T.infoBorder}`,display:"flex",justifyContent:"space-between",alignItems:"center",gap:8}}>
              <div style={{flex:1,minWidth:0}}>
                <span style={{fontSize:14,fontWeight:600,color:T.text}}>{d.checklistName}</span>
                <div style={{display:"flex",gap:8,marginTop:2,flexWrap:"wrap"}}>
                  {d.person&&<span style={{fontSize:11,color:T.textMut}}>by {d.person}</span>}
                  {d.workDate&&<span style={{fontSize:11,color:T.textMut}}>{d.workDate}</span>}
                  {d.updatedAt&&<span style={{fontSize:11,color:T.textMut}}>updated {formatDateTime(d.updatedAt)}</span>}
                </div>
              </div>
              <Btn small variant="secondary" onClick={()=>onResumeDraft(d)}><Icon name="edit" size={13} color={T.text}/> Continue</Btn>
              {(isAdmin||String(d.userId)===String(currentUser?.id))&&<button onClick={()=>{if(confirm("Delete this draft?"))onDeleteDraft(d.id)}} style={{background:"none",border:"none",cursor:"pointer",padding:6}}><Icon name="trash" size={15} color={T.danger}/></button>}
            </div>
          ))}
        </div>
      </>}

      <Section icon="clock" count={active.length}>Active Orders</Section>
      {active.length===0?<Empty icon="package" text="No active orders" sub="Create a new order to get started"/>:
        <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:32}}>{active.map((o,i)=><OrderCard key={o.id} order={o} checklists={checklists} orderTypes={orderTypes} customers={customers} isAdmin={isAdmin} onClick={()=>onSelect(o)} onDelete={()=>{if(confirm("Delete this order and all its checklists?"))onDeleteOrder(o.id)}} delay={i*60}/>)}</div>}
      <button onClick={()=>setShowDelivered(v=>!v)} style={{background:"none",border:"none",cursor:"pointer",padding:0,marginBottom:12,width:"100%",textAlign:"left"}}>
        <Section icon="checkCircle" count={delivered.length} action={<Icon name="chevron" size={16} color={T.textMut} style={{transform:showDelivered?"rotate(90deg)":"rotate(0)",transition:"transform .2s"}}/>}>Delivered Orders</Section>
      </button>
      {showDelivered && (delivered.length===0?<Empty icon="check" text="No delivered orders yet" sub="Orders move here after delivery"/>:
        <div style={{display:"flex",flexDirection:"column",gap:10}}>{delivered.map((o,i)=><OrderCard key={o.id} order={o} checklists={checklists} orderTypes={orderTypes} customers={customers} isAdmin={isAdmin} onClick={()=>onSelect(o)} onDelete={()=>{if(confirm("Delete this order and all its checklists?"))onDeleteOrder(o.id)}} completed delay={i*60}/>)}</div>)}

      {previewUt && <UntaggedPreviewModal ut={previewUt} checklists={checklists} inventoryItems={inventoryItems} isAdmin={isAdmin}
        onEdit={(ut)=>{setPreviewUt(null); if(onEditUntaggedResponse) onEditUntaggedResponse(ut);}}
        onClose={()=>setPreviewUt(null)}/>}
      {deleteTarget && <DeleteConfirmModal entryId={deleteTarget.label||deleteTarget.id} entityType={deleteTarget.entityType}
        onConfirm={(r)=>{setDeleteTarget(null); if(addToast) addToast("Entry deleted"+(r?.reversed?" and "+r.reversed+" inventory entries reversed":""), "success"); window.location.reload();}}
        onCancel={()=>setDeleteTarget(null)}/>}
    </div>
  );
}

function UntaggedPreviewModal({ ut, checklists, inventoryItems = [], isAdmin, onEdit, onClose, classifications }) {
  const ck = checklists.find(c => c.id === ut.checklistId);
  const nq = ck ? normalizeQuestions(ck.questions) : [];
  const respList = Array.isArray(ut.responses) ? ut.responses : [];
  // Build lookup by question text (stable across template reorders), fallback to index
  const respByText = {};
  const respByIdx = {};
  respList.forEach(r => {
    if (r && r.questionText) respByText[r.questionText] = r.response;
    if (r && r.questionIndex !== undefined) respByIdx[r.questionIndex] = r.response;
  });
  const getResp = (q, qi) => respByText[q.text] !== undefined ? respByText[q.text] : respByIdx[qi];
  // Detect multi-batch roast entries: roast_batches JSON stored in "Shipment number used" field
  const shipmentVal = getResp(nq[0] || {}, 0) || "";
  let roastBatchesData = null;
  if (ut.checklistId === "ck_roasted_beans" && typeof shipmentVal === "string" && shipmentVal.startsWith("[")) {
    try { roastBatchesData = JSON.parse(shipmentVal); } catch(e) {}
  }
  // Lazy-load classifications for roast batch display
  const [modalClassifications, setModalClassifications] = useState(classifications || null);
  useEffect(() => {
    if (roastBatchesData && !modalClassifications) {
      API.get("getClassifications").then(d => { if (d && !d.error) setModalClassifications(d); }).catch(() => {});
    }
  }, []);
  const isReadOnly = !!ut.accessControl?.isTaggedToStage;
  return (
    <div onClick={onClose} style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:200,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,padding:20,maxWidth:520,width:"100%",maxHeight:"85vh",overflowY:"auto",border:`1px solid ${T.border}`,boxShadow:"0 12px 40px rgba(0,0,0,0.6)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:16}}>
          <div>
            <h3 style={{fontSize:16,fontWeight:600,color:T.text}}>{ck?.name || ut.checklistName}</h3>
            {ut.autoId && <span style={{fontSize:12,fontFamily:T.mono,color:T.accent}}>{ut.autoId}</span>}
            <div style={{display:"flex",gap:8,marginTop:4,flexWrap:"wrap"}}>
              {ut.person && <span style={{fontSize:11,color:T.textMut}}>by {ut.person}</span>}
              {ut.date && <span style={{fontSize:11,color:T.textMut}}>{ut.date}</span>}
              {ut.lastEditedBy && ut.lastEditedAt && (
                <span style={{fontSize:11,color:T.textMut}}>· Edited by {ut.lastEditedBy} at {formatDateTime(ut.lastEditedAt)}</span>
              )}
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:4}}>
            {isAdmin && (
              isReadOnly ? (
                <span title="Untag from order to edit" style={{padding:6,display:"inline-flex"}}><Icon name="lock" size={16} color={T.textMut}/></span>
              ) : (
                <button onClick={()=>onEdit && onEdit(ut)} style={{background:"none",border:"none",cursor:"pointer",padding:6,display:"flex"}} title="Edit response">
                  <Icon name="edit" size={16} color={T.accent}/>
                </button>
              )
            )}
            <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
          </div>
        </div>
        {ut.totalQuantity > 0 && (
          <div style={{display:"flex",gap:12,marginBottom:12,padding:10,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
            <div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Total</span><span style={{fontSize:14,fontWeight:600,color:T.text}}>{ut.totalQuantity}</span></div>
            <div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Tagged</span><span style={{fontSize:14,fontWeight:600,color:T.warning}}>{ut.taggedQuantity || 0}</span></div>
            <div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Available</span><span style={{fontSize:14,fontWeight:600,color:(ut.remainingQuantity||0)>0?T.success:T.danger}}>{ut.remainingQuantity != null ? ut.remainingQuantity : ut.totalQuantity}</span></div>
          </div>
        )}
        {roastBatchesData && Array.isArray(roastBatchesData) && roastBatchesData.length > 0 && (
          <div style={{marginBottom:12}}>
            <span style={{fontSize:12,fontWeight:600,color:T.accent,marginBottom:6,display:"block"}}>Roast Batches</span>
            <RoastBatchTable batchesJson={roastBatchesData} classifications={modalClassifications}/>
          </div>
        )}
        <div style={{background:T.bg,borderRadius:T.radSm,padding:"10px 12px"}}>
          {nq.length === 0 ? <p style={{fontSize:12,color:T.textMut}}>No template found</p> :
            nq.map((q, qi) => {
              if (roastBatchesData && ["Shipment number used","Quantity input","Quantity output","Loss in weight"].includes(q.text)) return null;
              return (
                <div key={qi} style={{padding:"6px 0",borderBottom:qi<nq.length-1?`1px solid ${T.border}`:"none"}}>
                  <span style={{fontSize:11,color:T.textMut}}>{q.text}</span>
                  <div style={{fontSize:14,color:T.text,fontWeight:500,marginTop:2}}>{displayResponseValue(q, getResp(q, qi), inventoryItems)}</div>
                </div>
              );
            }).filter(Boolean)
          }
        </div>
      </div>
    </div>
  );
}

// ─── Order Preview Modal + Clickable Badge ──────────────────

function OrderBadge({ orderId, orders, orderTypes, customers }) {
  const [showModal, setShowModal] = useState(false);
  if (!orderId || orderId === "UNTAGGED") return null;
  return (
    <>
      <button onClick={(e)=>{e.stopPropagation();setShowModal(true)}} style={{background:"rgba(212,165,116,0.15)",border:"1px solid rgba(212,165,116,0.3)",borderRadius:12,padding:"2px 10px",cursor:"pointer",fontSize:12,fontFamily:T.mono,color:T.accent,fontWeight:600,display:"inline-flex",alignItems:"center",gap:4}} title="View order details">
        <Icon name="package" size={10} color={T.accent}/>{orderId}
      </button>
      {showModal && <OrderPreviewModal orderId={orderId} orders={orders} orderTypes={orderTypes} customers={customers} onClose={()=>setShowModal(false)}/>}
    </>
  );
}

function OrderPreviewModal({ orderId, orders, orderTypes, customers, onClose }) {
  const order = (orders || []).find(o => o.id === orderId);
  const custLabel = order ? (customers || []).find(c => c.id === order.customerId)?.label || "" : "";
  const otLabel = order ? (orderTypes || []).find(t => t.id === order.orderType)?.label || "" : "";
  const statusColors = { beans_not_roasted: T.warning, beans_roasted: T.info, packed: T.accent, completed: T.success, delivered: T.success, cancelled: T.danger };
  const ORDER_STATUS_LABELS_LOCAL = {"beans_not_roasted":"Beans not yet roasted","beans_roasted":"Beans roasted","packed":"Packed","completed":"Ready for delivery","delivered":"Delivered"};
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:300,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,padding:20,maxWidth:480,width:"100%",maxHeight:"85vh",overflowY:"auto",border:`1px solid ${T.border}`,boxShadow:"0 12px 40px rgba(0,0,0,0.6)"}}>
        {!order ? (
          <div style={{textAlign:"center",padding:20}}>
            <p style={{color:T.textMut}}>Order <b style={{color:T.accent}}>{orderId}</b> not found in current data.</p>
          </div>
        ) : (
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:16}}>
              <div>
                <h3 style={{fontSize:18,fontWeight:600,color:T.accent,fontFamily:T.mono}}>{order.id}</h3>
                <p style={{fontSize:14,color:T.text,marginTop:2}}>{order.name}</p>
              </div>
              <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
              <div style={{padding:"8px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                <div style={{fontSize:10,color:T.textMut,textTransform:"uppercase"}}>Client</div>
                <div style={{fontSize:13,color:T.text,fontWeight:500,marginTop:2}}>{custLabel || "—"}</div>
              </div>
              <div style={{padding:"8px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                <div style={{fontSize:10,color:T.textMut,textTransform:"uppercase"}}>Order Type</div>
                <div style={{fontSize:13,color:T.text,fontWeight:500,marginTop:2}}>{otLabel || "—"}</div>
              </div>
              <div style={{padding:"8px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                <div style={{fontSize:10,color:T.textMut,textTransform:"uppercase"}}>Status</div>
                <Badge variant="muted" style={{background:(statusColors[order.status]||T.textMut)+"20",color:statusColors[order.status]||T.textMut,marginTop:4}}>{ORDER_STATUS_LABELS_LOCAL[order.status]||order.status}</Badge>
              </div>
              <div style={{padding:"8px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                <div style={{fontSize:10,color:T.textMut,textTransform:"uppercase"}}>Created</div>
                <div style={{fontSize:13,color:T.text,fontWeight:500,marginTop:2}}>{formatDate(order.createdAt)}</div>
              </div>
            </div>
            {order.invoiceSo && <div style={{padding:"8px 12px",background:T.accentBg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`,marginBottom:12}}>
              <span style={{fontSize:11,color:T.textMut}}>Invoice/SO:</span> <span style={{fontSize:13,fontWeight:600,color:T.accent}}>{order.invoiceSo}</span>
            </div>}
            {Array.isArray(order.orderLines) && order.orderLines.length > 0 && (
              <div style={{marginBottom:12}}>
                <div style={{fontSize:12,fontWeight:600,color:T.textSec,marginBottom:6}}>Blend Lines</div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:T.surfaceHover}}>
                      <th style={{padding:"6px 8px",textAlign:"left",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Blend</th>
                      <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Qty</th>
                      <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Tagged</th>
                      <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Remaining</th>
                    </tr></thead>
                    <tbody>
                      {order.orderLines.map((l,i) => {
                        const tq = parseFloat(l.taggedQuantity) || 0;
                        const qty = parseFloat(l.quantity) || 0;
                        return <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
                          <td style={{padding:"6px 8px",color:T.text}}>{l.blend || "—"}</td>
                          <td style={{padding:"6px 8px",textAlign:"right",color:T.text}}>{qty}</td>
                          <td style={{padding:"6px 8px",textAlign:"right",color:T.warning}}>{tq}</td>
                          <td style={{padding:"6px 8px",textAlign:"right",color:(qty-tq)>0?T.success:T.textMut}}>{Math.round((qty-tq)*100)/100}</td>
                        </tr>;
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            {Array.isArray(order.stages) && order.stages.length > 0 && (
              <div>
                <div style={{fontSize:12,fontWeight:600,color:T.textSec,marginBottom:6}}>Stages</div>
                <div style={{display:"flex",flexDirection:"column",gap:4}}>
                  {order.stages.map((s,i) => {
                    const tagged = Array.isArray(s.taggedEntries) ? s.taggedEntries : [];
                    const done = tagged.length > 0;
                    return <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 8px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                      <Icon name={done?"check":"clock"} size={12} color={done?T.success:T.textMut}/>
                      <span style={{fontSize:12,color:T.text,flex:1}}>{s.name}</span>
                      <span style={{fontSize:11,color:done?T.success:T.textMut}}>{done?`${tagged.length} tagged`:"pending"}</span>
                    </div>;
                  })}
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ─── Searchable Dropdown (for linked entries + tag-to-order) ──

function SearchableDropdown({ options, value, onChange, placeholder, emptyMessage, disabled: fieldDisabled, style: outerStyle }) {
  const [open,setOpen]=useState(false);
  const [search,setSearch]=useState("");
  const ref=useRef(null);
  const filtered=(options||[]).filter(o=>{const s=search.toLowerCase();const label=typeof o==="string"?o:(o.label||"");return label.toLowerCase().includes(s)}).sort((a,b)=>{
    const aD=typeof a==="object"&&a.disabled;const bD=typeof b==="object"&&b.disabled;
    if(aD&&!bD)return 1;if(!aD&&bD)return -1;return 0;
  });
  useEffect(()=>{
    const handler=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false)};
    document.addEventListener("mousedown",handler);return()=>document.removeEventListener("mousedown",handler);
  },[]);
  const selectedOpt=(options||[]).find(o=>typeof o==="object"?o.value===value:o===value);
  const selectedLabel=selectedOpt?(typeof selectedOpt==="string"?selectedOpt:selectedOpt.label):(value||"");
  return (
    <div ref={ref} style={{position:"relative",...(outerStyle||{})}}>
      <button onClick={()=>{if(!fieldDisabled)setOpen(!open)}} style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:selectedLabel?T.text:T.textMut,fontSize:14,textAlign:"left",cursor:fieldDisabled?"not-allowed":"pointer",display:"flex",justifyContent:"space-between",alignItems:"center",opacity:fieldDisabled?0.7:1}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{selectedLabel||placeholder||"Select..."}</span>
        <Icon name="chevron" size={14} color={T.textMut} style={{transform:open?"rotate(90deg)":"rotate(0)",transition:"transform .2s",flexShrink:0}}/>
      </button>
      {open&&!fieldDisabled&&(
        <div style={{position:"absolute",top:"100%",left:0,right:0,zIndex:20,background:T.card,border:`1px solid ${T.border}`,borderRadius:T.radSm,marginTop:4,boxShadow:"0 8px 24px rgba(0,0,0,0.4)",maxHeight:250,display:"flex",flexDirection:"column"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search..." autoFocus
            style={{padding:"10px 12px",background:T.bg,border:"none",borderBottom:`1px solid ${T.border}`,color:T.text,fontSize:14,outline:"none",minHeight:44}}/>
          <div style={{overflowY:"auto",flex:1}}>
            {filtered.length===0?<p style={{padding:12,fontSize:13,color:T.textMut}}>{emptyMessage||"No options found"}</p>:
              filtered.map((o,i)=>{
                const label=typeof o==="string"?o:o.label;const val=typeof o==="string"?o:o.value;
                const isDisabled=typeof o==="object"&&o.disabled;
                const sublabel=typeof o==="object"?o.sublabel:"";
                return <button key={i} disabled={isDisabled} onClick={()=>{if(isDisabled)return;onChange(val);setOpen(false);setSearch("")}}
                  style={{display:"block",width:"100%",padding:"10px 12px",background:val===value?T.accentBg:"none",border:"none",borderBottom:`1px solid ${T.border}`,color:isDisabled?T.textMut:T.text,fontSize:14,cursor:isDisabled?"not-allowed":"pointer",textAlign:"left",minHeight:44,opacity:isDisabled?0.4:1}}>
                  <span>{label}</span>
                  {sublabel&&<span style={{display:"block",fontSize:11,color:isDisabled?T.danger:T.textMut}}>{sublabel}</span>}
                </button>;
              })
            }
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Preview Modal (view full linked entry details) ───────────

function PreviewModal({ entry, checklistName, sourceChecklistId, checklists, onClose }) {
  if (!entry) return null;
  const chainAutoId = entry.autoId || entry.linkedId || "";
  return (
    <div onClick={onClose} style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:200,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,padding:20,maxWidth:500,width:"100%",maxHeight:"80vh",overflowY:"auto",border:`1px solid ${T.border}`,boxShadow:"0 12px 40px rgba(0,0,0,0.6)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div>
            <h3 style={{fontSize:16,fontWeight:600,color:T.text}}>{checklistName}</h3>
            <span style={{fontSize:12,color:T.accent}}>{entry.linkedId}</span>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
        </div>
        <div style={{display:"flex",gap:16,marginBottom:12,flexWrap:"wrap"}}>
          {entry.person&&<div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Person</span><span style={{fontSize:13,color:T.textSec}}>{entry.person}</span></div>}
          {entry.date&&<div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Date</span><span style={{fontSize:13,color:T.textSec}}>{entry.date}</span></div>}
          {entry.orderId&&entry.orderId!=="UNTAGGED"&&<div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Order</span><span style={{fontSize:13,color:T.textSec}}>{entry.orderId}</span></div>}
        </div>
        {entry.totalQuantity!==undefined&&(
          <div style={{display:"flex",gap:12,marginBottom:12,padding:10,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
            <div><span style={{fontSize:11,color:T.textMut}}>Total</span><span style={{display:"block",fontSize:14,fontWeight:600,color:T.text}}>{entry.totalQuantity}</span></div>
            <div><span style={{fontSize:11,color:T.textMut}}>Used</span><span style={{display:"block",fontSize:14,fontWeight:600,color:T.warning}}>{entry.usedQuantity}</span></div>
            <div><span style={{fontSize:11,color:T.textMut}}>Remaining</span><span style={{display:"block",fontSize:14,fontWeight:600,color:entry.remainingQuantity>0?T.success:T.danger}}>{entry.remainingQuantity}</span></div>
          </div>
        )}
        <div style={{background:T.bg,borderRadius:T.radSm,padding:"10px 12px"}}>
          {entry.responses.map((r,i)=>(
            <div key={i} style={{padding:"6px 0",borderBottom:i<entry.responses.length-1?`1px solid ${T.border}`:"none"}}>
              <span style={{fontSize:12,color:T.textMut}}>{r.question}</span>
              <div style={{fontSize:14,color:T.text,fontWeight:500,marginTop:2}}>{r.response||"—"}</div>
              {r.remark&&<span style={{fontSize:11,color:T.warning,background:T.warningBg,padding:"2px 6px",borderRadius:8,marginTop:2,display:"inline-block"}}>Remark: {r.remark}</span>}
            </div>
          ))}
        </div>
        {sourceChecklistId && chainAutoId && <SourceChainDisplay checklistId={sourceChecklistId} autoId={chainAutoId} checklists={checklists}/>}
      </div>
    </div>
  );
}

// ─── My Account Modal (self-service profile + change password) ──

function MyAccountModal({ user, onClose, addToast }) {
  const [mode, setMode] = useState("view"); // "view" | "changePassword"
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState("");

  const handleChangePassword = async () => {
    setError("");
    if (!currentPassword || !newPassword || !confirmPassword) { setError("All fields are required"); return; }
    if (newPassword.length < 6) { setError("New password must be at least 6 characters"); return; }
    if (newPassword !== confirmPassword) { setError("New password and confirmation do not match"); return; }
    setSubmitting(true);
    try {
      await API.post("changePassword", { currentPassword, newPassword });
      addToast?.("Password changed successfully", "success");
      setCurrentPassword(""); setNewPassword(""); setConfirmPassword("");
      setMode("view");
    } catch (e) {
      setError(e.message || "Failed to change password");
    }
    setSubmitting(false);
  };

  return (
    <div onClick={onClose} style={{position:"fixed",top:0,left:0,right:0,bottom:0,zIndex:200,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,padding:20,maxWidth:420,width:"100%",maxHeight:"85vh",overflowY:"auto",border:`1px solid ${T.border}`,boxShadow:"0 12px 40px rgba(0,0,0,0.6)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <h3 style={{fontSize:16,fontWeight:600,color:T.text}}>{mode==="changePassword"?"Change Password":"My Account"}</h3>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
        </div>

        {mode === "view" && (
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <Field label="Display Name">
              <Input value={user?.displayName || ""} onChange={()=>{}} />
            </Field>
            <Field label="Username">
              <Input value={user?.username || ""} onChange={()=>{}} style={{opacity:0.7}}/>
            </Field>
            <Field label="Role">
              <Input value={user?.role || ""} onChange={()=>{}} style={{opacity:0.7,textTransform:"capitalize"}}/>
            </Field>
            <Btn onClick={()=>{setMode("changePassword");setError("")}} style={{width:"100%",marginTop:4}}>
              <Icon name="lock" size={16} color={T.bg}/> Change Password
            </Btn>
          </div>
        )}

        {mode === "changePassword" && (
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            {error && <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm,padding:"10px 14px"}}>
              <span style={{fontSize:13,color:T.danger}}>{error}</span>
            </div>}
            <Field label="Current Password">
              <Input value={currentPassword} onChange={setCurrentPassword} type="password" placeholder="Enter current password"/>
            </Field>
            <Field label="New Password">
              <Input value={newPassword} onChange={setNewPassword} type="password" placeholder="At least 6 characters"/>
            </Field>
            <Field label="Confirm New Password">
              <Input value={confirmPassword} onChange={setConfirmPassword} type="password" placeholder="Re-enter new password"/>
            </Field>
            <div style={{display:"flex",gap:8,marginTop:4}}>
              <Btn variant="secondary" onClick={()=>{setMode("view");setError("");setCurrentPassword("");setNewPassword("");setConfirmPassword("")}} style={{flex:1}}>Cancel</Btn>
              <Btn onClick={handleChangePassword} disabled={submitting||!currentPassword||!newPassword||!confirmPassword} style={{flex:1}}>
                {submitting ? "Saving..." : "Save"}
              </Btn>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Batch Selector (multi-batch tagging with per-batch quantity) ──

// ─── Multi-Batch Roasting Section ────────────────────────────

function RoastBatchSection({ entries, batches, onChange, classifications, onAddClassification, addToast }) {
  const safeEntries = entries || [];
  const rows = Array.isArray(batches) && batches.length > 0
    ? batches
    : [{ sourceAutoId: "", inputQty: "", outputQty: "", reasonForLoss: "", classificationId: "" }];

  const remainingFor = (entry, currentIdx) => {
    const baseRemaining = entry.remainingQuantity ?? ((entry.totalQuantity || entry.masterQuantity || 0) - (entry.usedQuantity || entry.allocatedQuantity || 0));
    const id = entry.autoId || entry.linkedId;
    let usedInForm = 0;
    rows.forEach((r, idx) => { if (idx !== currentIdx && r.sourceAutoId === id) usedInForm += parseFloat(r.inputQty) || 0; });
    return baseRemaining - usedInForm;
  };

  const updateRow = (i, patch) => {
    const next = rows.map((r, idx) => idx === i ? { ...r, ...patch } : r);
    onChange(next);
  };
  const removeRow = (i) => { if (rows.length <= 1) return; onChange(rows.filter((_, idx) => idx !== i)); };
  const addRow = () => {
    if (rows.length >= 6) return;
    onChange([...rows, { sourceAutoId: "", inputQty: "", outputQty: "", reasonForLoss: "", classificationId: "" }]);
  };

  const totalInput = rows.reduce((s, r) => s + (parseFloat(r.inputQty) || 0), 0);
  const totalOutput = rows.reduce((s, r) => s + (parseFloat(r.outputQty) || 0), 0);
  const totalLoss = Math.round((totalInput - totalOutput) * 100) / 100;
  const totalLossPct = totalInput > 0 ? Math.round((totalInput - totalOutput) / totalInput * 1000) / 10 : 0;

  // Inline add classification
  const [addingClass, setAddingClass] = useState(false);
  const [newClassName, setNewClassName] = useState("");
  const [savingClass, setSavingClass] = useState(false);
  const handleAddClass = async () => {
    if (!newClassName.trim()) return;
    setSavingClass(true);
    try {
      await onAddClassification(newClassName.trim());
      setNewClassName(""); setAddingClass(false);
    } catch (e) { if (addToast) addToast(e.message, "error"); }
    setSavingClass(false);
  };

  const roastDegrees = classifications?.roast_degree || [];

  if (safeEntries.length === 0) {
    return <div style={{padding:"12px 14px",borderRadius:T.radSm,background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",color:T.danger,fontSize:13}}>
      No approved Green Bean QC entries found. Please complete and approve a Green Beans Quality Check first.
    </div>;
  }

  return (
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      <div style={{fontSize:13,fontWeight:600,color:T.accent,display:"flex",alignItems:"center",gap:6}}>
        <Icon name="layers" size={14} color={T.accent}/> Roast Batches
      </div>
      {rows.map((row, i) => {
        const selected = safeEntries.find(e => (e.autoId || e.linkedId) === row.sourceAutoId);
        const maxQty = selected ? remainingFor(selected, i) : 0;
        const inputQty = parseFloat(row.inputQty) || 0;
        const outputQty = parseFloat(row.outputQty) || 0;
        const loss = inputQty > 0 ? Math.round((inputQty - outputQty) * 100) / 100 : 0;
        const lossPct = inputQty > 0 ? Math.round((inputQty - outputQty) / inputQty * 1000) / 10 : 0;
        const inputOverflow = inputQty > maxQty + 0.01 && selected;
        const outputOverflow = outputQty > inputQty + 0.01;
        return (
          <div key={i} style={{background:T.card,borderRadius:T.rad,padding:12,border:`1px solid ${T.border}`,display:"flex",flexDirection:"column",gap:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <span style={{fontSize:12,fontWeight:600,color:T.textSec}}>Batch {i + 1}</span>
              {rows.length > 1 && <button onClick={() => removeRow(i)} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={14} color={T.danger}/></button>}
            </div>
            {/* Source batch dropdown */}
            <Field label="Source Green Bean Batch">
              <SearchableDropdown
                options={[...safeEntries].sort((a,b)=>{
                  const af=a.fullyAllocated||(remainingFor(a,i)<=0&&(a.autoId||a.linkedId)!==row.sourceAutoId);
                  const bf=b.fullyAllocated||(remainingFor(b,i)<=0&&(b.autoId||b.linkedId)!==row.sourceAutoId);
                  return af&&!bf?1:!af&&bf?-1:0;
                }).map(e=>{
                  const id=e.autoId||e.linkedId;const rem=remainingFor(e,i);
                  const dis=(e.fullyAllocated||rem<=0)&&id!==row.sourceAutoId;
                  return {label:id+" — "+(rem>0?Math.round(rem*100)/100+"kg available":"fully allocated"),value:id,disabled:dis,sublabel:dis?"fully allocated":""};
                })}
                value={row.sourceAutoId} onChange={v=>updateRow(i,{sourceAutoId:v})} placeholder="-- Select batch --"/>
            </Field>
            {selected && <div style={{fontSize:11,color:T.textMut,marginTop:-4}}>{Math.round(maxQty*100)/100}kg available from this batch</div>}
            {/* Quantities row */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:8}}>
              <Field label="Input Qty (kg)">
                <input type="number" min="0" value={row.inputQty} onChange={e => updateRow(i, { inputQty: e.target.value })}
                  style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${inputOverflow?T.danger:T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
              </Field>
              <Field label="Output Qty (kg)">
                <input type="number" min="0" value={row.outputQty} onChange={e => updateRow(i, { outputQty: e.target.value })}
                  style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${outputOverflow?T.danger:T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
              </Field>
              <Field label="Loss">
                <div style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surfaceHover,border:`1px solid ${T.border}`,color:T.textMut,fontSize:13}}>{loss > 0 ? loss + " kg" : "--"}</div>
              </Field>
              <Field label="Loss %">
                <div style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surfaceHover,border:`1px solid ${T.border}`,color:loss>0?T.warning:T.textMut,fontSize:13}}>{lossPct > 0 ? lossPct + "%" : "--"}</div>
              </Field>
            </div>
            {inputOverflow && <div style={{fontSize:11,color:T.danger}}>Input exceeds available ({Math.round(maxQty*100)/100}kg)</div>}
            {outputOverflow && <div style={{fontSize:11,color:T.danger}}>Output cannot exceed input ({inputQty}kg)</div>}
            {/* Reason for loss */}
            {loss > 0 && <Field label="Reason for Loss">
              <input value={row.reasonForLoss || ""} onChange={e => updateRow(i, { reasonForLoss: e.target.value })} placeholder="Reason for weight loss..."
                style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
            </Field>}
            {/* Roast classification */}
            <Field label="Roast Classification">
              <div style={{display:"flex",gap:6}}>
                <SearchableDropdown
                  options={roastDegrees.map(c=>({label:c.name,value:c.id}))}
                  value={row.classificationId||""} onChange={v=>updateRow(i,{classificationId:v})} placeholder="-- Select --"
                  style={{flex:1}}/>
                {!addingClass && <button onClick={() => setAddingClass(true)} style={{padding:"6px 10px",borderRadius:T.radSm,background:T.accentBg,border:`1px solid ${T.accentBorder}`,color:T.accent,fontSize:12,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>+ New</button>}
              </div>
            </Field>
            {addingClass && (
              <div style={{display:"flex",gap:6,alignItems:"center"}}>
                <input value={newClassName} onChange={e=>setNewClassName(e.target.value)} placeholder="Classification name..." autoFocus
                  style={{flex:1,padding:"6px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.accentBorder}`,color:T.text,fontSize:12,outline:"none"}}/>
                <Btn small onClick={handleAddClass} disabled={savingClass||!newClassName.trim()}>{savingClass?"...":"Save"}</Btn>
                <Btn small variant="ghost" onClick={()=>{setAddingClass(false);setNewClassName("")}}>Cancel</Btn>
              </div>
            )}
          </div>
        );
      })}

      {rows.length < 6 && <Btn variant="ghost" small onClick={addRow}><Icon name="plus" size={12} color={T.textSec}/> Add Another Batch</Btn>}

      {/* Running totals */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,padding:"10px 12px",background:T.accentBg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
        <div style={{textAlign:"center"}}><div style={{fontSize:11,color:T.textMut}}>Total Input</div><div style={{fontSize:15,fontWeight:600,color:T.text}}>{totalInput} kg</div></div>
        <div style={{textAlign:"center"}}><div style={{fontSize:11,color:T.textMut}}>Total Output</div><div style={{fontSize:15,fontWeight:600,color:T.success}}>{totalOutput} kg</div></div>
        <div style={{textAlign:"center"}}><div style={{fontSize:11,color:T.textMut}}>Total Loss</div><div style={{fontSize:15,fontWeight:600,color:totalLoss>0?T.warning:T.textMut}}>{totalLoss} kg ({totalLossPct}%)</div></div>
      </div>
    </div>
  );
}

// ─── Roast Batch View (read-only table for viewing multi-batch entries) ──
function RoastBatchTable({ batchesJson, classifications }) {
  const batches = typeof batchesJson === "string" ? (() => { try { return JSON.parse(batchesJson); } catch(e) { return []; }})() : (Array.isArray(batchesJson) ? batchesJson : []);
  if (batches.length === 0) return null;
  const roastDegrees = classifications?.roast_degree || [];
  const classLabel = (id) => { const c = roastDegrees.find(r => r.id === id); return c ? c.name : ""; };
  const totalIn = batches.reduce((s, b) => s + (parseFloat(b.inputQty) || 0), 0);
  const totalOut = batches.reduce((s, b) => s + (parseFloat(b.outputQty) || 0), 0);
  return (
    <div style={{overflowX:"auto"}}>
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:T.surfaceHover}}>
          <th style={{padding:"6px 8px",textAlign:"left",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Source</th>
          <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>In (kg)</th>
          <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Out (kg)</th>
          <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Loss</th>
          <th style={{padding:"6px 8px",textAlign:"right",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Loss%</th>
          <th style={{padding:"6px 8px",textAlign:"left",color:T.textMut,borderBottom:`1px solid ${T.border}`}}>Classification</th>
        </tr></thead>
        <tbody>
          {batches.map((b, i) => {
            const inQ = parseFloat(b.inputQty) || 0;
            const outQ = parseFloat(b.outputQty) || 0;
            const loss = Math.round((inQ - outQ) * 100) / 100;
            const pct = inQ > 0 ? Math.round((inQ - outQ) / inQ * 1000) / 10 : 0;
            return <tr key={i} style={{borderBottom:`1px solid ${T.border}`}}>
              <td style={{padding:"6px 8px",fontFamily:T.mono,color:T.accent}}>{b.sourceAutoId}</td>
              <td style={{padding:"6px 8px",textAlign:"right",color:T.text}}>{inQ}</td>
              <td style={{padding:"6px 8px",textAlign:"right",color:T.success}}>{outQ}</td>
              <td style={{padding:"6px 8px",textAlign:"right",color:loss>0?T.warning:T.textMut}}>{loss}</td>
              <td style={{padding:"6px 8px",textAlign:"right",color:loss>0?T.warning:T.textMut}}>{pct}%</td>
              <td style={{padding:"6px 8px",color:T.textSec}}>{classLabel(b.classificationId)}</td>
            </tr>;
          })}
          <tr style={{background:T.surfaceHover,fontWeight:600}}>
            <td style={{padding:"6px 8px",color:T.textSec}}>Total</td>
            <td style={{padding:"6px 8px",textAlign:"right",color:T.text}}>{totalIn}</td>
            <td style={{padding:"6px 8px",textAlign:"right",color:T.success}}>{totalOut}</td>
            <td style={{padding:"6px 8px",textAlign:"right",color:T.warning}}>{Math.round((totalIn-totalOut)*100)/100}</td>
            <td style={{padding:"6px 8px",textAlign:"right",color:T.warning}}>{totalIn>0?Math.round((totalIn-totalOut)/totalIn*1000)/10:0}%</td>
            <td style={{padding:"6px 8px"}}></td>
          </tr>
        </tbody>
      </table>
    </div>
  );
}

function GrindClassAddBtn({ onAdd, addToast }) {
  const [open,setOpen]=useState(false);
  const [name,setName]=useState("");
  const [saving,setSaving]=useState(false);
  if(!open) return <button onClick={()=>setOpen(true)} style={{padding:"6px 10px",borderRadius:T.radSm,background:T.accentBg,border:`1px solid ${T.accentBorder}`,color:T.accent,fontSize:12,cursor:"pointer",whiteSpace:"nowrap",flexShrink:0}}>+ New</button>;
  return <div style={{display:"flex",gap:4,alignItems:"center",flexShrink:0}}>
    <input value={name} onChange={e=>setName(e.target.value)} placeholder="Name..." autoFocus style={{width:100,padding:"4px 8px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.accentBorder}`,color:T.text,fontSize:12,outline:"none"}}/>
    <Btn small onClick={async()=>{if(!name.trim())return;setSaving(true);try{await onAdd(name.trim())}catch(e){if(addToast)addToast(e.message,"error")}setSaving(false);setOpen(false);setName("")}} disabled={saving||!name.trim()}>{saving?"...":"Save"}</Btn>
    <Btn small variant="ghost" onClick={()=>{setOpen(false);setName("")}}>Cancel</Btn>
  </div>;
}

function BatchSelector({ entries, allocations, onChange, checklistName, emptyMessage }) {
  const safeEntries = entries || [];
  const rows = Array.isArray(allocations) && allocations.length > 0 ? allocations : [{ sourceAutoId: "", quantity: "" }];

  const updateRow = (i, patch) => {
    const next = rows.map((r, idx) => idx === i ? { ...r, ...patch } : r);
    onChange(next.filter(r => r.sourceAutoId || r.quantity));
  };
  const removeRow = (i) => {
    const next = rows.filter((_, idx) => idx !== i);
    onChange(next);
  };
  const addRow = () => {
    onChange([...rows, { sourceAutoId: "", quantity: "" }]);
  };

  // Compute remaining for each entry, accounting for amounts already allocated to that entry in OTHER rows of this form
  const remainingFor = (entry, currentIdx) => {
    const baseRemaining = entry.remainingQuantity ?? ((entry.totalQuantity || entry.masterQuantity || 0) - (entry.usedQuantity || entry.allocatedQuantity || 0));
    const id = entry.autoId || entry.linkedId;
    let usedInThisForm = 0;
    rows.forEach((r, idx) => {
      if (idx !== currentIdx && r.sourceAutoId === id) usedInThisForm += parseFloat(r.quantity) || 0;
    });
    return baseRemaining - usedInThisForm;
  };

  if (safeEntries.length === 0) {
    return (
      <div style={{padding:"12px 14px",borderRadius:T.radSm,background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",color:T.danger,fontSize:13}}>
        {emptyMessage || `No approved batches found in ${checklistName || "source"}`}
      </div>
    );
  }

  const total = rows.reduce((s, r) => s + (parseFloat(r.quantity) || 0), 0);
  const hasAvailable = safeEntries.some(e => {
    const rem = e.remainingQuantity ?? ((e.totalQuantity || e.masterQuantity || 0) - (e.usedQuantity || e.allocatedQuantity || 0));
    return rem > 0;
  });

  return (
    <div style={{display:"flex",flexDirection:"column",gap:8}}>
      {rows.map((row, i) => {
        const selected = safeEntries.find(e => (e.autoId || e.linkedId) === row.sourceAutoId);
        const maxQty = selected ? remainingFor(selected, i) : 0;
        const qtyVal = parseFloat(row.quantity) || 0;
        const overflow = selected && qtyVal > maxQty;
        return (
          <div key={i} style={{display:"flex",gap:8,alignItems:"flex-start",padding:8,background:T.bg,borderRadius:T.radSm,border:`1px solid ${overflow?T.danger:T.border}`}}>
            <div style={{flex:2,minWidth:0}}>
              <select value={row.sourceAutoId} onChange={e=>updateRow(i,{sourceAutoId:e.target.value})}
                style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13}}>
                <option value="">— Select batch —</option>
                {[...safeEntries].sort((a,b)=>{
                  const af=a.fullyAllocated||(remainingFor(a,i)<=0&&(a.autoId||a.linkedId)!==row.sourceAutoId);
                  const bf=b.fullyAllocated||(remainingFor(b,i)<=0&&(b.autoId||b.linkedId)!==row.sourceAutoId);
                  if(af&&!bf)return 1;if(!af&&bf)return -1;return 0;
                }).map((e,ei)=>{
                  const id=e.autoId||e.linkedId;
                  const rem=remainingFor(e,i);
                  const disabled=(e.fullyAllocated||rem<=0)&&id!==row.sourceAutoId;
                  return <option key={ei} value={id} disabled={disabled}>{id} — {rem>0?`${rem} available`:"\u2014 fully allocated"}</option>;
                })}
              </select>
              {selected && <div style={{fontSize:11,color:T.textMut,marginTop:4}}>{maxQty} available from this batch</div>}
            </div>
            <input type="number" min="0" value={row.quantity} onChange={e=>{ const v=e.target.value; const n=parseFloat(v); if(!isNaN(n)&&n<0){ updateRow(i,{quantity:"0"}); return; } updateRow(i,{quantity:v}); }} onBlur={e=>{ const n=parseFloat(e.target.value); if(!isNaN(n)&&n<0) updateRow(i,{quantity:"0"}); }} placeholder="Qty"
              style={{width:90,padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${overflow?T.danger:T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
            <button onClick={()=>removeRow(i)} style={{background:"none",border:"none",cursor:"pointer",padding:6,marginTop:2}} title="Remove">
              <Icon name="x" size={14} color={T.danger}/>
            </button>
          </div>
        );
      })}
      <Btn variant="ghost" small onClick={addRow} disabled={!hasAvailable}>
        <Icon name="plus" size={12} color={T.textSec}/> Add another batch
      </Btn>
      <div style={{padding:"8px 12px",background:T.accentBg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`,fontSize:12,color:T.accent,fontWeight:600}}>
        Total from selected batches: {total} kgs
      </div>
    </div>
  );
}

// ─── Linked Dropdown (with quantity + preview) ────────────────

function LinkedDropdown({ entries, value, onChange, checklistName, sourceChecklistId, checklists, placeholder, emptyMessage }) {
  const [open,setOpen]=useState(false);
  const [search,setSearch]=useState("");
  const [preview,setPreview]=useState(null);
  const ref=useRef(null);
  const safeEntries=entries||[];
  const filtered=safeEntries.filter(e=>{const s=search.toLowerCase();const id=e.autoId||e.linkedId||"";return id.toLowerCase().includes(s)||(e.linkedId||"").toLowerCase().includes(s)||(e.orderId||"").toLowerCase().includes(s)||(e.orderName||"").toLowerCase().includes(s)}).sort((a,b)=>{
    const aFull=a.fullyAllocated||(a.totalQuantity>0&&a.remainingQuantity!==undefined&&a.remainingQuantity<=0);
    const bFull=b.fullyAllocated||(b.totalQuantity>0&&b.remainingQuantity!==undefined&&b.remainingQuantity<=0);
    if(aFull&&!bFull)return 1;if(!aFull&&bFull)return -1;return 0;
  });
  useEffect(()=>{
    const handler=e=>{if(ref.current&&!ref.current.contains(e.target))setOpen(false)};
    document.addEventListener("mousedown",handler);return()=>document.removeEventListener("mousedown",handler);
  },[]);
  const selectedEntry=safeEntries.find(e=>(e.autoId&&e.autoId===value)||e.linkedId===value);
  const displayId=(e)=>e.autoId||e.linkedId;

  // Empty state: no approved entries at all
  if(safeEntries.length===0){
    return (
      <div>
        <div style={{padding:"12px 14px",borderRadius:T.radSm,background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",color:T.danger,fontSize:13}}>
          {emptyMessage||"No approved entries found in "+checklistName}
          <p style={{fontSize:12,marginTop:4,color:T.textMut}}>Please complete and approve a {checklistName} entry first.</p>
        </div>
      </div>
    );
  }

  return (
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={()=>setOpen(!open)} style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${value?T.accentBorder:T.border}`,color:value?T.text:T.textMut,fontSize:14,textAlign:"left",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{(selectedEntry&&displayId(selectedEntry))||value||placeholder||"Select..."}</span>
        <Icon name="chevron" size={14} color={T.textMut} style={{transform:open?"rotate(90deg)":"rotate(0)",transition:"transform .2s",flexShrink:0}}/>
      </button>
      {/* Quantity tracker bar */}
      {selectedEntry&&selectedEntry.totalQuantity!==undefined&&selectedEntry.totalQuantity>0&&(
        <div style={{display:"flex",gap:10,marginTop:6,padding:"8px 10px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`,fontSize:12}}>
          <span style={{color:T.textSec}}>Total: <b style={{color:T.text}}>{selectedEntry.totalQuantity}</b></span>
          <span style={{color:T.warning}}>Used: <b>{selectedEntry.usedQuantity}</b></span>
          <span style={{color:selectedEntry.remainingQuantity>0?T.success:T.danger}}>Remaining: <b>{selectedEntry.remainingQuantity}</b></span>
        </div>
      )}
      {/* View details link */}
      {selectedEntry&&(
        <button onClick={()=>setPreview(selectedEntry)} style={{background:"none",border:"none",cursor:"pointer",padding:"6px 0",marginTop:2,fontSize:13,color:T.info,display:"flex",alignItems:"center",gap:6}}>
          <Icon name="clipboard" size={14} color={T.info}/> View {displayId(selectedEntry)} details
        </button>
      )}
      {open&&(
        <div style={{position:"absolute",top:"100%",left:0,right:0,zIndex:20,background:T.card,border:`1px solid ${T.border}`,borderRadius:T.radSm,marginTop:4,boxShadow:"0 8px 24px rgba(0,0,0,0.4)",maxHeight:280,display:"flex",flexDirection:"column"}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search..." autoFocus
            style={{padding:"10px 12px",background:T.bg,border:"none",borderBottom:`1px solid ${T.border}`,color:T.text,fontSize:14,outline:"none",minHeight:44}}/>
          <div style={{overflowY:"auto",flex:1}}>
            {filtered.length===0?<p style={{padding:12,fontSize:13,color:T.textMut}}>No matching entries</p>:
              filtered.map((e,i)=>{
                const id=displayId(e);
                const isFull=e.fullyAllocated||(e.totalQuantity>0&&e.remainingQuantity!==undefined&&e.remainingQuantity<=0);
                return <button key={i} disabled={isFull} onClick={()=>{if(isFull)return;onChange(id);setOpen(false);setSearch("")}}
                  style={{display:"block",width:"100%",padding:"10px 12px",background:id===value?T.accentBg:"none",border:"none",borderBottom:`1px solid ${T.border}`,color:isFull?T.textMut:T.text,fontSize:14,cursor:isFull?"not-allowed":"pointer",textAlign:"left",minHeight:44,opacity:isFull?0.4:1}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:500,textDecoration:isFull?"line-through":"none"}}>{id}</span>
                    {e.remainingQuantity!==undefined&&e.totalQuantity>0&&<span style={{fontSize:11,color:isFull?T.danger:T.success}}>{isFull?"\u2014 fully allocated":`${e.remainingQuantity} available`}</span>}
                  </div>
                  <div style={{display:"flex",gap:8,marginTop:2,flexWrap:"wrap"}}>
                    {e.person&&<span style={{fontSize:11,color:T.textMut}}>by {e.person}</span>}
                    {e.date&&<span style={{fontSize:11,color:T.textMut}}>{e.date}</span>}
                    {e.orderId&&e.orderId!=="UNTAGGED"&&<span style={{fontSize:11,color:T.textMut}}>Order: <span style={{color:T.accent,fontWeight:500}}>{e.orderId}</span></span>}
                  </div>
                </button>;
              })
            }
          </div>
        </div>
      )}
      {preview&&<PreviewModal entry={preview} checklistName={checklistName} sourceChecklistId={sourceChecklistId} checklists={checklists} onClose={()=>setPreview(null)}/>}
    </div>
  );
}

// ─── Quick Fill View (Free-form Checklist) ────────────────────

function QuickFillView({ checklists, orders, customers, currentUser, approvedEntries, inventoryItems, inventoryCategories, onSubmit, onSaveDraft, resumeDraft, allOrders, addToast }) {
  const [selCkId,setSelCkId]=useState(resumeDraft?.checklistId||"");
  const [formData,setFormData]=useState(()=>{
    if(resumeDraft){
      return {
        date: resumeDraft.workDate || new Date().toISOString().split("T")[0],
        person: resumeDraft.person || currentUser?.displayName || "",
        responses: resumeDraft.responses || {},
        remarks: resumeDraft.remarks || {},
        batchAllocations: resumeDraft.batchAllocations || {},
      };
    }
    return {date:new Date().toISOString().split("T")[0],person:currentUser?.displayName||"",responses:{},remarks:{},batchAllocations:{}};
  });
  const [draftId,setDraftId]=useState(resumeDraft?.id||null);
  const [invItemId,setInvItemId]=useState("");
  const [invOutputItemId,setInvOutputItemId]=useState("");
  const [orderId,setOrderId]=useState("");
  const [submitting,setSubmitting]=useState(false);
  const [savingDraft,setSavingDraft]=useState(false);
  const [invError,setInvError]=useState({idx:null,message:""});
  const [roastBatches,setRoastBatches]=useState([{sourceAutoId:"",inputQty:"",outputQty:"",reasonForLoss:"",classificationId:""}]);
  const [classifications,setClassifications]=useState(null);
  const [grindClassificationId,setGrindClassificationId]=useState("");

  const ck=checklists.find(c=>c.id===selCkId);
  const nq=ck?normalizeQuestions(ck.questions):[];

  // Batch allocation change handler — updates allocations and auto-fills any OUT-linked field with the new total
  const onBatchAllocChange=(qi,allocs)=>{
    setFormData(p=>{
      const nextBA={...(p.batchAllocations||{}),[qi]:allocs};
      let total=0;
      Object.values(nextBA).forEach(a=>{if(Array.isArray(a))a.forEach(x=>total+=parseFloat(x.quantity)||0);});
      const nextResponses={...p.responses};
      if(ck){
        normalizeQuestions(ck.questions).forEach((qq,i)=>{
          if(qq.inventoryLink?.enabled&&qq.inventoryLink.txType==="OUT") nextResponses[i]=String(total);
        });
      }
      return {...p,batchAllocations:nextBA,responses:nextResponses};
    });
  };

  const getFieldValue=(checklistRef,questionIdx)=>{
    if(checklistRef==="self") return formData.responses[questionIdx]||"";
    return "";
  };

  // Load classifications when ck_roasted_beans or ck_grinding is selected
  useEffect(()=>{
    if((selCkId==="ck_roasted_beans"||selCkId==="ck_grinding")&&!classifications){
      API.get("getClassifications").then(d=>{if(d&&!d.error)setClassifications(d)}).catch(()=>{});
    }
  },[selCkId]);

  const isMultiBatchRoast = selCkId === "ck_roasted_beans";
  const hasValidRoastBatches = isMultiBatchRoast && roastBatches.some(b => b.sourceAutoId && (parseFloat(b.inputQty) || 0) > 0);

  const orderOptions=[{label:"Untagged (no order)",value:""},...orders.filter(o=>o.canTag!==false&&o.status!=="delivered"&&o.status!=="cancelled").map(o=>({label:`${o.id} — ${o.name}`,value:o.id}))];

  const handleSubmit=async()=>{
    setSubmitting(true);
    setInvError({idx:null,message:""});
    // Required inventory-tracking fields must have a value
    const computedResponsesForInv = {};
    nq.forEach((q, qi) => {
      if (q.formula) {
        const computed = evaluateFormula(q.formula, getFieldValue);
        computedResponsesForInv[qi] = computed !== null ? String(computed) : "";
      } else {
        computedResponsesForInv[qi] = formData.responses[qi] !== undefined ? formData.responses[qi] : "";
      }
    });
    // When multi-batch roast is active with valid batches, skip inventory field validation
    // — RoastBatchSection handles all inventory tracking for those fields
    if (!hasValidRoastBatches) {
      const invCheck = validateRequiredInventoryFields(nq, computedResponsesForInv, formData.batchAllocations);
      if (!invCheck.ok) {
        setInvError({idx: invCheck.firstIdx, message: invCheck.message});
        setSubmitting(false);return;
      }
    }
    // Validate linked dropdown fields (skip for multi-batch roast — source batches handled by RoastBatchSection)
    for(let qi=0;qi<nq.length;qi++){
      const q=nq[qi];
      if(q.linkedSource&&q.linkedSource.checklistId){
        if (isMultiBatchRoast && q.linkedSource.checklistId === "ck_green_beans" && hasValidRoastBatches) continue;
        const entries=approvedEntries[q.linkedSource.checklistId]||[];
        const srcCk=checklists.find(c=>c.id===q.linkedSource.checklistId);
        const srcName=srcCk?.name||"source checklist";
        if(entries.length===0){
          alert("Cannot submit — no approved entries available from \""+srcName+"\". Please complete and approve a "+srcName+" first.");
          setSubmitting(false);return;
        }
        const hasDirectValue = (formData.responses[qi]||"").trim();
        const batchAllocs = formData.batchAllocations?.[qi];
        const hasValidBatch = Array.isArray(batchAllocs) && batchAllocs.some(a => a.sourceAutoId && (parseFloat(a.quantity) || 0) > 0);
        if(!hasDirectValue && !hasValidBatch){
          alert("Please select a "+q.text+" before submitting.");
          setSubmitting(false);return;
        }
      }
    }
    // Check required remarks
    for(let qi=0;qi<nq.length;qi++){
      const q=nq[qi];
      if(q.remarkCondition && q.ideal){
        const val=formData.responses[qi]||"";
        const idealVal=evaluateFormula(q.ideal,getFieldValue);
        if(checkRemarkCondition(val,idealVal,q.remarkCondition) && !(formData.remarks[qi]||"").trim()){
          alert("Please provide remarks for: "+q.text);
          setSubmitting(false);return;
        }
      }
    }
    // Forced remark targets — another field's deviation requires this field to be filled
    for(let qi=0;qi<nq.length;qi++){
      const q=nq[qi];
      if(!q.remarkCondition||q.remarksTargetIdx==null||!q.formula||!q.ideal) continue;
      const aVal=evaluateFormula(q.formula,getFieldValue);
      const iVal=evaluateFormula(q.ideal,getFieldValue);
      if(aVal!=null&&iVal!=null&&checkRemarkCondition(aVal,iVal,q.remarkCondition)){
        const targetVal=formData.responses[q.remarksTargetIdx]||"";
        if(!String(targetVal).trim()){
          const targetText=nq[q.remarksTargetIdx]?.text||`Q${q.remarksTargetIdx+1}`;
          alert(`Please fill "${targetText}" — ${q.remarkCondition.message||"value differs from ideal"}`);
          setSubmitting(false);return;
        }
      }
    }
    // Date comparison validation
    for(let qi=0;qi<nq.length;qi++){
      const q=nq[qi];
      if(q.type==="date"&&q.dateComparison&&q.dateComparison.compareToFieldIdx!==""&&q.dateComparison.compareToFieldIdx!==undefined){
        const val=formData.responses[qi]||"";
        const cmpVal=formData.responses[Number(q.dateComparison.compareToFieldIdx)]||"";
        if(val&&cmpVal){
          const d1=new Date(val),d2=new Date(cmpVal);
          if(!isNaN(d1.getTime())&&!isNaN(d2.getTime())){
            const t1=d1.setHours(0,0,0,0),t2=d2.setHours(0,0,0,0);
            const op=q.dateComparison.operator;
            const cmpField=nq[Number(q.dateComparison.compareToFieldIdx)]?.text||"the other date";
            let err=null;
            if(op==="gte"&&t1<t2) err=q.dateComparison.errorMessage||`${q.text} cannot be before ${cmpField}`;
            else if(op==="lte"&&t1>t2) err=q.dateComparison.errorMessage||`${q.text} cannot be after ${cmpField}`;
            else if(op==="eq"&&t1!==t2) err=q.dateComparison.errorMessage||`${q.text} must be the same as ${cmpField}`;
            if(err){alert(err);setSubmitting(false);return;}
          }
        }
      }
    }
    const responses=nq.map((q,qi)=>{
      let resp=formData.responses[qi]||"";
      // Compute formula fields fresh from current state
      if(q.formula){
        const computed=evaluateFormula(q.formula,getFieldValue);
        resp = computed!==null?String(computed):"";
      }
      if(q.linkedSource&&Array.isArray(formData.batchAllocations?.[qi])){
        const ids=formData.batchAllocations[qi].map(a=>a.sourceAutoId).filter(Boolean);
        if(ids.length>0) resp=ids.join(", ");
      }
      return {questionIndex:qi,questionText:q.text,response:resp};
    });
    const payload = {checklistId:selCkId,date:formData.date,person:formData.person,responses,remarks:formData.remarks,orderId:orderId||"",inventoryItemId:invItemId,inventoryOutputItemId:invOutputItemId,batchAllocations:formData.batchAllocations||{},grindClassificationId:grindClassificationId||""};
    if (isMultiBatchRoast) {
      // Validate roast batches before submitting
      const validBatches = roastBatches.filter(b => b.sourceAutoId && (parseFloat(b.inputQty) || 0) > 0);
      if (validBatches.length === 0) { alert("Please add at least one roast batch with a source and input quantity."); setSubmitting(false); return; }
      for (let bi = 0; bi < validBatches.length; bi++) {
        const b = validBatches[bi];
        if ((parseFloat(b.outputQty)||0) > (parseFloat(b.inputQty)||0)) { alert("Batch "+(bi+1)+": output cannot exceed input."); setSubmitting(false); return; }
        if ((parseFloat(b.inputQty)||0) - (parseFloat(b.outputQty)||0) > 0 && !(b.reasonForLoss||"").trim()) { alert("Batch "+(bi+1)+": please provide a reason for loss."); setSubmitting(false); return; }
      }
      payload.roast_batches = validBatches;
    }
    await onSubmit(payload);
    setSubmitting(false);
  };

  const handleSaveDraft=async()=>{
    if(!selCkId){alert("Pick a checklist first");return;}
    if(typeof onSaveDraft!=="function") return;
    setSavingDraft(true);
    try{
      const r=await onSaveDraft({
        id: draftId || undefined,
        checklistId: selCkId,
        checklistName: ck?.name || "",
        responses: formData.responses,
        remarks: formData.remarks,
        batchAllocations: formData.batchAllocations,
        person: formData.person,
        workDate: formData.date,
        linkedOrders: orderId ? [orderId] : [],
      });
      if(r?.id) setDraftId(r.id);
    }catch{}
    setSavingDraft(false);
  };

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:20}}>
      <Field label="Select Checklist">
        <SearchableDropdown options={checklists.map(c=>({label:c.name,value:c.id}))} value={selCkId} onChange={v=>{setSelCkId(v);setFormData(p=>({...p,responses:{},remarks:{}}))}} placeholder="— Pick a checklist —"/>
      </Field>

      {selCkId && <>
        {ck?.autoIdConfig?.enabled && (() => {
          const preview = buildAutoIdPreview(ck, formData.responses, formData.date, inventoryItems);
          return preview ? (
            <div style={{padding:"10px 12px",background:T.accentBg,border:`1px solid ${T.accentBorder}`,borderRadius:T.radSm,display:"flex",alignItems:"center",gap:8}}>
              <Icon name="clipboard" size={14} color={T.accent}/>
              <span style={{fontSize:11,color:T.textMut}}>Auto ID:</span>
              <span style={{fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600}}>{preview}</span>
            </div>
          ) : null;
        })()}

        <Field label="Tag to Order (optional)">
          <SearchableDropdown options={orderOptions} value={orderId} onChange={setOrderId} placeholder="Untagged (no order)" emptyMessage="No orders found"/>
        </Field>

        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <Field label="Date">
            <input type="date" value={formData.date} onChange={e=>setFormData(p=>({...p,date:e.target.value}))}
              style={{width:"100%",padding:"12px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:15,outline:"none",colorScheme:"dark"}}/>
          </Field>
          <Field label="Person">
            <Input value={formData.person} onChange={()=>{}} readOnly placeholder="Who is filling this?" style={{fontSize:15,padding:"12px 14px"}}/>
          </Field>
        </div>

        {/* ── Inventory Item Selection ── */}
        {(()=>{
          const hasNewInvLink=nq.some(qq=>qq.inventoryLink&&qq.inventoryLink.enabled);
          if(hasNewInvLink) return null;
          const invCatMap={"ck_green_beans":"Green Beans","ck_roasted_beans":"Green Beans","ck_grinding":"Roasted Beans"};
          const invOutCatMap={"ck_roasted_beans":"Roasted Beans","ck_grinding":"Packing Items"};
          const inCat=invCatMap[selCkId];const outCat=invOutCatMap[selCkId];
          if(!inCat&&!outCat) return null;
          const inItems=(inventoryItems||[]).filter(i=>i.category===inCat&&i.isActive);
          const allOutItems=outCat?(inventoryItems||[]).filter(i=>i.category===outCat&&i.isActive):[];
          // Determine output item behaviour based on input item's equivalents
          // equivalentItems is an array of { category, itemId } objects
          const selectedInItem=invItemId?(inventoryItems||[]).find(i=>i.id===invItemId):null;
          const inEqList=selectedInItem&&Array.isArray(selectedInItem.equivalentItems)?selectedInItem.equivalentItems:[];
          const inEqItemIds=inEqList.map(e=>e.itemId);
          const linkedOutItems=inEqItemIds.length>0?allOutItems.filter(i=>inEqItemIds.includes(i.id)):[];
          const outputLocked=inEqItemIds.length===1&&linkedOutItems.length===1;
          const outputChoices=inEqItemIds.length>0&&linkedOutItems.length>0?linkedOutItems:allOutItems;
          return <div style={{display:"flex",flexDirection:"column",gap:12,padding:12,background:T.accentBg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
            <span style={{fontSize:12,fontWeight:600,color:T.accent}}>Inventory Tracking</span>
            {inCat&&<Field label={selCkId==="ck_green_beans"?"Green Bean Item (IN)":"Input Item (OUT)"}>
              <select value={invItemId} onChange={e=>{
                const newId=e.target.value;
                setInvItemId(newId);
                if(!newId){setInvOutputItemId("");return;}
                const selItem=(inventoryItems||[]).find(i=>i.id===newId);
                const eqL=selItem&&Array.isArray(selItem.equivalentItems)?selItem.equivalentItems:[];
                const eqItemIds=eqL.map(eq=>eq.itemId);
                const eqOutItems=eqItemIds.length>0?allOutItems.filter(i=>eqItemIds.includes(i.id)):[];
                if(eqItemIds.length===1&&eqOutItems.length===1){setInvOutputItemId(eqOutItems[0].id)}
                else if(eqItemIds.length>1&&eqOutItems.length>0){if(!eqOutItems.find(i=>i.id===invOutputItemId))setInvOutputItemId("")}
                else{/* no equivalents — keep current selection */}
              }}
                style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14}}>
                <option value="">— None (skip inventory) —</option>
                {inItems.map(i=><option key={i.id} value={i.id}>{i.name} ({i.currentStock} {i.unit})</option>)}
              </select>
            </Field>}
            {outCat&&<Field label={<span>Output Item (IN){outputLocked&&<span style={{fontSize:10,color:T.textMut,marginLeft:6,fontWeight:400}}>Auto-filled from linked equivalent</span>}</span>}>
              <select value={invOutputItemId} onChange={e=>setInvOutputItemId(e.target.value)} disabled={outputLocked}
                style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:outputLocked?"rgba(255,255,255,0.05)":T.bg,border:`1px solid ${T.border}`,color:outputLocked?T.textMut:T.text,fontSize:14,opacity:outputLocked?0.7:1,cursor:outputLocked?"not-allowed":"pointer"}}>
                <option value="">— None (skip inventory) —</option>
                {outputChoices.map(i=><option key={i.id} value={i.id}>{i.name} ({i.currentStock} {i.unit})</option>)}
              </select>
            </Field>}
          </div>;
        })()}

        {/* ── Multi-batch roasting section (replaces Shipment/Qty fields for ck_roasted_beans) ── */}
        {isMultiBatchRoast && (
          <RoastBatchSection
            entries={approvedEntries?.["ck_green_beans"]||[]}
            batches={roastBatches}
            onChange={setRoastBatches}
            classifications={classifications}
            onAddClassification={async(name)=>{
              const r=await API.post("addClassification",{name,type:"roast_degree"});
              if(r&&!r.error){const d=await API.get("getClassifications");if(d&&!d.error)setClassifications(d);}
              else throw new Error(r?.error||"Failed");
            }}
            addToast={addToast}
          />
        )}

        {/* ── Grind Size Classification (ck_grinding only) ── */}
        {selCkId==="ck_grinding"&&classifications&&(()=>{
          const grindSizes=classifications.grind_size||[];
          return <div style={{padding:12,background:T.accentBg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
            <Field label="Grind Size Classification (optional)">
              <div style={{display:"flex",gap:6}}>
                <SearchableDropdown
                  options={[{label:"— None —",value:""},...grindSizes.map(c=>({label:c.name,value:c.id}))]}
                  value={grindClassificationId} onChange={setGrindClassificationId} placeholder="— Select grind size —"
                  style={{flex:1}}/>
                <GrindClassAddBtn onAdd={async(name)=>{
                  const r=await API.post("addClassification",{name,type:"grind_size"});
                  if(r&&!r.error){const d=await API.get("getClassifications");if(d&&!d.error){setClassifications(d);setGrindClassificationId(r.id)}}
                  else throw new Error(r?.error||"Failed");
                }} addToast={addToast}/>
              </div>
            </Field>
          </div>;
        })()}

        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          {nq.map((q,qi)=>{
            // For multi-batch roast, skip the fields that are now handled by RoastBatchSection
            if(isMultiBatchRoast){
              const skipTexts=["Shipment number used","Type of Beans","Quantity input","Quantity output","Loss in weight","Reason for loss"];
              if(skipTexts.some(t=>q.text===t)||q.linkedSource?.checklistId==="ck_green_beans") return null;
            }
            let autoVal=null;
            if(q.formula) autoVal=evaluateFormula(q.formula,getFieldValue);
            const currentVal=q.formula
              ? (autoVal!==null?String(autoVal):"")
              : (formData.responses[qi]!==undefined?formData.responses[qi]:"");
            let idealVal=null;
            if(q.ideal) idealVal=evaluateFormula(q.ideal,getFieldValue);
            const needsRemark=q.remarkCondition&&idealVal!==null&&checkRemarkCondition(currentVal,idealVal,q.remarkCondition);
            const required = isInventoryRequiredQuestion(q) && !(isMultiBatchRoast && hasValidRoastBatches);
            const invalid = invError.idx === qi && !hasValidRoastBatches;
            return (
              <div key={qi} style={invalid ? {border:`1px solid ${T.danger}`,borderRadius:T.radSm,padding:8} : undefined}>
                {required && (
                  <div style={{fontSize:11,color:T.danger,marginBottom:4,display:"flex",alignItems:"center",gap:4}}>
                    <span style={{color:T.danger,fontWeight:700}}>*</span> Required for inventory tracking
                  </div>
                )}
                <QuestionInputField q={q} qi={qi} currentVal={currentVal} idealVal={idealVal} needsRemark={needsRemark}
                  formData={formData} setFormData={setFormData} approvedEntries={approvedEntries} checklists={checklists} getFieldValue={getFieldValue}
                  orders={selCkId==="ck_grinding"?orders:null} customers={customers}
                  inventoryItems={inventoryItems} onBatchAllocChange={onBatchAllocChange} allQuestions={nq}
                  onInventoryAutoFill={(item)=>{
                    if(!item){setInvItemId("");setInvOutputItemId("");return;}
                    setInvItemId(item.id);
                    const eqL=Array.isArray(item.equivalentItems)?item.equivalentItems:[];
                    if(eqL.length>0){
                      const eqItem=(inventoryItems||[]).find(it=>eqL.some(e=>e.itemId===it.id));
                      if(eqItem) setInvOutputItemId(eqItem.id);
                    }
                  }}/>
              </div>
            );
          }).filter(Boolean)}
        </div>

        {invError.message && (
          <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"10px 14px"}}>
            <span style={{fontSize:13,color:T.danger}}>{invError.message}</span>
          </div>
        )}

        <div style={{display:"flex",gap:8}}>
          <Btn variant="secondary" onClick={handleSaveDraft} disabled={savingDraft||!selCkId} style={{flex:1}}>
            <Icon name="edit" size={14} color={T.text}/> {savingDraft?"Saving...":(draftId?"Update Draft":"Save as Draft")}
          </Btn>
          <Btn variant="success" onClick={handleSubmit} disabled={submitting||!formData.person.trim()} style={{flex:1}}>
            <Icon name="check" size={16} color={T.success}/> {submitting?"Submitting...":"Submit"}
          </Btn>
        </div>
      </>}
    </div>
  );
}

const ORDER_STATUSES = ["beans_not_roasted","beans_roasted","packed","completed","delivered"];
const ORDER_STATUS_LABELS = {"beans_not_roasted":"Beans not yet roasted","beans_roasted":"Beans roasted","packed":"Packed","completed":"Ready for delivery","delivered":"Delivered"};

const PRODUCT_TYPES = ["Roasted Beans", "Roast & Ground", "Instant Coffee", "Others"];

const PRODUCT_TYPE_COLORS = {
  "Roasted Beans": { bg: "rgba(212,165,116,0.12)", color: "#D4A574", border: "rgba(212,165,116,0.25)" },
  "Roast & Ground": { bg: "rgba(107,203,119,0.12)", color: "#6BCB77", border: "rgba(107,203,119,0.25)" },
  "Instant Coffee": { bg: "rgba(91,156,246,0.12)", color: "#5B9CF6", border: "rgba(91,156,246,0.25)" },
  "Others": { bg: "rgba(155,150,160,0.12)", color: "#9B96A0", border: "rgba(155,150,160,0.25)" },
};

// ─── Source Chain Display Component ──────────────────────────
function SourceChainDisplay({ checklistId, autoId, checklists }) {
  const [chain, setChain] = useState(null);
  const [loading, setLoading] = useState(false);
  const [expanded, setExpanded] = useState(true);
  const [modalItem, setModalItem] = useState(null);

  useEffect(() => {
    if (!checklistId || !autoId) { setChain(null); return; }
    setLoading(true);
    API.get("getResponseChain", { checklistId, responseId: autoId }).then(data => {
      // data is the full chain starting from the selected source entry
      // Show all entries — each is an upstream link the user wants to inspect
      if (Array.isArray(data) && data.length > 0 && !data.error) setChain(data);
      else setChain(null);
    }).catch(() => setChain(null)).finally(() => setLoading(false));
  }, [checklistId, autoId]);

  if (loading) return <div style={{padding:"8px 12px",background:T.surface,borderRadius:T.radSm,marginTop:8}}><span style={{fontSize:12,color:T.textMut,animation:"pulse 1.5s infinite"}}>Loading source chain...</span></div>;
  if (!chain || chain.length === 0) return null;

  return (
    <div style={{marginTop:8}}>
      <button onClick={()=>setExpanded(!expanded)} style={{display:"flex",alignItems:"center",gap:6,background:"none",border:"none",cursor:"pointer",padding:0,marginBottom:expanded?8:0}}>
        <Icon name="link" size={14} color={T.info}/>
        <span style={{fontSize:12,fontWeight:500,color:T.info}}>Source Chain ({chain.length})</span>
        <Icon name="chevron" size={12} color={T.info} style={{transform:expanded?"rotate(90deg)":"rotate(0)",transition:"transform .2s"}}/>
      </button>
      {expanded && <div style={{display:"flex",flexDirection:"column",gap:4,paddingLeft:8,borderLeft:`2px solid ${T.infoBorder}`}}>
        {chain.map((item, i) => (
          <button key={i} onClick={()=>item.error?null:setModalItem(item)}
            style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`,cursor:item.error?"not-allowed":"pointer",opacity:item.error?0.5:1,textAlign:"left",width:"100%"}}>
            <Icon name="clipboard" size={14} color={item.error?T.textMut:T.accent}/>
            <div style={{flex:1}}>
              <span style={{fontSize:13,fontWeight:500,color:item.error?T.textMut:T.text}}>{item.checklistName} — {item.autoId}</span>
              {item.error && <span style={{fontSize:11,color:T.textMut,marginLeft:8}}>unavailable</span>}
            </div>
            {!item.error && <Icon name="chevron" size={14} color={T.textMut}/>}
          </button>
        ))}
      </div>}
      {modalItem && <ChainEntryModal item={modalItem} onClose={()=>setModalItem(null)}/>}
    </div>
  );
}

function ChainEntryModal({ item, onClose, zIndex }) {
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:zIndex||1000,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,border:`1px solid ${T.border}`,maxWidth:500,width:"100%",maxHeight:"80vh",overflow:"auto",padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div>
            <h3 style={{fontSize:16,fontWeight:600,color:T.text,margin:0}}>{item.checklistName}</h3>
            <span style={{fontSize:12,fontFamily:T.mono,color:T.accent}}>{item.autoId}</span>
          </div>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:2}}>
          {(item.fields||[]).map((f,i) => (
            <div key={i} style={{padding:"8px 0",borderBottom:i<item.fields.length-1?`1px solid ${T.border}`:"none"}}>
              <span style={{fontSize:12,color:T.textMut,display:"block",marginBottom:2}}>{f.question}</span>
              <span style={{fontSize:14,color:T.text,fontWeight:500}}>{f.response || "—"}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// Modal for viewing a tagged checklist response + its source chain.
// Fetches response-chain, uses first entry as the main view, rest as upstream chain.
function TaggedEntryModal({ checklistId, autoId, checklists, isAdmin, onClose }) {
  const [chain, setChain] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [expandedId, setExpandedId] = useState(null); // upstream chain row expanded inline

  useEffect(() => {
    if (!checklistId || !autoId) { setLoading(false); return; }
    setLoading(true); setError("");
    API.get("getResponseChain", { checklistId, responseId: autoId }).then(data => {
      if (Array.isArray(data) && data.length > 0) setChain(data);
      else if (data && data.error) setError(data.error);
      else setError("Response not found");
    }).catch(e => setError(e.message || "Failed to load response")).finally(() => setLoading(false));
  }, [checklistId, autoId]);

  const main = chain && chain.length > 0 ? chain[0] : null;
  const upstream = chain && chain.length > 1 ? chain.slice(1) : [];

  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:900,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,border:`1px solid ${T.border}`,maxWidth:560,width:"100%",maxHeight:"85vh",overflow:"auto",padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
          <div>
            <h3 style={{fontSize:16,fontWeight:600,color:T.text,margin:0}}>{main?.checklistName || "Response"}</h3>
            <span style={{fontSize:12,fontFamily:T.mono,color:T.accent}}>{autoId}</span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:4}}>
            {isAdmin && (
              <span title="Untag from order to edit" style={{padding:6,display:"inline-flex"}}>
                <Icon name="lock" size={16} color={T.textMut}/>
              </span>
            )}
            <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
          </div>
        </div>

        {loading && <p style={{fontSize:13,color:T.textMut,textAlign:"center",padding:20,animation:"pulse 1.5s infinite"}}>Loading response...</p>}
        {!loading && error && <div style={{padding:"12px 14px",background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm}}><span style={{fontSize:13,color:T.danger}}>{error}</span></div>}
        {!loading && !error && main && (
          <>
            {main.error ? <div style={{padding:"10px 12px",background:T.dangerBg,borderRadius:T.radSm,marginBottom:12}}><span style={{fontSize:12,color:T.danger}}>Response unavailable</span></div> :
              <div style={{background:T.bg,borderRadius:T.radSm,padding:"10px 12px",marginBottom:12}}>
                {(main.fields || []).map((f, i) => (
                  <div key={i} style={{padding:"6px 0",borderBottom:i<main.fields.length-1?`1px solid ${T.border}`:"none"}}>
                    <span style={{fontSize:12,color:T.textMut,display:"block",marginBottom:2}}>{f.question}</span>
                    <span style={{fontSize:14,color:T.text,fontWeight:500}}>{f.response || "—"}</span>
                  </div>
                ))}
              </div>
            }

            {upstream.length > 0 && (
              <div style={{marginTop:8,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
                  <Icon name="link" size={14} color={T.info}/>
                  <span style={{fontSize:12,fontWeight:600,color:T.info}}>Source Chain ({upstream.length})</span>
                </div>
                <div style={{display:"flex",flexDirection:"column",gap:6,paddingLeft:8,borderLeft:`2px solid ${T.infoBorder}`}}>
                  {upstream.map((item, idx) => {
                    const key = item.checklistId + "::" + item.autoId;
                    const isExp = expandedId === key;
                    return <div key={idx}>
                      <button onClick={()=>item.error?null:setExpandedId(isExp?null:key)}
                        style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`,cursor:item.error?"not-allowed":"pointer",opacity:item.error?0.5:1,textAlign:"left",width:"100%"}}>
                        <Icon name="clipboard" size={14} color={item.error?T.textMut:T.accent}/>
                        <div style={{flex:1,minWidth:0}}>
                          <span style={{fontSize:13,fontWeight:500,color:item.error?T.textMut:T.text}}>{item.checklistName} — {item.autoId}</span>
                          {item.error && <span style={{fontSize:11,color:T.textMut,marginLeft:8}}>unavailable</span>}
                        </div>
                        {!item.error && <Icon name="chevron" size={14} color={T.textMut} style={{transform:isExp?"rotate(90deg)":"rotate(0)",transition:"transform .2s"}}/>}
                      </button>
                      {isExp && !item.error && (
                        <div style={{marginTop:4,marginLeft:4,padding:"10px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                          {(item.fields || []).map((f, fi) => (
                            <div key={fi} style={{padding:"6px 0",borderBottom:fi<item.fields.length-1?`1px solid ${T.border}`:"none"}}>
                              <span style={{fontSize:11,color:T.textMut,display:"block",marginBottom:2}}>{f.question}</span>
                              <span style={{fontSize:13,color:T.text,fontWeight:500}}>{f.response || "—"}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>;
                  })}
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function OrderCard({order,checklists,orderTypes,customers,isAdmin,onClick,onDelete,completed,delay=0}){
  const type=orderTypes.find(t=>t.id===order.orderType);
  const cust=customers.find(c=>c.id===order.customerId);
  const statusIdx=ORDER_STATUSES.indexOf(order.status||"beans_not_roasted");
  const statusLabel=ORDER_STATUS_LABELS[order.status||"beans_not_roasted"]||order.status;
  const statusColor=order.status==="delivered"?T.success:order.status==="completed"?T.info:T.accent;
  const pct=((statusIdx+1)/ORDER_STATUSES.length)*100;
  const blendCount=(order.orderLines||[]).length;
  return (
    <div className="slide-in" style={{background:T.card,borderRadius:T.rad,padding:"16px 18px",border:`1px solid ${completed?T.successBorder:T.border}`,transition:"all .2s",animationDelay:`${delay}ms`,animationFillMode:"backwards"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10,cursor:"pointer"}} onClick={onClick}>
        <div style={{flex:1}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
            <span style={{fontSize:15,fontWeight:600,color:T.text}}>{order.name}</span>
            <Badge variant={order.status==="delivered"?"success":order.status==="completed"?"info":"default"} style={{fontSize:10}}>{statusLabel}</Badge>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            {type&&<Badge>{type.label}</Badge>}
            {cust&&<Badge variant="info"><Icon name="user" size={10} color={T.info}/> {cust.label}</Badge>}
            {order.productType&&(()=>{const ptc=PRODUCT_TYPE_COLORS[order.productType]||PRODUCT_TYPE_COLORS.Others;return <span style={{display:"inline-flex",alignItems:"center",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:500,letterSpacing:".02em",background:ptc.bg,color:ptc.color,border:`1px solid ${ptc.border}`,whiteSpace:"nowrap"}}>{order.productType}</span>})()}
            {order.invoiceSo&&<Badge variant="muted">{order.invoiceSo}</Badge>}
            {order.orderTypeDetail&&<Badge variant={order.orderTypeDetail==="Sample Order"?"danger":"success"}>{order.orderTypeDetail}</Badge>}
            {blendCount>0&&<Badge variant="muted">{blendCount} blend{blendCount>1?"s":""}</Badge>}
            {order.assignedTo&&<span style={{fontSize:12,color:T.textMut}}>→ {order.assignedTo}</span>}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:4}}>
          {isAdmin && <button onClick={e=>{e.stopPropagation();onDelete()}} style={{background:"none",border:"none",cursor:"pointer",padding:4,display:"flex"}}><Icon name="trash" size={16} color={T.textMut}/></button>}
          <Icon name="chevron" size={18} color={T.textMut}/>
        </div>
      </div>
      <div style={{display:"flex",alignItems:"center",gap:10,cursor:"pointer"}} onClick={onClick}>
        <div style={{flex:1,height:4,borderRadius:2,background:T.surfaceHover,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",borderRadius:2,background:statusColor,transition:"width .5s ease"}}/></div>
        <span style={{fontSize:11,fontWeight:500,color:T.textSec}}>{statusLabel}</span>
      </div>
    </div>
  );
}

// ─── New Order View ───────────────────────────────────────────

function NewOrderView({orderTypes,customers,checklists,currentUser,blends,orderStageTemplates,onAddCustomer,onCreate}){
  const [name,setName]=useState("");
  const [custId,setCustId]=useState(customers[0]?.id||"");
  const [newCust,setNewCust]=useState("");
  const [type,setType]=useState(orderTypes[0]?.id||"");
  const [assignedTo,setAssignedTo]=useState("");
  const [invoiceSo,setInvoiceSo]=useState("");
  const [orderTypeDetail,setOrderTypeDetail]=useState("Client Order");
  const [productType,setProductType]=useState("");
  const [stages,setStages]=useState([]);
  const [selCk,setSelCk]=useState([]);
  const [showNewCust,setShowNewCust]=useState(false);
  const [orderLines,setOrderLines]=useState([{blendId:"",blend:"",blendComponents:[],quantity:"",deliveryDate:""}]);

  const customerLabel = customers.find(c=>c.id===custId)?.label || "";

  useEffect(()=>{
    if(type) API.get("resolveChecklists",{order_type_id:type,customer_id:custId}).then(ids=>setSelCk(ids)).catch(()=>{});
  },[type,custId]);

  // Auto-populate stages when product type changes (only if stages list is empty)
  useEffect(()=>{
    if(productType && stages.length === 0 && orderStageTemplates){
      const tpl = orderStageTemplates[productType] || [];
      if(tpl.length>0){
        setStages(tpl.map((s,i)=>({
          id: "stage_new_"+i+"_"+Date.now(),
          name: s.name || ("Stage "+(i+1)),
          checklistId: s.checklistId || "",
          quantityField: s.quantityField || "",
          requiredQty: parseFloat(s.requiredQty)||0,
          position: i,
          taggedEntries: [],
          advanced: false,
        })));
      }
    }
  },[productType]);

  const addStage = () => setStages(p=>[...p,{id:"stage_new_"+p.length+"_"+Date.now(),name:"Stage "+(p.length+1),checklistId:"",quantityField:"",requiredQty:0,position:p.length,taggedEntries:[],advanced:false}]);
  const removeStage = (i) => setStages(p=>p.filter((_,idx)=>idx!==i).map((s,idx)=>({...s,position:idx})));
  const updateStage = (i,patch) => setStages(p=>p.map((s,idx)=>idx===i?{...s,...patch}:s));

  const toggleCk=id=>setSelCk(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  const handleAddCust=async()=>{if(!newCust.trim())return;const c=await onAddCustomer(newCust.trim());setCustId(c.id);setNewCust("");setShowNewCust(false)};
  const addLine=()=>setOrderLines(p=>[...p,{blendId:"",blend:"",blendComponents:[],quantity:"",deliveryDate:""}]);
  const removeLine=(i)=>setOrderLines(p=>p.filter((_,idx)=>idx!==i));
  const updateLine=(i,patch)=>setOrderLines(p=>p.map((l,idx)=>idx===i?{...l,...patch}:l));
  const totalQty=orderLines.reduce((sum,l)=>sum+(parseFloat(l.quantity)||0),0);

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:20}}>
      <Field label="Order Name"><Input value={name} onChange={setName} placeholder="e.g., ORD-2025-001"/></Field>
      <Field label="Invoice / SO Number"><Input value={invoiceSo} onChange={setInvoiceSo} placeholder="e.g., INV-001 or SO-123"/></Field>
      <Field label="Order Type Detail">
        <div style={{display:"flex",gap:8}}>
          <Chip label="Client Order" active={orderTypeDetail==="Client Order"} onClick={()=>setOrderTypeDetail("Client Order")}/>
          <Chip label="Sample Order" active={orderTypeDetail==="Sample Order"} onClick={()=>setOrderTypeDetail("Sample Order")}/>
        </div>
      </Field>
      <Field label="Product Type">
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          {PRODUCT_TYPES.map(pt=><Chip key={pt} label={pt} active={productType===pt} onClick={()=>setProductType(productType===pt?"":pt)}/>)}
        </div>
      </Field>
      <Field label="Customer">
        {!showNewCust?<div style={{display:"flex",flexDirection:"column",gap:8}}>
          <SearchableDropdown options={customers.map(c=>({label:c.label,value:c.id}))} value={custId} onChange={setCustId} placeholder="— Select customer —"/>
          {currentUser?.role==="admin"&&<Btn variant="ghost" small onClick={()=>setShowNewCust(true)} style={{alignSelf:"flex-start"}}><Icon name="plus" size={14} color={T.textSec}/> New Customer</Btn>}
        </div>:<div style={{display:"flex",gap:8}}>
          <Input value={newCust} onChange={setNewCust} placeholder="Customer name..." style={{flex:1}}/>
          <Btn small onClick={handleAddCust}>Add</Btn><Btn variant="ghost" small onClick={()=>setShowNewCust(false)}>Cancel</Btn>
        </div>}
      </Field>

      {/* ── Blend Lines ── */}
      <Field label="Blend Lines">
        <p style={{fontSize:12,color:T.textMut,marginBottom:10}}>Add one or more blend items for this order. Total: <b style={{color:T.accent}}>{totalQty}</b></p>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {orderLines.map((line,i)=>(
            <div key={i} style={{background:T.bg,borderRadius:T.radSm,padding:12,border:`1px solid ${T.border}`}}>
              <div style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
                <span style={{fontSize:12,color:T.textMut,fontFamily:T.mono,flexShrink:0}}>{String(i+1).padStart(2,"0")}</span>
                <div style={{flex:1}}>
                  <BlendSelector blends={blends} customerLabel={customerLabel} value={line.blendId}
                    onChange={(b)=>updateLine(i,{blendId:b?.id||"",blend:b?.name||"",blendComponents:b?.components||[]})}/>
                </div>
                {orderLines.length>1&&<button onClick={()=>removeLine(i)} style={{background:"none",border:"none",cursor:"pointer",padding:4,flexShrink:0}}><Icon name="trash" size={16} color={T.danger}/></button>}
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                <div>
                  <label style={{fontSize:11,color:T.textMut,marginBottom:4,display:"block"}}>Quantity (kgs)</label>
                  <Input value={line.quantity} onChange={v=>updateLine(i,{quantity:v})} type="number" placeholder="0"/>
                </div>
                <div>
                  <label style={{fontSize:11,color:T.textMut,marginBottom:4,display:"block"}}>Delivery Date</label>
                  <input type="date" value={line.deliveryDate} onChange={e=>updateLine(i,{deliveryDate:e.target.value})}
                    style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14,outline:"none",colorScheme:"dark"}}/>
                </div>
              </div>
            </div>
          ))}
        </div>
        <Btn variant="ghost" small onClick={addLine} style={{marginTop:6}}><Icon name="plus" size={14} color={T.textSec}/> Add Blend Line</Btn>
      </Field>

      <Field label="Assigned To"><Input value={assignedTo} onChange={setAssignedTo} placeholder="Team member name"/></Field>
      <Field label="Order Type">
        <SearchableDropdown options={orderTypes.map(ot=>({label:ot.label,value:ot.id}))} value={type} onChange={setType} placeholder="— Select order type —"/>
      </Field>
      <Field label="Order Stages">
        <p style={{fontSize:12,color:T.textMut,marginBottom:10}}>Auto-populated from the product type template. Customize as needed.</p>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {stages.map((s,i)=>{
            const stageCk = s.checklistId ? checklists.find(c=>c.id===s.checklistId) : null;
            const qtyFieldOptions = stageCk ? (stageCk.questions || []).filter(q => q.type === "number" || q.type === "text_number") : [];
            return (
            <div key={s.id} style={{background:T.bg,borderRadius:T.radSm,padding:10,border:`1px solid ${T.border}`}}>
              <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:6}}>
                <span style={{fontSize:11,fontFamily:T.mono,color:T.textMut,width:24}}>{String(i+1).padStart(2,"0")}</span>
                <Input value={s.name} onChange={v=>updateStage(i,{name:v})} placeholder="Stage name..." style={{flex:1,fontSize:13,padding:"8px 10px"}}/>
                <button onClick={()=>removeStage(i)} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={14} color={T.danger}/></button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:6,marginBottom:6}}>
                <select value={s.checklistId||""} onChange={e=>{const v=e.target.value; updateStage(i,{checklistId:v,quantityField:""});}}
                  style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
                  <option value="">— Optional checklist —</option>
                  {checklists.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                </select>
                <Input value={s.requiredQty||0} onChange={v=>updateStage(i,{requiredQty:parseFloat(v)||0})} type="number" placeholder="Req qty" style={{fontSize:12,padding:"8px 10px"}}/>
              </div>
              {stageCk && (
                <select value={s.quantityField||""} onChange={e=>updateStage(i,{quantityField:e.target.value})}
                  style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
                  <option value="">— Quantity field from checklist (optional) —</option>
                  {qtyFieldOptions.map((q,qi)=><option key={qi} value={q.text}>{q.text}</option>)}
                </select>
              )}
            </div>
          );})}
        </div>
        <Btn variant="ghost" small onClick={addStage} style={{marginTop:6}}><Icon name="plus" size={12} color={T.textSec}/> Add Stage</Btn>
      </Field>
      <Field label="Checklists">
        <p style={{fontSize:12,color:T.textMut,marginBottom:10}}>Auto-selected from assignment rules for this order type + customer combo. Adjust as needed.</p>
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {checklists.map(ck=>{
            const active=selCk.includes(ck.id);
            return <button key={ck.id} onClick={()=>toggleCk(ck.id)} style={{display:"flex",alignItems:"center",gap:12,padding:"12px 14px",borderRadius:T.radSm,border:`1px solid ${active?T.accentBorder:T.border}`,background:active?T.accentBg:"transparent",cursor:"pointer",textAlign:"left",transition:"all .2s"}}>
              <div style={{width:20,height:20,borderRadius:4,border:`2px solid ${active?T.accent:T.borderLight}`,background:active?T.accent:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{active&&<Icon name="check" size={14} color={T.bg}/>}</div>
              <div><span style={{fontSize:14,fontWeight:500,color:T.text}}>{ck.name}</span>{ck.subtitle&&<span style={{display:"block",fontSize:12,color:T.textMut}}>{ck.subtitle}</span>}<span style={{display:"block",fontSize:11,color:T.textMut,marginTop:2}}>{ck.questions.length} questions</span></div>
            </button>;
          })}
        </div>
      </Field>
      <Btn onClick={()=>onCreate({name,customerId:custId,assignedTo,orderType:type,invoiceSo,orderTypeDetail,productType,stages,createdAt:new Date().toISOString(),orderLines:orderLines.filter(l=>(l.blend||"").trim()).map(l=>({blendId:l.blendId||"",blend:l.blend,blendComponents:l.blendComponents||[],quantity:parseFloat(l.quantity)||0,deliveryDate:l.deliveryDate})),checklists:selCk.map(ckId=>({checklistId:ckId,status:"pending",completedAt:null,completedBy:null}))})} disabled={!name.trim()} style={{width:"100%",marginTop:8}}>Create Order</Btn>
    </div>
  );
}

// ─── Stages Panel (per-order stages with tag / untag / advance) ────

// ─── Blend Line Stage Section — per-ingredient tagging for a single blend line at a single stage ───
function BlendLineStageSection({
  order, stage, blendLine, lineIndex, checklists, approvedEntries, untaggedChecklists, inventoryItems, blends,
  onTagIngredient, onTagMixed, onUntag, onTaggedEntryClick, isAdmin,
}) {
  const [mixedMode, setMixedMode] = useState(false);
  const [picker, setPicker] = useState(null); // { type: "ingredient"|"mixed", componentItemId?, componentItemName? }
  const [selectedId, setSelectedId] = useState("");
  const [qty, setQty] = useState("");
  const [autoReadValue, setAutoReadValue] = useState("");
  const [mixedQty, setMixedQty] = useState("");
  const [mixedInventoryItemId, setMixedInventoryItemId] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const allTags = Array.isArray(stage.taggedEntries) ? stage.taggedEntries : [];
  const lineTags = allTags.filter(te => Number(te.blendLineIndex) === lineIndex);
  const analysis = analyzeBlendLineAtStage(blendLine, lineTags);
  const lineQty = parseFloat(blendLine.quantity) || 0;
  const progressPct = lineQty > 0 ? Math.min(100, (analysis.perIngredient.reduce((s,p)=>s+p.totalTagged,0) / lineQty * 100)) : 0;

  // Available approved entries for a given ingredient.
  // Returns { matched: [...entries], totalWithStock: N, canFilter: bool }
  // - canFilter=false when the checklist has no inventoryLink (can't tell which item each entry produced) — show all with a note.
  // - totalWithStock is the count of entries of this checklist type that have remaining quantity, regardless of ingredient match.
  const availableForIngredient = (ingredient) => {
    const reqCkId = stage.checklistId;
    const canFilter = reqCkId ? checklistHasInventoryLink(checklists, reqCkId) : false;
    const matched = [];
    const all = []; // all entries with remaining stock (before ingredient filter)
    if (reqCkId) {
      const entries = approvedEntries?.[reqCkId] || [];
      entries.forEach(e => {
        const rem = (e.remainingQuantity !== undefined) ? e.remainingQuantity : (e.remainingMasterQuantity !== undefined ? e.remainingMasterQuantity : (e.totalQuantity || 0));
        if (rem <= 0) return;
        const rec = { autoId: e.autoId || e.linkedId, checklistId: reqCkId, remaining: rem, responses: e.responses, date: e.date };
        all.push(rec);
        if (!canFilter) return;
        const te = { checklistId: reqCkId, autoId: e.autoId || e.linkedId, responses: e.responses };
        const resolved = resolveTaggedEntryItem(te, checklists, inventoryItems, approvedEntries, untaggedChecklists);
        if (!resolved) return;
        const matches = (ingredient.itemId && resolved.id && resolved.id === ingredient.itemId) ||
                        (resolved.name && ingredient.itemName && String(resolved.name).toLowerCase().trim() === String(ingredient.itemName).toLowerCase().trim());
        if (matches) matched.push(rec);
      });
    }
    (untaggedChecklists || []).forEach(u => {
      if (reqCkId && String(u.checklistId) !== String(reqCkId)) return;
      const rem = parseFloat(u.remainingQuantity);
      if (rem <= 0 || !u.autoId) return;
      const rec = { autoId: u.autoId, checklistId: u.checklistId, remaining: rem, responses: u.responses, date: u.date };
      all.push(rec);
      if (!canFilter) return;
      const te = { checklistId: u.checklistId, autoId: u.autoId, responses: u.responses };
      const resolved = resolveTaggedEntryItem(te, checklists, inventoryItems, approvedEntries, untaggedChecklists);
      if (!resolved) return;
      const matches = (ingredient.itemId && resolved.id && resolved.id === ingredient.itemId) ||
                      (resolved.name && ingredient.itemName && String(resolved.name).toLowerCase().trim() === String(ingredient.itemName).toLowerCase().trim());
      if (matches) matched.push(rec);
    });
    const dedup = (arr) => { const seen = {}; return arr.filter(o => { if(seen[o.autoId]) return false; seen[o.autoId]=true; return true; }); };
    return { matched: dedup(matched), all: dedup(all), canFilter, totalWithStock: dedup(all).length };
  };

  const readQuantityFieldValue = (picked) => {
    if (!stage.quantityField || !picked?.responses) return "";
    const rs = picked.responses;
    const target = String(stage.quantityField);
    // Array shape: [{question|questionText, response}] (preferred)
    if (Array.isArray(rs)) {
      for (let i = 0; i < rs.length; i++) {
        const r = rs[i];
        const q = r.question || r.questionText || "";
        if (String(q) === target) {
          const v = r.response;
          return (v === undefined || v === null) ? "" : String(v);
        }
      }
      return "";
    }
    // Object shape: { questionText: value } fallback
    if (typeof rs === "object") {
      const v = rs[target];
      if (v !== undefined && v !== null) return String(v);
    }
    return "";
  };

  // Candidate pre-blended inventory items for Mixed Blend picker: filter by those with ratios matching this blend line
  const mixedCandidates = (() => {
    const matchingBlends = findMatchingBlendsForLine(blendLine, blends);
    if (matchingBlends.length === 0) return [];
    // For each matching blend, find inventory items whose name matches the blend name
    const items = [];
    matchingBlends.forEach(mb => {
      (inventoryItems || []).forEach(it => {
        if (String(it.name).toLowerCase().trim() === String(mb.name).toLowerCase().trim() && it.isActive !== false) {
          items.push({ inventoryItem: it, blend: mb });
        }
      });
    });
    return items;
  })();

  const lineHeader = `${blendLine.blend || ("Blend "+(lineIndex+1))} — ${lineQty} kg`;

  return (
    <div style={{background:T.bg,borderRadius:T.radSm,border:`1px solid ${analysis.lineComplete?T.successBorder:T.border}`,padding:12}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,marginBottom:8}}>
        <div>
          <span style={{fontSize:13,fontWeight:600,color:T.text}}>{lineHeader}</span>
          {analysis.lineComplete && <Badge variant="success" style={{marginLeft:8,fontSize:10}}>Complete</Badge>}
        </div>
        {isAdmin && order.status !== "delivered" && (
          <button onClick={()=>{setMixedMode(m=>!m);setPicker(null);setError("")}}
            style={{background:mixedMode?T.accentBg:"transparent",border:`1px solid ${mixedMode?T.accent:T.border}`,borderRadius:T.radSm,padding:"4px 10px",cursor:"pointer",fontSize:11,color:mixedMode?T.accent:T.textSec}}>
            {mixedMode ? "Mixed Blend: ON" : "Tag as Mixed Blend"}
          </button>
        )}
      </div>

      {/* Progress bar */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
        <div style={{flex:1,height:4,borderRadius:2,background:T.surfaceHover,overflow:"hidden"}}>
          <div style={{width:`${progressPct}%`,height:"100%",borderRadius:2,background:analysis.lineComplete?T.success:T.accent,transition:"width .4s ease"}}/>
        </div>
        <span style={{fontSize:11,color:T.textSec,fontFamily:T.mono,whiteSpace:"nowrap"}}>{Math.round(progressPct)}%</span>
      </div>

      {/* Live ratio indicator */}
      {analysis.directTotal > 0 && !mixedMode && (
        <div style={{padding:"6px 10px",background:analysis.ratioOk?T.successBg:T.dangerBg,border:`1px solid ${analysis.ratioOk?T.successBorder:"rgba(232,93,93,0.25)"}`,borderRadius:T.radSm,marginBottom:10}}>
          <span style={{fontSize:11,color:analysis.ratioOk?T.success:T.danger,fontWeight:500}}>
            Current ratio: {analysis.actualRatioParts.map(p=>`${p.pct}% ${p.itemName}`).join(" : ")} | Required: {analysis.requiredRatioParts.map(p=>`${p.pct}% ${p.itemName}`).join(" : ")} {analysis.ratioOk?"✓":"✗"}
          </span>
        </div>
      )}

      {!mixedMode ? (
        // ── Per-ingredient tagging ──
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {analysis.perIngredient.map((ing, ii) => {
            const isPicking = picker?.type === "ingredient" && picker.componentKey === ing.key;
            const availObj = isPicking ? availableForIngredient(ing.component) : { matched: [], all: [], canFilter: true, totalWithStock: 0 };
            // When the checklist can be filtered, show only matched; otherwise fall back to "all" with a note.
            const avail = availObj.canFilter ? availObj.matched : availObj.all;
            const pickedEntry = isPicking && selectedId ? avail.find(a=>a.autoId===selectedId) : null;
            const ckName = stage.checklistId ? (checklists.find(c=>c.id===stage.checklistId)?.name || "checklist") : "checklist";
            return (
              <div key={ii} style={{background:T.surface,borderRadius:T.radSm,border:`1px solid ${ing.remaining<=0.01?T.successBorder:T.border}`,padding:10}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,flexWrap:"wrap"}}>
                  <div style={{flex:1,minWidth:180}}>
                    <span style={{fontSize:13,fontWeight:500,color:T.text}}>{ing.component.itemName || "Unknown"}</span>
                    <div style={{display:"flex",gap:10,fontSize:11,color:T.textMut,marginTop:2,flexWrap:"wrap",fontFamily:T.mono}}>
                      <span>Req: <b style={{color:T.text}}>{ing.required}kg</b></span>
                      <span>Tagged: <b style={{color:ing.remaining<=0.01?T.success:T.warning}}>{ing.totalTagged}kg</b></span>
                      <span>Rem: <b style={{color:T.text}}>{ing.remaining}kg</b></span>
                    </div>
                  </div>
                  {isAdmin && order.status !== "delivered" && ing.remaining > 0.01 && !isPicking && (
                    <Btn variant="secondary" small onClick={()=>{setPicker({type:"ingredient",componentKey:ing.key,componentItemId:ing.component.itemId,componentItemName:ing.component.itemName});setSelectedId("");setQty("");setAutoReadValue("");setError("")}}>
                      <Icon name="plus" size={12} color={T.text}/> Tag
                    </Btn>
                  )}
                </div>
                {/* Show individual tagged entries */}
                {ing.entries.length > 0 && (
                  <div style={{marginTop:8,display:"flex",flexDirection:"column",gap:4}}>
                    {ing.entries.map((te, ei) => (
                      <div key={ei} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 8px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                        <Icon name="clipboard" size={12} color={T.accent}/>
                        <button onClick={()=>onTaggedEntryClick && onTaggedEntryClick(te.checklistId, te.autoId||te.responseId)}
                          style={{background:"none",border:"none",padding:0,cursor:"pointer",fontSize:12,fontFamily:T.mono,color:T.accent,fontWeight:600,textDecoration:"underline",textDecorationStyle:"dotted"}}>
                          {te.autoId || te.responseId}
                        </button>
                        <span style={{flex:1,fontSize:11,color:T.textMut,marginLeft:4}}>{te.qty}kg</span>
                        {isAdmin && <button onClick={()=>onUntag(stage.id, { responseId: te.responseId || te.autoId, blendLineIndex: lineIndex, componentItemId: te.componentItemId || "" })}
                          style={{background:"none",border:"none",cursor:"pointer",padding:2}} title="Untag"><Icon name="x" size={12} color={T.danger}/></button>}
                      </div>
                    ))}
                  </div>
                )}
                {/* Picker for this ingredient */}
                {isPicking && (
                  <div style={{marginTop:8,padding:8,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
                    {!availObj.canFilter && availObj.all.length > 0 && (
                      <div style={{padding:"6px 10px",background:T.warningBg,border:`1px solid ${T.warningBorder}`,borderRadius:T.radSm,marginBottom:8}}>
                        <span style={{fontSize:11,color:T.warning}}>Cannot filter by ingredient — no inventory link configured on "{ckName}". Showing all entries; pick carefully.</span>
                      </div>
                    )}
                    {avail.length === 0 ? (
                      <p style={{fontSize:12,color:T.textMut}}>
                        No approved {ckName} entries found for <b>{ing.component.itemName}</b>.
                        {availObj.totalWithStock > 0 && <> Available entries: {availObj.totalWithStock} total but none match this ingredient.</>}
                      </p>
                    ) :
                      <>
                        <select value={selectedId} onChange={e=>{
                          const v = e.target.value;
                          setSelectedId(v);
                          const picked = avail.find(a=>a.autoId===v);
                          if (picked) {
                            const autoVal = readQuantityFieldValue(picked);
                            setAutoReadValue(autoVal);
                            const autoNum = parseFloat(autoVal);
                            let prefill;
                            if (!isNaN(autoNum) && autoNum > 0) prefill = Math.min(picked.remaining, autoNum, ing.remaining);
                            else prefill = Math.min(picked.remaining, ing.remaining);
                            setQty(String(Math.round(prefill*100)/100));
                          } else { setAutoReadValue(""); setQty(""); }
                        }} style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12,marginBottom:6}}>
                          <option value="">— Select approved entry —</option>
                          {avail.map(a=><option key={a.autoId} value={a.autoId}>{a.autoId} — {a.remaining}kg available</option>)}
                        </select>
                        {pickedEntry && stage.quantityField && autoReadValue !== "" && (
                          <div style={{padding:"4px 8px",background:T.infoBg,border:`1px solid ${T.infoBorder}`,borderRadius:T.radSm,marginBottom:6}}>
                            <span style={{fontSize:11,color:T.info}}>Quantity from this checklist: <b>{autoReadValue}kg</b></span>
                          </div>
                        )}
                        <label style={{fontSize:11,color:T.textMut,display:"block",marginBottom:4}}>Qty to tag (can adjust down, max is the smaller of remaining or available):</label>
                        <div style={{display:"flex",gap:6}}>
                          <Input value={qty} onChange={v=>{
                            const num = parseFloat(v);
                            const maxAllowed = pickedEntry ? Math.min(pickedEntry.remaining, ing.remaining) : ing.remaining;
                            if (!isNaN(num) && num > maxAllowed) {
                              setError(`Max is ${Math.round(maxAllowed*100)/100}kg (can't tag more than available or needed)`);
                              setQty(String(Math.round(maxAllowed*100)/100));
                            } else { setError(""); setQty(v); }
                          }} type="number" placeholder="Qty" style={{flex:1,fontSize:12,padding:"6px 10px"}}/>
                          <Btn small disabled={busy||!selectedId||!(parseFloat(qty)>0)} onClick={async()=>{
                            const picked = avail.find(a=>a.autoId===selectedId);
                            if (!picked) return;
                            setBusy(true); setError("");
                            try {
                              await onTagIngredient({
                                stageId: stage.id, autoId: picked.autoId, sourceChecklistId: picked.checklistId,
                                quantity: parseFloat(qty), quantityFieldValue: autoReadValue,
                                blendLineIndex: lineIndex,
                                componentItemId: ing.component.itemId,
                                componentItemName: ing.component.itemName,
                              });
                              setPicker(null); setSelectedId(""); setQty(""); setAutoReadValue("");
                            } catch(e) { setError(e.message || "Tag failed"); }
                            setBusy(false);
                          }}>Tag</Btn>
                          <Btn variant="ghost" small onClick={()=>{setPicker(null);setSelectedId("");setQty("");setAutoReadValue("");setError("")}}>Cancel</Btn>
                        </div>
                        {error && <div style={{marginTop:6,fontSize:11,color:T.danger}}>{error}</div>}
                      </>
                    }
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        // ── Mixed Blend tagging ──
        <div style={{padding:10,background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
          <p style={{fontSize:11,color:T.textMut,marginBottom:8}}>Tag a pre-blended stock item. Only items with exactly matching ratios are shown.</p>
          {mixedCandidates.length === 0 ? (
            <p style={{fontSize:12,color:T.danger}}>No pre-blended stock items match this blend's ratios. Create a Blend recipe + inventory item with matching composition.</p>
          ) : (
            <>
              <select value={mixedInventoryItemId} onChange={e=>setMixedInventoryItemId(e.target.value)}
                style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13,marginBottom:8}}>
                <option value="">— Select pre-blended stock —</option>
                {mixedCandidates.map((mc,idx)=><option key={idx} value={mc.inventoryItem.id}>{mc.inventoryItem.name} (stock: {mc.inventoryItem.currentStock}{mc.inventoryItem.unit||"kg"})</option>)}
              </select>
              <label style={{fontSize:11,color:T.textMut,display:"block",marginBottom:4}}>Quantity to tag (kg):</label>
              <div style={{display:"flex",gap:6}}>
                <Input value={mixedQty} onChange={setMixedQty} type="number" placeholder="Qty" style={{flex:1,fontSize:13,padding:"8px 10px"}}/>
                <Btn small disabled={busy || !mixedInventoryItemId || !(parseFloat(mixedQty) > 0)} onClick={async()=>{
                  const chosen = mixedCandidates.find(mc => mc.inventoryItem.id === mixedInventoryItemId);
                  if (!chosen) return;
                  setBusy(true); setError("");
                  try {
                    await onTagMixed({
                      stageId: stage.id,
                      blendLineIndex: lineIndex,
                      mixedInventoryItemId: chosen.inventoryItem.id,
                      mixedInventoryItemName: chosen.inventoryItem.name,
                      mixedBlendId: chosen.blend.id,
                      quantity: parseFloat(mixedQty),
                    });
                    setMixedQty(""); setMixedInventoryItemId(""); setMixedMode(false);
                  } catch(e) { setError(e.message || "Tag failed"); }
                  setBusy(false);
                }}>Tag Mixed</Btn>
              </div>
              {error && <div style={{marginTop:6,fontSize:11,color:T.danger}}>{error}</div>}
            </>
          )}
          {/* Show existing mixed tags */}
          {analysis.mixedTags.length > 0 && (
            <div style={{marginTop:10,display:"flex",flexDirection:"column",gap:4}}>
              <span style={{fontSize:11,fontWeight:600,color:T.textSec}}>Tagged mixed stock:</span>
              {analysis.mixedTags.map((mt, mi) => (
                <div key={mi} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 8px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                  <Icon name="layers" size={12} color={T.accent}/>
                  <span style={{flex:1,fontSize:12,color:T.text,fontWeight:500}}>{mt.mixedItemName}</span>
                  <span style={{fontSize:11,color:T.textMut}}>{mt.qty}kg</span>
                  {isAdmin && <button onClick={()=>onUntag(stage.id, { isMixedBlend: true, mixedItemId: mt.mixedItemId, blendLineIndex: lineIndex })}
                    style={{background:"none",border:"none",cursor:"pointer",padding:2}} title="Untag"><Icon name="x" size={12} color={T.danger}/></button>}
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function StagesPanel({ order, checklists, approvedEntries, untaggedChecklists, inventoryItems, blends, onTag, onTagMixed, onUntag, onDeliver, isAdmin, onTaggedEntryClick }) {
  const [tagPicking, setTagPicking] = useState(null); // stageId
  const [tagQty, setTagQty] = useState("");
  const [tagSelectedId, setTagSelectedId] = useState("");
  const [autoReadValue, setAutoReadValue] = useState(""); // value read from configured quantityField
  const [busyAction, setBusyAction] = useState(false);
  const [pendingTagWarning, setPendingTagWarning] = useState(null); // { message, onConfirm }
  const stages = Array.isArray(order.stages) ? order.stages : [];

  if (stages.length === 0) return null;

  // Build list of available entries (approved + untagged) with remaining quantity, filtered by stage.checklistId
  const availableForStage = (stage) => {
    const out = [];
    const reqCkId = stage.checklistId;
    if (reqCkId) {
      const entries = approvedEntries?.[reqCkId] || [];
      entries.forEach(e => {
        const rem = (e.remainingQuantity !== undefined) ? e.remainingQuantity : (e.remainingMasterQuantity !== undefined ? e.remainingMasterQuantity : (e.totalQuantity || 0));
        if (rem > 0) out.push({ autoId: e.autoId || e.linkedId, checklistId: reqCkId, remaining: rem, source: "approved", orderId: e.orderId, date: e.date, responses: e.responses });
      });
    }
    (untaggedChecklists || []).forEach(u => {
      if (reqCkId && String(u.checklistId) !== String(reqCkId)) return;
      const rem = parseFloat(u.remainingQuantity);
      if (rem > 0 && u.autoId) out.push({ autoId: u.autoId, checklistId: u.checklistId, remaining: rem, source: "untagged", orderId: u.taggedOrderId||"", date: u.date, responses: u.responses });
    });
    // Dedup by autoId
    const seen = {};
    return out.filter(o => { if(seen[o.autoId]) return false; seen[o.autoId]=true; return true; });
  };

  // Look up the configured quantityField's value in a selected entry's responses.
  // approvedEntries responses: [{question, response}]. Untagged responses: [{questionIndex, questionText, response}].
  const readQuantityFieldValue = (stage, pickedEntry) => {
    if (!stage.quantityField || !pickedEntry?.responses) return "";
    const rs = pickedEntry.responses;
    for (let i = 0; i < rs.length; i++) {
      const r = rs[i];
      const q = r.question || r.questionText || "";
      if (String(q) === String(stage.quantityField)) return String(r.response || "");
    }
    return "";
  };

  const stageTaggedTotal = (stage) => (stage.taggedEntries || []).reduce((s,t)=>s+(parseFloat(t.qty)||0),0);
  const stageSatisfied = (stage) => isStageComplete(stage, order);
  const blendOrder = isBlendOrder(order);
  const blendLines = Array.isArray(order.orderLines) ? order.orderLines.filter(l => Array.isArray(l.blendComponents) && l.blendComponents.length > 0) : [];
  const blendLineIndicesWithComponents = []; // keep original indices for tagging
  if (Array.isArray(order.orderLines)) {
    order.orderLines.forEach((l, idx) => { if (Array.isArray(l.blendComponents) && l.blendComponents.length > 0) blendLineIndicesWithComponents.push(idx); });
  }

  const allStagesSatisfied = stages.every(stageSatisfied);

  return (
    <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.border}`,marginBottom:24}}>
      <Section icon="layers">Order Stages</Section>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {stages.map((stage, si) => {
          const ck = stage.checklistId ? checklists.find(c=>c.id===stage.checklistId) : null;
          const totalTagged = stageTaggedTotal(stage);
          const req = parseFloat(stage.requiredQty) || 0;
          const satisfied = stageSatisfied(stage);
          const isPicking = tagPicking === stage.id;
          const avail = isPicking ? availableForStage(stage) : [];
          const pickedEntry = isPicking && tagSelectedId ? avail.find(a=>a.autoId===tagSelectedId) : null;

          return <div key={stage.id} style={{background:T.bg,borderRadius:T.radSm,border:`1px solid ${satisfied?T.successBorder:T.border}`,padding:12}}>
            {/* Stage header — always visible */}
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  <span style={{fontSize:11,fontFamily:T.mono,color:T.textMut}}>{String(si+1).padStart(2,"0")}</span>
                  <span style={{fontSize:14,fontWeight:600,color:T.text}}>{stage.name}</span>
                  {satisfied ? <Badge variant="success" style={{fontSize:10}}>Ready</Badge> : <Badge variant="muted" style={{fontSize:10}}>{req>0 ? `${totalTagged}/${req} kg` : "Pending"}</Badge>}
                </div>
                {ck && <span style={{fontSize:11,color:T.textMut,display:"block",marginTop:2}}>Requires: {ck.name}{stage.quantityField?` · qty from "${stage.quantityField}"`:""}</span>}
              </div>
            </div>

            {/* Progress bar (non-blend orders only — blend orders have per-line progress) */}
            {!blendOrder && req > 0 && (() => {
              const overTagged = totalTagged > req;
              const barColor = overTagged ? T.warning : (satisfied ? T.success : T.accent);
              return <>
                <div style={{display:"flex",alignItems:"center",gap:8,marginTop:8}}>
                  <div style={{flex:1,height:4,borderRadius:2,background:T.surfaceHover,overflow:"hidden"}}>
                    <div style={{width:`${Math.min(100,(totalTagged/req)*100)}%`,height:"100%",borderRadius:2,background:barColor,transition:"width .4s ease"}}/>
                  </div>
                  <span style={{fontSize:11,color:overTagged?T.warning:T.textSec,fontFamily:T.mono,whiteSpace:"nowrap",fontWeight:overTagged?600:400}}>Tagged: {totalTagged} / Required: {req} kg</span>
                </div>
                {overTagged && (
                  <div style={{marginTop:6,padding:"6px 10px",background:T.warningBg,border:`1px solid ${T.warningBorder}`,borderRadius:T.radSm}}>
                    <span style={{fontSize:11,color:T.warning,fontWeight:500}}>Warning: Total tagged ({totalTagged}kg) exceeds required ({req}kg). Please verify.</span>
                  </div>
                )}
              </>;
            })()}

            {blendOrder ? (
              // ── Blend-order stage body: one section per blend line ──
              <div style={{marginTop:10,display:"flex",flexDirection:"column",gap:10}}>
                {blendLineIndicesWithComponents.map(li => (
                  <BlendLineStageSection key={li}
                    order={order} stage={stage} blendLine={order.orderLines[li]} lineIndex={li}
                    checklists={checklists} approvedEntries={approvedEntries} untaggedChecklists={untaggedChecklists} inventoryItems={inventoryItems} blends={blends}
                    isAdmin={isAdmin}
                    onTagIngredient={async(payload)=>{ await onTag(payload.stageId, payload.autoId, payload.sourceChecklistId, payload.quantity, payload.quantityFieldValue, { blendLineIndex: payload.blendLineIndex, componentItemId: payload.componentItemId, componentItemName: payload.componentItemName }); }}
                    onTagMixed={async(payload)=>{ await onTagMixed(payload); }}
                    onUntag={async(stageId, spec)=>{
                      if (!confirm("Untag this from the stage?")) return;
                      await onUntag(stageId, spec);
                    }}
                    onTaggedEntryClick={onTaggedEntryClick}
                  />
                ))}
                {/* Show any legacy/uncategorized tagged entries (no blendLineIndex) so admin can untag them */}
                {(stage.taggedEntries || []).filter(te => te.blendLineIndex === undefined || te.blendLineIndex === null).length > 0 && (
                  <div style={{padding:10,background:T.surface,borderRadius:T.radSm,border:`1px dashed ${T.border}`}}>
                    <span style={{fontSize:11,color:T.textMut,fontWeight:600,display:"block",marginBottom:6}}>Uncategorized tags (pre-blend-update entries — untag and re-tag to assign ingredients)</span>
                    {stage.taggedEntries.filter(te => te.blendLineIndex === undefined || te.blendLineIndex === null).map((te, ti) => {
                      const teAutoId = te.autoId || te.responseId;
                      const clickable = !!(teAutoId && te.checklistId && typeof onTaggedEntryClick === "function");
                      return <div key={ti} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 8px",marginBottom:4,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                        <Icon name="clipboard" size={12} color={T.textMut}/>
                        <div style={{flex:1,minWidth:0}}>
                          {clickable ? (
                            <button onClick={()=>onTaggedEntryClick(te.checklistId, teAutoId)} style={{background:"none",border:"none",padding:0,cursor:"pointer",fontSize:12,fontFamily:T.mono,color:T.accent,textDecoration:"underline",textDecorationStyle:"dotted"}}>{teAutoId}</button>
                          ) : (<span style={{fontSize:12,fontFamily:T.mono,color:T.accent}}>{teAutoId}</span>)}
                        </div>
                        <span style={{fontSize:11,color:T.textMut}}>{te.qty}kg</span>
                        {isAdmin && <button onClick={()=>onUntag(stage.id, te.responseId || te.autoId)} style={{background:"none",border:"none",cursor:"pointer",padding:2}} title="Untag"><Icon name="x" size={12} color={T.danger}/></button>}
                      </div>;
                    })}
                  </div>
                )}
              </div>
            ) : (
              <>
            {/* Tagged entries — ALWAYS visible (non-blend orders) */}
            <div style={{marginTop:10}}>
              {(stage.taggedEntries || []).length === 0 ? (
                <p style={{fontSize:12,color:T.textMut,margin:0}}>No checklists tagged yet.</p>
              ) : (
                <div style={{display:"flex",flexDirection:"column",gap:6}}>
                  {stage.taggedEntries.map((te, ti) => {
                    const teCk = checklists.find(c=>c.id===te.checklistId);
                    const teAutoId = te.autoId || te.responseId;
                    const clickable = !!(teAutoId && te.checklistId && typeof onTaggedEntryClick === "function");
                    return <div key={ti} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                      <Icon name="clipboard" size={14} color={T.accent}/>
                      <div style={{flex:1,minWidth:0}}>
                        {clickable ? (
                          <button onClick={()=>onTaggedEntryClick(te.checklistId, teAutoId)}
                            style={{background:"none",border:"none",padding:0,cursor:"pointer",fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600,textDecoration:"underline",textDecorationStyle:"dotted"}}
                            title="View response details and source chain">{teAutoId}</button>
                        ) : (
                          <span style={{fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600}}>{teAutoId}</span>
                        )}
                        <span style={{fontSize:11,color:T.textMut,marginLeft:6}}>{teCk?.name||""}</span>
                      </div>
                      <span style={{fontSize:12,color:T.textSec,fontWeight:500}}>{te.qty} kg</span>
                      {isAdmin && <button onClick={()=>onUntag(stage.id, te.responseId || te.autoId)} style={{background:"none",border:"none",cursor:"pointer",padding:4}} title="Untag"><Icon name="x" size={14} color={T.danger}/></button>}
                    </div>;
                  })}
                </div>
              )}
            </div>

            {/* Tag picker (non-blend orders only) */}
            <div style={{marginTop:10}}>
              {!isPicking ? (
                isAdmin && order.status !== "delivered" && (
                  <Btn variant="secondary" small onClick={()=>{setTagPicking(stage.id);setTagSelectedId("");setTagQty("");setAutoReadValue("")}}><Icon name="plus" size={12} color={T.text}/> Tag Checklist to this Stage</Btn>
                )
              ) : (
                <div style={{padding:10,background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
                  {avail.length === 0 ? <p style={{fontSize:12,color:T.textMut}}>No available checklists of this type with remaining quantity.</p> :
                    <>
                      <select value={tagSelectedId} onChange={e=>{
                        const v = e.target.value;
                        setTagSelectedId(v);
                        const picked = avail.find(a=>a.autoId===v);
                        if (picked) {
                          // Try to auto-read the qty from the configured quantityField
                          const autoVal = readQuantityFieldValue(stage, picked);
                          setAutoReadValue(autoVal);
                          const autoNum = parseFloat(autoVal);
                          let prefill;
                          if (!isNaN(autoNum) && autoNum > 0) {
                            prefill = Math.min(picked.remaining, autoNum);
                          } else {
                            const stillNeeded = req > 0 ? Math.max(0, req - totalTagged) : picked.remaining;
                            prefill = Math.min(picked.remaining, stillNeeded || picked.remaining);
                          }
                          setTagQty(String(prefill));
                        } else {
                          setAutoReadValue("");
                          setTagQty("");
                        }
                      }} style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13,marginBottom:8}}>
                        <option value="">— Select approved entry —</option>
                        {avail.map(a=><option key={a.autoId} value={a.autoId}>{a.autoId} — {a.remaining} kg available</option>)}
                      </select>
                      {pickedEntry && stage.quantityField && autoReadValue !== "" && (
                        <div style={{padding:"6px 10px",background:T.infoBg,border:`1px solid ${T.infoBorder}`,borderRadius:T.radSm,marginBottom:8}}>
                          <span style={{fontSize:11,color:T.info}}>Quantity from this checklist ({stage.quantityField}): <b>{autoReadValue} kg</b></span>
                        </div>
                      )}
                      <label style={{fontSize:11,color:T.textMut,display:"block",marginBottom:4}}>Qty to tag (editable):</label>
                      <div style={{display:"flex",gap:6}}>
                        <Input value={tagQty} onChange={setTagQty} type="number" placeholder="Qty" style={{flex:1,fontSize:13,padding:"8px 10px"}}/>
                        <Btn small disabled={busyAction||!tagSelectedId||!(parseFloat(tagQty)>0)} onClick={async()=>{
                          const picked = avail.find(a=>a.autoId===tagSelectedId);
                          if(!picked) return;
                          const qtyToTag = parseFloat(tagQty);

                          // ── Blend composition pre-check (Fix 5) ──
                          // Resolve what item this checklist is for (via its inventory input item in the response).
                          const te = { checklistId: picked.checklistId, autoId: picked.autoId, responseId: picked.autoId };
                          const resolvedItem = resolveTaggedEntryItem(te, checklists, inventoryItems, approvedEntries, untaggedChecklists);
                          const blendLines = Array.isArray(order.orderLines) ? order.orderLines.filter(l => Array.isArray(l.blendComponents) && l.blendComponents.length > 0) : [];

                          // Proceed with actual tag — extracted so both "tag anyway" and no-warning paths share it
                          const proceed = async () => {
                            setBusyAction(true);
                            try {
                              await onTag(stage.id, picked.autoId, picked.checklistId, qtyToTag, autoReadValue);
                              setTagPicking(null); setTagSelectedId(""); setTagQty(""); setAutoReadValue("");
                            } catch(e) { /* handled upstream */ }
                            setBusyAction(false);
                          };

                          if (blendLines.length > 0 && resolvedItem) {
                            // Build expected-per-item map across the whole order (from blend lines)
                            const expected = {};
                            blendLines.forEach(line => {
                              const q = parseFloat(line.quantity) || 0;
                              (line.blendComponents || []).forEach(c => {
                                const k = blendItemKey(c.itemId, c.itemName);
                                if (!expected[k]) expected[k] = { itemId: c.itemId, itemName: c.itemName, qty: 0 };
                                expected[k].qty += (parseFloat(c.percentage) || 0) / 100 * q;
                              });
                            });
                            const itemKey = blendItemKey(resolvedItem.id, resolvedItem.name);
                            const expRow = expected[itemKey];
                            if (!expRow) {
                              const componentNames = Object.values(expected).map(e => e.itemName).join(", ");
                              setPendingTagWarning({
                                message: `This checklist uses "${resolvedItem.name}" which is not a required ingredient for this order's blend (${componentNames}). Tag anyway?`,
                                onConfirm: proceed,
                              });
                              return;
                            }
                            // Is: already-tagged-for-this-item + new qty
                            const taggedMap = computeTaggedByItem(order, checklists, inventoryItems, approvedEntries, untaggedChecklists);
                            const alreadyTagged = parseFloat(taggedMap[itemKey]) || 0;
                            const projected = alreadyTagged + qtyToTag;
                            const limit = expRow.qty * 1.1; // 10% over
                            if (projected > limit + 0.0001) {
                              setPendingTagWarning({
                                message: `Adding this will tag ${Math.round(projected*100)/100}kg total, exceeding the required ${Math.round(expRow.qty*100)/100}kg for "${expRow.itemName}". Tag anyway?`,
                                onConfirm: proceed,
                              });
                              return;
                            }
                          }
                          await proceed();
                        }}>Tag</Btn>
                        <Btn variant="ghost" small onClick={()=>{setTagPicking(null);setTagSelectedId("");setTagQty("");setAutoReadValue("")}}>Cancel</Btn>
                      </div>
                    </>
                  }
                </div>
              )}
            </div>
              </>
            )}
          </div>;
        })}
      </div>
      {isAdmin && order.status !== "delivered" && <Btn variant="success" small onClick={onDeliver} disabled={!allStagesSatisfied}
        title={!allStagesSatisfied ? "All stages must have required checklists tagged with sufficient quantity" : "Mark delivered — will deduct inventory"}
        style={{marginTop:12,width:"100%",opacity:allStagesSatisfied?1:0.5}}>
        <Icon name="check" size={14} color={T.success}/> Mark Delivered{!allStagesSatisfied?" (stages incomplete)":""}
      </Btn>}

      {pendingTagWarning && (
        <div onClick={()=>setPendingTagWarning(null)} style={{position:"fixed",inset:0,zIndex:1050,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
          <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,border:`1px solid ${T.warningBorder}`,maxWidth:460,width:"100%",padding:20}}>
            <h3 style={{fontSize:15,fontWeight:600,color:T.warning,margin:"0 0 8px",display:"flex",alignItems:"center",gap:8}}>
              <Icon name="clipboard" size={18} color={T.warning}/> Blend ingredient warning
            </h3>
            <p style={{fontSize:13,color:T.textSec,marginBottom:16}}>{pendingTagWarning.message}</p>
            <div style={{display:"flex",gap:8}}>
              <Btn variant="secondary" onClick={()=>setPendingTagWarning(null)} style={{flex:1}}>Cancel</Btn>
              <Btn onClick={async()=>{ const fn = pendingTagWarning.onConfirm; setPendingTagWarning(null); if (fn) await fn(); }} style={{flex:1}}>Tag Anyway</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Stage Editor (inline editor on order detail, admin only) ───
function StageEditor({ order, checklists, onSave, onCancel }) {
  // Clone stages to draft — preserve id + taggedEntries when editing. Add a local `_error` field (not persisted)
  const [draft, setDraft] = useState(() => (Array.isArray(order.stages) ? order.stages : []).map(s => ({
    id: s.id || ("stage_" + Date.now() + "_" + Math.random().toString(36).slice(2,7)),
    name: s.name || "",
    checklistId: s.checklistId || "",
    quantityField: s.quantityField || "",
    requiredQty: parseFloat(s.requiredQty) || 0,
    position: s.position,
    taggedEntries: Array.isArray(s.taggedEntries) ? s.taggedEntries : [],
    advanced: s.advanced === true,
    _error: "",
  })));
  const [saving, setSaving] = useState(false);
  const [validationError, setValidationError] = useState("");

  const updateStage = (i, patch) => setDraft(p => p.map((s, idx) => idx === i ? { ...s, ...patch, _error: "" } : s));
  const addStage = () => setDraft(p => [...p, { id: "stage_" + Date.now() + "_" + Math.random().toString(36).slice(2,7), name: "Stage " + (p.length + 1), checklistId: "", quantityField: "", requiredQty: 0, position: p.length, taggedEntries: [], advanced: false, _error: "" }]);
  const deleteStage = (i) => {
    const stage = draft[i];
    if (stage.taggedEntries && stage.taggedEntries.length > 0) {
      setDraft(p => p.map((s, idx) => idx === i ? { ...s, _error: "Cannot delete — this stage has " + stage.taggedEntries.length + " tagged checklist(s). Untag them first." } : s));
      return;
    }
    setDraft(p => p.filter((_, idx) => idx !== i).map((s, idx) => ({ ...s, position: idx })));
  };
  const moveStage = (i, dir) => {
    const ni = i + dir;
    if (ni < 0 || ni >= draft.length) return;
    setDraft(p => {
      const a = [...p];
      const tmp = a[i]; a[i] = a[ni]; a[ni] = tmp;
      return a.map((s, idx) => ({ ...s, position: idx, _error: "" }));
    });
  };

  const handleSave = async () => {
    setValidationError("");
    // Validation: names must be non-empty
    for (let i = 0; i < draft.length; i++) {
      if (!String(draft[i].name || "").trim()) { setValidationError("Stage " + (i+1) + " is missing a name."); return; }
    }
    setSaving(true);
    try {
      // Strip the local _error field before saving
      const clean = draft.map(s => ({
        id: s.id, name: s.name, checklistId: s.checklistId || "",
        quantityField: s.quantityField || "",
        requiredQty: parseFloat(s.requiredQty) || 0,
        position: s.position, taggedEntries: s.taggedEntries || [], advanced: s.advanced === true,
      }));
      await onSave(clean);
    } catch (e) { setValidationError(e.message || "Save failed"); }
    setSaving(false);
  };

  return (
    <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.accentBorder}`,marginBottom:24}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <Section icon="edit">Edit Stages</Section>
      </div>
      {isBlendOrder(order) && (
        <div style={{padding:"8px 12px",background:T.infoBg,border:`1px solid ${T.infoBorder}`,borderRadius:T.radSm,marginBottom:10}}>
          <span style={{fontSize:12,color:T.info}}>This order has blend lines — required quantities per stage are auto-derived from the blend recipe. No manual quantity input needed.</span>
        </div>
      )}
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {draft.length === 0 && <p style={{fontSize:12,color:T.textMut}}>No stages. Add one below.</p>}
        {draft.map((s, i) => {
          const stageCk = s.checklistId ? checklists.find(c=>c.id===s.checklistId) : null;
          const qtyFieldOptions = stageCk ? (stageCk.questions || []).filter(q => q.type === "number" || q.type === "text_number") : [];
          const hideQty = isBlendOrder(order);
          return (
          <div key={s.id} style={{background:T.bg,borderRadius:T.radSm,padding:10,border:`1px solid ${s._error?"rgba(232,93,93,0.4)":T.border}`}}>
            <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:6}}>
              <span style={{fontSize:11,fontFamily:T.mono,color:T.textMut,width:24}}>{String(i+1).padStart(2,"0")}</span>
              <Input value={s.name} onChange={v=>updateStage(i,{name:v})} placeholder="Stage name..." style={{flex:1,fontSize:13,padding:"8px 10px"}}/>
              <button onClick={()=>moveStage(i,-1)} disabled={i===0} style={{background:"none",border:"none",cursor:i===0?"not-allowed":"pointer",padding:4,opacity:i===0?0.3:1}} title="Move up"><span style={{fontSize:14,color:T.textSec}}>↑</span></button>
              <button onClick={()=>moveStage(i,1)} disabled={i===draft.length-1} style={{background:"none",border:"none",cursor:i===draft.length-1?"not-allowed":"pointer",padding:4,opacity:i===draft.length-1?0.3:1}} title="Move down"><span style={{fontSize:14,color:T.textSec}}>↓</span></button>
              <button onClick={()=>deleteStage(i)} style={{background:"none",border:"none",cursor:"pointer",padding:4}} title="Delete"><Icon name="trash" size={14} color={T.danger}/></button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:hideQty?"1fr":"2fr 1fr",gap:6,marginBottom:6}}>
              <select value={s.checklistId||""} onChange={e=>{const v=e.target.value; updateStage(i,{checklistId:v,quantityField:""});}}
                style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
                <option value="">— Optional checklist —</option>
                {checklists.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
              {!hideQty && <Input value={s.requiredQty||0} onChange={v=>updateStage(i,{requiredQty:parseFloat(v)||0})} type="number" placeholder="Req qty" style={{fontSize:12,padding:"8px 10px"}}/>}
            </div>
            {stageCk && (
              <select value={s.quantityField||""} onChange={e=>updateStage(i,{quantityField:e.target.value})}
                style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
                <option value="">— Quantity field from checklist (optional) —</option>
                {qtyFieldOptions.map((q,qi)=><option key={qi} value={q.text}>{q.text}</option>)}
              </select>
            )}
            {s.taggedEntries && s.taggedEntries.length > 0 && <p style={{fontSize:11,color:T.textMut,marginTop:6}}>{s.taggedEntries.length} tagged checklist(s) preserved</p>}
            {s._error && <p style={{fontSize:11,color:T.danger,marginTop:6}}>{s._error}</p>}
          </div>
        );})}
      </div>
      <Btn variant="ghost" small onClick={addStage} style={{marginTop:8}}><Icon name="plus" size={12} color={T.textSec}/> Add Stage</Btn>
      {validationError && <div style={{marginTop:10,padding:"8px 12px",background:T.dangerBg,border:"1px solid rgba(232,93,93,0.2)",borderRadius:T.radSm}}><span style={{fontSize:12,color:T.danger}}>{validationError}</span></div>}
      <div style={{display:"flex",gap:8,marginTop:12}}>
        <Btn variant="secondary" onClick={onCancel} disabled={saving} style={{flex:1}}>Cancel</Btn>
        <Btn onClick={handleSave} disabled={saving} style={{flex:1}}>{saving?"Saving...":"Save Stages"}</Btn>
      </div>
    </div>
  );
}

// ─── Tagged Checklists Grouped (all tagged entries across stages, grouped by checklist type) ─
function TaggedChecklistsGrouped({ order, checklists, approvedEntries, untaggedChecklists, inventoryItems, onEntryClick }) {
  const stages = Array.isArray(order.stages) ? order.stages : [];
  // Flatten all tagged entries with their stage name
  const flat = [];
  stages.forEach(s => {
    (s.taggedEntries || []).forEach(te => {
      flat.push({
        ...te,
        stageName: s.name,
        stageId: s.id,
      });
    });
  });
  if (flat.length === 0) return null;

  // Helper to find the submitted date for a tagged entry (via approvedEntries or untagged cache)
  const findEntryMeta = (te) => {
    const aeArr = approvedEntries?.[te.checklistId] || [];
    const aeMatch = aeArr.find(e => (e.autoId||e.linkedId) === te.autoId);
    if (aeMatch) return { date: aeMatch.date || "", person: aeMatch.person || "", responses: aeMatch.responses };
    const utMatch = (untaggedChecklists || []).find(u => u.autoId === te.autoId);
    if (utMatch) return { date: utMatch.date || "", person: utMatch.person || "", responses: utMatch.responses };
    return { date: te.taggedAt ? String(te.taggedAt).split("T")[0] : "", person: te.taggedBy || "", responses: [] };
  };

  // Group by checklistId
  const groups = {};
  flat.forEach(te => {
    const key = te.checklistId;
    if (!groups[key]) groups[key] = [];
    groups[key].push(te);
  });

  // ── Blend composition check — use the shared helper which resolves items from inventoryLink (no duplicates)
  const blendAnalysis = computeOrderBlendAnalysis(order, checklists, inventoryItems, approvedEntries, untaggedChecklists);

  const statusColor = (s) => {
    if (s === "ok") return T.success;
    if (s === "missing") return T.danger;
    if (s === "under") return T.danger;
    if (s === "over") return T.warning;
    if (s === "unexpected") return T.warning;
    return T.textSec;
  };
  const statusLabel = (s) => {
    if (s === "ok") return "✓";
    if (s === "missing") return "Missing";
    if (s === "under") return "Under";
    if (s === "over") return "Over";
    if (s === "unexpected") return "Unexpected";
    return "";
  };

  return (
    <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.border}`,marginBottom:24}}>
      <Section icon="clipboard" count={flat.length}>Tagged Checklists</Section>

      {blendAnalysis && blendAnalysis.length > 0 && (
        <div style={{marginBottom:14,padding:10,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
          <span style={{fontSize:12,fontWeight:600,color:T.textSec,display:"block",marginBottom:6}}>Blend composition check</span>
          <div style={{display:"flex",flexDirection:"column",gap:4}}>
            {blendAnalysis.map((b,i) => (
              <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:12,padding:"4px 0",borderBottom:i<blendAnalysis.length-1?`1px solid ${T.border}`:"none"}}>
                <span style={{color:T.text,fontWeight:500}}>{b.itemName}</span>
                <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",justifyContent:"flex-end"}}>
                  <span style={{color:T.textMut}}>Expected: {b.expected.toFixed(1)} kg</span>
                  <span style={{color:statusColor(b.status),fontWeight:600,display:"inline-flex",alignItems:"center",gap:4}}>Tagged: {b.actual.toFixed(1)} kg {b.status==="ok"?"✓":"⚠"} <span style={{fontSize:10,padding:"1px 6px",borderRadius:8,background:"rgba(255,255,255,0.04)",border:`1px solid ${statusColor(b.status)}`}}>{statusLabel(b.status)}</span></span>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        {Object.keys(groups).map(ckId => {
          const ck = checklists.find(c=>c.id===ckId);
          const items = groups[ckId];
          return (
            <div key={ckId}>
              <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
                <span style={{fontSize:13,fontWeight:600,color:T.accent}}>{ck?.name || ckId}</span>
                <Badge variant="muted" style={{fontSize:10}}>{items.length}</Badge>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {items.map((te, idx) => {
                  const meta = findEntryMeta(te);
                  const teAutoId = te.autoId || te.responseId;
                  const clickable = !!(teAutoId && te.checklistId && typeof onEntryClick === "function");
                  return (
                    <div key={idx} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                      <Icon name="clipboard" size={14} color={T.accent}/>
                      <div style={{flex:1,minWidth:0}}>
                        {clickable ? (
                          <button onClick={()=>onEntryClick(te.checklistId, teAutoId)}
                            style={{background:"none",border:"none",padding:0,cursor:"pointer",fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600,textDecoration:"underline",textDecorationStyle:"dotted"}}
                            title="View response details and source chain">{teAutoId}</button>
                        ) : (
                          <span style={{fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600}}>{teAutoId}</span>
                        )}
                        <div style={{display:"flex",gap:8,marginTop:2,flexWrap:"wrap"}}>
                          {meta.date && <span style={{fontSize:11,color:T.textMut}}>{meta.date}</span>}
                          <span style={{fontSize:11,color:T.textMut}}>Stage: {te.stageName}</span>
                        </div>
                      </div>
                      <span style={{fontSize:12,color:T.textSec,fontWeight:500,whiteSpace:"nowrap"}}>{te.qty} kg</span>
                    </div>
                  );
                })}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function DeliverConfirmModal({ preview, onConfirm, onCancel, busy }) {
  const [overrideAcknowledged, setOverrideAcknowledged] = useState(false);
  if (!preview) return null;
  const blendWarnings = Array.isArray(preview.blendWarnings) ? preview.blendWarnings : [];
  const hasWarnings = blendWarnings.length > 0;
  const canConfirm = !busy && (!hasWarnings || overrideAcknowledged);
  return (
    <div onClick={onCancel} style={{position:"fixed",inset:0,zIndex:1000,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,border:`1px solid ${T.border}`,maxWidth:520,width:"100%",maxHeight:"85vh",overflow:"auto",padding:20}}>
        <h3 style={{fontSize:16,fontWeight:600,color:T.text,margin:"0 0 4px"}}>Mark as Delivered?</h3>
        <p style={{fontSize:13,color:T.textSec,marginBottom:12}}>The following inventory deductions will be applied atomically:</p>
        <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:16}}>
          {(preview.deductions || []).length === 0 ? <p style={{fontSize:12,color:T.textMut}}>No inventory deductions.</p> :
            (() => {
              // Group by blendLineIndex when present; otherwise show "Other"
              const groups = {};
              preview.deductions.forEach(d => {
                const key = (d.blendLineIndex !== undefined && d.blendLineIndex !== null) ? `L${d.blendLineIndex}` : "other";
                if (!groups[key]) groups[key] = [];
                groups[key].push(d);
              });
              return Object.keys(groups).map(k => (
                <div key={k}>
                  <span style={{fontSize:12,fontWeight:600,color:T.textSec,display:"block",marginBottom:4}}>{k === "other" ? "Deductions" : "Blend Line " + (Number(k.slice(1))+1)}</span>
                  <div style={{display:"flex",flexDirection:"column",gap:4}}>
                    {groups[k].map((d, i) => (
                      <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"8px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
                        <div>
                          <span style={{fontSize:13,fontWeight:500,color:T.text}}>{d.itemName || d.itemId}</span>
                          <span style={{fontSize:11,color:T.textMut,marginLeft:6}}>({d.stageName}{d.checklistAutoId?" — "+d.checklistAutoId:""}{d.isMixed?" · mixed":""})</span>
                        </div>
                        <span style={{fontSize:13,fontWeight:600,color:T.danger}}>-{d.qty} kg</span>
                      </div>
                    ))}
                  </div>
                </div>
              ));
            })()
          }
        </div>

        {hasWarnings && (
          <div style={{padding:"10px 12px",background:T.warningBg,border:`1px solid ${T.warningBorder}`,borderRadius:T.radSm,marginBottom:16}}>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6}}>
              <Icon name="clipboard" size={14} color={T.warning}/>
              <span style={{fontSize:13,fontWeight:600,color:T.warning}}>Blend composition warning</span>
            </div>
            <p style={{fontSize:12,color:T.warning,marginBottom:8}}>The following ingredient(s) are over-tagged by more than 10% compared to the blend recipe:</p>
            <ul style={{fontSize:12,color:T.text,paddingLeft:18,margin:"0 0 10px"}}>
              {blendWarnings.map((w, i) => (
                <li key={i} style={{marginBottom:4}}>
                  <b>{w.itemName}</b>: expected {w.expected}kg, tagged {w.actual}kg
                </li>
              ))}
            </ul>
            <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:12,color:T.text}}>
              <input type="checkbox" checked={overrideAcknowledged} onChange={e=>setOverrideAcknowledged(e.target.checked)} style={{width:16,height:16,cursor:"pointer"}}/>
              <span>I acknowledge the over-tagged quantities and want to proceed with delivery</span>
            </label>
          </div>
        )}

        <div style={{display:"flex",gap:8}}>
          <Btn variant="secondary" onClick={onCancel} style={{flex:1}}>Cancel</Btn>
          <Btn variant="success" onClick={onConfirm} disabled={!canConfirm} style={{flex:1,opacity:canConfirm?1:0.5}}>
            {busy ? "Applying..." : "Confirm & Deliver"}
          </Btn>
        </div>
      </div>
    </div>
  );
}

// ─── Order Detail with Checklist Input ─────────────────────────

function OrderDetailView({order,checklists,customers,isAdmin,currentUser,approvedEntries,inventoryItems,inventoryCategories,untaggedChecklists,blends,onUpdate,onEditOrder,onDeleteOrder,onRevertChecklist,onEditResponses,onUpdateStatus,onTagStage,onTagMixedStage,onUntagStage,onDeliver}){
  const [expandedId,setExpandedId]=useState(null);
  const [formData,setFormData]=useState({date:"",person:"",responses:{},remarks:{},batchAllocations:{}});
  const [invItemId,setInvItemId]=useState("");
  const [invOutputItemId,setInvOutputItemId]=useState("");
  const [submitting,setSubmitting]=useState(false);
  const [viewingId,setViewingId]=useState(null);
  const [viewData,setViewData]=useState(null);
  const [loadingView,setLoadingView]=useState(false);
  const [editing,setEditing]=useState(false);
  const [editForm,setEditForm]=useState({});
  const [completedResponses,setCompletedResponses]=useState({}); // checklistId -> responses[]
  const [deliverPreview,setDeliverPreview]=useState(null);
  const [delivering,setDelivering]=useState(false);
  const [editingStages,setEditingStages]=useState(false);
  const [taggedEntryView,setTaggedEntryView]=useState(null); // {checklistId, autoId}
  const [invError,setInvError]=useState({idx:null,message:""});

  const pending=order.checklists.filter(c=>c.status==="pending");
  const completed=order.checklists.filter(c=>c.status==="completed");
  const getCk=id=>checklists.find(c=>c.id===id);

  const handleExpand=(item)=>{
    if(expandedId===item.id){setExpandedId(null);return}
    setExpandedId(item.id);
    setFormData({date:new Date().toISOString().split("T")[0],person:currentUser?.displayName||order.assignedTo||"",responses:{},remarks:{},batchAllocations:{}});
    // Pre-fetch completed checklist responses for cross-checklist formulas
    completed.forEach(c=>{
      if(!completedResponses[c.checklistId]){
        API.get("getResponses",{id:c.id}).then(data=>{
          if(data?.responses) setCompletedResponses(prev=>({...prev,[c.checklistId]:data.responses}));
        }).catch(()=>{});
      }
    });
  };

  // Formula field value resolver for the currently-being-filled checklist
  const getFieldValue=(checklistRef,questionIdx)=>{
    if(checklistRef==="self") return formData.responses[questionIdx]||"";
    // Cross-checklist: look up from completedResponses
    const resps=completedResponses[checklistRef];
    if(resps) { const r=resps.find(r=>r.questionIndex===questionIdx); return r?r.response:""; }
    return "";
  };

  const handleSubmit=async(item,ck)=>{
    setSubmitting(true);
    setInvError({idx:null,message:""});
    try {
      const nq=normalizeQuestions(ck.questions);
      // Required inventory-tracking fields must have a value
      const computedResponsesForInv = {};
      nq.forEach((q, qi) => {
        if (q.formula) {
          const computed = evaluateFormula(q.formula, getFieldValue);
          computedResponsesForInv[qi] = computed !== null ? String(computed) : "";
        } else {
          computedResponsesForInv[qi] = formData.responses[qi] !== undefined ? formData.responses[qi] : "";
        }
      });
      const invCheck = validateRequiredInventoryFields(nq, computedResponsesForInv, formData.batchAllocations);
      if (!invCheck.ok) {
        setInvError({idx: invCheck.firstIdx, message: invCheck.message});
        setSubmitting(false);return;
      }
      // Validate linked dropdown fields
      for(let qi=0;qi<nq.length;qi++){
        const q=nq[qi];
        if(q.linkedSource&&q.linkedSource.checklistId){
          const entries=approvedEntries[q.linkedSource.checklistId]||[];
          const srcCk=checklists.find(c=>c.id===q.linkedSource.checklistId);
          const srcName=srcCk?.name||"source checklist";
          if(entries.length===0){
            alert("Cannot submit — no approved entries available from \""+srcName+"\". Please complete and approve a "+srcName+" first.");
            setSubmitting(false);return;
          }
          // Check if value is satisfied via either the direct response (LinkedDropdown)
          // or via batch allocations (BatchSelector) — at least one valid batch entry
          const hasDirectValue = (formData.responses[qi]||"").trim();
          const batchAllocs = formData.batchAllocations?.[qi];
          const hasValidBatch = Array.isArray(batchAllocs) && batchAllocs.some(a => a.sourceAutoId && (parseFloat(a.quantity) || 0) > 0);
          if(!hasDirectValue && !hasValidBatch){
            alert("Please select a "+q.text+" before submitting.");
            setSubmitting(false);return;
          }
        }
      }
      // Check required remarks
      for(let qi=0;qi<nq.length;qi++){
        const q=nq[qi];
        if(q.remarkCondition && q.ideal){
          const val=formData.responses[qi]||"";
          const idealVal=evaluateFormula(q.ideal,getFieldValue);
          if(checkRemarkCondition(val,idealVal,q.remarkCondition) && !(formData.remarks[qi]||"").trim()){
            alert("Please provide remarks for: "+q.text);
            setSubmitting(false);return;
          }
        }
      }
      // Forced remark targets — another field's deviation requires this field to be filled
      for(let qi=0;qi<nq.length;qi++){
        const q=nq[qi];
        if(!q.remarkCondition||q.remarksTargetIdx==null||!q.formula||!q.ideal) continue;
        const aVal=evaluateFormula(q.formula,getFieldValue);
        const iVal=evaluateFormula(q.ideal,getFieldValue);
        if(aVal!=null&&iVal!=null&&checkRemarkCondition(aVal,iVal,q.remarkCondition)){
          const targetVal=formData.responses[q.remarksTargetIdx]||"";
          if(!String(targetVal).trim()){
            const targetText=nq[q.remarksTargetIdx]?.text||`Q${q.remarksTargetIdx+1}`;
            alert(`Please fill "${targetText}" — ${q.remarkCondition.message||"value differs from ideal"}`);
            setSubmitting(false);return;
          }
        }
      }
      // Date comparison validation
      for(let qi=0;qi<nq.length;qi++){
        const q2=nq[qi];
        if(q2.type==="date"&&q2.dateComparison&&q2.dateComparison.compareToFieldIdx!==""&&q2.dateComparison.compareToFieldIdx!==undefined){
          const v1=formData.responses[qi]||"";
          const v2=formData.responses[Number(q2.dateComparison.compareToFieldIdx)]||"";
          if(v1&&v2){
            const da=new Date(v1),db=new Date(v2);
            if(!isNaN(da.getTime())&&!isNaN(db.getTime())){
              const ta=da.setHours(0,0,0,0),tb=db.setHours(0,0,0,0);
              const op2=q2.dateComparison.operator;
              const cf=nq[Number(q2.dateComparison.compareToFieldIdx)]?.text||"the other date";
              let de=null;
              if(op2==="gte"&&ta<tb) de=q2.dateComparison.errorMessage||`${q2.text} cannot be before ${cf}`;
              else if(op2==="lte"&&ta>tb) de=q2.dateComparison.errorMessage||`${q2.text} cannot be after ${cf}`;
              else if(op2==="eq"&&ta!==tb) de=q2.dateComparison.errorMessage||`${q2.text} must be the same as ${cf}`;
              if(de){alert(de);setSubmitting(false);return;}
            }
          }
        }
      }
      const responses=nq.map((q,qi)=>{
        let resp=formData.responses[qi]||"";
        if(q.formula){
          const computed=evaluateFormula(q.formula,getFieldValue);
          resp = computed!==null?String(computed):"";
        }
        if(q.linkedSource&&Array.isArray(formData.batchAllocations?.[qi])){
          const ids=formData.batchAllocations[qi].map(a=>a.sourceAutoId).filter(Boolean);
          if(ids.length>0) resp=ids.join(", ");
        }
        return { questionIndex:qi, questionText:q.text, response:resp };
      });
      await API.post("submitChecklist",{id:item.id,date:formData.date,person:formData.person,responses,remarks:formData.remarks,inventoryItemId:invItemId,inventoryOutputItemId:invOutputItemId,batchAllocations:formData.batchAllocations||{}});
      // Optimistic update: move checklist from pending to completed locally
      const now = new Date().toISOString();
      const updatedChecklists = order.checklists.map(c =>
        c.id === item.id ? { ...c, status:"completed", completedAt:now, completedBy:formData.person, workDate:formData.date } : c
      );
      const updated = { ...order, checklists: updatedChecklists };
      onUpdate(updated);
      setExpandedId(null);
    } catch(e) { alert("Failed to submit: " + e.message); }
    setSubmitting(false);
  };

  const handleViewResponses=async(item)=>{
    if(viewingId===item.id){setViewingId(null);setViewData(null);return}
    setViewingId(item.id); setLoadingView(true);
    try { const data=await API.get("getResponses",{id:item.id}); setViewData(data); } catch {setViewData(null)}
    setLoadingView(false);
  };

  const handleSaveEdit = async () => {
    await onEditOrder({ id: order.id, name: editForm.name, customerId: editForm.customerId, assignedTo: editForm.assignedTo, invoiceSo: editForm.invoiceSo, orderTypeDetail: editForm.orderTypeDetail, productType: editForm.productType });
    setEditing(false);
  };


  return (
    <div className="fade-up">
      {/* ── Order Header Card ── */}
      <div style={{background:T.card,borderRadius:T.rad,padding:18,border:`1px solid ${T.border}`,marginBottom:24}}>
        {editing ? (
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <Field label="Order Name"><Input value={editForm.name} onChange={v=>setEditForm(p=>({...p,name:v}))}/></Field>
            <Field label="Customer"><div style={{display:"flex",flexWrap:"wrap",gap:8}}>
              {customers.map(c=><Chip key={c.id} label={c.label} active={editForm.customerId===c.id} onClick={()=>setEditForm(p=>({...p,customerId:c.id}))}/>)}
            </div></Field>
            <Field label="Assigned To"><Input value={editForm.assignedTo} onChange={v=>setEditForm(p=>({...p,assignedTo:v}))}/></Field>
            <Field label="Invoice / SO"><Input value={editForm.invoiceSo} onChange={v=>setEditForm(p=>({...p,invoiceSo:v}))} placeholder="INV-001"/></Field>
            <Field label="Order Type Detail">
              <div style={{display:"flex",gap:8}}>
                <Chip label="Client Order" active={editForm.orderTypeDetail==="Client Order"} onClick={()=>setEditForm(p=>({...p,orderTypeDetail:"Client Order"}))}/>
                <Chip label="Sample Order" active={editForm.orderTypeDetail==="Sample Order"} onClick={()=>setEditForm(p=>({...p,orderTypeDetail:"Sample Order"}))}/>
              </div>
            </Field>
            <Field label="Product Type">
              <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                {PRODUCT_TYPES.map(pt=><Chip key={pt} label={pt} active={editForm.productType===pt} onClick={()=>setEditForm(p=>({...p,productType:editForm.productType===pt?"":pt}))}/>)}
              </div>
            </Field>
            <div style={{display:"flex",gap:8}}>
              <Btn variant="secondary" small onClick={()=>setEditing(false)} style={{flex:1}}>Cancel</Btn>
              <Btn small onClick={handleSaveEdit} style={{flex:1}}>Save</Btn>
            </div>
          </div>
        ) : (
          <>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <h2 style={{fontSize:18,fontWeight:600,marginBottom:8}}>{order.name}</h2>
              <div style={{display:"flex",gap:4}}>
                <button onClick={()=>{setEditing(true);setEditForm({name:order.name,customerId:order.customerId,assignedTo:order.assignedTo||"",invoiceSo:order.invoiceSo||"",orderTypeDetail:order.orderTypeDetail||"Client Order",productType:order.productType||""})}} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="edit" size={16} color={T.textSec}/></button>
                {isAdmin && <button onClick={()=>{if(confirm("Delete this order and all its checklists?"))onDeleteOrder(order.id)}} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={16} color={T.danger}/></button>}
              </div>
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:6}}>
              {order.assignedTo&&<div style={{display:"flex",alignItems:"center",gap:6}}><Icon name="user" size={14} color={T.textMut}/><span style={{fontSize:13,color:T.textSec}}>{order.assignedTo}</span></div>}
              {order.invoiceSo&&<Badge variant="muted">{order.invoiceSo}</Badge>}
              {order.orderTypeDetail&&<Badge variant={order.orderTypeDetail==="Sample Order"?"danger":"success"}>{order.orderTypeDetail}</Badge>}
              {order.productType&&(()=>{const ptc=PRODUCT_TYPE_COLORS[order.productType]||PRODUCT_TYPE_COLORS.Others;return <span style={{display:"inline-flex",alignItems:"center",padding:"3px 10px",borderRadius:20,fontSize:11,fontWeight:500,background:ptc.bg,color:ptc.color,border:`1px solid ${ptc.border}`,whiteSpace:"nowrap"}}>{order.productType}</span>})()}
            </div>
            <span style={{fontSize:12,color:T.textMut}}>{formatDate(order.createdAt)}</span>

            {/* ── Status Workflow ── */}
            {(()=>{
              const currentStatus=order.status||"beans_not_roasted";
              const statusIdx=ORDER_STATUSES.indexOf(currentStatus);
              const orderLines=order.orderLines||[];
              const totalRequired=orderLines.reduce((s,l)=>s+(parseFloat(l.quantity)||0),0);
              const totalTagged=orderLines.reduce((s,l)=>s+(parseFloat(l.taggedQuantity)||0),0);
              const allFullyTagged=totalRequired>0?totalTagged>=totalRequired:false;
              return <div style={{marginTop:14}}>
                <div style={{display:"flex",alignItems:"center",gap:0,marginBottom:8,overflowX:"auto",paddingBottom:4}}>
                  {ORDER_STATUSES.map((s,i)=>{
                    const isActive=i<=statusIdx;const isCurrent=i===statusIdx;
                    const isCompleted=s==="completed"&&!allFullyTagged;
                    return <div key={s} style={{display:"flex",alignItems:"center",flex:1,minWidth:0}}>
                      <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:3,flex:1,minWidth:0}}>
                        <div style={{width:20,height:20,borderRadius:"50%",background:isCurrent?T.accent:isActive?T.success:"transparent",border:`2px solid ${isActive?T.success:T.border}`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                          {isActive&&i<statusIdx&&<Icon name="check" size={12} color={T.bg}/>}
                        </div>
                        <span style={{fontSize:9,color:isCurrent?T.accent:isActive?T.success:T.textMut,textAlign:"center",lineHeight:1.1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:60}}>{ORDER_STATUS_LABELS[s]}</span>
                      </div>
                      {i<ORDER_STATUSES.length-1&&<div style={{height:2,flex:1,background:isActive&&i<statusIdx?T.success:T.border,minWidth:4}}/>}
                    </div>;
                  })}
                </div>
                {currentStatus!=="delivered"&&<div style={{display:"flex",gap:6,marginTop:8}}>
                  {statusIdx>0&&<Btn variant="ghost" small onClick={()=>onUpdateStatus(order.id,ORDER_STATUSES[statusIdx-1])} style={{flex:1,fontSize:11}}>← Back</Btn>}
                  {statusIdx<ORDER_STATUSES.length-1&&(
                    ORDER_STATUSES[statusIdx+1]==="completed"&&!allFullyTagged?
                    <Btn variant="secondary" small disabled style={{flex:1,fontSize:11,opacity:0.4}}>Complete (tag all quantities first)</Btn>:
                    ORDER_STATUSES[statusIdx+1]==="delivered"?
                    <Btn variant="success" small onClick={()=>{if(confirm("Mark as delivered? This will update inventory."))onUpdateStatus(order.id,"delivered")}} style={{flex:1,fontSize:11}}>Mark Delivered</Btn>:
                    <Btn small onClick={()=>onUpdateStatus(order.id,ORDER_STATUSES[statusIdx+1])} style={{flex:1,fontSize:11}}>Advance →</Btn>
                  )}
                </div>}
              </div>;
            })()}

            {/* ── Blend Lines Progress ── */}
            {order.orderLines&&order.orderLines.length>0&&(()=>{
              const taggedByItemId = computeTaggedByItem(order, checklists, inventoryItems, approvedEntries, untaggedChecklists);
              return (
              <div style={{marginTop:14,background:T.bg,borderRadius:T.radSm,padding:12,border:`1px solid ${T.border}`}}>
                <span style={{fontSize:12,fontWeight:600,color:T.textSec,display:"block",marginBottom:8}}>Blend Lines</span>
                {order.orderLines.map((line,li)=>{
                  const req=parseFloat(line.quantity)||0;
                  // Compute tagged qty for this line: sum of tagged quantities for the components in this line.
                  const comps = Array.isArray(line.blendComponents) ? line.blendComponents : [];
                  let tagged = 0;
                  if (comps.length > 0) {
                    // Sum proportional tagged per component (capped at the line's required qty for that component)
                    comps.forEach(c => {
                      const lineCompReq = ((parseFloat(c.percentage)||0)/100) * req;
                      const k = blendItemKey(c.itemId, c.itemName);
                      const totalTaggedForItem = parseFloat(taggedByItemId[k]) || 0;
                      // Count up to the component's requirement for this line
                      tagged += Math.min(totalTaggedForItem, lineCompReq);
                    });
                  } else {
                    tagged = parseFloat(line.taggedQuantity) || 0;
                  }
                  tagged = Math.round(tagged*100)/100;
                  const pct=req>0?Math.min(100,(tagged/req)*100):0;
                  return <div key={li} style={{marginBottom:li<order.orderLines.length-1?10:0}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                      <span style={{fontSize:13,fontWeight:500,color:T.text}}>{line.blend||"Blend "+(li+1)}</span>
                      {line.deliveryDate&&<span style={{fontSize:11,color:T.textMut}}>Due: {line.deliveryDate}</span>}
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <div style={{flex:1,height:4,borderRadius:2,background:T.surfaceHover,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",borderRadius:2,background:pct>=100?T.success:T.accent,transition:"width .5s ease"}}/></div>
                      <span style={{fontSize:11,color:T.textSec,fontFamily:T.mono,whiteSpace:"nowrap"}}>{tagged}/{req} kg</span>
                    </div>
                    {comps.length>0&&<OrderBlendBreakdown line={line} allocations={[]} taggedByItemId={taggedByItemId}/>}
                  </div>;
                })}
              </div>
            );})()}
          </>
        )}
      </div>

      {/* ── Stages Panel ── */}
      {Array.isArray(order.stages) && order.stages.length > 0 && typeof onTagStage === "function" && !editingStages && (
        <StagesPanel order={order} checklists={checklists} approvedEntries={approvedEntries} untaggedChecklists={untaggedChecklists} inventoryItems={inventoryItems} blends={blends} isAdmin={isAdmin}
          onTag={async(stageId, autoId, sourceCkId, qty, quantityFieldValue, blendExtras)=>{
            await onTagStage(order.id, stageId, autoId, sourceCkId, qty, quantityFieldValue, blendExtras);
          }}
          onTagMixed={async(payload)=>{
            await onTagMixedStage(order.id, payload);
          }}
          onUntag={async(stageId, specOrAutoId)=>{
            // Accept either a string autoId (legacy) or a spec object (blend flow)
            if (typeof specOrAutoId === "string") {
              if(!confirm("Untag this checklist from the stage?")) return;
              await onUntagStage(order.id, stageId, specOrAutoId);
            } else {
              await onUntagStage(order.id, stageId, specOrAutoId.responseId || "", specOrAutoId);
            }
          }}
          onTaggedEntryClick={(checklistId, autoId)=>setTaggedEntryView({checklistId, autoId})}
          onDeliver={async()=>{
            try {
              const preview = await onDeliver(order.id, false);
              setDeliverPreview(preview);
            } catch(e) { /* handled upstream */ }
          }}/>
      )}

      {/* Admin: Edit Stages button (hidden when the editor is open) */}
      {isAdmin && Array.isArray(order.stages) && !editingStages && order.status !== "delivered" && typeof onEditOrder === "function" && (
        <div style={{marginBottom:24,marginTop:-16}}>
          <Btn variant="secondary" small onClick={()=>setEditingStages(true)} style={{width:"100%"}}><Icon name="edit" size={14} color={T.text}/> Edit Stages</Btn>
        </div>
      )}

      {/* Inline stage editor */}
      {editingStages && (
        <StageEditor order={order} checklists={checklists}
          onCancel={()=>setEditingStages(false)}
          onSave={async(newStages)=>{
            await onEditOrder({ id: order.id, stages: newStages });
            setEditingStages(false);
          }}/>
      )}

      {deliverPreview && <DeliverConfirmModal preview={deliverPreview} busy={delivering}
        onCancel={()=>setDeliverPreview(null)}
        onConfirm={async()=>{
          setDelivering(true);
          try { await onDeliver(order.id, true); setDeliverPreview(null); } catch(e) { /* handled */ }
          setDelivering(false);
        }}/>}

      {taggedEntryView && <TaggedEntryModal checklistId={taggedEntryView.checklistId} autoId={taggedEntryView.autoId} checklists={checklists} isAdmin={isAdmin} onClose={()=>setTaggedEntryView(null)}/>}

      {/* ── Tagged Checklists grouped by type ── */}
      {Array.isArray(order.stages) && order.stages.some(s => (s.taggedEntries||[]).length > 0) && !editingStages && (
        <TaggedChecklistsGrouped order={order} checklists={checklists} approvedEntries={approvedEntries} untaggedChecklists={untaggedChecklists} inventoryItems={inventoryItems}
          onEntryClick={(checklistId, autoId)=>setTaggedEntryView({checklistId, autoId})}/>
      )}

      {/* ── Pending Checklists ── */}
      <Section icon="clock" count={pending.length}>Pending Checklists</Section>
      {pending.length===0?<Empty icon="checkCircle" text="All checklists completed!"/>:
        <div style={{display:"flex",flexDirection:"column",gap:10,marginBottom:32}}>
          {pending.map((item,i)=>{
            const ck=getCk(item.checklistId);if(!ck)return null;
            const isExpanded=expandedId===item.id;
            const hasForm=!!ck.formUrl;
            return <div key={item.id} className="slide-in" style={{background:T.card,borderRadius:T.rad,padding:"16px 18px",border:`1px solid ${isExpanded?T.accentBorder:T.border}`,animationDelay:`${i*60}ms`,animationFillMode:"backwards",transition:"border .2s"}}>
              <div onClick={()=>handleExpand(item)} style={{cursor:"pointer",marginBottom:isExpanded?16:0}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <span style={{fontSize:15,fontWeight:600,color:T.text}}>{ck.name}</span>
                    {ck.subtitle&&<span style={{display:"block",fontSize:12,color:T.accent,marginTop:2}}>{ck.subtitle}</span>}
                    <span style={{display:"block",fontSize:12,color:T.textMut,marginTop:2}}>{ck.questions.length} items to verify</span>
                  </div>
                  <Icon name="chevron" size={18} color={T.textMut} style={{transform:isExpanded?"rotate(90deg)":"rotate(0)",transition:"transform .2s"}}/>
                </div>
              </div>
              {isExpanded&&<div style={{borderTop:`1px solid ${T.border}`,paddingTop:16}}>
                {item.checklistId==="ck_grinding"&&Array.isArray(order.orderLines)&&order.orderLines.length>0&&(
                  <div style={{padding:"10px 12px",background:T.accentBg,border:`1px solid ${T.accentBorder}`,borderRadius:T.radSm,marginBottom:12}}>
                    <span style={{fontSize:12,fontWeight:600,color:T.accent,display:"block",marginBottom:6}}>Blend Requirements</span>
                    {order.orderLines.map((line,li)=>(
                      <div key={li} style={{marginBottom:li<order.orderLines.length-1?6:0}}>
                        <span style={{fontSize:12,color:T.text,fontWeight:500}}>{line.blend||"Blend "+(li+1)} — {line.quantity} kg</span>
                        {Array.isArray(line.blendComponents)&&line.blendComponents.length>0?
                          <OrderBlendBreakdown line={line} allocations={[]}/>:
                          <p style={{fontSize:11,color:T.textMut,marginTop:2}}>No blend recipe attached</p>
                        }
                      </div>
                    ))}
                  </div>
                )}
                {ck.autoIdConfig?.enabled && (() => {
                  const preview = buildAutoIdPreview(ck, formData.responses, formData.date, inventoryItems);
                  return preview ? (
                    <div style={{padding:"10px 12px",background:T.accentBg,border:`1px solid ${T.accentBorder}`,borderRadius:T.radSm,marginBottom:12,display:"flex",alignItems:"center",gap:8}}>
                      <Icon name="clipboard" size={14} color={T.accent}/>
                      <span style={{fontSize:11,color:T.textMut}}>Auto ID:</span>
                      <span style={{fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600}}>{preview}</span>
                    </div>
                  ) : null;
                })()}
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:20}}>
                  <Field label="Date">
                    <input type="date" value={formData.date} onChange={e=>setFormData(p=>({...p,date:e.target.value}))}
                      style={{width:"100%",padding:"12px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:15,outline:"none",colorScheme:"dark"}}/>
                  </Field>
                  <Field label="Person">
                    <Input value={formData.person} onChange={()=>{}} readOnly placeholder="Who is filling this?" style={{fontSize:15,padding:"12px 14px"}}/>
                  </Field>
                </div>
                {/* ── Inventory Item Selection (for checklists that affect stock) ── */}
                {(()=>{
                  const ckId=item.checklistId;
                  // If the checklist has any per-question inventoryLink configured, suppress the legacy picker
                  const hasNewInvLink=normalizeQuestions(ck.questions).some(qq=>qq.inventoryLink&&qq.inventoryLink.enabled);
                  if(hasNewInvLink) return null;
                  const invCatMap={"ck_green_beans":"Green Beans","ck_roasted_beans":"Green Beans","ck_grinding":"Roasted Beans"};
                  const invOutCatMap={"ck_roasted_beans":"Roasted Beans","ck_grinding":"Packing Items"};
                  const inCat=invCatMap[ckId];const outCat=invOutCatMap[ckId];
                  if(!inCat&&!outCat) return null;
                  const inItems=(inventoryItems||[]).filter(i=>i.category===inCat&&i.isActive);
                  const allOutItems=outCat?(inventoryItems||[]).filter(i=>i.category===outCat&&i.isActive):[];
                  // Determine output item behaviour based on input item's equivalents
                  // equivalentItems is an array of { category, itemId } objects
                  const selectedInItem=invItemId?(inventoryItems||[]).find(i=>i.id===invItemId):null;
                  const inEqList=selectedInItem&&Array.isArray(selectedInItem.equivalentItems)?selectedInItem.equivalentItems:[];
                  const inEqItemIds=inEqList.map(e=>e.itemId);
                  const linkedOutItems=inEqItemIds.length>0?allOutItems.filter(i=>inEqItemIds.includes(i.id)):[];
                  const outputLocked=inEqItemIds.length===1&&linkedOutItems.length===1;
                  const outputChoices=inEqItemIds.length>0&&linkedOutItems.length>0?linkedOutItems:allOutItems;
                  return <div style={{display:"flex",flexDirection:"column",gap:12,marginBottom:16,padding:12,background:T.accentBg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
                    <span style={{fontSize:12,fontWeight:600,color:T.accent}}>Inventory Tracking</span>
                    {inCat&&<Field label={ckId==="ck_green_beans"?"Green Bean Item (IN)":"Input Item (OUT)"}>
                      <select value={invItemId} onChange={e=>{
                        const newId=e.target.value;
                        setInvItemId(newId);
                        if(!newId){setInvOutputItemId("");return;}
                        const selItem=(inventoryItems||[]).find(i=>i.id===newId);
                        const eqL=selItem&&Array.isArray(selItem.equivalentItems)?selItem.equivalentItems:[];
                        const eqItemIds=eqL.map(eq=>eq.itemId);
                        const eqOutItems=eqItemIds.length>0?allOutItems.filter(i=>eqItemIds.includes(i.id)):[];
                        if(eqItemIds.length===1&&eqOutItems.length===1){setInvOutputItemId(eqOutItems[0].id)}
                        else if(eqItemIds.length>1&&eqOutItems.length>0){if(!eqOutItems.find(i=>i.id===invOutputItemId))setInvOutputItemId("")}
                        else{/* no equivalents — keep current selection */}
                      }}
                        style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14}}>
                        <option value="">— None (skip inventory) —</option>
                        {inItems.map(i=><option key={i.id} value={i.id}>{i.name} ({i.currentStock} {i.unit})</option>)}
                      </select>
                    </Field>}
                    {outCat&&<Field label={<span>Output Item (IN){outputLocked&&<span style={{fontSize:10,color:T.textMut,marginLeft:6,fontWeight:400}}>Auto-filled from linked equivalent</span>}</span>}>
                      <select value={invOutputItemId} onChange={e=>setInvOutputItemId(e.target.value)} disabled={outputLocked}
                        style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:outputLocked?"rgba(255,255,255,0.05)":T.bg,border:`1px solid ${T.border}`,color:outputLocked?T.textMut:T.text,fontSize:14,opacity:outputLocked?0.7:1,cursor:outputLocked?"not-allowed":"pointer"}}>
                        <option value="">— None (skip inventory) —</option>
                        {outputChoices.map(i=><option key={i.id} value={i.id}>{i.name} ({i.currentStock} {i.unit})</option>)}
                      </select>
                    </Field>}
                  </div>;
                })()}
                <div style={{display:"flex",flexDirection:"column",gap:14,marginBottom:20}}>
                  {(()=>{
                    const ckNq=normalizeQuestions(ck.questions);
                    const onBatchAllocChange=(qi,allocs)=>{
                      setFormData(p=>{
                        const nextBA={...(p.batchAllocations||{}),[qi]:allocs};
                        let total=0;
                        Object.values(nextBA).forEach(a=>{if(Array.isArray(a))a.forEach(x=>total+=parseFloat(x.quantity)||0);});
                        const nextResponses={...p.responses};
                        ckNq.forEach((qq,i)=>{
                          if(qq.inventoryLink?.enabled&&qq.inventoryLink.txType==="OUT") nextResponses[i]=String(total);
                        });
                        return {...p,batchAllocations:nextBA,responses:nextResponses};
                      });
                    };
                    return ckNq.map((q,qi)=>{
                      let autoVal=null;
                      if(q.formula) autoVal=evaluateFormula(q.formula,getFieldValue);
                      // Formula fields are always derived from current state — never read from formData.responses
                      const currentVal=q.formula
                        ? (autoVal!==null?String(autoVal):"")
                        : (formData.responses[qi]!==undefined?formData.responses[qi]:"");
                      let idealVal=null;
                      if(q.ideal) idealVal=evaluateFormula(q.ideal,getFieldValue);
                      const needsRemark=q.remarkCondition&&idealVal!==null&&checkRemarkCondition(currentVal,idealVal,q.remarkCondition);
                      const required = isInventoryRequiredQuestion(q);
                      const invalid = invError.idx === qi;
                      return (
                        <div key={qi} style={invalid ? {border:`1px solid ${T.danger}`,borderRadius:T.radSm,padding:8} : undefined}>
                          {required && (
                            <div style={{fontSize:11,color:T.danger,marginBottom:4,display:"flex",alignItems:"center",gap:4}}>
                              <span style={{color:T.danger,fontWeight:700}}>*</span> Required for inventory tracking
                            </div>
                          )}
                          <QuestionInputField q={q} qi={qi} currentVal={currentVal} idealVal={idealVal} needsRemark={needsRemark}
                            formData={formData} setFormData={setFormData} approvedEntries={approvedEntries} checklists={checklists} getFieldValue={getFieldValue}
                            orders={item.checklistId==="ck_grinding"?null:null} customers={customers}
                            inventoryItems={inventoryItems} onBatchAllocChange={onBatchAllocChange} allQuestions={ckNq}
                            onInventoryAutoFill={(item2)=>{
                              if(!item2){setInvItemId("");setInvOutputItemId("");return;}
                              setInvItemId(item2.id);
                              // equivalentItems is [{category, itemId}]
                              const eqL=Array.isArray(item2.equivalentItems)?item2.equivalentItems:[];
                              if(eqL.length>0){
                                const eqItem=(inventoryItems||[]).find(it=>eqL.some(e=>e.itemId===it.id));
                                if(eqItem) setInvOutputItemId(eqItem.id);
                              }
                            }}/>
                        </div>
                      );
                    });
                  })()}
                </div>
                {invError.message && (
                  <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"10px 14px",marginBottom:12}}>
                    <span style={{fontSize:13,color:T.danger}}>{invError.message}</span>
                  </div>
                )}
                <div style={{display:"flex",gap:8}}>
                  {hasForm&&<Btn variant="secondary" onClick={()=>window.open(ck.formUrl,"_blank")} style={{flex:1}}>
                    <Icon name="externalLink" size={16} color={T.text}/> Fill Google Form
                  </Btn>}
                  <Btn variant="success" onClick={()=>handleSubmit(item,ck)} disabled={submitting} style={{flex:1}}>
                    <Icon name="check" size={16} color={T.success}/> {submitting?"Submitting...":"Submit Checklist"}
                  </Btn>
                </div>
              </div>}
            </div>;
          })}
        </div>}

      {/* ── Completed Checklists ── */}
      <Section icon="checkCircle" count={completed.length}>Completed Checklists</Section>
      {completed.length===0?<Empty icon="clock" text="No checklists completed yet"/>:
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {completed.map(item=>{
            const ck=getCk(item.checklistId);if(!ck)return null;
            const isViewing=viewingId===item.id;
            return <div key={item.id} style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.successBorder}`,transition:"all .2s"}}>
              <div onClick={()=>handleViewResponses(item)} style={{display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer"}}>
                <div style={{display:"flex",alignItems:"center",gap:10}}>
                  <div style={{width:24,height:24,borderRadius:"50%",background:T.successBg,display:"flex",alignItems:"center",justifyContent:"center"}}><Icon name="check" size={14} color={T.success}/></div>
                  <div><span style={{fontSize:14,fontWeight:500,color:T.text}}>{ck.name}</span>
                    <div style={{display:"flex",gap:8,marginTop:2}}>
                      {item.completedBy&&<span style={{fontSize:11,color:T.textMut}}>by {item.completedBy}</span>}
                      {item.completedAt&&<span style={{fontSize:11,color:T.textMut}}>{formatDateTime(item.completedAt)}</span>}
                    </div>
                  </div>
                </div>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <Badge variant="success">Submitted</Badge>
                  <Icon name="chevron" size={16} color={T.textMut} style={{transform:isViewing?"rotate(90deg)":"rotate(0)",transition:"transform .2s"}}/>
                </div>
              </div>

              {isViewing&&<div onClick={e=>e.stopPropagation()} style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
                {loadingView?<p style={{fontSize:13,color:T.textMut,textAlign:"center",padding:16}}>Loading responses...</p>:
                viewData&&viewData.responses.length>0?<>
                  {(viewData.workDate||viewData.person||viewData.autoId)&&<div style={{display:"flex",gap:16,marginBottom:12,flexWrap:"wrap"}}>
                    {viewData.autoId&&<div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Auto ID</span><span style={{fontSize:12,fontFamily:T.mono,color:T.accent,fontWeight:600}}>{viewData.autoId}</span></div>}
                    {viewData.workDate&&<div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Date</span><span style={{fontSize:13,color:T.textSec}}>{viewData.workDate}</span></div>}
                    {viewData.person&&<div><span style={{fontSize:11,color:T.textMut,display:"block"}}>Person</span><span style={{fontSize:13,color:T.textSec}}>{viewData.person}</span></div>}
                  </div>}
                  {viewData.lastEditedBy && viewData.lastEditedAt && (
                    <div style={{fontSize:11,color:T.textMut,marginBottom:10}}>Last edited by <b style={{color:T.textSec}}>{viewData.lastEditedBy}</b> at {formatDateTime(viewData.lastEditedAt)}</div>
                  )}
                  <div style={{background:T.bg,borderRadius:T.radSm,padding:"10px 12px",maxHeight:300,overflowY:"auto"}}>
                    {normalizeQuestions(ck.questions).map((q,qi)=>{
                      const resp=viewData.responses.find(r=>r.questionIndex===qi);
                      return <div key={qi} style={{padding:"8px 0",borderBottom:qi<ck.questions.length-1?`1px solid ${T.border}`:"none"}}>
                        <div style={{display:"flex",alignItems:"flex-start",gap:8,marginBottom:4}}>
                          <span style={{fontSize:11,color:T.textMut,fontFamily:T.mono,flexShrink:0}}>{String(qi+1).padStart(2,"0")}</span>
                          <span style={{fontSize:12,color:T.textMut}}>{q.text}</span>
                        </div>
                        <div style={{paddingLeft:26}}><span style={{fontSize:14,color:T.text,fontWeight:500}}>{displayResponseValue(q,resp?.response,inventoryItems)}</span></div>
                        {resp?.remark&&<div style={{paddingLeft:26,marginTop:4,padding:"4px 8px",background:T.warningBg,borderRadius:T.radSm,border:`1px solid ${T.warningBorder}`}}>
                          <span style={{fontSize:11,color:T.warning}}>Remark: {resp.remark}</span>
                        </div>}
                        {q.linkedSource?.checklistId && resp?.response && <div style={{paddingLeft:26}}><SourceChainDisplay checklistId={q.linkedSource.checklistId} autoId={String(resp.response).split(",")[0].trim()} checklists={checklists}/></div>}
                      </div>;
                    })}
                  </div>
                </>:<p style={{fontSize:13,color:T.textMut,textAlign:"center",padding:16}}>No in-app responses recorded (may have been filled via Google Form)</p>}

                {/* Admin actions on completed checklists */}
                {isAdmin && <div style={{display:"flex",gap:8,marginTop:12}}>
                  {viewData?.accessControl?.isTaggedToStage ? (
                    <Btn variant="secondary" small disabled style={{flex:1,opacity:0.6}} title="Untag from order stage to edit">
                      <Icon name="lock" size={14} color={T.textMut}/> Tagged (read-only)
                    </Btn>
                  ) : (
                    <Btn variant="secondary" small onClick={()=>onEditResponses(item.id,item.checklistId)} style={{flex:1}}>
                      <Icon name="edit" size={14} color={T.text}/> Edit Responses
                    </Btn>
                  )}
                  <Btn variant="danger" small onClick={()=>{if(confirm("Revert to pending? All responses will be deleted."))onRevertChecklist(item.id)}} style={{flex:1}}>
                    <Icon name="undo" size={14} color={T.danger}/> Revert to Pending
                  </Btn>
                </div>}
              </div>}
            </div>;
          })}
        </div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Edit Response View ───────────────────────────────────────

// Shared inventory-field validation for submit / save. Returns { ok, message, firstIdx }.
function validateRequiredInventoryFields(nq, responses, batchAllocations) {
  for (let qi = 0; qi < nq.length; qi++) {
    const q = nq[qi];
    const isInLinked = q.inventoryLink && q.inventoryLink.enabled && q.inventoryLink.txType === "IN";
    const isMasterQty = !!q.isMasterQuantity;
    if (!isInLinked && !isMasterQty) continue;
    // Formula-derived fields: check computed value
    let val = responses?.[qi];
    if (val === undefined || val === null) val = "";
    if (isMasterQty) {
      const allocs = batchAllocations?.[qi];
      const hasBatch = Array.isArray(allocs) && allocs.some(a => a.sourceAutoId && (parseFloat(a.quantity) || 0) > 0);
      const numVal = parseFloat(val) || 0;
      if (!hasBatch && !(numVal > 0)) {
        return { ok: false, firstIdx: qi, message: "This field is required for inventory tracking — please enter a quantity" };
      }
      continue;
    }
    // IN-linked: numeric > 0 required
    const numVal = parseFloat(val) || 0;
    if (!(numVal > 0)) {
      return { ok: false, firstIdx: qi, message: "This field is required for inventory tracking — please enter a quantity" };
    }
  }
  return { ok: true };
}

// Label an inventory-tracking field with a red asterisk when required.
function isInventoryRequiredQuestion(q) {
  return !!(q && ((q.inventoryLink && q.inventoryLink.enabled && q.inventoryLink.txType === "IN") || q.isMasterQuantity));
}

function EditResponseView({ orderChecklistId, checklistId, checklists, approvedEntries, inventoryItems, customers, isUntagged, onSave, onCancel }) {
  const [loading,setLoading]=useState(true);
  const [formData,setFormData]=useState({date:"",person:"",responses:{},remarks:{},batchAllocations:{}});
  const [saving,setSaving]=useState(false);
  const [meta,setMeta]=useState({autoId:"",lastEditedBy:"",lastEditedAt:"",accessControl:{isTaggedToStage:false,stageRefs:[],linkedByOthers:[],editable:true}});
  const [error,setError]=useState("");
  const [inventoryError,setInventoryError]=useState({idx:null,message:""});
  const ck = checklists.find(c => c.id === checklistId);
  const nq = ck ? normalizeQuestions(ck.questions) : [];

  useEffect(() => {
    const action = isUntagged ? "getUntaggedResponse" : "getResponses";
    API.get(action, { id: orderChecklistId }).then(data => {
      if (data && !data.error) {
        const map = {}, rmk = {};
        if (data.responses) {
          // Match by question text (stable across template reorders), falling back
          // to question index only when the text lookup doesn't hit.
          const textToIdx = {};
          nq.forEach((q, qi) => { if (q.text) textToIdx[q.text] = qi; });
          data.responses.forEach(r => {
            let idx = (r.questionText && textToIdx[r.questionText] !== undefined)
              ? textToIdx[r.questionText]
              : r.questionIndex;
            map[idx] = r.response || "";
            if (r.remark) rmk[idx] = r.remark;
          });
        }
        setFormData({ date: data.workDate || data.date || "", person: data.person || "", responses: map, remarks: rmk, batchAllocations: {} });
        setMeta({
          autoId: data.autoId || "",
          lastEditedBy: data.lastEditedBy || "",
          lastEditedAt: data.lastEditedAt || "",
          accessControl: data.accessControl || {isTaggedToStage:false,stageRefs:[],linkedByOthers:[],editable:true},
        });
      } else if (data && data.error) {
        setError(data.error);
      }
      setLoading(false);
    }).catch((err) => { setError(err.message || "Failed to load"); setLoading(false); });
  }, [orderChecklistId, isUntagged]);

  const getFieldValue = (checklistRef, questionIdx) => {
    if (checklistRef === "self") return formData.responses[questionIdx] || "";
    return "";
  };

  const readOnly = !!meta.accessControl?.isTaggedToStage;

  const handleSave = async () => {
    if (!ck || readOnly) return;
    setInventoryError({idx:null,message:""});
    // Compute responses including formula-derived values for validation
    const computedResponses = {};
    nq.forEach((q, qi) => {
      if (q.formula) {
        const computed = evaluateFormula(q.formula, getFieldValue);
        computedResponses[qi] = computed !== null ? String(computed) : "";
      } else {
        computedResponses[qi] = formData.responses[qi] !== undefined ? formData.responses[qi] : "";
      }
    });
    const v = validateRequiredInventoryFields(nq, computedResponses, formData.batchAllocations);
    if (!v.ok) {
      setInventoryError({idx: v.firstIdx, message: v.message});
      return;
    }
    setSaving(true);
    const respArray = nq.map((q, qi) => {
      let resp = computedResponses[qi];
      if (q.linkedSource && Array.isArray(formData.batchAllocations?.[qi])) {
        const ids = formData.batchAllocations[qi].map(a => a.sourceAutoId).filter(Boolean);
        if (ids.length > 0) resp = ids.join(", ");
      }
      return { questionIndex: qi, questionText: q.text, response: resp };
    });
    await onSave({ id: orderChecklistId, date: formData.date, person: formData.person, responses: respArray, remarks: formData.remarks, isUntagged: !!isUntagged });
    setSaving(false);
  };

  if (loading) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.textSec,animation:"pulse 1.5s infinite"}}>Loading responses...</p></div>;
  if (error) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.danger}}>{error}</p></div>;
  if (!ck) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.danger}}>Checklist template not found (id: {checklistId || "missing"})</p></div>;
  if (!orderChecklistId) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.danger}}>Cannot edit — missing id. Open this entry from the order page instead.</p></div>;

  const linkedBy = meta.accessControl?.linkedByOthers || [];
  const stageRefs = meta.accessControl?.stageRefs || [];

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
          <div>
            <h3 style={{fontSize:15,fontWeight:600,color:T.text,marginBottom:4}}>{ck.name}</h3>
            {ck.subtitle && <p style={{fontSize:12,color:T.accent}}>{ck.subtitle}</p>}
          </div>
          {readOnly && <div style={{display:"flex",alignItems:"center",gap:4,color:T.textMut,fontSize:11}}><Icon name="lock" size={14} color={T.textMut}/> Read-only</div>}
        </div>
        {meta.autoId && (
          <div style={{marginTop:10,padding:"8px 10px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:8}}>
            <Icon name="lock" size={12} color={T.textMut}/>
            <span style={{fontSize:11,color:T.textMut}}>Auto ID (read-only):</span>
            <span style={{fontSize:13,fontFamily:T.mono,color:T.accent,fontWeight:600}}>{meta.autoId}</span>
          </div>
        )}
        {meta.lastEditedBy && meta.lastEditedAt && (
          <div style={{marginTop:8,fontSize:11,color:T.textMut}}>
            Last edited by <b style={{color:T.textSec}}>{meta.lastEditedBy}</b> at {formatDateTime(meta.lastEditedAt)}
          </div>
        )}
      </div>

      {readOnly && (
        <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"12px 14px",display:"flex",gap:10,alignItems:"flex-start"}}>
          <Icon name="lock" size={16} color={T.danger}/>
          <div>
            <div style={{fontSize:13,color:T.danger,fontWeight:600,marginBottom:2}}>This entry is tagged to an order stage</div>
            <div style={{fontSize:12,color:T.textSec}}>Tagged to: {stageRefs.join(", ")}. Untag it from the order to edit.</div>
          </div>
        </div>
      )}

      {!readOnly && linkedBy.length > 0 && (
        <div style={{background:T.warningBg,border:`1px solid ${T.warningBorder}`,borderRadius:T.radSm,padding:"12px 14px",display:"flex",gap:10,alignItems:"flex-start"}}>
          <Icon name="alert-triangle" size={16} color={T.warning}/>
          <div>
            <div style={{fontSize:13,color:T.warning,fontWeight:600,marginBottom:2}}>This entry is referenced by downstream responses</div>
            <div style={{fontSize:12,color:T.textSec}}>Referenced by: {linkedBy.join(", ")}. Changes will affect linked data.</div>
          </div>
        </div>
      )}

      {inventoryError.message && (
        <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"10px 14px"}}>
          <span style={{fontSize:13,color:T.danger}}>{inventoryError.message}</span>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <Field label="Date">
          <input type="date" value={formData.date} onChange={e=>!readOnly&&setFormData(p=>({...p,date:e.target.value}))} disabled={readOnly}
            style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:readOnly?T.surfaceHover:T.bg,border:`1px solid ${T.border}`,color:readOnly?T.textMut:T.text,fontSize:14,outline:"none",colorScheme:"dark",cursor:readOnly?"not-allowed":"text"}}/>
        </Field>
        <Field label="Person"><Input value={formData.person} onChange={()=>{}} readOnly placeholder="Person..."/></Field>
      </div>

      <div style={{display:"flex",flexDirection:"column",gap:14, pointerEvents: readOnly ? "none" : "auto", opacity: readOnly ? 0.7 : 1}}>
        {nq.map((q, qi) => {
          let autoVal = null;
          if (q.formula) autoVal = evaluateFormula(q.formula, getFieldValue);
          const currentVal = q.formula
            ? (autoVal !== null ? String(autoVal) : "")
            : (formData.responses[qi] !== undefined ? formData.responses[qi] : "");
          let idealVal = null;
          if (q.ideal) idealVal = evaluateFormula(q.ideal, getFieldValue);
          const needsRemark = q.remarkCondition && idealVal !== null && checkRemarkCondition(currentVal, idealVal, q.remarkCondition);
          const required = isInventoryRequiredQuestion(q);
          const invalid = inventoryError.idx === qi;
          return (
            <div key={qi} style={invalid ? {border:`1px solid ${T.danger}`,borderRadius:T.radSm,padding:8} : undefined}>
              {required && (
                <div style={{fontSize:11,color:T.danger,marginBottom:4,display:"flex",alignItems:"center",gap:4}}>
                  <span style={{color:T.danger,fontWeight:700}}>*</span> Required for inventory tracking
                </div>
              )}
              <QuestionInputField q={q} qi={qi} currentVal={currentVal} idealVal={idealVal} needsRemark={needsRemark}
                formData={formData} setFormData={setFormData} approvedEntries={approvedEntries} checklists={checklists} getFieldValue={getFieldValue}
                customers={customers} inventoryItems={inventoryItems} allQuestions={nq}/>
            </div>
          );
        })}
      </div>

      <div style={{display:"flex",gap:8}}>
        <Btn variant="secondary" onClick={onCancel} style={{flex:1}}>{readOnly ? "Close" : "Cancel"}</Btn>
        {!readOnly && (
          <Btn onClick={handleSave} disabled={saving} style={{flex:1}}>
            {saving ? "Saving..." : "Save Changes"}
          </Btn>
        )}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Responses Log View (Admin) ───────────────────────────────

function ResponsesLogView({ checklists, inventoryItems, isAdmin, addToast, onEditResponses, onRevertChecklist }) {
  const [entries,setEntries]=useState([]);
  const [loading,setLoading]=useState(true);
  const [expandedId,setExpandedId]=useState(null);
  const [reverting,setReverting]=useState(null);

  useEffect(() => {
    API.get("getAllResponses").then(data => { setEntries(data || []); setLoading(false); }).catch((e) => { setLoading(false); addToast("Failed to load responses", "error"); });
  }, []);

  const handleRevert = async (ocId) => {
    if (!confirm("Revert this checklist to pending? All responses will be deleted.")) return;
    setReverting(ocId);
    try { await onRevertChecklist(ocId); setEntries(prev => prev.filter(e => e.orderChecklistId !== ocId)); } catch {}
    setReverting(null);
  };

  if (loading) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.textSec,animation:"pulse 1.5s infinite"}}>Loading responses...</p></div>;

  return (
    <div className="fade-up">
      <p style={{fontSize:13,color:T.textSec,marginBottom:16}}>All submitted checklist responses across all orders.</p>
      {entries.length===0 ? <Empty icon="clipboard" text="No completed checklists yet" sub="Responses will appear here after checklists are submitted"/> :
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {entries.map(entry => {
          const isExpanded = expandedId === entry.orderChecklistId;
          return (
            <div key={entry.orderChecklistId} style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.border}`}}>
              <div onClick={() => setExpandedId(isExpanded ? null : entry.orderChecklistId)} style={{cursor:"pointer"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div>
                    <span style={{fontSize:14,fontWeight:600,color:T.text}}>{entry.checklistName}</span>
                    <div style={{display:"flex",gap:8,marginTop:4,flexWrap:"wrap"}}>
                      <Badge>{entry.orderName}</Badge>
                      {entry.customer && <Badge variant="info">{entry.customer}</Badge>}
                    </div>
                    <div style={{display:"flex",gap:10,marginTop:4,flexWrap:"wrap"}}>
                      {entry.autoId && <span style={{fontSize:11,fontFamily:T.mono,color:T.accent}}>{entry.autoId}</span>}
                      {entry.person && <span style={{fontSize:11,color:T.textMut}}>by {entry.person}</span>}
                      {entry.date && <span style={{fontSize:11,color:T.textMut}}>{entry.date}</span>}
                      {entry.editedAt && <Badge variant="muted" style={{fontSize:9,padding:"1px 6px"}}>Edited</Badge>}
                    </div>
                    {entry.editedBy && entry.editedAt && (
                      <div style={{fontSize:10,color:T.textMut,marginTop:2}}>Last edited by {entry.editedBy} at {formatDateTime(entry.editedAt)}</div>
                    )}
                  </div>
                  <Icon name="chevron" size={16} color={T.textMut} style={{transform:isExpanded?"rotate(90deg)":"rotate(0)",transition:"transform .2s",flexShrink:0,marginTop:4}}/>
                </div>
              </div>

              {isExpanded && (
                <div style={{marginTop:12,paddingTop:12,borderTop:`1px solid ${T.border}`}}>
                  {entry.responses.length > 0 ? (
                    <div style={{background:T.bg,borderRadius:T.radSm,padding:"10px 12px",marginBottom:12,maxHeight:300,overflowY:"auto"}}>
                      {(()=>{
                        const ck=checklists.find(c=>c.name===entry.checklistName);
                        const nq=ck?normalizeQuestions(ck.questions):[];
                        return entry.responses.map((r, i) => {
                          const q = nq.find(qq => qq.text === r.question);
                          return (
                            <div key={i} style={{padding:"6px 0",borderBottom:i<entry.responses.length-1?`1px solid ${T.border}`:"none"}}>
                              <span style={{fontSize:12,color:T.textMut}}>{r.question}</span>
                              <div style={{display:"flex",alignItems:"center",gap:8,marginTop:2,flexWrap:"wrap"}}>
                                <span style={{fontSize:14,color:T.text,fontWeight:500}}>{displayResponseValue(q, r.response, inventoryItems)}</span>
                                {r.originalResponse && <Badge variant="muted" style={{fontSize:9,padding:"1px 6px"}}>was: {displayResponseValue(q, r.originalResponse, inventoryItems)}</Badge>}
                                {r.remark && <span style={{fontSize:11,color:T.warning,background:T.warningBg,padding:"2px 6px",borderRadius:8}}>Remark: {r.remark}</span>}
                              </div>
                            </div>
                          );
                        });
                      })()}
                    </div>
                  ) : <p style={{fontSize:13,color:T.textMut,marginBottom:12}}>No responses recorded</p>}
                  {isAdmin && <div style={{display:"flex",gap:8}}>
                    {entry.accessControl?.isTaggedToStage ? (
                      <Btn variant="secondary" small disabled style={{flex:1,opacity:0.6}} title="Untag from order to edit">
                        <Icon name="lock" size={14} color={T.textMut}/> Tagged (read-only)
                      </Btn>
                    ) : (
                      <Btn variant="secondary" small onClick={() => onEditResponses(entry.orderChecklistId, entry.checklistName)} style={{flex:1}}>
                        <Icon name="edit" size={14} color={T.text}/> Edit
                      </Btn>
                    )}
                    <Btn variant="danger" small onClick={() => handleRevert(entry.orderChecklistId)} disabled={reverting===entry.orderChecklistId} style={{flex:1}}>
                      <Icon name="undo" size={14} color={T.danger}/> {reverting===entry.orderChecklistId ? "Reverting..." : "Revert to Pending"}
                    </Btn>
                  </div>}
                </div>
              )}
            </div>
          );
        })}
      </div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Delete Confirmation Modal ───────────────────────────────

function DeleteConfirmModal({ entryId, entityType, onConfirm, onCancel }) {
  const [reason, setReason] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const handleDelete = async () => {
    if (!reason.trim()) { setError("Please provide a reason for deletion."); return; }
    setBusy(true); setError("");
    try {
      const r = await API.post("softDeleteChecklist", { id: entryId, entityType, reason: reason.trim() });
      if (r?.error) { setError(r.error); setBusy(false); return; }
      onConfirm(r);
    } catch (e) { setError(e.message); }
    setBusy(false);
  };
  return (
    <div style={{position:"fixed",inset:0,zIndex:999,background:"rgba(0,0,0,0.6)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={onCancel}>
      <div style={{background:T.card,borderRadius:T.rad,padding:20,maxWidth:420,width:"100%",border:`1px solid ${T.border}`}} onClick={e=>e.stopPropagation()}>
        <h3 style={{fontSize:16,fontWeight:600,color:T.danger,marginBottom:8}}>Delete {entryId}?</h3>
        <p style={{fontSize:13,color:T.textSec,marginBottom:12}}>This action cannot be undone from the main app. All inventory movements from this entry will be reversed.</p>
        {error && <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"8px 12px",marginBottom:10}}><span style={{fontSize:12,color:T.danger}}>{error}</span></div>}
        <Field label="Reason for deletion (required)">
          <textarea value={reason} onChange={e=>setReason(e.target.value)} placeholder="Why is this being deleted?" rows={2}
            style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none",resize:"vertical",fontFamily:"inherit"}}/>
        </Field>
        <div style={{display:"flex",gap:8,marginTop:12}}>
          <Btn variant="secondary" onClick={onCancel} style={{flex:1}} disabled={busy}>Cancel</Btn>
          <Btn variant="danger" onClick={handleDelete} disabled={busy||!reason.trim()} style={{flex:1,background:T.danger,color:"#fff"}}>{busy?"Deleting...":"Delete"}</Btn>
        </div>
      </div>
    </div>
  );
}

// ─── Audit Log Viewer (used in Settings) ─────────────────────

function AuditLogSection() {
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(true);
  const [total, setTotal] = useState(0);
  const [hasMore, setHasMore] = useState(false);
  const [offset, setOffset] = useState(0);
  const [filterAction, setFilterAction] = useState("");
  const [filterSearch, setFilterSearch] = useState("");
  const [expandedId, setExpandedId] = useState(null);

  const load = (off, append) => {
    setLoading(true);
    const params = { offset: String(off), limit: "50" };
    if (filterAction) params.action = filterAction;
    if (filterSearch) params.entityId = filterSearch;
    API.get("getAuditLog", params).then(data => {
      if (data && !data.error) {
        setEntries(prev => append ? [...prev, ...(data.entries || [])] : (data.entries || []));
        setTotal(data.total || 0);
        setHasMore(data.hasMore || false);
      }
      setLoading(false);
    }).catch(() => setLoading(false));
  };

  useEffect(() => { setOffset(0); load(0, false); }, [filterAction, filterSearch]);

  const actionColors = { delete: T.danger, edit: T.warning, edit_response: T.warning, submit: T.success, tag: T.success, create: T.info, deactivate: T.textMut, revert: T.danger };

  return (
    <div>
      <Section icon="clipboard">Audit Log</Section>
      <p style={{fontSize:12,color:T.textMut,marginTop:-10,marginBottom:12}}>Activity log for all checklist, inventory, and admin actions.</p>
      <div style={{display:"flex",gap:8,marginBottom:12}}>
        <select value={filterAction} onChange={e=>{setFilterAction(e.target.value)}}
          style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
          <option value="">All Actions</option>
          <option value="submit">Submit</option>
          <option value="edit">Edit</option>
          <option value="edit_response">Edit Response</option>
          <option value="delete">Delete</option>
          <option value="tag">Tag</option>
          <option value="create">Create</option>
          <option value="revert">Revert</option>
        </select>
        <input value={filterSearch} onChange={e=>setFilterSearch(e.target.value)} placeholder="Search by entry ID..."
          style={{flex:1,padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12,outline:"none"}}/>
      </div>
      {loading && entries.length === 0 ? <p style={{fontSize:12,color:T.textMut,textAlign:"center",padding:20}}>Loading audit log...</p> : (
        <div style={{display:"flex",flexDirection:"column",gap:4}}>
          {entries.length === 0 && <p style={{fontSize:12,color:T.textMut,textAlign:"center",padding:20}}>No entries found.</p>}
          {entries.map(e => (
            <div key={e.id} style={{background:T.card,borderRadius:T.radSm,border:`1px solid ${T.border}`,overflow:"hidden"}}>
              <button onClick={()=>setExpandedId(expandedId===e.id?null:e.id)}
                style={{width:"100%",padding:"8px 12px",background:"none",border:"none",cursor:"pointer",textAlign:"left",display:"flex",alignItems:"center",gap:8}}>
                <span style={{fontSize:10,fontFamily:T.mono,color:T.textMut,minWidth:60}}>{formatDate(e.timestamp)}</span>
                <Badge variant="muted" style={{background:actionColors[e.action]?actionColors[e.action]+"20":"transparent",color:actionColors[e.action]||T.textSec,fontSize:10}}>{e.action}</Badge>
                <span style={{fontSize:12,fontFamily:T.mono,color:T.accent,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{e.entityId}</span>
                <span style={{fontSize:11,color:T.textMut}}>{e.performedBy}</span>
              </button>
              {expandedId === e.id && (
                <div style={{padding:"8px 12px",background:T.bg,borderTop:`1px solid ${T.border}`,fontSize:12}}>
                  <div style={{color:T.textMut,marginBottom:4}}>
                    <b>Time:</b> {formatDateTime(e.timestamp)}
                  </div>
                  <div style={{color:T.textMut,marginBottom:4}}><b>Type:</b> {e.entityType}</div>
                  {e.details && <div style={{color:T.textSec,whiteSpace:"pre-wrap",fontFamily:T.mono,fontSize:11,maxHeight:200,overflowY:"auto",background:T.surfaceHover,padding:8,borderRadius:T.radSm,marginTop:4}}>{e.details}</div>}
                </div>
              )}
            </div>
          ))}
          {hasMore && <Btn variant="ghost" small onClick={()=>{const no=offset+50;setOffset(no);load(no,true)}} disabled={loading}>{loading?"Loading...":"Load More"}</Btn>}
          {total > 0 && <p style={{fontSize:11,color:T.textMut,textAlign:"center"}}>Showing {entries.length} of {total}</p>}
        </div>
      )}
    </div>
  );
}

// ─── Classifications Management (used in Settings) ────���──────

function ClassificationsSection({ addToast }) {
  const [data, setData] = useState({ roast_degree: [], grind_size: [] });
  const [loading, setLoading] = useState(true);
  const [addForm, setAddForm] = useState(null); // { type, name, description }
  const [editId, setEditId] = useState(null);
  const [editName, setEditName] = useState("");
  const [editDesc, setEditDesc] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const load = () => API.get("getClassifications").then(d => { if (d && !d.error) setData(d); }).catch(() => {}).finally(() => setLoading(false));
  useEffect(() => { load(); }, []);

  const handleAdd = async () => {
    if (!addForm || !addForm.name.trim()) return;
    setSaving(true); setError("");
    try {
      await API.post("addClassification", { name: addForm.name.trim(), type: addForm.type, description: addForm.description || "" });
      setAddForm(null); await load(); addToast("Classification added", "success");
    } catch (e) { setError(e.message); }
    setSaving(false);
  };

  const handleEdit = async (id) => {
    if (!editName.trim()) return;
    setSaving(true); setError("");
    try {
      await API.post("editClassification", { id, name: editName.trim(), description: editDesc });
      setEditId(null); await load(); addToast("Classification updated", "success");
    } catch (e) { setError(e.message); }
    setSaving(false);
  };

  const handleDeactivate = async (id, name) => {
    if (!confirm(`Deactivate "${name}"? This cannot be undone if used in inventory.`)) return;
    setSaving(true); setError("");
    try {
      await API.post("deactivateClassification", { id });
      await load(); addToast("Classification deactivated", "success");
    } catch (e) { setError(e.message); }
    setSaving(false);
  };

  const renderGroup = (type, label) => {
    const items = data[type] || [];
    return (
      <div style={{background:T.card,borderRadius:T.rad,padding:14,border:`1px solid ${T.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <span style={{fontSize:14,fontWeight:600,color:T.text}}>{label}</span>
          <Badge variant="muted">{items.length}</Badge>
        </div>
        {items.length === 0 && <p style={{fontSize:12,color:T.textMut,marginBottom:8}}>No classifications defined.</p>}
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {items.map(c => (
            <div key={c.id} style={{display:"flex",alignItems:"center",gap:8,background:T.bg,borderRadius:T.radSm,padding:"8px 12px",border:`1px solid ${T.border}`}}>
              {editId === c.id ? (
                <div style={{flex:1,display:"flex",flexDirection:"column",gap:6}}>
                  <input value={editName} onChange={e=>setEditName(e.target.value)} placeholder="Name..."
                    style={{padding:"6px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
                  <input value={editDesc} onChange={e=>setEditDesc(e.target.value)} placeholder="Description (optional)..."
                    style={{padding:"6px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12,outline:"none"}}/>
                  <div style={{display:"flex",gap:6}}>
                    <Btn small onClick={()=>handleEdit(c.id)} disabled={saving||!editName.trim()}>{saving?"Saving...":"Save"}</Btn>
                    <Btn small variant="ghost" onClick={()=>setEditId(null)}>Cancel</Btn>
                  </div>
                </div>
              ) : (
                <>
                  <div style={{flex:1}}>
                    <span style={{fontSize:13,fontWeight:500,color:T.text}}>{c.name}</span>
                    {c.description && <span style={{display:"block",fontSize:11,color:T.textMut,marginTop:2}}>{c.description}</span>}
                  </div>
                  <button onClick={()=>{setEditId(c.id);setEditName(c.name);setEditDesc(c.description||"")}} style={{background:"none",border:"none",cursor:"pointer",padding:4}}>
                    <Icon name="edit" size={14} color={T.textSec}/>
                  </button>
                  <button onClick={()=>handleDeactivate(c.id,c.name)} disabled={saving} style={{background:"none",border:"none",cursor:"pointer",padding:4}}>
                    <Icon name="trash" size={14} color={T.danger}/>
                  </button>
                </>
              )}
            </div>
          ))}
        </div>
        {addForm && addForm.type === type ? (
          <div style={{marginTop:8,background:T.bg,borderRadius:T.radSm,padding:10,border:`1px solid ${T.accentBorder}`,display:"flex",flexDirection:"column",gap:6}}>
            <input value={addForm.name} onChange={e=>setAddForm(p=>({...p,name:e.target.value}))} placeholder="Name..." autoFocus
              style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
            <input value={addForm.description} onChange={e=>setAddForm(p=>({...p,description:e.target.value}))} placeholder="Description (optional)..."
              style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12,outline:"none"}}/>
            <div style={{display:"flex",gap:6}}>
              <Btn small onClick={handleAdd} disabled={saving||!addForm.name.trim()}>{saving?"Adding...":"Save"}</Btn>
              <Btn small variant="ghost" onClick={()=>{setAddForm(null);setError("")}}>Cancel</Btn>
            </div>
          </div>
        ) : (
          <Btn variant="ghost" small onClick={()=>{setAddForm({type,name:"",description:""});setError("")}} style={{marginTop:8}}>
            <Icon name="plus" size={12} color={T.textSec}/> Add New
          </Btn>
        )}
      </div>
    );
  };

  if (loading) return <div style={{padding:20,textAlign:"center"}}><p style={{color:T.textMut,fontSize:13}}>Loading classifications...</p></div>;

  return (
    <div>
      <Section icon="layers">Roast & Grind Classifications</Section>
      <p style={{fontSize:12,color:T.textMut,marginTop:-10,marginBottom:12}}>Manage roast degree and grind size classifications used in inventory items.</p>
      {error && <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"8px 12px",marginBottom:10}}>
        <span style={{fontSize:12,color:T.danger}}>{error}</span>
      </div>}
      <div style={{display:"flex",flexDirection:"column",gap:12}}>
        {renderGroup("roast_degree", "Roast Degrees")}
        {renderGroup("grind_size", "Grind Sizes")}
      </div>
    </div>
  );
}

// ─── Admin / Settings View ────────────────────────────────────

function AdminView({checklists,orderTypes,customers,rules,isAdmin,addToast,onEditChecklist,onNewChecklist,onEditRules,onDeleteChecklist,onAddOrderType,onDeleteOrderType,onAddCustomer,onDeleteCustomer,onArchive,orderStageTemplates,onSaveOrderStageTemplates}){
  const [newType,setNewType]=useState("");
  const [newCust,setNewCust]=useState("");
  const [archiveDays,setArchiveDays]=useState("30");
  const [archiving,setArchiving]=useState(false);
  const [archiveMsg,setArchiveMsg]=useState(null);
  const [archives,setArchives]=useState([]);
  // Order Stage Templates state
  const [stageTpl, setStageTpl] = useState(orderStageTemplates || {});
  const [stageTplSaving, setStageTplSaving] = useState(false);
  useEffect(()=>setStageTpl(orderStageTemplates || {}), [orderStageTemplates]);

  const addStage = (pt) => {
    setStageTpl(prev => {
      const arr = [...(prev[pt] || [])];
      arr.push({ name: "Stage " + (arr.length + 1), checklistId: "", quantityField: "", requiredQty: 0, position: arr.length });
      return { ...prev, [pt]: arr };
    });
  };
  const removeStage = (pt, idx) => {
    setStageTpl(prev => {
      const arr = (prev[pt] || []).filter((_,i)=>i!==idx);
      return { ...prev, [pt]: arr.map((s,i)=>({...s,position:i})) };
    });
  };
  const updateStage = (pt, idx, patch) => {
    setStageTpl(prev => {
      const arr = (prev[pt] || []).map((s,i)=>i===idx?{...s,...patch}:s);
      return { ...prev, [pt]: arr };
    });
  };
  const moveStage = (pt, idx, dir) => {
    setStageTpl(prev => {
      const arr = [...(prev[pt] || [])];
      const ni = idx + dir;
      if (ni < 0 || ni >= arr.length) return prev;
      const tmp = arr[idx]; arr[idx] = arr[ni]; arr[ni] = tmp;
      return { ...prev, [pt]: arr.map((s,i)=>({...s,position:i})) };
    });
  };
  const handleSaveStageTpl = async () => {
    setStageTplSaving(true);
    try { await onSaveOrderStageTemplates(stageTpl); addToast("Stage templates saved", "success"); } catch(e) { addToast(e.message, "error"); }
    setStageTplSaving(false);
  };

  useEffect(() => {
    if (isAdmin) API.get("getArchives").then(setArchives).catch(() => {});
  }, [isAdmin]);

  const handleArchive = async () => {
    if (!confirm(`Archive completed orders older than ${archiveDays} days?`)) return;
    setArchiving(true); setArchiveMsg(null);
    try {
      const r = await onArchive(archiveDays);
      setArchiveMsg(`Archived ${r.archived} orders successfully.`);
      API.get("getArchives").then(setArchives).catch(() => {});
    } catch (e) { setArchiveMsg(null); }
    setArchiving(false);
  };

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:28}}>
      <div>
        <Section icon="clipboard" count={checklists.length} action={<Btn small onClick={onNewChecklist}><Icon name="plus" size={14} color={T.bg}/> New</Btn>}>Checklist Templates</Section>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {checklists.map(ck=><div key={ck.id} style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
              <div style={{flex:1}}>
                <span style={{fontSize:14,fontWeight:600,color:T.text}}>{ck.name}</span>
                {ck.subtitle&&<span style={{display:"block",fontSize:12,color:T.accent,marginTop:2}}>{ck.subtitle}</span>}
                <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                  <Badge variant="muted">{ck.questions.length} questions</Badge>
                  {ck.formUrl?<Badge variant="success"><Icon name="link" size={11} color={T.success}/> Form linked</Badge>:<Badge variant="danger">No form</Badge>}
                </div>
              </div>
              <div style={{display:"flex",gap:4}}>
                <button onClick={()=>onEditChecklist(ck)} style={{background:"none",border:"none",cursor:"pointer",padding:6,borderRadius:6}}><Icon name="edit" size={16} color={T.textSec}/></button>
                <button onClick={()=>{if(confirm("Delete this checklist?"))onDeleteChecklist(ck.id)}} style={{background:"none",border:"none",cursor:"pointer",padding:6,borderRadius:6}}><Icon name="trash" size={16} color={T.danger}/></button>
              </div>
            </div>
          </div>)}
        </div>
      </div>

      <div>
        <Section icon="layers" count={rules.length} action={<Btn small variant="secondary" onClick={onEditRules}><Icon name="edit" size={14} color={T.text}/> Manage</Btn>}>Assignment Rules</Section>
        <p style={{fontSize:12,color:T.textMut,marginTop:-10,marginBottom:12}}>Rules determine which checklists auto-assign based on order type AND/OR customer.</p>
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {rules.slice(0,5).map(r=>{
            const ot=orderTypes.find(t=>t.id===r.orderTypeId);const cu=customers.find(c=>c.id===r.customerId);
            return <div key={r.id} style={{background:T.card,borderRadius:T.radSm,padding:"10px 14px",border:`1px solid ${T.border}`,fontSize:13}}>
              <span style={{color:T.accent,fontWeight:600}}>{r.orderTypeId==="any"?"Any Type":ot?.label||"?"}</span>
              <span style={{color:T.textMut}}> + </span>
              <span style={{color:T.info,fontWeight:500}}>{r.customerId==="any"?"Any Customer":cu?.label||"?"}</span>
              <span style={{color:T.textMut}}> → </span>
              <span style={{color:T.textSec}}>{r.checklistIds.length} checklists</span>
            </div>;
          })}
          {rules.length>5&&<span style={{fontSize:12,color:T.textMut}}>+{rules.length-5} more rules</span>}
        </div>
      </div>

      <div>
        <Section icon="package" count={orderTypes.length}>Order Types</Section>
        <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12}}>
          {orderTypes.map(ot=><div key={ot.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:T.card,borderRadius:T.radSm,padding:"10px 14px",border:`1px solid ${T.border}`}}>
            <span style={{fontSize:14,color:T.text}}>{ot.label}</span>
            <button onClick={()=>{if(confirm("Delete?"))onDeleteOrderType(ot.id)}} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={15} color={T.textMut}/></button>
          </div>)}
        </div>
        <div style={{display:"flex",gap:8}}><Input value={newType} onChange={setNewType} placeholder="Add order type..." style={{flex:1}}/><Btn small disabled={!newType.trim()} onClick={()=>{onAddOrderType(newType.trim());setNewType("")}}>Add</Btn></div>
      </div>

      <div>
        <Section icon="users" count={customers.length}>Customers</Section>
        <p style={{fontSize:12,color:T.textMut,marginTop:-10,marginBottom:12}}>Manage customers to create customer-specific checklist rules.</p>
        <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12}}>
          {customers.map(c=><div key={c.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:T.card,borderRadius:T.radSm,padding:"10px 14px",border:`1px solid ${T.border}`}}>
            <span style={{fontSize:14,color:T.text}}>{c.label}</span>
            <button onClick={()=>{if(confirm("Delete?"))onDeleteCustomer(c.id)}} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={15} color={T.textMut}/></button>
          </div>)}
        </div>
        <div style={{display:"flex",gap:8}}><Input value={newCust} onChange={setNewCust} placeholder="Add customer..." style={{flex:1}}/><Btn small disabled={!newCust.trim()} onClick={()=>{onAddCustomer(newCust.trim());setNewCust("")}}>Add</Btn></div>
      </div>

      {/* ── Order Stages Configuration ── */}
      {isAdmin && <div>
        <Section icon="layers">Order Stages</Section>
        <p style={{fontSize:12,color:T.textMut,marginTop:-10,marginBottom:12}}>Configure stage templates per product type. Orders of that type will auto-populate these stages.</p>
        <div style={{display:"flex",flexDirection:"column",gap:16}}>
          {PRODUCT_TYPES.map(pt => {
            const arr = stageTpl[pt] || [];
            return <div key={pt} style={{background:T.card,borderRadius:T.rad,padding:14,border:`1px solid ${T.border}`}}>
              <span style={{fontSize:14,fontWeight:600,color:T.text,display:"block",marginBottom:8}}>{pt}</span>
              {arr.length === 0 && <p style={{fontSize:12,color:T.textMut,marginBottom:8}}>No stages defined.</p>}
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {arr.map((s,i)=>{
                  const stageCk = s.checklistId ? checklists.find(c=>c.id===s.checklistId) : null;
                  const qtyFieldOptions = stageCk ? (stageCk.questions || []).filter(q => q.type === "number" || q.type === "text_number") : [];
                  return (
                  <div key={i} style={{background:T.bg,borderRadius:T.radSm,padding:10,border:`1px solid ${T.border}`}}>
                    <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:6}}>
                      <span style={{fontSize:11,fontFamily:T.mono,color:T.textMut,width:24}}>{String(i+1).padStart(2,"0")}</span>
                      <Input value={s.name} onChange={v=>updateStage(pt,i,{name:v})} placeholder="Stage name..." style={{flex:1,fontSize:13,padding:"8px 10px"}}/>
                      <button onClick={()=>moveStage(pt,i,-1)} disabled={i===0} style={{background:"none",border:"none",cursor:i===0?"not-allowed":"pointer",padding:4,opacity:i===0?0.3:1}}><span style={{fontSize:14,color:T.textSec}}>↑</span></button>
                      <button onClick={()=>moveStage(pt,i,1)} disabled={i===arr.length-1} style={{background:"none",border:"none",cursor:i===arr.length-1?"not-allowed":"pointer",padding:4,opacity:i===arr.length-1?0.3:1}}><span style={{fontSize:14,color:T.textSec}}>↓</span></button>
                      <button onClick={()=>removeStage(pt,i)} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={14} color={T.danger}/></button>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:6,marginBottom:6}}>
                      <select value={s.checklistId||""} onChange={e=>{const v=e.target.value; updateStage(pt,i,{checklistId:v,quantityField:""});}}
                        style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
                        <option value="">— Optional: required checklist —</option>
                        {checklists.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}
                      </select>
                      <Input value={s.requiredQty||0} onChange={v=>updateStage(pt,i,{requiredQty:parseFloat(v)||0})} type="number" placeholder="Req qty" style={{fontSize:12,padding:"8px 10px"}}/>
                    </div>
                    {stageCk && (
                      <select value={s.quantityField||""} onChange={e=>updateStage(pt,i,{quantityField:e.target.value})}
                        style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:12}}>
                        <option value="">— Quantity field from checklist (optional) —</option>
                        {qtyFieldOptions.map((q,qi)=><option key={qi} value={q.text}>{q.text}</option>)}
                      </select>
                    )}
                  </div>
                );})}
              </div>
              <Btn variant="ghost" small onClick={()=>addStage(pt)} style={{marginTop:8}}><Icon name="plus" size={12} color={T.textSec}/> Add Stage</Btn>
            </div>;
          })}
        </div>
        <Btn small onClick={handleSaveStageTpl} disabled={stageTplSaving} style={{marginTop:12,width:"100%"}}>{stageTplSaving?"Saving...":"Save Stage Templates"}</Btn>
      </div>}

      {/* ── Roast & Grind Classifications (Admin only) ── */}
      {isAdmin && <ClassificationsSection addToast={addToast}/>}

      {/* ── Audit Log (Admin only) ── */}
      {isAdmin && <AuditLogSection/>}

      {/* ── Data Archiving (Admin only) ── */}
      {isAdmin && <div>
        <Section icon="archive">Data Archiving</Section>
        <p style={{fontSize:12,color:T.textMut,marginTop:-10,marginBottom:12}}>Archive completed orders older than a specified number of days. Archived data is moved to separate tabs in the same spreadsheet.</p>
        <div style={{display:"flex",gap:8,marginBottom:12}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flex:1}}>
            <Input value={archiveDays} onChange={setArchiveDays} type="number" placeholder="Days..." style={{width:100}}/>
            <span style={{fontSize:13,color:T.textMut,whiteSpace:"nowrap"}}>days old</span>
          </div>
          <Btn small onClick={handleArchive} disabled={archiving}>
            <Icon name="archive" size={14} color={T.bg}/> {archiving ? "Archiving..." : "Archive"}
          </Btn>
        </div>
        {archiveMsg && <div style={{background:T.successBg,border:`1px solid ${T.successBorder}`,borderRadius:T.radSm,padding:"10px 14px",marginBottom:12}}>
          <span style={{fontSize:13,color:T.success}}>{archiveMsg}</span>
        </div>}
        {archives.length > 0 && <>
          <p style={{fontSize:12,color:T.textSec,fontWeight:500,marginBottom:8}}>Archive History</p>
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {archives.map(a => (
              <div key={a.id} style={{background:T.card,borderRadius:T.radSm,padding:"10px 14px",border:`1px solid ${T.border}`,fontSize:13}}>
                <span style={{color:T.accent,fontWeight:500}}>{a.ordersCount} orders</span>
                <span style={{color:T.textMut}}> archived on </span>
                <span style={{color:T.textSec}}>{formatDate(a.createdAt)}</span>
                <span style={{color:T.textMut}}> by {a.createdBy}</span>
              </div>
            ))}
          </div>
        </>}
      </div>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Formula Builder Component ────────────────────────────────

function FormulaBuilder({ fields, onChange, allChecklists, selfQuestions, label }) {
  const [adding, setAdding] = useState(null); // "field"|"op"|"const"|null
  const [selCk, setSelCk] = useState("self");
  const [selQ, setSelQ] = useState(0);
  const [constVal, setConstVal] = useState("");

  const addField = () => {
    const ck = selCk === "self" ? null : allChecklists?.find(c => c.id === selCk);
    const qs = selCk === "self" ? selfQuestions : (ck ? normalizeQuestions(ck.questions) : []);
    const qObj = qs[selQ];
    onChange([...(fields || []), { checklist: selCk, question: selQ, label: qObj ? getQText(qObj) : `Q${selQ+1}` }]);
    setAdding(null);
  };
  const addOp = (op) => { onChange([...(fields || []), { type: "operator", value: op }]); setAdding(null); };
  const addConst = () => { if (constVal) { onChange([...(fields || []), { type: "constant", value: parseFloat(constVal) || 0 }]); setConstVal(""); setAdding(null); } };
  const remove = (idx) => onChange((fields || []).filter((_, i) => i !== idx));

  const getFieldOptions = () => {
    const qs = selCk === "self" ? selfQuestions : normalizeQuestions(allChecklists?.find(c => c.id === selCk)?.questions || []);
    return qs.map((q, i) => ({ i, text: getQText(q), type: (typeof q === "string" ? "text" : q.type) || "text" }));
  };

  return (
    <div style={{ background: T.bg, borderRadius: T.radSm, padding: 12, border: `1px solid ${T.border}` }}>
      <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 8 }}>{label || "Formula"}</span>
      {/* Token display */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 4, marginBottom: 8, minHeight: 28 }}>
        {(fields || []).map((f, i) => (
          <span key={i} onClick={() => remove(i)} title="Click to remove" style={{
            padding: "3px 8px", borderRadius: 12, fontSize: 12, cursor: "pointer",
            background: f.type === "operator" ? T.accentBg : f.type === "constant" ? T.infoBg : T.successBg,
            color: f.type === "operator" ? T.accent : f.type === "constant" ? T.info : T.success,
            border: `1px solid ${f.type === "operator" ? T.accentBorder : f.type === "constant" ? T.infoBorder : T.successBorder}`,
          }}>
            {f.type === "operator" ? f.value : f.type === "constant" ? f.value : f.label || `Q${f.question + 1}`}
          </span>
        ))}
        {(!fields || fields.length === 0) && <span style={{ fontSize: 12, color: T.textMut }}>No formula set</span>}
      </div>
      {/* Preview */}
      {fields?.length > 0 && <p style={{ fontSize: 11, color: T.textSec, marginBottom: 8 }}>= {formulaPreview(fields, allChecklists, selfQuestions)}</p>}
      {/* Add buttons */}
      {!adding && (
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <Btn variant="ghost" small onClick={() => setAdding("field")} style={{ fontSize: 11, padding: "4px 8px" }}>+ Field</Btn>
          <Btn variant="ghost" small onClick={() => setAdding("op")} style={{ fontSize: 11, padding: "4px 8px" }}>+ Operator</Btn>
          <Btn variant="ghost" small onClick={() => setAdding("const")} style={{ fontSize: 11, padding: "4px 8px" }}>+ Number</Btn>
          {fields?.length > 0 && <Btn variant="ghost" small onClick={() => onChange([])} style={{ fontSize: 11, padding: "4px 8px", color: T.danger }}>Clear</Btn>}
        </div>
      )}
      {/* Field picker */}
      {adding === "field" && (
        <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 6 }}>
          <div style={{ display: "flex", gap: 6 }}>
            <select value={selCk} onChange={e => { setSelCk(e.target.value); setSelQ(0); }}
              style={{ flex: 1, padding: "6px 8px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 12 }}>
              <option value="self">This checklist</option>
              {(allChecklists || []).map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            <select value={selQ} onChange={e => setSelQ(parseInt(e.target.value))}
              style={{ flex: 1, padding: "6px 8px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 12 }}>
              {getFieldOptions().map(o => <option key={o.i} value={o.i}>{o.text} ({o.type})</option>)}
            </select>
            <Btn small onClick={addField} style={{ fontSize: 11 }}>Add</Btn>
            <Btn variant="ghost" small onClick={() => setAdding(null)} style={{ fontSize: 11 }}>Cancel</Btn>
          </div>
        </div>
      )}
      {/* Operator picker */}
      {adding === "op" && (
        <div style={{ display: "flex", gap: 4, marginTop: 6 }}>
          {["+", "-", "×", "÷", "%"].map(op => (
            <button key={op} onClick={() => addOp(op)} style={{ width: 32, height: 32, borderRadius: 8, border: `1px solid ${T.border}`, background: T.surface, color: T.accent, fontSize: 16, fontWeight: 600, cursor: "pointer" }}>{op}</button>
          ))}
          <Btn variant="ghost" small onClick={() => setAdding(null)} style={{ fontSize: 11 }}>Cancel</Btn>
        </div>
      )}
      {/* Constant input */}
      {adding === "const" && (
        <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
          <Input value={constVal} onChange={setConstVal} placeholder="Enter number..." type="number" style={{ flex: 1, fontSize: 12, padding: "6px 8px" }} />
          <Btn small onClick={addConst} style={{ fontSize: 11 }}>Add</Btn>
          <Btn variant="ghost" small onClick={() => setAdding(null)} style={{ fontSize: 11 }}>Cancel</Btn>
        </div>
      )}
    </div>
  );
}

// ─── Edit Checklist View ──────────────────────────────────────

function AutoFillAddRow({ questions, srcQ, linkedIdx, onAdd }) {
  const [open, setOpen] = useState(false);
  const [targetIdx, setTargetIdx] = useState("");
  const [sourceIdx, setSourceIdx] = useState("");
  const [readOnly, setReadOnly] = useState(true);
  const unmapped = questions.map((q, qi) => ({ q, qi })).filter(({ q, qi }) => qi !== linkedIdx && !q.autoFillMapping);
  if (unmapped.length === 0) return null;
  if (!open) return <button onClick={() => setOpen(true)} style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: T.radSm, padding: "4px 10px", fontSize: 11, color: T.accent, cursor: "pointer", marginTop: 6 }}>+ Add Mapping</button>;
  return (
    <div style={{ display: "flex", gap: 6, alignItems: "center", padding: "6px 8px", background: T.bg, borderRadius: T.radSm, border: `1px dashed ${T.accentBorder}`, marginTop: 6, flexWrap: "wrap" }}>
      <select value={targetIdx} onChange={e => setTargetIdx(e.target.value)} style={{ flex: 1, minWidth: 90, padding: "3px 6px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 11 }}>
        <option value="">— Target field —</option>
        {unmapped.map(({ q, qi }) => <option key={qi} value={qi}>{q.text || `Q${qi + 1}`}</option>)}
      </select>
      <span style={{ fontSize: 10, color: T.textMut }}>pulls from:</span>
      <select value={sourceIdx} onChange={e => setSourceIdx(e.target.value)} style={{ flex: 1, minWidth: 90, padding: "3px 6px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 11 }}>
        <option value="">— Source field —</option>
        {srcQ.map((sq, si) => <option key={si} value={si}>{sq.text || `Q${si + 1}`}</option>)}
      </select>
      <label style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 10, color: T.textSec, cursor: "pointer", whiteSpace: "nowrap" }}>
        <input type="checkbox" checked={readOnly} onChange={e => setReadOnly(e.target.checked)} style={{ accentColor: T.accent }} />
        Read-only
      </label>
      <Btn small disabled={targetIdx === "" || sourceIdx === ""} onClick={() => { onAdd(Number(targetIdx), sourceIdx, readOnly); setOpen(false); setTargetIdx(""); setSourceIdx(""); }}>Save</Btn>
      <Btn small variant="ghost" onClick={() => { setOpen(false); setTargetIdx(""); setSourceIdx(""); }}>Cancel</Btn>
    </div>
  );
}

function EditChecklistView({ checklist, allChecklists, onSave, inventoryItems, inventoryCategories }) {
  const [name, setName] = useState(checklist?.name || "");
  const [subtitle, setSubtitle] = useState(checklist?.subtitle || "");
  const [formUrl, setFormUrl] = useState(checklist?.formUrl || "");
  const [questions, setQuestions] = useState(() => {
    const raw = checklist?.questions || [{ text: "", type: "text", formula: null, ideal: null, remarkCondition: null }];
    return normalizeQuestions(raw);
  });
  const [expandedQ, setExpandedQ] = useState(null); // index of question with expanded advanced options
  const [autoIdConfig, setAutoIdConfig] = useState(() => {
    const cfg = checklist?.autoIdConfig || null;
    return {
      enabled: cfg?.enabled || false,
      prefix: cfg?.prefix || getDefaultPrefixForChecklist(checklist?.name || "") || "",
      dateFieldIdx: cfg?.dateFieldIdx ?? "",
      itemCodeFieldIdx: cfg?.itemCodeFieldIdx ?? "",
    };
  });
  const [canTagTo, setCanTagTo] = useState(() => Array.isArray(checklist?.canTagTo) ? checklist.canTagTo : []);
  const toggleCanTagTo = (key) => setCanTagTo(prev => prev.includes(key) ? prev.filter(x => x !== key) : [...prev, key]);

  const updateQ = (i, patch) => setQuestions(prev => prev.map((q, idx) => idx === i ? { ...q, ...patch } : q));
  const removeQ = (i) => setQuestions(prev => prev.filter((_, idx) => idx !== i));
  const addQ = () => setQuestions(prev => [...prev, { text: "", type: "text", formula: null, ideal: null, remarkCondition: null, isApprovalGate: false, linkedSource: null, inventoryLink: null, isMasterQuantity: false, autoFillMapping: null, dateComparison: null }]);
  const moveQ = (i, dir) => setQuestions(prev => {
    const arr = [...prev]; const j = i + dir;
    if (j < 0 || j >= arr.length) return arr;
    [arr[i], arr[j]] = [arr[j], arr[i]];
    // Update expandedQ to follow the moved question
    if (expandedQ === i) setExpandedQ(j);
    else if (expandedQ === j) setExpandedQ(i);
    return arr;
  });

  const handleSave = () => {
    const cleaned = questions.filter(q => q.text.trim());
    const cfgToSave = autoIdConfig.enabled ? {
      enabled: true,
      prefix: (autoIdConfig.prefix || "").toUpperCase(),
      dateFieldIdx: autoIdConfig.dateFieldIdx === "" ? null : Number(autoIdConfig.dateFieldIdx),
      itemCodeFieldIdx: autoIdConfig.itemCodeFieldIdx === "" ? null : Number(autoIdConfig.itemCodeFieldIdx),
    } : null;
    onSave({ id: checklist?.id || "ck_" + Date.now(), name, subtitle, formUrl, questions: cleaned, autoIdConfig: cfgToSave, canTagTo });
  };

  const isNumeric = (q) => q.type === "number" || q.type === "text_number";
  const dateQuestions = questions.map((q, i) => ({ ...q, _idx: i })).filter(q => q.type === "date");
  const previewChecklist = { name, autoIdConfig: { ...autoIdConfig, enabled: true } };
  const previewStr = autoIdConfig.enabled ? buildAutoIdPreview(previewChecklist, {}, new Date(), inventoryItems) : "";

  return (
    <div className="fade-up" style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      <Field label="Checklist Name"><Input value={name} onChange={setName} placeholder="e.g., Quality Inspection" /></Field>
      <Field label="Subtitle / Context"><Input value={subtitle} onChange={setSubtitle} placeholder="e.g., Per roast batch" /></Field>
      <Field label="Google Form URL">
        <Input value={formUrl} onChange={setFormUrl} placeholder="https://docs.google.com/forms/d/e/..." />
        <p style={{ fontSize: 12, color: T.textMut, marginTop: 6 }}>Paste the shareable link.</p>
      </Field>

      {/* ── Auto ID Configuration ── */}
      <div style={{ background: T.card, borderRadius: T.rad, padding: 14, border: `1px solid ${autoIdConfig.enabled ? T.accentBorder : T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
          <span style={{ fontSize: 14, fontWeight: 600, color: T.text }}>Auto ID Configuration</span>
          <button onClick={() => setAutoIdConfig(p => ({ ...p, enabled: !p.enabled, prefix: p.prefix || getDefaultPrefixForChecklist(name) || "" }))}
            style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${autoIdConfig.enabled ? T.accent : T.borderLight}`, background: autoIdConfig.enabled ? T.accent : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer" }}>
            {autoIdConfig.enabled && <Icon name="check" size={12} color={T.bg} />}
          </button>
        </div>
        {autoIdConfig.enabled && (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <Field label="Prefix">
              <Input value={autoIdConfig.prefix} onChange={v => setAutoIdConfig(p => ({ ...p, prefix: v.toUpperCase() }))} placeholder="e.g., GBS" style={{ textTransform: "uppercase" }} />
            </Field>
            <Field label="Date Field">
              <select value={autoIdConfig.dateFieldIdx} onChange={e => setAutoIdConfig(p => ({ ...p, dateFieldIdx: e.target.value }))}
                style={{ width: "100%", padding: "10px 14px", borderRadius: T.radSm, background: T.bg, border: `1px solid ${T.border}`, color: T.text, fontSize: 14 }}>
                <option value="">— Use submission date —</option>
                {dateQuestions.map(q => <option key={q._idx} value={q._idx}>{q.text || `Q${q._idx + 1}`}</option>)}
              </select>
            </Field>
            <Field label="Item Code Field">
              <select value={autoIdConfig.itemCodeFieldIdx} onChange={e => setAutoIdConfig(p => ({ ...p, itemCodeFieldIdx: e.target.value }))}
                style={{ width: "100%", padding: "10px 14px", borderRadius: T.radSm, background: T.bg, border: `1px solid ${T.border}`, color: T.text, fontSize: 14 }}>
                <option value="">— None (use 'X') —</option>
                {questions.map((q, i) => q.type === "inventory_item" ? <option key={i} value={i}>{q.text || `Q${i + 1}`}</option> : null)}
              </select>
              <p style={{ fontSize: 11, color: T.textMut, marginTop: 4 }}>Only Inventory Item fields can be picked — the chosen item's abbreviation becomes the code.</p>
            </Field>
            <div style={{ padding: "8px 12px", background: T.bg, borderRadius: T.radSm, border: `1px solid ${T.border}` }}>
              <span style={{ fontSize: 11, color: T.textMut, display: "block" }}>Preview</span>
              <span style={{ fontSize: 14, fontFamily: T.mono, color: T.accent }}>{previewStr || "—"}</span>
            </div>
            <p style={{ fontSize: 11, color: T.textMut }}>When Auto ID is enabled, the generated ID is also used as the Linked ID downstream — no separate Linked ID question is required.</p>
          </div>
        )}
      </div>

      {/* ── Can Tag To (allowed downstream destinations) ── */}
      <div style={{ background: T.card, borderRadius: T.rad, padding: 14, border: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <span style={{ fontSize: 14, fontWeight: 600, color: T.text }}>Can Tag To</span>
          <span style={{ fontSize: 11, color: T.textMut }}>Empty = all + Invoice</span>
        </div>
        <p style={{ fontSize: 12, color: T.textMut, marginBottom: 10 }}>Limit which destinations submissions of this checklist can be tagged to. Leave empty for no restriction.</p>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          <Chip label="Invoice / Order" active={canTagTo.includes("order")} onClick={() => toggleCanTagTo("order")}/>
          {(allChecklists || []).filter(c => c.id !== (checklist?.id || "")).map(c => (
            <Chip key={c.id} label={c.name} active={canTagTo.includes(c.id)} onClick={() => toggleCanTagTo(c.id)}/>
          ))}
        </div>
      </div>

      <Field label="Questions">
        <p style={{ fontSize: 12, color: T.textMut, marginBottom: 10 }}>Define questions, field types, formulas, and remark conditions.</p>
        {questions.map((q, i) => {
          const isExp = expandedQ === i;
          return (
            <div key={i} style={{ background: T.card, borderRadius: T.radSm, padding: 12, border: `1px solid ${isExp ? T.accentBorder : T.border}`, marginBottom: 10, transition: "border .2s" }}>
              {/* Question text + reorder + delete */}
              <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 8 }}>
                <div style={{ display: "flex", flexDirection: "column", gap: 2, flexShrink: 0 }}>
                  {i > 0 && <button onClick={() => moveQ(i, -1)} style={{ background: "none", border: "none", cursor: "pointer", padding: 2, minWidth: 44, minHeight: 22, display: "flex", alignItems: "center", justifyContent: "center" }} title="Move up">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={T.textSec} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M18 15l-6-6-6 6"/></svg>
                  </button>}
                  {i < questions.length - 1 && <button onClick={() => moveQ(i, 1)} style={{ background: "none", border: "none", cursor: "pointer", padding: 2, minWidth: 44, minHeight: 22, display: "flex", alignItems: "center", justifyContent: "center" }} title="Move down">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={T.textSec} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M6 9l6 6 6-6"/></svg>
                  </button>}
                </div>
                <span style={{ fontSize: 12, color: T.textMut, fontFamily: T.mono, width: 22, flexShrink: 0 }}>{String(i + 1).padStart(2, "0")}</span>
                <Input value={q.text} onChange={v => updateQ(i, { text: v })} placeholder="Question text..." style={{ flex: 1 }} />
                {questions.length > 1 && <button onClick={() => removeQ(i)} style={{ background: "none", border: "none", cursor: "pointer", padding: 4, flexShrink: 0 }}><Icon name="trash" size={16} color={T.danger} /></button>}
              </div>
              {/* Type selector */}
              <div style={{ display: "flex", gap: 6, marginBottom: 8, paddingLeft: 30, flexWrap: "wrap" }}>
                {[["text", "Text"], ["number", "Number"], ["text_number", "Text & Number"], ["date", "Date"], ["yesno", "Yes / No"], ["inventory_item", "Inventory Item"]].map(([val, lbl]) => (
                  <Chip key={val} label={lbl} active={q.type === val} onClick={() => {
                    const patch = { type: val };
                    if (val === "text" || val === "date" || val === "yesno" || val === "inventory_item") { patch.formula = null; patch.ideal = null; patch.remarkCondition = null; }
                    updateQ(i, patch);
                  }} />
                ))}
                {isNumeric(q) && <Btn variant="ghost" small onClick={() => setExpandedQ(isExp ? null : i)} style={{ fontSize: 11, marginLeft: "auto" }}>
                  {isExp ? "Collapse" : "Advanced"}
                </Btn>}
              </div>
              {/* ── Date Comparison Rule ── */}
              {q.type === "date" && (
                <div style={{ paddingLeft: 30, marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                    <button onClick={() => {
                      if (q.dateComparison) { updateQ(i, { dateComparison: null }); }
                      else { updateQ(i, { dateComparison: { operator: "gte", compareToFieldIdx: "", errorMessage: "" } }); }
                    }} style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${q.dateComparison ? T.warning : T.borderLight}`, background: q.dateComparison ? T.warning : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
                      {q.dateComparison && <Icon name="check" size={12} color={T.bg} />}
                    </button>
                    <span style={{ fontSize: 12, color: T.textSec }}>Add date comparison rule</span>
                  </div>
                  {q.dateComparison && (
                    <div style={{ marginLeft: 26, background: T.bg, borderRadius: T.radSm, padding: 10, border: `1px solid ${T.border}`, display: "flex", flexDirection: "column", gap: 8 }}>
                      <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                        <span style={{ fontSize: 12, color: T.textSec }}>This date must be</span>
                        <select value={q.dateComparison.operator || "gte"} onChange={e => updateQ(i, { dateComparison: { ...q.dateComparison, operator: e.target.value } })}
                          style={{ padding: "6px 8px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 12 }}>
                          <option value="gte">on or after</option>
                          <option value="lte">on or before</option>
                          <option value="eq">the same as</option>
                        </select>
                        <span style={{ fontSize: 12, color: T.textSec }}>the date in field:</span>
                        <select value={q.dateComparison.compareToFieldIdx ?? ""} onChange={e => updateQ(i, { dateComparison: { ...q.dateComparison, compareToFieldIdx: e.target.value } })}
                          style={{ padding: "6px 8px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 12, minWidth: 120 }}>
                          <option value="">— Select field —</option>
                          {questions.filter((oq, oi) => oi !== i && oq.type === "date").map((oq, oi) => {
                            const realIdx = questions.indexOf(oq);
                            return <option key={realIdx} value={realIdx}>{oq.text || `Q${realIdx + 1}`}</option>;
                          })}
                        </select>
                      </div>
                      <Input value={q.dateComparison.errorMessage || ""} onChange={v => updateQ(i, { dateComparison: { ...q.dateComparison, errorMessage: v } })} placeholder="Custom error message (optional)" style={{ fontSize: 12 }} />
                    </div>
                  )}
                </div>
              )}
              {q.type === "inventory_item" && (
                <div style={{ paddingLeft: 30, marginBottom: 8 }}>
                  <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 4 }}>Inventory Category (filter)</span>
                  <select value={q.inventoryCategory || ""} onChange={e => updateQ(i, { inventoryCategory: e.target.value })}
                    style={{ width: "100%", padding: "8px 10px", borderRadius: T.radSm, background: T.bg, border: `1px solid ${T.border}`, color: T.text, fontSize: 13 }}>
                    <option value="">— Any —</option>
                    {(inventoryCategories || []).map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
                  </select>
                </div>
              )}
              {/* Advanced options (number types) — always shown if a formula already exists, otherwise hidden behind the toggle */}
              {(isExp || q.formula) && isNumeric(q) && (
                <div style={{ paddingLeft: 30, display: "flex", flexDirection: "column", gap: 10 }}>
                  {/* Formula */}
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                      <span style={{ fontSize: 12, fontWeight: 500, color: T.textSec }}>Auto-Calculate Formula</span>
                      {!q.formula && <Btn variant="ghost" small onClick={() => updateQ(i, { formula: { fields: [] } })} style={{ fontSize: 11, padding: "2px 8px" }}>+ Add</Btn>}
                      {q.formula && <Btn variant="ghost" small onClick={() => updateQ(i, { formula: null, ideal: null, remarkCondition: null })} style={{ fontSize: 11, padding: "2px 8px", color: T.danger }}>Remove</Btn>}
                    </div>
                    {q.formula && <FormulaBuilder fields={q.formula.fields} onChange={fields => updateQ(i, { formula: { fields } })} allChecklists={allChecklists} selfQuestions={questions} label="Calculate value as:" />}
                  </div>
                  {/* Ideal value (only if formula exists) */}
                  {q.formula && (
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                        <span style={{ fontSize: 12, fontWeight: 500, color: T.textSec }}>Ideal Value</span>
                        {!q.ideal && <Btn variant="ghost" small onClick={() => updateQ(i, { ideal: { fields: [], suffix: "" } })} style={{ fontSize: 11, padding: "2px 8px" }}>+ Set Ideal</Btn>}
                        {q.ideal && <Btn variant="ghost" small onClick={() => updateQ(i, { ideal: null, remarkCondition: null })} style={{ fontSize: 11, padding: "2px 8px", color: T.danger }}>Remove</Btn>}
                      </div>
                      {q.ideal && (
                        <>
                          <FormulaBuilder fields={q.ideal.fields} onChange={fields => updateQ(i, { ideal: { ...q.ideal, fields } })} allChecklists={allChecklists} selfQuestions={questions} label="Ideal value formula:" />
                          <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
                            <Input value={q.idealLabel || ""} onChange={v => updateQ(i, { idealLabel: v })} placeholder="Ideal label (e.g., Expected roast loss 20%)" style={{ flex: 2, fontSize: 12, padding: "6px 8px" }} />
                            <Input value={q.idealUnit || q.ideal.suffix || ""} onChange={v => updateQ(i, { idealUnit: v, ideal: { ...q.ideal, suffix: v } })} placeholder="Unit (e.g., kgs)" style={{ flex: 1, fontSize: 12, padding: "6px 8px" }} />
                          </div>
                        </>
                      )}
                    </div>
                  )}
                  {/* Remark Condition (only if ideal exists) */}
                  {q.ideal && (
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                        <span style={{ fontSize: 12, fontWeight: 500, color: T.textSec }}>Remark Condition</span>
                        {!q.remarkCondition && <Btn variant="ghost" small onClick={() => updateQ(i, { remarkCondition: { type: "ne_ideal", value: 0, message: "Value differs from ideal. Please provide reason." } })} style={{ fontSize: 11, padding: "2px 8px" }}>+ Add Condition</Btn>}
                        {q.remarkCondition && <Btn variant="ghost" small onClick={() => updateQ(i, { remarkCondition: null })} style={{ fontSize: 11, padding: "2px 8px", color: T.danger }}>Remove</Btn>}
                      </div>
                      {q.remarkCondition && (
                        <div style={{ background: T.warningBg, border: `1px solid ${T.warningBorder}`, borderRadius: T.radSm, padding: 10 }}>
                          <div style={{ display: "flex", gap: 6, marginBottom: 8 }}>
                            <select value={q.remarkCondition.type} onChange={e => updateQ(i, { remarkCondition: { ...q.remarkCondition, type: e.target.value } })}
                              style={{ flex: 1, padding: "6px 8px", borderRadius: T.radSm, background: T.bg, border: `1px solid ${T.border}`, color: T.text, fontSize: 12 }}>
                              <option value="ne_ideal">Any difference</option>
                              <option value="gt_ideal">Greater than ideal</option>
                              <option value="lt_ideal">Less than ideal</option>
                              <option value="differs_by_percent">More than X%</option>
                              <option value="differs_by_units">More than X units</option>
                            </select>
                            {(q.remarkCondition.type === "differs_by_percent" || q.remarkCondition.type === "differs_by_units") && (
                              <Input value={q.remarkCondition.value || ""} onChange={v => updateQ(i, { remarkCondition: { ...q.remarkCondition, value: parseFloat(v) || 0 } })} placeholder={q.remarkCondition.type === "differs_by_percent" ? "%" : "units"} type="number" style={{ width: 80, fontSize: 12, padding: "6px 8px" }} />
                            )}
                          </div>
                          <Input value={q.remarkCondition.message || ""} onChange={v => updateQ(i, { remarkCondition: { ...q.remarkCondition, message: v } })} placeholder="Warning message to show..." style={{ fontSize: 12, padding: "6px 8px" }} />
                          <div style={{ marginTop: 8 }}>
                            <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 4 }}>Required remarks target (optional)</span>
                            <select value={q.remarksTargetIdx ?? ""} onChange={e => updateQ(i, { remarksTargetIdx: e.target.value === "" ? null : Number(e.target.value) })}
                              style={{ width: "100%", padding: "6px 8px", borderRadius: T.radSm, background: T.bg, border: `1px solid ${T.border}`, color: T.text, fontSize: 12 }}>
                              <option value="">— Inline remark on this field —</option>
                              {questions.map((qq, qqi) => qqi !== i ? <option key={qqi} value={qqi}>{qq.text || `Q${qqi + 1}`}</option> : null)}
                            </select>
                          </div>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
              {/* ── Inventory Link (number fields only) ── */}
              {isNumeric(q) && (
                <div style={{ paddingLeft: 30, display: "flex", flexDirection: "column", gap: 8, marginTop: 8 }}>
                  {/* Link to Inventory toggle */}
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <button onClick={() => {
                      if (q.inventoryLink) updateQ(i, { inventoryLink: null });
                      else updateQ(i, { inventoryLink: { enabled: true, txType: "IN", category: "", itemSource: { type: "fixed", itemId: "", fieldIdx: null } } });
                    }} style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${q.inventoryLink ? T.info : T.borderLight}`, background: q.inventoryLink ? T.info : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
                      {q.inventoryLink && <Icon name="check" size={12} color={T.bg} />}
                    </button>
                    <span style={{ fontSize: 12, color: T.textSec }}>Link to Inventory</span>
                    {q.inventoryLink && <Badge variant="info" style={{ fontSize: 9 }}>Active</Badge>}
                  </div>
                  {q.inventoryLink && (
                    <div style={{ marginLeft: 26, background: T.bg, borderRadius: T.radSm, padding: 10, border: `1px solid ${T.border}`, display: "flex", flexDirection: "column", gap: 8 }}>
                      <div>
                        <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 4 }}>Transaction Type</span>
                        <div style={{ display: "flex", gap: 6 }}>
                          <Chip label="IN (addition)" active={q.inventoryLink.txType === "IN"} onClick={() => updateQ(i, { inventoryLink: { ...q.inventoryLink, txType: "IN" } })} />
                          <Chip label="OUT (subtraction)" active={q.inventoryLink.txType === "OUT"} onClick={() => updateQ(i, { inventoryLink: { ...q.inventoryLink, txType: "OUT" } })} />
                        </div>
                      </div>
                      <div>
                        <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 4 }}>Inventory Category</span>
                        <select value={q.inventoryLink.category || ""} onChange={e => updateQ(i, { inventoryLink: { ...q.inventoryLink, category: e.target.value } })}
                          style={{ width: "100%", padding: "8px 10px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 13 }}>
                          <option value="">— Any —</option>
                          {(inventoryCategories || []).map(c => <option key={c.id} value={c.name}>{c.name}</option>)}
                        </select>
                      </div>
                      <div>
                        <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 4 }}>Item Source</span>
                        <div style={{ display: "flex", gap: 6, marginBottom: 6 }}>
                          <Chip label="From field in this form" active={q.inventoryLink.itemSource?.type === "field"} onClick={() => updateQ(i, { inventoryLink: { ...q.inventoryLink, itemSource: { type: "field", fieldIdx: null, itemId: "" } } })} />
                          <Chip label="Fixed item" active={q.inventoryLink.itemSource?.type === "fixed"} onClick={() => updateQ(i, { inventoryLink: { ...q.inventoryLink, itemSource: { type: "fixed", itemId: "", fieldIdx: null } } })} />
                        </div>
                        {q.inventoryLink.itemSource?.type === "field" && (
                          <select value={q.inventoryLink.itemSource.fieldIdx ?? ""} onChange={e => updateQ(i, { inventoryLink: { ...q.inventoryLink, itemSource: { ...q.inventoryLink.itemSource, fieldIdx: e.target.value === "" ? null : Number(e.target.value) } } })}
                            style={{ width: "100%", padding: "8px 10px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 13 }}>
                            <option value="">— Select field —</option>
                            {questions.map((qq, qi) => <option key={qi} value={qi}>{qq.text || `Q${qi + 1}`}</option>)}
                          </select>
                        )}
                        {q.inventoryLink.itemSource?.type === "fixed" && (
                          <select value={q.inventoryLink.itemSource.itemId || ""} onChange={e => updateQ(i, { inventoryLink: { ...q.inventoryLink, itemSource: { ...q.inventoryLink.itemSource, itemId: e.target.value } } })}
                            style={{ width: "100%", padding: "8px 10px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 13 }}>
                            <option value="">— Select item —</option>
                            {(inventoryItems || []).filter(it => !q.inventoryLink.category || it.category === q.inventoryLink.category).map(it => <option key={it.id} value={it.id}>{it.name}{it.abbreviation ? ` (${it.abbreviation})` : ""}</option>)}
                          </select>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* ── Approval Gate / Linked ID / Linked Source (always visible, not just advanced) ── */}
              <div style={{ paddingLeft: 30, display: "flex", flexDirection: "column", gap: 8, marginTop: 8 }}>
                {/* Approval Gate toggle */}
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button onClick={() => {
                    const newVal = !q.isApprovalGate;
                    // Only one approval gate per checklist
                    if (newVal) {
                      setQuestions(prev => prev.map((qq, idx) => ({ ...qq, isApprovalGate: idx === i ? true : false })));
                    } else { updateQ(i, { isApprovalGate: false }); }
                  }} style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${q.isApprovalGate ? T.success : T.borderLight}`, background: q.isApprovalGate ? T.success : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
                    {q.isApprovalGate && <Icon name="check" size={12} color={T.bg} />}
                  </button>
                  <span style={{ fontSize: 12, color: T.textSec }}>Mark as Approval Gate</span>
                  {q.isApprovalGate && <Badge variant="success" style={{ fontSize: 9 }}>Active</Badge>}
                </div>

                {/* Link to Approved Entries */}
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button onClick={() => {
                    if (q.linkedSource) { updateQ(i, { linkedSource: null }); }
                    else { updateQ(i, { linkedSource: { checklistId: "", type: "approved_only" } }); }
                  }} style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${q.linkedSource ? T.accent : T.borderLight}`, background: q.linkedSource ? T.accent : "transparent", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}>
                    {q.linkedSource && <Icon name="check" size={12} color={T.bg} />}
                  </button>
                  <span style={{ fontSize: 12, color: T.textSec }}>Link to Approved Entries</span>
                  {q.linkedSource && <Badge style={{ fontSize: 9 }}>Active</Badge>}
                </div>
                {q.linkedSource && (
                  <div style={{ marginLeft: 26, background: T.bg, borderRadius: T.radSm, padding: 10, border: `1px solid ${T.border}` }}>
                    <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 6 }}>Source Checklist (pull approved entries from):</span>
                    <select value={q.linkedSource.checklistId || ""} onChange={e => updateQ(i, { linkedSource: { ...q.linkedSource, checklistId: e.target.value } })}
                      style={{ width: "100%", padding: "8px 10px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 13 }}>
                      <option value="">— Select source checklist —</option>
                      {(allChecklists || []).filter(c => c.id !== (checklist?.id || "")).map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                    </select>
                    {q.linkedSource.checklistId && (() => {
                      const src = (allChecklists || []).find(c => c.id === q.linkedSource.checklistId);
                      if (!src) return null;
                      const srcQ = normalizeQuestions(src.questions);
                      const hasGate = srcQ.some(sq => sq.isApprovalGate);
                      const hasAutoId = src.autoIdConfig && src.autoIdConfig.enabled;
                      return <div style={{ display: "flex", flexDirection: "column", gap: 6, marginTop: 6 }}>
                        {hasGate && hasAutoId ? <Badge variant="success" style={{ fontSize: 10 }}>Source has Approval Gate + Auto ID ✓</Badge>
                          : <Badge variant="danger" style={{ fontSize: 10 }}>{!hasGate && !hasAutoId ? "Source needs both Approval Gate and Auto ID enabled" : !hasGate ? "Source needs Approval Gate configured" : "Source checklist must have Auto ID enabled for linking to work"}</Badge>}
                      </div>;
                    })()}
                    {/* Auto-fill mappings — centralized builder */}
                    {q.linkedSource.checklistId && (() => {
                      const src = (allChecklists || []).find(c => c.id === q.linkedSource.checklistId);
                      if (!src) return null;
                      const srcQ = normalizeQuestions(src.questions);
                      const existingMappings = questions.map((tq, ti) => {
                        if (ti === i || !tq.autoFillMapping || tq.autoFillMapping.sourceFieldIdx === "" || tq.autoFillMapping.sourceFieldIdx === undefined) return null;
                        return { targetIdx: ti, targetText: tq.text, sourceFieldIdx: tq.autoFillMapping.sourceFieldIdx, readOnly: tq.autoFillMapping.readOnly !== false };
                      }).filter(Boolean);
                      const unmappedQuestions = questions.filter((tq, ti) => ti !== i && !tq.autoFillMapping);
                      return <div style={{ marginTop: 8, padding: 10, background: T.surface, borderRadius: T.radSm, border: `1px solid ${T.border}` }}>
                        <span style={{ fontSize: 11, fontWeight: 600, color: T.textSec, display: "block", marginBottom: 6 }}>Auto-fill Mappings</span>
                        {existingMappings.length === 0 && <span style={{ fontSize: 11, color: T.textMut, display: "block", marginBottom: 6 }}>No mappings configured yet.</span>}
                        <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                          {existingMappings.map((m, mi) => {
                            const srcField = srcQ[Number(m.sourceFieldIdx)];
                            return <div key={mi} style={{ display: "flex", gap: 6, alignItems: "center", padding: "6px 8px", background: T.bg, borderRadius: T.radSm, border: `1px solid ${T.border}`, flexWrap: "wrap" }}>
                              <span style={{ fontSize: 11, color: T.text, fontWeight: 500, minWidth: 80 }}>{m.targetText}</span>
                              <span style={{ fontSize: 10, color: T.textMut }}>pulls from:</span>
                              <select value={m.sourceFieldIdx} onChange={e => updateQ(m.targetIdx, { autoFillMapping: { ...questions[m.targetIdx].autoFillMapping, sourceFieldIdx: e.target.value } })}
                                style={{ flex: 1, minWidth: 100, padding: "3px 6px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 11 }}>
                                <option value="">— Select —</option>
                                {srcQ.map((sq, si) => <option key={si} value={si}>{sq.text || `Q${si + 1}`}</option>)}
                              </select>
                              <label style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 10, color: T.textSec, cursor: "pointer", whiteSpace: "nowrap" }}>
                                <input type="checkbox" checked={m.readOnly} onChange={e => updateQ(m.targetIdx, { autoFillMapping: { ...questions[m.targetIdx].autoFillMapping, readOnly: e.target.checked } })} style={{ accentColor: T.accent }} />
                                Read-only
                              </label>
                              <button onClick={() => updateQ(m.targetIdx, { autoFillMapping: null })} style={{ background: "none", border: "none", cursor: "pointer", padding: 2 }}><Icon name="x" size={12} color={T.danger} /></button>
                            </div>;
                          })}
                        </div>
                        {unmappedQuestions.length > 0 && <AutoFillAddRow questions={questions} srcQ={srcQ} linkedIdx={i} onAdd={(targetIdx, sourceFieldIdx, readOnly) => updateQ(targetIdx, { autoFillMapping: { sourceFieldIdx: String(sourceFieldIdx), readOnly } })} />}
                      </div>;
                    })()}
                  </div>
                )}

                {/* ── Per-question Auto-fill from Source ── */}
                {!q.linkedSource && (() => {
                  const linkedQ = questions.find(qq => qq.linkedSource && qq.linkedSource.checklistId);
                  if (!linkedQ) return null;
                  const src = (allChecklists || []).find(c => c.id === linkedQ.linkedSource.checklistId);
                  if (!src) return null;
                  const srcQ = normalizeQuestions(src.questions);
                  const mapping = q.autoFillMapping;
                  const hasMapping = mapping && mapping.sourceFieldIdx !== "" && mapping.sourceFieldIdx !== undefined;
                  return <div style={{ marginTop: 4 }}>
                    {!mapping ? (
                      <button onClick={() => updateQ(i, { autoFillMapping: { sourceFieldIdx: "", readOnly: true } })}
                        style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: T.radSm, padding: "4px 10px", fontSize: 11, color: T.info, cursor: "pointer" }}>
                        + Auto-fill from {src.name}
                      </button>
                    ) : (
                      <div style={{ background: T.bg, borderRadius: T.radSm, padding: 8, border: `1px solid ${T.infoBorder}`, display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
                        <span style={{ fontSize: 11, color: T.textMut, flexShrink: 0 }}>Pull value from:</span>
                        <select value={mapping.sourceFieldIdx ?? ""} onChange={e => updateQ(i, { autoFillMapping: { ...mapping, sourceFieldIdx: e.target.value } })}
                          style={{ flex: 1, minWidth: 120, padding: "4px 8px", borderRadius: T.radSm, background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontSize: 11 }}>
                          <option value="">— Select source field —</option>
                          {srcQ.map((sq, si) => <option key={si} value={si}>{sq.text || `Q${si + 1}`}</option>)}
                        </select>
                        <label style={{ display: "flex", alignItems: "center", gap: 3, fontSize: 10, color: T.textSec, cursor: "pointer", whiteSpace: "nowrap" }}>
                          <input type="checkbox" checked={mapping.readOnly !== false} onChange={e => updateQ(i, { autoFillMapping: { ...mapping, readOnly: e.target.checked } })} style={{ accentColor: T.accent }} />
                          Read-only
                        </label>
                        <button onClick={() => updateQ(i, { autoFillMapping: null })} style={{ background: "none", border: "none", cursor: "pointer", padding: 2 }}><Icon name="x" size={12} color={T.danger} /></button>
                      </div>
                    )}
                  </div>;
                })()}
              </div>
            </div>
          );
        })}
        <Btn variant="ghost" small onClick={addQ} style={{ marginTop: 4 }}><Icon name="plus" size={14} color={T.textSec} /> Add Question</Btn>
      </Field>
      <Btn onClick={handleSave} disabled={!name.trim()} style={{ width: "100%", marginTop: 8 }}>{checklist ? "Save Changes" : "Create Checklist"}</Btn>
    </div>
  );
}

// ─── Rules View ───────────────────────────────────────────────

function RulesView({rules,orderTypes,customers,checklists,onAddRule,onEditRule,onDeleteRule}){
  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:16}}>
      <p style={{fontSize:13,color:T.textSec,lineHeight:1.5}}>Rules auto-assign checklists when creating orders. Specific rules (matching both order type AND customer) take priority over general ones.</p>
      <Btn onClick={onAddRule} style={{width:"100%"}}><Icon name="plus" size={18} color={T.bg}/> Add Rule</Btn>
      <div style={{display:"flex",flexDirection:"column",gap:10}}>
        {rules.map(r=>{
          const ot=orderTypes.find(t=>t.id===r.orderTypeId);const cu=customers.find(c=>c.id===r.customerId);
          const cks=r.checklistIds.map(id=>checklists.find(c=>c.id===id)).filter(Boolean);
          return <div key={r.id} style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
              <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                <Badge>{r.orderTypeId==="any"?"Any Order Type":ot?.label||"?"}</Badge>
                <span style={{fontSize:12,color:T.textMut}}>+</span>
                <Badge variant="info">{r.customerId==="any"?"Any Customer":cu?.label||"?"}</Badge>
              </div>
              <div style={{display:"flex",gap:4}}>
                <button onClick={()=>onEditRule(r)} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="edit" size={15} color={T.textSec}/></button>
                <button onClick={()=>{if(confirm("Delete this rule?"))onDeleteRule(r.id)}} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={15} color={T.danger}/></button>
              </div>
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {cks.length===0?<span style={{fontSize:12,color:T.textMut}}>No checklists</span>:cks.map(ck=><Badge key={ck.id} variant="muted">{ck.name}</Badge>)}
            </div>
          </div>;
        })}
      </div>
    </div>
  );
}

// ─── Edit Rule View ───────────────────────────────────────────

function EditRuleView({rule,orderTypes,customers,checklists,onSave}){
  const [typeId,setTypeId]=useState(rule?.orderTypeId||"any");
  const [custId,setCustId]=useState(rule?.customerId||"any");
  const [ckIds,setCkIds]=useState(rule?.checklistIds||[]);
  const toggleCk=id=>setCkIds(p=>p.includes(id)?p.filter(x=>x!==id):[...p,id]);
  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:20}}>
      <Field label="Order Type">
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          <Chip label="Any Order Type" active={typeId==="any"} onClick={()=>setTypeId("any")}/>
          {orderTypes.map(ot=><Chip key={ot.id} label={ot.label} active={typeId===ot.id} onClick={()=>setTypeId(ot.id)}/>)}
        </div>
      </Field>
      <Field label="Customer">
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          <Chip label="Any Customer" active={custId==="any"} onClick={()=>setCustId("any")}/>
          {customers.map(c=><Chip key={c.id} label={c.label} active={custId===c.id} onClick={()=>setCustId(c.id)}/>)}
        </div>
      </Field>
      <Field label="Checklists to Assign">
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {checklists.map(ck=><Toggle key={ck.id} active={ckIds.includes(ck.id)} onClick={()=>toggleCk(ck.id)} label={ck.name} sub={ck.subtitle}/>)}
        </div>
      </Field>
      <Btn onClick={()=>onSave({id:rule?.id||"rule_"+Date.now(),orderTypeId:typeId,customerId:custId,checklistIds:ckIds})} style={{width:"100%",marginTop:8}}>{rule?"Save Rule":"Create Rule"}</Btn>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Inventory View ───────────────────────────────────────────

// Equivalent items editor — chip list of currently-linked items + an inline add row.
function EquivalentItemsEditor({ items, currentItem, value, onChange }) {
  const safeValue = Array.isArray(value) ? value : [];
  const [adding, setAdding] = useState(false);
  const [pendingCategory, setPendingCategory] = useState("");
  const [pendingItemId, setPendingItemId] = useState("");

  const allCategories = Array.from(new Set((items || []).map(i => i.category).filter(c => c && c !== currentItem?.category)));
  const candidates = (items || []).filter(i => i.category === pendingCategory && i.isActive && i.id !== currentItem?.id);

  const itemFor = (id) => (items || []).find(i => i.id === id);

  const removeLink = (idx) => {
    const next = safeValue.filter((_, i) => i !== idx);
    onChange(next);
  };

  const addLink = () => {
    if (!pendingCategory || !pendingItemId) return;
    // Avoid duplicates of (category,itemId)
    const exists = safeValue.some(v => v.category === pendingCategory && v.itemId === pendingItemId);
    if (exists) { setAdding(false); setPendingCategory(""); setPendingItemId(""); return; }
    onChange([...safeValue, { category: pendingCategory, itemId: pendingItemId }]);
    setAdding(false); setPendingCategory(""); setPendingItemId("");
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
        {safeValue.length === 0 && <span style={{fontSize:11,color:T.textMut}}>No equivalents linked yet.</span>}
        {safeValue.map((link, i) => {
          const it = itemFor(link.itemId);
          const label = it ? `${link.category}: ${it.name}` : `${link.category}: (missing)`;
          return (
            <span key={i} style={{display:"inline-flex",alignItems:"center",gap:6,padding:"4px 10px",borderRadius:14,background:T.accentBg,border:`1px solid ${T.accentBorder}`,fontSize:12,color:T.accent}}>
              {label}
              <button onClick={()=>removeLink(i)} style={{background:"none",border:"none",cursor:"pointer",padding:0,color:T.danger,fontSize:14,lineHeight:1}}>✕</button>
            </span>
          );
        })}
      </div>

      {!adding ? (
        <Btn variant="ghost" small onClick={()=>{setAdding(true);setPendingCategory(allCategories[0]||"");setPendingItemId("")}} style={{alignSelf:"flex-start"}}>
          <Icon name="plus" size={12} color={T.textSec}/> Link Equivalent Item
        </Btn>
      ) : (
        <div style={{display:"flex",flexDirection:"column",gap:8,padding:10,background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.accentBorder}`}}>
          <select value={pendingCategory} onChange={e=>{setPendingCategory(e.target.value);setPendingItemId("")}}
            style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13}}>
            <option value="">— Category —</option>
            {allCategories.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={pendingItemId} onChange={e=>setPendingItemId(e.target.value)} disabled={!pendingCategory}
            style={{padding:"8px 10px",borderRadius:T.radSm,background:T.surface,border:`1px solid ${T.border}`,color:T.text,fontSize:13}}>
            <option value="">— Item —</option>
            {candidates.map(it => <option key={it.id} value={it.id}>{it.name}{it.abbreviation?` (${it.abbreviation})`:""}</option>)}
          </select>
          <div style={{display:"flex",gap:6}}>
            <Btn variant="ghost" small onClick={()=>{setAdding(false);setPendingCategory("");setPendingItemId("")}} style={{flex:1}}>Cancel</Btn>
            <Btn small onClick={addLink} disabled={!pendingCategory||!pendingItemId} style={{flex:1}}>Add</Btn>
          </div>
        </div>
      )}
    </div>
  );
}

function InventoryView({ items, categories, summary, isAdmin, addToast, onViewLedger, onCreateItem, onUpdateItem, onCreateCategory }) {
  const [activeTab,setActiveTab]=useState(categories[0]?.name||"Green Beans");
  const [showAdd,setShowAdd]=useState(false);
  const [newItem,setNewItem]=useState({name:"",category:"",unit:"kg",openingStock:"",minStockAlert:"",abbreviation:""});
  const [newCat,setNewCat]=useState("");
  const [editId,setEditId]=useState(null);
  const [editData,setEditData]=useState({});

  const tabItems=items.filter(i=>i.category===activeTab&&i.isActive);

  const getStockColor=(item)=>{
    if(item.minStockAlert>0&&item.currentStock<item.minStockAlert) return T.danger;
    if(item.minStockAlert>0&&item.currentStock<item.minStockAlert*1.5) return T.warning;
    return T.success;
  };

  const handleCreate=()=>{
    if(!newItem.name.trim()||!newItem.category) return;
    const ab=String(newItem.abbreviation||"").toUpperCase();
    if(!/^[A-Z0-9]{2,6}$/.test(ab)){addToast?.("Abbreviation must be 2-6 uppercase letters/digits","error");return;}
    onCreateItem({...newItem,abbreviation:ab,openingStock:parseFloat(newItem.openingStock)||0,minStockAlert:parseFloat(newItem.minStockAlert)||0});
    setNewItem({name:"",category:"",unit:"kg",openingStock:"",minStockAlert:"",abbreviation:""});setShowAdd(false);
  };

  const handleUpdate=(id)=>{
    onUpdateItem({id,...editData});setEditId(null);
  };

  const negativeItems = items.filter(i => i.isActive && (parseFloat(i.currentStock) || 0) < 0);

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:20}}>
      {/* ── Negative Stock Warning Banner ── */}
      {negativeItems.length > 0 && (
        <div style={{background:T.dangerBg,border:`1px solid ${T.danger}`,borderRadius:T.rad,padding:"12px 14px",display:"flex",alignItems:"flex-start",gap:10}}>
          <Icon name="alert-triangle" size={18} color={T.danger}/>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:13,fontWeight:600,color:T.danger,marginBottom:2}}>Warning: {negativeItems.length} item{negativeItems.length===1?"":"s"} ha{negativeItems.length===1?"s":"ve"} negative stock. Please review.</div>
            <div style={{fontSize:11,color:T.textSec,fontFamily:T.mono,overflowWrap:"anywhere"}}>
              {negativeItems.map(i => `${i.name} (${i.currentStock} ${i.unit})`).join(" · ")}
            </div>
          </div>
        </div>
      )}
      {/* ── Summary Cards ── */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <div style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.successBorder}`}}>
          <span style={{fontSize:11,color:T.textMut,display:"block"}}>Green Beans</span>
          <span style={{fontSize:20,fontWeight:700,color:T.success}}>{summary.greenBeans} <span style={{fontSize:12,fontWeight:400}}>kg</span></span>
        </div>
        <div style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.accentBorder}`}}>
          <span style={{fontSize:11,color:T.textMut,display:"block"}}>Roasted Beans</span>
          <span style={{fontSize:20,fontWeight:700,color:T.accent}}>{summary.roastedBeans} <span style={{fontSize:12,fontWeight:400}}>kg</span></span>
        </div>
        <div style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.infoBorder}`}}>
          <span style={{fontSize:11,color:T.textMut,display:"block"}}>Packed Goods</span>
          <span style={{fontSize:20,fontWeight:700,color:T.info}}>{summary.packedGoods} <span style={{fontSize:12,fontWeight:400}}>units</span></span>
        </div>
        <div style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${summary.lowStockCount>0?"rgba(232,93,93,0.2)":T.border}`}}>
          <span style={{fontSize:11,color:T.textMut,display:"block"}}>Low Stock Alerts</span>
          <span style={{fontSize:20,fontWeight:700,color:summary.lowStockCount>0?T.danger:T.textSec}}>{summary.lowStockCount}</span>
        </div>
      </div>

      {/* ── Category Tabs ── */}
      <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:4}}>
        {categories.map(c=>(
          <Chip key={c.id} label={c.name} active={activeTab===c.name} onClick={()=>setActiveTab(c.name)}/>
        ))}
        {isAdmin&&<Btn variant="ghost" small onClick={()=>{const n=prompt("New category name:");if(n&&n.trim())onCreateCategory(n.trim())}} style={{fontSize:11,whiteSpace:"nowrap"}}>+ Category</Btn>}
      </div>

      {/* ── Add Item Button ── */}
      {isAdmin&&<Btn small onClick={()=>{setShowAdd(!showAdd);setNewItem(p=>({...p,category:activeTab}))}} style={{width:"100%"}}>
        <Icon name="plus" size={16} color={T.bg}/> Add Inventory Item
      </Btn>}

      {/* ── Add Item Form ── */}
      {showAdd&&(
        <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.accentBorder}`,display:"flex",flexDirection:"column",gap:12}}>
          <Field label="Item Name"><Input value={newItem.name} onChange={v=>setNewItem(p=>({...p,name:v}))} placeholder="e.g., Arabica AA Grade"/></Field>
          <Field label="Abbreviation (2-6 chars, uppercase)">
            <Input value={newItem.abbreviation} onChange={v=>setNewItem(p=>({...p,abbreviation:v.toUpperCase().replace(/[^A-Z0-9]/g,"").slice(0,6)}))} placeholder="e.g., ACAB" style={{textTransform:"uppercase",fontFamily:T.mono}}/>
          </Field>
          <Field label="Category">
            <select value={newItem.category} onChange={e=>setNewItem(p=>({...p,category:e.target.value}))}
              style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14}}>
              <option value="">Select category</option>
              {categories.map(c=><option key={c.id} value={c.name}>{c.name}</option>)}
            </select>
          </Field>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <Field label="Unit">
              <select value={newItem.unit} onChange={e=>setNewItem(p=>({...p,unit:e.target.value}))}
                style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14}}>
                {["kg","grams","pieces","packs"].map(u=><option key={u} value={u}>{u}</option>)}
              </select>
            </Field>
            <Field label="Opening Stock"><Input value={newItem.openingStock} onChange={v=>setNewItem(p=>({...p,openingStock:v}))} type="number" placeholder="0"/></Field>
          </div>
          <Field label="Min Stock Alert (optional)"><Input value={newItem.minStockAlert} onChange={v=>setNewItem(p=>({...p,minStockAlert:v}))} type="number" placeholder="Alert when below..."/></Field>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="secondary" small onClick={()=>setShowAdd(false)} style={{flex:1}}>Cancel</Btn>
            <Btn small onClick={handleCreate} disabled={!newItem.name.trim()} style={{flex:1}}>Create Item</Btn>
          </div>
        </div>
      )}

      {/* ── Item Cards ── */}
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {tabItems.length===0?<Empty icon="layers" text={`No ${activeTab} items yet`} sub={isAdmin?"Tap + to add an inventory item":"Admin can add items"}/>:
          tabItems.map(item=>(
            <div key={item.id} style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.border}`,cursor:"pointer"}}
              onClick={()=>{if(editId!==item.id)onViewLedger(item)}}>
              {editId===item.id?(
                <div onClick={e=>e.stopPropagation()} style={{display:"flex",flexDirection:"column",gap:10}}>
                  <Field label="Name"><Input value={editData.name||""} onChange={v=>setEditData(p=>({...p,name:v}))}/></Field>
                  <Field label="Abbreviation (2-6 chars, uppercase)">
                    <Input value={editData.abbreviation||""} onChange={v=>setEditData(p=>({...p,abbreviation:v.toUpperCase().replace(/[^A-Z0-9]/g,"").slice(0,6)}))} placeholder="e.g., ACAB" style={{textTransform:"uppercase",fontFamily:T.mono}}/>
                  </Field>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    <Field label="Unit">
                      <select value={editData.unit||"kg"} onChange={e=>setEditData(p=>({...p,unit:e.target.value}))}
                        style={{width:"100%",padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13}}>
                        {["kg","grams","pieces","packs"].map(u=><option key={u} value={u}>{u}</option>)}
                      </select>
                    </Field>
                    <Field label="Min Alert"><Input value={editData.minStockAlert||""} onChange={v=>setEditData(p=>({...p,minStockAlert:v}))} type="number"/></Field>
                  </div>
                  <Field label="Linked Items in Other Categories">
                    <EquivalentItemsEditor items={items} currentItem={item} value={editData.equivalentItems||[]} onChange={v=>setEditData(p=>({...p,equivalentItems:v}))}/>
                  </Field>
                  <div style={{display:"flex",gap:8}}>
                    <Btn variant="secondary" small onClick={()=>setEditId(null)} style={{flex:1}}>Cancel</Btn>
                    <Btn small onClick={()=>handleUpdate(item.id)} style={{flex:1}}>Save</Btn>
                    <Btn variant="danger" small onClick={()=>{onUpdateItem({id:item.id,isActive:false});setEditId(null)}} style={{flex:1}}>Deactivate</Btn>
                  </div>
                </div>
              ):(
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                  <div>
                    <span style={{fontSize:14,fontWeight:600,color:T.text}}>{item.name}</span>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                      <span style={{fontSize:18,fontWeight:700,color:getStockColor(item)}}>{item.currentStock}</span>
                      <span style={{fontSize:12,color:T.textMut}}>{item.unit}</span>
                      {item.minStockAlert>0&&item.currentStock<item.minStockAlert&&<Badge variant="danger" style={{fontSize:9}}>LOW</Badge>}
                    </div>
                  </div>
                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                    {isAdmin&&<button onClick={e=>{e.stopPropagation();setEditId(item.id);setEditData({name:item.name,unit:item.unit,minStockAlert:item.minStockAlert||"",abbreviation:item.abbreviation||"",equivalentItems:Array.isArray(item.equivalentItems)?item.equivalentItems:[]})}} style={{background:"none",border:"none",cursor:"pointer",padding:6}}>
                      <Icon name="edit" size={15} color={T.textSec}/>
                    </button>}
                    <Icon name="chevron" size={16} color={T.textMut}/>
                  </div>
                </div>
              )}
            </div>
          ))
        }
      </div>
    </div>
  );
}

// ─── Inventory Ledger View ────────────────────────────────────

function InventoryLedgerView({ item, isAdmin, addToast, onAdjust }) {
  const [entries,setEntries]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showAdjust,setShowAdjust]=useState(false);
  const [adjType,setAdjType]=useState("addition");
  const [adjQty,setAdjQty]=useState("");
  const [adjNotes,setAdjNotes]=useState("");
  const [saving,setSaving]=useState(false);

  useEffect(()=>{
    API.get("getInventoryLedger",{item_id:item.id}).then(data=>{setEntries(data||[]);setLoading(false)}).catch(()=>setLoading(false));
  },[item.id]);

  const handleAdjust=async()=>{
    const n = parseFloat(adjQty);
    if(adjQty===""||isNaN(n))return;
    // IN adjustments (addition) must be > 0. OUT adjustments (reduction) accept any non-zero
    // value — a negative reduction is a manual correction that restores stock.
    if(adjType==="addition" && n<=0) return;
    if(adjType==="reduction" && n===0) return;
    setSaving(true);
    try{
      await onAdjust({itemId:item.id,adjustmentType:adjType,quantity:n,notes:adjNotes});
      // Refresh ledger
      const data=await API.get("getInventoryLedger",{item_id:item.id});
      setEntries(data||[]);
      setShowAdjust(false);setAdjQty("");setAdjNotes("");
    }catch{}
    setSaving(false);
  };

  if(loading) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.textSec,animation:"pulse 1.5s infinite"}}>Loading ledger...</p></div>;

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:16}}>
      {/* ── Item Header ── */}
      <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.border}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div>
            <h3 style={{fontSize:16,fontWeight:600,color:T.text}}>{item.name}</h3>
            <Badge variant="muted">{item.category}</Badge>
          </div>
          <div style={{textAlign:"right"}}>
            <span style={{fontSize:22,fontWeight:700,color:T.accent}}>{item.currentStock}</span>
            <span style={{fontSize:12,color:T.textMut,display:"block"}}>{item.unit} in stock</span>
          </div>
        </div>
      </div>

      {/* ── Add Adjustment ── */}
      {isAdmin&&<Btn small onClick={()=>setShowAdjust(!showAdjust)} style={{width:"100%"}}>
        <Icon name="plus" size={16} color={T.bg}/> Add Adjustment
      </Btn>}

      {showAdjust&&(
        <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.accentBorder}`,display:"flex",flexDirection:"column",gap:12}}>
          <Field label="Adjustment Type">
            <div style={{display:"flex",gap:8}}>
              <Chip label="Addition (+)" active={adjType==="addition"} onClick={()=>setAdjType("addition")}/>
              <Chip label="Reduction (-)" active={adjType==="reduction"} onClick={()=>setAdjType("reduction")}/>
            </div>
          </Field>
          <Field label="Quantity">
            <Input value={adjQty} onChange={setAdjQty} type="number"
              allowNegative={adjType==="reduction"}
              placeholder={adjType==="reduction"?"Enter quantity (negative allowed for corrections)...":"Enter quantity..."}/>
            {adjType==="reduction" && parseFloat(adjQty)<0 && (
              <div style={{marginTop:6,fontSize:11,color:T.textMut}}>Negative reduction restores stock (manual correction)</div>
            )}
          </Field>
          <Field label="Reason / Notes"><Input value={adjNotes} onChange={setAdjNotes} placeholder="Why this adjustment?"/></Field>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="secondary" small onClick={()=>setShowAdjust(false)} style={{flex:1}}>Cancel</Btn>
            <Btn small onClick={handleAdjust} disabled={saving||adjQty===""||isNaN(parseFloat(adjQty))||(adjType==="addition"&&parseFloat(adjQty)<=0)||(adjType==="reduction"&&parseFloat(adjQty)===0)} style={{flex:1}}>{saving?"Saving...":"Save Adjustment"}</Btn>
          </div>
        </div>
      )}

      {/* ── Ledger Entries ── */}
      <Section icon="clipboard" count={entries.length}>Transaction Ledger</Section>
      {entries.length===0?<Empty icon="clipboard" text="No transactions yet" sub="Stock movements will appear here"/>:
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {entries.map(e=>{
            const isIn=e.type==="IN"||(e.type==="ADJUSTMENT"&&e.quantity>0);
            return (
              <div key={e.id} style={{background:T.card,borderRadius:T.radSm,padding:"12px 14px",border:`1px solid ${T.border}`,fontSize:13}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:14,fontWeight:600,color:isIn?T.success:T.danger}}>{isIn?"+":""}{e.quantity} {item.unit}</span>
                    <Badge variant={e.type==="IN"?"success":e.type==="OUT"?"danger":"muted"}>{e.type}</Badge>
                  </div>
                  <span style={{fontSize:12,color:T.textMut,fontFamily:T.mono}}>Bal: {e.balanceAfter}</span>
                </div>
                <div style={{display:"flex",gap:12,color:T.textMut,fontSize:11,flexWrap:"wrap"}}>
                  {e.date&&<span>{e.date}</span>}
                  {e.doneBy&&<span>by {e.doneBy}</span>}
                  {e.referenceType==="checklist"&&<span>Ref: checklist</span>}
                </div>
                {e.notes&&<p style={{fontSize:12,color:T.textSec,marginTop:4}}>{e.notes}</p>}
              </div>
            );
          })}
        </div>
      }
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Blends ────────────────────────────────────────────────────

const blendComponentSummary = (components) => {
  if (!Array.isArray(components) || components.length === 0) return "No components";
  return components.map(c => `${c.percentage}% ${c.itemName || c.category || "Item"}`).join(" + ");
};

const blendTotalPercent = (components) => {
  if (!Array.isArray(components)) return 0;
  return components.reduce((s, c) => s + (parseFloat(c.percentage) || 0), 0);
};

// Returns an array of {component, requiredQty} for an order line of `quantity` kgs using a blend snapshot.
const computeBlendRequirements = (blendOrSnapshot, quantity) => {
  const qty = parseFloat(quantity) || 0;
  const components = blendOrSnapshot?.components || [];
  return components.map(c => ({
    component: c,
    requiredQty: Math.round(((parseFloat(c.percentage) || 0) / 100) * qty * 100) / 100,
  }));
};

// Resolve the OUTPUT inventory item for a stage's tagged entry / approved checklist response.
// Priority:
//   1. An IN-tx inventoryLink question (the item this checklist produces). If itemSource.type="fixed",
//      use itemSource.itemId directly. If "field", read the field's value (usually an input-category item)
//      and map to the link's target category via the resolved item's equivalent_items.
//   2. A non-IN inventoryLink question (fallback — rarely needed).
//   3. Any question of type "inventory_item".
// Works with both approvedEntries responses ([{question, response, remark}]) and
// untaggedChecklists responses ([{questionIndex, questionText, response}]).
// Returns { id, name } or null.
const resolveTaggedEntryItem = (te, checklists, inventoryItems, approvedEntries, untaggedChecklists) => {
  if (!te || !te.checklistId) return null;
  const ck = (checklists || []).find(c => c.id === te.checklistId);
  if (!ck || !Array.isArray(ck.questions)) return null;
  const nq = ck.questions;

  // Prefer an IN-tx inventoryLink (that's the OUTPUT item of this checklist)
  let chosenQ = null;
  for (let i = 0; i < nq.length; i++) {
    const q = nq[i];
    if (q?.inventoryLink?.enabled && q.inventoryLink.txType === "IN") { chosenQ = { q, idx: i }; break; }
  }
  // Fallback: any inventoryLink
  if (!chosenQ) {
    for (let i = 0; i < nq.length; i++) {
      const q = nq[i];
      if (q?.inventoryLink?.enabled) { chosenQ = { q, idx: i }; break; }
    }
  }
  // Fallback: an inventory_item question
  if (!chosenQ) {
    for (let i = 0; i < nq.length; i++) {
      if (nq[i]?.type === "inventory_item") { chosenQ = { q: nq[i], idx: i, plainInventory: true }; break; }
    }
  }
  if (!chosenQ) return null;

  const q = chosenQ.q;
  const link = q.inventoryLink || null;
  const targetCategory = link?.category || "";

  // Direct fixed item
  if (link?.itemSource?.type === "fixed" && link.itemSource.itemId) {
    const fixed = (inventoryItems || []).find(it => it.id === link.itemSource.itemId);
    if (fixed) return { id: fixed.id, name: fixed.name };
    return { id: link.itemSource.itemId, name: link.itemSource.itemId };
  }

  // Otherwise read the source field value from responses
  let itemFieldIdx = -1;
  if (link?.itemSource?.type === "field" && link.itemSource.fieldIdx !== undefined && link.itemSource.fieldIdx !== null && link.itemSource.fieldIdx !== "") {
    itemFieldIdx = Number(link.itemSource.fieldIdx);
  } else if (chosenQ.plainInventory) {
    itemFieldIdx = chosenQ.idx;
  }
  if (itemFieldIdx < 0) return null;

  const autoId = te.autoId || te.responseId;
  let responses = null;
  const aeArr = approvedEntries?.[te.checklistId] || [];
  const aeMatch = aeArr.find(e => (e.autoId || e.linkedId) === autoId);
  if (aeMatch && Array.isArray(aeMatch.responses)) responses = aeMatch.responses;
  if (!responses) {
    const utMatch = (untaggedChecklists || []).find(u => u.autoId === autoId);
    if (utMatch && Array.isArray(utMatch.responses)) responses = utMatch.responses;
  }
  // Support passing responses directly via te.responses (e.g., when caller already has them)
  if (!responses && Array.isArray(te.responses)) responses = te.responses;
  // Support object-style { questionText: value } fallback
  if (!Array.isArray(responses) && responses && typeof responses === "object") {
    const targetText = nq[itemFieldIdx]?.text;
    const v = responses[targetText];
    if (v !== undefined) {
      const item = (inventoryItems || []).find(it => it.id === v || it.name === v);
      if (item && targetCategory && item.category !== targetCategory) {
        const eqs = Array.isArray(item.equivalentItems) ? item.equivalentItems : [];
        const eq = eqs.find(e => e.category === targetCategory);
        if (eq?.itemId) {
          const eqItem = (inventoryItems || []).find(it => it.id === eq.itemId);
          if (eqItem) return { id: eqItem.id, name: eqItem.name };
        }
      }
      if (item) return { id: item.id, name: item.name };
      return { id: "", name: String(v) };
    }
  }
  if (!Array.isArray(responses)) return null;

  const targetQuestionText = nq[itemFieldIdx]?.text;
  let itemVal = "";
  for (let i = 0; i < responses.length; i++) {
    const r = responses[i];
    const qIdx = r.questionIndex;
    const qText = r.question || r.questionText;
    if ((qIdx !== undefined && Number(qIdx) === itemFieldIdx) || (qText && String(qText) === String(targetQuestionText))) {
      itemVal = String(r.response || "").trim();
      break;
    }
  }
  if (!itemVal) return null;

  // Resolve the raw value to an inventory item
  const raw = (inventoryItems || []).find(it => it.id === itemVal || it.name === itemVal);
  if (!raw) return { id: "", name: itemVal };

  // If the link has a target category and the resolved item is in a different category,
  // use equivalent_items to find the item in the target category (matches backend processInventoryLinks behavior).
  if (targetCategory && raw.category && raw.category !== targetCategory) {
    const eqs = Array.isArray(raw.equivalentItems) ? raw.equivalentItems : [];
    const eq = eqs.find(e => e.category === targetCategory);
    if (eq?.itemId) {
      const eqItem = (inventoryItems || []).find(it => it.id === eq.itemId);
      if (eqItem) return { id: eqItem.id, name: eqItem.name };
      return { id: eq.itemId, name: eq.itemId };
    }
  }
  return { id: raw.id, name: raw.name };
};

// Helper: does the given checklist have any configured inventoryLink (so we can filter entries by ingredient)?
const checklistHasInventoryLink = (checklists, checklistId) => {
  const ck = (checklists || []).find(c => c.id === checklistId);
  if (!ck || !Array.isArray(ck.questions)) return false;
  return ck.questions.some(q => q?.inventoryLink?.enabled) || ck.questions.some(q => q?.type === "inventory_item");
};

// Compute blend composition analysis for an order.
// Returns a deduped array: [{ itemId, itemName, expected, actual, status: "ok"|"missing"|"under"|"over"|"unexpected", overPct }]
// Or null if the order has no blend lines.
const computeOrderBlendAnalysis = (order, checklists, inventoryItems, approvedEntries, untaggedChecklists) => {
  const lines = Array.isArray(order?.orderLines) ? order.orderLines.filter(l => Array.isArray(l.blendComponents) && l.blendComponents.length > 0) : [];
  if (lines.length === 0) return null;

  const keyFor = (itemId, itemName) => itemId ? `id:${itemId}` : `n:${String(itemName || "").toLowerCase().trim()}`;
  const expected = {}; // key -> { itemId, itemName, qty }
  lines.forEach(line => {
    const qty = parseFloat(line.quantity) || 0;
    (line.blendComponents || []).forEach(c => {
      const k = keyFor(c.itemId, c.itemName);
      if (!expected[k]) expected[k] = { itemId: c.itemId || "", itemName: c.itemName || k, qty: 0 };
      expected[k].qty += (parseFloat(c.percentage) || 0) / 100 * qty;
    });
  });

  const actual = {};
  const stages = Array.isArray(order?.stages) ? order.stages : [];
  stages.forEach(s => {
    (s.taggedEntries || []).forEach(te => {
      const item = resolveTaggedEntryItem(te, checklists, inventoryItems, approvedEntries, untaggedChecklists);
      if (!item) return;
      const k = keyFor(item.id, item.name);
      if (!actual[k]) actual[k] = { itemId: item.id, itemName: item.name, qty: 0 };
      actual[k].qty += parseFloat(te.qty) || 0;
    });
  });

  const allKeys = new Set([...Object.keys(expected), ...Object.keys(actual)]);
  const rows = [];
  allKeys.forEach(k => {
    const e = expected[k] || { itemId: "", itemName: (actual[k]?.itemName) || k, qty: 0 };
    const a = actual[k] || { itemId: e.itemId, itemName: e.itemName, qty: 0 };
    const expQty = Math.round(e.qty * 100) / 100;
    const actQty = Math.round(a.qty * 100) / 100;
    const tolerance = Math.max(expQty * 0.05, 0.5);
    let status;
    if (expQty === 0 && actQty === 0) status = "ok";
    else if (expQty === 0) status = "unexpected";
    else if (actQty === 0) status = "missing";
    else if (Math.abs(actQty - expQty) <= tolerance) status = "ok";
    else if (actQty < expQty) status = "under";
    else status = "over";
    const overPct = expQty > 0 ? ((actQty - expQty) / expQty) * 100 : 0;
    rows.push({ itemId: e.itemId || a.itemId, itemName: e.itemName || a.itemName, expected: expQty, actual: actQty, status, overPct });
  });
  return rows;
};

// Compute per-item tagged quantity across all stages: { [itemId or lowerName]: qty }
const computeTaggedByItem = (order, checklists, inventoryItems, approvedEntries, untaggedChecklists) => {
  const map = {};
  const stages = Array.isArray(order?.stages) ? order.stages : [];
  stages.forEach(s => {
    (s.taggedEntries || []).forEach(te => {
      const item = resolveTaggedEntryItem(te, checklists, inventoryItems, approvedEntries, untaggedChecklists);
      if (!item) return;
      const k = item.id ? `id:${item.id}` : `n:${String(item.name || "").toLowerCase().trim()}`;
      map[k] = (map[k] || 0) + (parseFloat(te.qty) || 0);
    });
  });
  return map;
};

const blendItemKey = (itemId, itemName) => itemId ? `id:${itemId}` : `n:${String(itemName || "").toLowerCase().trim()}`;

// Returns true if an order has at least one blend line with defined components
const isBlendOrder = (order) => {
  const lines = Array.isArray(order?.orderLines) ? order.orderLines : [];
  return lines.some(l => Array.isArray(l.blendComponents) && l.blendComponents.length > 0);
};

// Analyze a single blend line at a stage level. Returns per-ingredient detail + ratio status.
// taggedEntriesForLine: taggedEntries scoped to this stage AND this blend line (blendLineIndex === lineIndex).
const analyzeBlendLineAtStage = (blendLine, taggedEntriesForLine) => {
  const lineQty = parseFloat(blendLine.quantity) || 0;
  const components = Array.isArray(blendLine.blendComponents) ? blendLine.blendComponents : [];
  const mixedTags = taggedEntriesForLine.filter(te => te.isMixed === true);
  const mixedTotal = mixedTags.reduce((s, t) => s + (parseFloat(t.qty) || 0), 0);

  const directPerKey = {}; // key -> { entries, total }
  taggedEntriesForLine.forEach(te => {
    if (te.isMixed === true) return;
    const k = blendItemKey(te.componentItemId, te.componentItemName);
    if (!directPerKey[k]) directPerKey[k] = { entries: [], total: 0 };
    directPerKey[k].entries.push(te);
    directPerKey[k].total += parseFloat(te.qty) || 0;
  });
  const directTotal = Object.values(directPerKey).reduce((s, r) => s + r.total, 0);

  const perIngredient = components.map(c => {
    const key = blendItemKey(c.itemId, c.itemName);
    const pct = parseFloat(c.percentage) || 0;
    const required = Math.round(pct / 100 * lineQty * 100) / 100;
    const directRec = directPerKey[key] || { entries: [], total: 0 };
    const mixedContribution = mixedTotal * pct / 100;
    const totalTagged = Math.round((directRec.total + mixedContribution) * 100) / 100;
    const remaining = Math.max(0, Math.round((required - totalTagged) * 100) / 100);
    return {
      component: c, key, percentage: pct, required,
      directTagged: Math.round(directRec.total * 100) / 100,
      mixedContribution: Math.round(mixedContribution * 100) / 100,
      totalTagged, remaining,
      entries: directRec.entries,
    };
  });

  // Ratio check (only for direct tags; mixed are pre-validated at tag time)
  let ratioOk = true;
  let actualRatioParts = [], requiredRatioParts = [];
  if (directTotal > 0) {
    requiredRatioParts = components.map(c => ({ itemName: c.itemName, pct: parseFloat(c.percentage) || 0 }));
    actualRatioParts = perIngredient.map(p => ({ itemName: p.component.itemName, pct: directTotal > 0 ? Math.round(p.directTagged / directTotal * 1000) / 10 : 0 }));
    for (let i = 0; i < components.length; i++) {
      const reqP = parseFloat(components[i].percentage) || 0;
      const actP = directTotal > 0 ? (perIngredient[i].directTagged / directTotal * 100) : 0;
      if (Math.abs(actP - reqP) > 0.5) { ratioOk = false; break; }
    }
  }

  const allIngredientsComplete = perIngredient.every(p => p.remaining <= 0.01);
  const lineComplete = allIngredientsComplete && ratioOk;

  return {
    perIngredient, mixedTags, mixedTotal, directTotal,
    ratioOk, actualRatioParts, requiredRatioParts,
    allIngredientsComplete, lineComplete,
  };
};

// Check whether a stage is complete given an order.
// For blend orders: all blend lines must be lineComplete at this stage.
// For non-blend orders: tagged total >= requiredQty (or has tags when no requiredQty).
const isStageComplete = (stage, order) => {
  if (isBlendOrder(order)) {
    const lines = Array.isArray(order.orderLines) ? order.orderLines : [];
    const allTags = Array.isArray(stage.taggedEntries) ? stage.taggedEntries : [];
    for (let li = 0; li < lines.length; li++) {
      const line = lines[li];
      if (!Array.isArray(line.blendComponents) || line.blendComponents.length === 0) continue;
      const lineTags = allTags.filter(te => Number(te.blendLineIndex) === li);
      const a = analyzeBlendLineAtStage(line, lineTags);
      if (!a.lineComplete) return false;
    }
    return true;
  }
  const req = parseFloat(stage.requiredQty) || 0;
  const tagged = (stage.taggedEntries || []).reduce((s, t) => s + (parseFloat(t.qty) || 0), 0);
  if (req > 0) return tagged >= req;
  if (stage.checklistId) return (stage.taggedEntries || []).length > 0;
  return true;
};

// Return blends whose component ratios exactly match a given blend line's ratios.
const findMatchingBlendsForLine = (blendLine, blends) => {
  if (!blendLine || !Array.isArray(blendLine.blendComponents)) return [];
  const keyFor = (c) => blendItemKey(c.itemId, c.itemName);
  const lineMap = {};
  (blendLine.blendComponents || []).forEach(c => { lineMap[keyFor(c)] = (lineMap[keyFor(c)] || 0) + (parseFloat(c.percentage)||0); });
  return (blends || []).filter(b => {
    const bComps = Array.isArray(b.components) ? b.components : [];
    const bMap = {};
    bComps.forEach(c => { bMap[keyFor(c)] = (bMap[keyFor(c)] || 0) + (parseFloat(c.percentage)||0); });
    const allKeys = new Set([...Object.keys(lineMap), ...Object.keys(bMap)]);
    for (const k of allKeys) {
      if (Math.abs((lineMap[k]||0) - (bMap[k]||0)) > 0.01) return false;
    }
    return true;
  });
};

function BlendsPage({ blends, customers, isAdmin, inventoryItems, onCreate, onEdit, onDelete, onImport, addToast }) {
  const [filter, setFilter] = useState("All");
  const [importPayload, setImportPayload] = useState(null); // { parsed, errors } for preview modal
  const [importError, setImportError] = useState("");
  const fileInputRef = useRef(null);
  const safeBlends = blends || [];
  const customerNames = Array.from(new Set(safeBlends.map(b => b.customer).filter(c => c && c !== "General")));
  const filtered = safeBlends.filter(b => {
    if (filter === "All") return true;
    if (filter === "General") return (b.customer || "General") === "General";
    return b.customer === filter;
  });

  const handleExport = () => {
    try {
      exportBlendsToExcel(safeBlends);
      if (addToast) addToast("Blends exported", "success");
    } catch (e) {
      if (addToast) addToast("Export failed: " + e.message, "error");
    }
  };

  const handleImportPick = () => {
    setImportError("");
    if (fileInputRef.current) { fileInputRef.current.value = ""; fileInputRef.current.click(); }
  };

  const handleImportFile = async (e) => {
    let file = null;
    try {
      file = e && e.target && e.target.files ? e.target.files[0] : null;
    } catch (_) { file = null; }
    if (!file) return;
    setImportError("");
    try {
      const parsed = await parseBlendsExcel(file, inventoryItems || []);
      if (!parsed || !parsed.ok) {
        const msg = (parsed && parsed.error) || "Failed to parse file";
        setImportError(msg);
        if (addToast) addToast("Import failed: " + msg.split("\n")[0], "error");
        return;
      }
      setImportPayload(parsed);
    } catch (err) {
      const msg = (err && err.message) ? err.message : String(err || "Unknown error");
      setImportError("Import failed: " + msg);
      if (addToast) addToast("Import failed: " + msg, "error");
    }
  };

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:16}}>
      {isAdmin && (
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <Btn onClick={onCreate} style={{flex:"1 1 160px"}}><Icon name="plus" size={18} color={T.bg}/> Create Blend</Btn>
          <Btn variant="secondary" onClick={handleExport} disabled={safeBlends.length===0} style={{flex:"1 1 120px"}}>
            <Icon name="externalLink" size={16} color={T.text}/> Export
          </Btn>
          <Btn variant="secondary" onClick={handleImportPick} style={{flex:"1 1 120px"}}>
            <Icon name="clipboard" size={16} color={T.text}/> Import
          </Btn>
          <input ref={fileInputRef} type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            onChange={handleImportFile} style={{display:"none"}}/>
        </div>
      )}

      {importError && (
        <div style={{background:T.dangerBg,border:"1px solid rgba(232,93,93,0.25)",borderRadius:T.radSm,padding:"10px 14px",display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
          <span style={{fontSize:13,color:T.danger,whiteSpace:"pre-wrap"}}>{importError}</span>
          <button onClick={()=>setImportError("")} style={{background:"none",border:"none",cursor:"pointer",padding:2}}><Icon name="x" size={14} color={T.danger}/></button>
        </div>
      )}

      <div style={{display:"flex",gap:6,overflowX:"auto",paddingBottom:4}}>
        <Chip label="All" active={filter==="All"} onClick={()=>setFilter("All")}/>
        <Chip label="General" active={filter==="General"} onClick={()=>setFilter("General")}/>
        {customerNames.map(c => <Chip key={c} label={c} active={filter===c} onClick={()=>setFilter(c)}/>)}
      </div>

      {filtered.length === 0 ? <Empty icon="coffee" text="No blends yet" sub={isAdmin?"Tap Create Blend to define one":"Admin can add blends"}/> :
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {filtered.map(b => <BlendCard key={b.id} blend={b} isAdmin={isAdmin} onEdit={()=>onEdit(b)} onDelete={()=>{if(confirm("Delete (deactivate) this blend?"))onDelete(b.id)}}/>)}
        </div>
      }

      {importPayload && (
        <ImportBlendsPreviewModal
          payload={importPayload}
          existingBlends={safeBlends}
          onCancel={()=>setImportPayload(null)}
          onConfirm={async (selections)=>{
            try {
              const result = await onImport(selections);
              setImportPayload(null);
              const r = result || {created:0, updated:0, skipped:0};
              if (addToast) addToast(`Import complete — ${r.created} created, ${r.updated} updated, ${r.skipped} skipped`, "success");
            } catch (err) {
              const msg = (err && err.message) ? err.message : String(err || "Unknown error");
              if (addToast) addToast("Import failed: " + msg, "error");
              setImportPayload(null);
            }
          }}/>
      )}
    </div>
  );
}

// ── Export blends to .xlsx (one row per component) ──
function exportBlendsToExcel(blends) {
  const rows = [];
  (blends || []).forEach(b => {
    const comps = Array.isArray(b.components) ? b.components : [];
    const sortedComps = comps.slice().sort((a, b2) => (parseFloat(b2.percentage)||0) - (parseFloat(a.percentage)||0));
    if (sortedComps.length === 0) {
      rows.push({
        "Blend Name": b.name || "",
        "Customer": b.customer || "General",
        "Description": b.description || "",
        "Component Name": "",
        "Component Category": "",
        "Percentage": "",
        "Active": b.isActive === false ? "No" : "Yes",
      });
    } else {
      sortedComps.forEach(c => {
        rows.push({
          "Blend Name": b.name || "",
          "Customer": b.customer || "General",
          "Description": b.description || "",
          "Component Name": c.itemName || "",
          "Component Category": c.category || "",
          "Percentage": parseFloat(c.percentage) || 0,
          "Active": b.isActive === false ? "No" : "Yes",
        });
      });
    }
  });
  rows.sort((a, b) => {
    const byName = String(a["Blend Name"]).localeCompare(String(b["Blend Name"]));
    if (byName !== 0) return byName;
    return (parseFloat(b["Percentage"])||0) - (parseFloat(a["Percentage"])||0);
  });
  const ws = XLSX.utils.json_to_sheet(rows, {
    header: ["Blend Name","Customer","Description","Component Name","Component Category","Percentage","Active"]
  });
  ws["!cols"] = [{wch:28},{wch:20},{wch:32},{wch:28},{wch:20},{wch:12},{wch:8}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Blends");
  const now = new Date();
  const dd = String(now.getDate()).padStart(2,"0");
  const mm = String(now.getMonth()+1).padStart(2,"0");
  const yyyy = now.getFullYear();
  XLSX.writeFile(wb, `Sunoha_Blends_${dd}${mm}${yyyy}.xlsx`);
}

// ── Parse + validate an .xlsx into grouped blends ──
async function parseBlendsExcel(file, inventoryItems) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) return { ok: false, error: "Excel file has no sheets" };
  const ws = wb.Sheets[sheetName];
  const rawRows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (rawRows.length === 0) return { ok: false, error: "No data rows found in Excel file" };

  const required = ["Blend Name","Customer","Component Name","Component Category","Percentage"];
  const firstRow = rawRows[0];
  const missing = required.filter(col => !(col in firstRow));
  if (missing.length > 0) return { ok: false, error: "Missing required column(s): " + missing.join(", ") };

  // Group by blend name, preserving first-seen Customer / Description / Active
  const groupsMap = {};
  const order = [];
  const errors = [];
  rawRows.forEach((r, idx) => {
    const rowNum = idx + 2; // excel row number (1-indexed + header)
    const blendName = String(r["Blend Name"] || "").trim();
    if (!blendName) {
      errors.push(`Row ${rowNum}: Blend Name must not be empty`);
      return;
    }
    const pct = parseFloat(r["Percentage"]);
    if (isNaN(pct) || pct <= 0) {
      errors.push(`Row ${rowNum}: Percentage must be a positive number (got "${r["Percentage"]}") for "${blendName}"`);
      return;
    }
    const compName = String(r["Component Name"] || "").trim();
    const compCat = String(r["Component Category"] || "").trim();
    if (!compName) {
      errors.push(`Row ${rowNum}: Component Name must not be empty for "${blendName}"`);
      return;
    }
    if (!groupsMap[blendName]) {
      groupsMap[blendName] = {
        name: blendName,
        customer: String(r["Customer"] || "General").trim() || "General",
        description: String(r["Description"] || "").trim(),
        isActive: String(r["Active"] || "Yes").trim().toLowerCase() !== "no",
        components: [],
      };
      order.push(blendName);
    }
    // Resolve itemId via inventoryItems by case-insensitive name match
    const match = (inventoryItems || []).find(it => String(it.name || "").toLowerCase().trim() === compName.toLowerCase());
    groupsMap[blendName].components.push({
      category: compCat,
      itemId: match ? match.id : "",
      itemName: compName,
      percentage: pct,
    });
  });

  // Validate each blend totals to 100%
  order.forEach(name => {
    const g = groupsMap[name];
    const total = g.components.reduce((s, c) => s + (parseFloat(c.percentage)||0), 0);
    if (Math.abs(total - 100) > 0.001) {
      errors.push(`"${name}": components total ${Math.round(total*100)/100}% (must equal 100%)`);
    }
  });

  if (errors.length > 0) {
    return { ok: false, error: "Validation failed:\n• " + errors.join("\n• ") };
  }

  const blends = order.map(n => groupsMap[n]);
  return { ok: true, blends };
}

function ImportBlendsPreviewModal({ payload, existingBlends, onCancel, onConfirm }) {
  const existingByName = {};
  (existingBlends || []).forEach(b => { if (b && b.name) existingByName[String(b.name).toLowerCase()] = b; });
  const items = payload.blends.map(b => {
    const existing = existingByName[String(b.name).toLowerCase()];
    return {
      blend: b,
      existing: existing || null,
      isNew: !existing,
      accept: !existing, // default: NEW checked, UPDATE unchecked
    };
  });
  const [state, setState] = useState(items);
  const [submitting, setSubmitting] = useState(false);

  const toggle = (i) => setState(prev => prev.map((it, idx) => idx === i ? { ...it, accept: !it.accept } : it));
  const newCount = state.filter(s => s.isNew).length;
  const updateCount = state.filter(s => !s.isNew).length;
  const selectedCount = state.filter(s => s.accept).length;

  const handleConfirm = async () => {
    setSubmitting(true);
    try { await onConfirm(state); } finally { setSubmitting(false); }
  };

  return (
    <div onClick={onCancel} style={{position:"fixed",inset:0,zIndex:900,background:"rgba(0,0,0,0.7)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div onClick={e=>e.stopPropagation()} style={{background:T.card,borderRadius:T.rad,padding:20,maxWidth:720,width:"100%",maxHeight:"90vh",overflowY:"auto",border:`1px solid ${T.border}`,boxShadow:"0 12px 40px rgba(0,0,0,0.6)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <h3 style={{fontSize:16,fontWeight:600,color:T.text}}>Import Blends — Preview</h3>
          <button onClick={onCancel} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="x" size={20} color={T.textSec}/></button>
        </div>

        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {state.map((it, i) => (
            <div key={i} style={{background:T.bg,borderRadius:T.radSm,padding:12,border:`1px solid ${it.isNew?T.successBorder:T.warningBorder}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:8}}>
                <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",flex:1,minWidth:0}}>
                  <input type="checkbox" checked={it.accept} onChange={()=>toggle(i)}
                    style={{width:16,height:16,cursor:"pointer",accentColor:T.accent}}/>
                  <div style={{minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                      <span style={{fontSize:14,fontWeight:600,color:T.text}}>{it.blend.name}</span>
                      {it.isNew
                        ? <Badge variant="success" style={{fontSize:10}}>NEW</Badge>
                        : <span style={{display:"inline-flex",alignItems:"center",padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:500,letterSpacing:".02em",background:T.warningBg,color:T.warning,border:`1px solid ${T.warningBorder}`,whiteSpace:"nowrap"}}>UPDATE</span>}
                      <Badge variant={it.blend.customer==="General"?"info":"default"} style={{fontSize:10}}>{it.blend.customer}</Badge>
                    </div>
                    {it.blend.description && <p style={{fontSize:11,color:T.textMut,marginTop:2}}>{it.blend.description}</p>}
                  </div>
                </label>
              </div>

              {it.isNew ? (
                <div style={{paddingLeft:24}}>
                  {it.blend.components.map((c, ci) => (
                    <div key={ci} style={{display:"flex",justifyContent:"space-between",padding:"4px 8px",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`,marginBottom:4,fontSize:12}}>
                      <span style={{color:T.textSec}}>{c.itemName}{c.category?` (${c.category})`:""}</span>
                      <span style={{color:T.accent,fontFamily:T.mono,fontWeight:600}}>{c.percentage}%</span>
                    </div>
                  ))}
                </div>
              ) : (
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,paddingLeft:24}}>
                  <div>
                    <p style={{fontSize:10,color:T.textMut,fontWeight:600,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.05em"}}>Current</p>
                    {(it.existing.components||[]).map((c, ci) => (
                      <div key={ci} style={{display:"flex",justifyContent:"space-between",padding:"4px 8px",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.border}`,marginBottom:4,fontSize:11}}>
                        <span style={{color:T.textSec,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.itemName}{c.category?` (${c.category})`:""}</span>
                        <span style={{color:T.textMut,fontFamily:T.mono,flexShrink:0,marginLeft:4}}>{c.percentage}%</span>
                      </div>
                    ))}
                  </div>
                  <div>
                    <p style={{fontSize:10,color:T.warning,fontWeight:600,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.05em"}}>Incoming</p>
                    {it.blend.components.map((c, ci) => (
                      <div key={ci} style={{display:"flex",justifyContent:"space-between",padding:"4px 8px",background:T.surface,borderRadius:T.radSm,border:`1px solid ${T.warningBorder}`,marginBottom:4,fontSize:11}}>
                        <span style={{color:T.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.itemName}{c.category?` (${c.category})`:""}</span>
                        <span style={{color:T.accent,fontFamily:T.mono,fontWeight:600,flexShrink:0,marginLeft:4}}>{c.percentage}%</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          ))}
        </div>

        <div style={{marginTop:16,padding:"10px 12px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`,fontSize:12,color:T.textSec}}>
          {newCount} new blend{newCount===1?"":"s"}, {updateCount} update{updateCount===1?"":"s"} — <b style={{color:T.accent}}>{selectedCount} selected for import</b>
        </div>

        <div style={{display:"flex",gap:8,marginTop:12}}>
          <Btn variant="secondary" onClick={onCancel} style={{flex:1}} disabled={submitting}>Cancel</Btn>
          <Btn onClick={handleConfirm} disabled={submitting || selectedCount===0} style={{flex:1}}>
            {submitting ? "Importing..." : `Confirm Import (${selectedCount})`}
          </Btn>
        </div>
      </div>
    </div>
  );
}

function BlendCard({ blend, isAdmin, onEdit, onDelete }) {
  const total = blendTotalPercent(blend.components);
  return (
    <div style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.border}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <span style={{fontSize:15,fontWeight:600,color:T.text}}>{blend.name}</span>
            <Badge variant={blend.customer==="General"?"info":"default"}>{blend.customer || "General"}</Badge>
          </div>
          {blend.description && <p style={{fontSize:12,color:T.textMut,marginTop:4}}>{blend.description}</p>}
          <p style={{fontSize:12,color:T.textSec,marginTop:6}}>{blendComponentSummary(blend.components)}</p>
          <BlendPercentageBar components={blend.components}/>
        </div>
        {isAdmin && <div style={{display:"flex",gap:4}}>
          <button onClick={onEdit} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="edit" size={15} color={T.textSec}/></button>
          <button onClick={onDelete} style={{background:"none",border:"none",cursor:"pointer",padding:4}}><Icon name="trash" size={15} color={T.danger}/></button>
        </div>}
      </div>
      {Math.abs(total - 100) > 0.001 && <Badge variant="danger" style={{marginTop:6,fontSize:10}}>⚠ Total: {total}%</Badge>}
    </div>
  );
}

function BlendPercentageBar({ components }) {
  const palette = [T.accent, T.info, T.success, T.warning, T.danger];
  const safe = Array.isArray(components) ? components : [];
  return (
    <div style={{display:"flex",height:6,borderRadius:3,overflow:"hidden",marginTop:6,background:T.surfaceHover}}>
      {safe.map((c, i) => (
        <div key={i} title={`${c.percentage}% ${c.itemName||c.category||""}`}
          style={{width:`${parseFloat(c.percentage)||0}%`,background:palette[i % palette.length]}}/>
      ))}
    </div>
  );
}

function CreateEditBlendForm({ blend, customers, inventoryItems, inventoryCategories, onSave }) {
  const [name, setName] = useState(blend?.name || "");
  const [customer, setCustomer] = useState(blend?.customer || "General");
  const [description, setDescription] = useState(blend?.description || "");
  const [components, setComponents] = useState(() => {
    if (Array.isArray(blend?.components) && blend.components.length > 0) return blend.components.map(c => ({...c, percentage: String(c.percentage)}));
    return [{ category: "Roasted Beans", itemId: "", itemName: "", percentage: "" }];
  });

  const allowedCats = ["Roasted Beans", "Others"];
  const itemsForCategory = (cat) => (inventoryItems || []).filter(it => it.isActive && it.category === cat);

  const updateRow = (i, patch) => setComponents(prev => prev.map((c, idx) => idx === i ? { ...c, ...patch } : c));
  const removeRow = (i) => setComponents(prev => prev.filter((_, idx) => idx !== i));
  const addRow = () => setComponents(prev => [...prev, { category: "Roasted Beans", itemId: "", itemName: "", percentage: "" }]);

  const total = components.reduce((s, c) => s + (parseFloat(c.percentage) || 0), 0);
  const valid = name.trim() && Math.abs(total - 100) < 0.001 && components.every(c => (parseFloat(c.percentage) || 0) > 0);

  const handleSave = () => {
    const cleaned = components.map(c => ({
      category: c.category || "",
      itemId: c.itemId || "",
      itemName: c.itemName || "",
      percentage: parseFloat(c.percentage) || 0,
    }));
    onSave({ name: name.trim(), customer, description, components: cleaned });
  };

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:16}}>
      <Field label="Blend Name"><Input value={name} onChange={setName} placeholder="e.g., House Blend 70-30"/></Field>
      <Field label="Customer">
        <select value={customer} onChange={e=>setCustomer(e.target.value)}
          style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14}}>
          <option value="General">General (available to all customers)</option>
          {(customers || []).map(c => <option key={c.id} value={c.label}>{c.label}</option>)}
        </select>
      </Field>
      <Field label="Description (optional)">
        <textarea value={description} onChange={e=>setDescription(e.target.value)} rows={2}
          style={{width:"100%",padding:"10px 14px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:14,outline:"none",resize:"vertical",fontFamily:T.font}}/>
      </Field>

      <div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
          <span style={{fontSize:13,fontWeight:600,color:T.textSec}}>Components</span>
          <Btn variant="ghost" small onClick={addRow}><Icon name="plus" size={12} color={T.textSec}/> Add Component</Btn>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {components.map((c, i) => {
            const items = itemsForCategory(c.category);
            return (
              <div key={i} style={{background:T.card,borderRadius:T.radSm,padding:10,border:`1px solid ${T.border}`,display:"flex",flexDirection:"column",gap:8}}>
                <div style={{display:"flex",gap:6}}>
                  {allowedCats.map(cat => (
                    <Chip key={cat} label={cat} active={c.category===cat} onClick={()=>updateRow(i,{category:cat,itemId:"",itemName:""})}/>
                  ))}
                  <button onClick={()=>removeRow(i)} style={{background:"none",border:"none",cursor:"pointer",padding:4,marginLeft:"auto"}}><Icon name="x" size={14} color={T.danger}/></button>
                </div>
                <div style={{display:"flex",gap:8}}>
                  <select value={c.itemId} onChange={e=>{
                      const it = items.find(x => x.id === e.target.value);
                      updateRow(i, { itemId: e.target.value, itemName: it ? it.name : "" });
                    }}
                    style={{flex:2,padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13}}>
                    <option value="">— Select item —</option>
                    {items.map(it => <option key={it.id} value={it.id}>{it.name}{it.abbreviation?` (${it.abbreviation})`:""}</option>)}
                  </select>
                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                    <input type="number" min="0" max="100" value={c.percentage} onChange={e=>{ const v=e.target.value; const n=parseFloat(v); if(!isNaN(n)&&n<0){ updateRow(i,{percentage:"0"}); return; } updateRow(i,{percentage:v}); }} onBlur={e=>{ const n=parseFloat(e.target.value); if(!isNaN(n)&&n<0) updateRow(i,{percentage:"0"}); }} placeholder="%"
                      style={{width:70,padding:"8px 10px",borderRadius:T.radSm,background:T.bg,border:`1px solid ${T.border}`,color:T.text,fontSize:13,outline:"none"}}/>
                    <span style={{fontSize:13,color:T.textMut}}>%</span>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
        <BlendPercentageBar components={components}/>
        <div style={{marginTop:6,padding:"6px 10px",background:Math.abs(total-100)<0.001?T.successBg:T.warningBg,border:`1px solid ${Math.abs(total-100)<0.001?T.successBorder:T.warningBorder}`,borderRadius:T.radSm,fontSize:12,color:Math.abs(total-100)<0.001?T.success:T.warning,fontWeight:600}}>
          Total: {total}% {Math.abs(total-100)>=0.001 && "(must equal 100% to save)"}
        </div>
      </div>

      <Btn onClick={handleSave} disabled={!valid} style={{width:"100%",marginTop:8}}>{blend ? "Save Changes" : "Create Blend"}</Btn>
    </div>
  );
}

// Order-line blend selector — dropdown filtered by customer
function BlendSelector({ blends, customerLabel, value, onChange }) {
  const safe = blends || [];
  const filtered = safe.filter(b => b.isActive !== false && ((b.customer || "General") === "General" || b.customer === customerLabel));
  const selected = filtered.find(b => b.id === value);
  return (
    <div style={{display:"flex",flexDirection:"column",gap:6}}>
      <SearchableDropdown
        options={filtered.map(b=>({label:b.name+(b.customer&&b.customer!=="General"?` (${b.customer})`:""),value:b.id}))}
        value={value||""} onChange={v=>{const b=filtered.find(x=>x.id===v);onChange(b||null)}} placeholder="— Select blend —"/>
      {selected && (
        <div style={{padding:"8px 10px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${T.border}`}}>
          <p style={{fontSize:11,color:T.textMut}}>{blendComponentSummary(selected.components)}</p>
          <BlendPercentageBar components={selected.components}/>
        </div>
      )}
    </div>
  );
}

// Per-line blend breakdown shown in OrderDetailView
function OrderBlendBreakdown({ line, allocations, taggedByItemId }) {
  const components = line?.blendComponents || [];
  if (!Array.isArray(components) || components.length === 0) return null;
  const lineQty = parseFloat(line.quantity) || 0;
  const reqs = computeBlendRequirements({ components }, lineQty);
  return (
    <div style={{marginTop:8,display:"flex",flexDirection:"column",gap:6}}>
      {reqs.map((r, i) => {
        // Prefer the taggedByItemId map when provided (stage-tagging source of truth).
        let tagged = 0;
        if (taggedByItemId) {
          const k = blendItemKey(r.component.itemId, r.component.itemName);
          tagged = parseFloat(taggedByItemId[k]) || 0;
        } else {
          tagged = (allocations || []).filter(a => a.componentItemId === r.component.itemId).reduce((s, a) => s + (parseFloat(a.allocated_quantity) || 0), 0);
        }
        const remaining = Math.max(0, r.requiredQty - tagged);
        const complete = remaining <= 0.0001;
        return (
          <div key={i} style={{fontSize:11,padding:"6px 8px",background:T.bg,borderRadius:T.radSm,border:`1px solid ${complete?T.successBorder:T.border}`}}>
            <div style={{display:"flex",justifyContent:"space-between",gap:6}}>
              <span style={{color:T.textSec}}>{r.component.itemName || r.component.category}</span>
              <span style={{color:complete?T.success:T.textMut,fontFamily:T.mono}}>{complete?"✓ Complete":`Req: ${r.requiredQty} • Tagged: ${Math.round(tagged*100)/100} • Rem: ${Math.round(remaining*100)/100}`}</span>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// ─── Users View (Admin) ───────────────────────────────────────

function UsersView({ addToast }) {
  const [users,setUsers]=useState([]);
  const [loading,setLoading]=useState(true);
  const [showAdd,setShowAdd]=useState(false);
  const [editingId,setEditingId]=useState(null);
  const [resetId,setResetId]=useState(null);
  const [newUser,setNewUser]=useState({username:"",displayName:"",password:"",role:"user"});
  const [editData,setEditData]=useState({});
  const [newPassword,setNewPassword]=useState("");
  const [error,setError]=useState(null);
  const [msg,setMsg]=useState(null);

  const loadUsers = () => API.get("getUsers").then(setUsers).catch(e=>setError(e.message)).finally(()=>setLoading(false));
  useEffect(() => { loadUsers(); }, []);

  const handleCreate = async () => {
    setError(null);
    try {
      await API.post("createUser", newUser);
      setNewUser({username:"",displayName:"",password:"",role:"user"});
      setShowAdd(false); await loadUsers();
      addToast("User created", "success");
    } catch (e) { setError(e.message); }
  };

  const handleUpdate = async (id) => {
    setError(null);
    try {
      await API.post("updateUser", { id, ...editData });
      setEditingId(null); await loadUsers();
      addToast("User updated", "success");
    } catch (e) { setError(e.message); }
  };

  const handleResetPassword = async (id) => {
    setError(null);
    try {
      await API.post("resetPassword", { id, newPassword });
      setResetId(null); setNewPassword("");
      addToast("Password reset successfully", "success");
    } catch (e) { setError(e.message); }
  };

  if (loading) return <div style={{textAlign:"center",padding:40}}><p style={{color:T.textSec,animation:"pulse 1.5s infinite"}}>Loading users...</p></div>;

  return (
    <div className="fade-up" style={{display:"flex",flexDirection:"column",gap:16}}>
      {error && <div style={{background:T.dangerBg,borderRadius:T.radSm,padding:"10px 14px",border:"1px solid rgba(232,93,93,0.2)"}}>
        <span style={{fontSize:13,color:T.danger}}>{error}</span>
      </div>}
      {msg && <div style={{background:T.successBg,borderRadius:T.radSm,padding:"10px 14px",border:`1px solid ${T.successBorder}`}}>
        <span style={{fontSize:13,color:T.success}}>{msg}</span>
      </div>}

      <Btn onClick={()=>setShowAdd(!showAdd)} style={{width:"100%"}}><Icon name="plus" size={18} color={T.bg}/> Add User</Btn>

      {showAdd && (
        <div style={{background:T.card,borderRadius:T.rad,padding:16,border:`1px solid ${T.accentBorder}`,display:"flex",flexDirection:"column",gap:12}}>
          <Field label="Username"><Input value={newUser.username} onChange={v=>setNewUser(p=>({...p,username:v}))} placeholder="username (lowercase)"/></Field>
          <Field label="Display Name"><Input value={newUser.displayName} onChange={v=>setNewUser(p=>({...p,displayName:v}))} placeholder="Full Name"/></Field>
          <Field label="Password"><Input value={newUser.password} onChange={v=>setNewUser(p=>({...p,password:v}))} type="password" placeholder="Initial password"/></Field>
          <Field label="Role">
            <div style={{display:"flex",gap:8}}>
              <Chip label="User" active={newUser.role==="user"} onClick={()=>setNewUser(p=>({...p,role:"user"}))}/>
              <Chip label="Admin" active={newUser.role==="admin"} onClick={()=>setNewUser(p=>({...p,role:"admin"}))}/>
            </div>
          </Field>
          <div style={{display:"flex",gap:8}}>
            <Btn variant="secondary" onClick={()=>setShowAdd(false)} style={{flex:1}}>Cancel</Btn>
            <Btn onClick={handleCreate} disabled={!newUser.username.trim()||!newUser.displayName.trim()||!newUser.password} style={{flex:1}}>Create User</Btn>
          </div>
        </div>
      )}

      <Section icon="users" count={users.length}>Users</Section>
      <div style={{display:"flex",flexDirection:"column",gap:8}}>
        {users.map(u => (
          <div key={u.id} style={{background:T.card,borderRadius:T.rad,padding:"14px 16px",border:`1px solid ${T.border}`}}>
            {editingId === u.id ? (
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                <Field label="Display Name"><Input value={editData.displayName||""} onChange={v=>setEditData(p=>({...p,displayName:v}))}/></Field>
                <Field label="Role">
                  <div style={{display:"flex",gap:8}}>
                    <Chip label="User" active={editData.role==="user"} onClick={()=>setEditData(p=>({...p,role:"user"}))}/>
                    <Chip label="Admin" active={editData.role==="admin"} onClick={()=>setEditData(p=>({...p,role:"admin"}))}/>
                  </div>
                </Field>
                <Field label="Status">
                  <div style={{display:"flex",gap:8}}>
                    <Chip label="Active" active={editData.status==="active"} onClick={()=>setEditData(p=>({...p,status:"active"}))}/>
                    <Chip label="Inactive" active={editData.status==="inactive"} onClick={()=>setEditData(p=>({...p,status:"inactive"}))}/>
                  </div>
                </Field>
                <div style={{display:"flex",gap:8}}>
                  <Btn variant="secondary" small onClick={()=>setEditingId(null)} style={{flex:1}}>Cancel</Btn>
                  <Btn small onClick={()=>handleUpdate(u.id)} style={{flex:1}}>Save</Btn>
                </div>
              </div>
            ) : resetId === u.id ? (
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                <span style={{fontSize:14,fontWeight:500,color:T.text}}>Reset password for {u.displayName}</span>
                <Field label="New Password"><Input value={newPassword} onChange={setNewPassword} type="password" placeholder="New password"/></Field>
                <div style={{display:"flex",gap:8}}>
                  <Btn variant="secondary" small onClick={()=>{setResetId(null);setNewPassword("")}} style={{flex:1}}>Cancel</Btn>
                  <Btn small onClick={()=>handleResetPassword(u.id)} disabled={!newPassword} style={{flex:1}}>Reset</Btn>
                </div>
              </div>
            ) : (
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <div>
                  <div style={{display:"flex",alignItems:"center",gap:8}}>
                    <span style={{fontSize:14,fontWeight:600,color:T.text}}>{u.displayName}</span>
                    <Badge variant={u.role==="admin"?"default":"muted"}>{u.role}</Badge>
                    {u.status==="inactive"&&<Badge variant="danger">Inactive</Badge>}
                  </div>
                  <span style={{fontSize:12,color:T.textMut}}>@{u.username}</span>
                </div>
                <div style={{display:"flex",gap:4}}>
                  <button onClick={()=>{setEditingId(u.id);setEditData({displayName:u.displayName,role:u.role,status:u.status||"active"})}} style={{background:"none",border:"none",cursor:"pointer",padding:6}}><Icon name="edit" size={15} color={T.textSec}/></button>
                  <button onClick={()=>setResetId(u.id)} style={{background:"none",border:"none",cursor:"pointer",padding:6}} title="Reset password"><Icon name="key" size={15} color={T.textSec}/></button>
                </div>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
