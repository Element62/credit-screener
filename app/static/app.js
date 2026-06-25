const state = {
  issuers: [],
  filteredIssuers: [],
  filters: {},
  metadata: {},
  selectedIssuer: null,
  instrumentMap: new Map(),
  detailRows: [],
  rawMoversRows: [],
  moversRows: [],
  moversSectorSummary: [],
  abnormalPriceRows: [],
  excludedRows: [],
  loanRows: [],
  filteredLoanRows: [],
  loanSortField: "AMT_OUTSTANDING_MM",
  loanSortDirection: "desc",
  loanScreenMode: null,
  loanActiveFilters: { yieldMin: true, yieldMax: true, priceMax: true },
  loanFilterValues: { yieldMin: 9, yieldMax: 50, priceMax: 95 },
  upsideMode: "52w",
  moveMode: "3m",
  mvAbsSort: true,
  faceStrikeZonePct: false,
  sortField: "52W PEAK UPSIDE SECURED TARGET ($MM)",
  sortDirection: "desc",
  coverageMap: {},
};

const loginCard = document.getElementById("loginCard");
const app = document.getElementById("app");
const loginForm = document.getElementById("loginForm");
const loginError = document.getElementById("loginError");
const logoutButton = document.getElementById("logoutButton");
const loadingBanner = document.getElementById("loadingBanner");
const issuerTabButton = document.getElementById("issuerTabButton");
const loansTabButton = document.getElementById("loansTabButton");
const moversTabButton = document.getElementById("moversTabButton");
const exceptionsTabButton = document.getElementById("exceptionsTabButton");
const issuerTabPanel = document.getElementById("issuerTabPanel");
const loansTabPanel = document.getElementById("loansTabPanel");
const moversTabPanel = document.getElementById("moversTabPanel");
const exceptionsTabPanel = document.getElementById("exceptionsTabPanel");
const documentationTabButton = document.getElementById("documentationTabButton");
const documentationTabPanel = document.getElementById("documentationTabPanel");
const issuerHead = document.getElementById("issuerHead");
const issuerBody = document.getElementById("issuerBody");
const moversHead = document.getElementById("moversHead");
const moversBody = document.getElementById("moversBody");
const moversStatus = document.getElementById("moversStatus");
const moversChartTitle = document.getElementById("moversChartTitle");
const moversChart = document.getElementById("moversChart");
const moversDirectionFilter = document.getElementById("moversDirectionFilter");
const moversChartMetric = document.getElementById("moversChartMetric");
const moversSortField = document.getElementById("moversSortField");
const moversSortDirection = document.getElementById("moversSortDirection");
const abnormalPriceSubtabButton = document.getElementById("abnormalPriceSubtabButton");
const abnormalBbgSubtabButton = document.getElementById("abnormalBbgSubtabButton");
const exclusionsSubtabButton = document.getElementById("exclusionsSubtabButton");
const abnormalPricePanel = document.getElementById("abnormalPricePanel");
const abnormalBbgPanel = document.getElementById("abnormalBbgPanel");
const exclusionsPanel = document.getElementById("exclusionsPanel");
const abnormalPriceHead = document.getElementById("abnormalPriceHead");
const abnormalPriceBody = document.getElementById("abnormalPriceBody");
const abnormalPriceStatus = document.getElementById("abnormalPriceStatus");
const exclusionsHead = document.getElementById("exclusionsHead");
const exclusionsBody = document.getElementById("exclusionsBody");
const exclusionsStatus = document.getElementById("exclusionsStatus");
const loansHead = document.getElementById("loansHead");
const loansBody = document.getElementById("loansBody");
const loansStatus = document.getElementById("loansStatus");
const loansSearchInput = document.getElementById("loansSearchInput");
const clearLoansSearchButton = document.getElementById("clearLoansSearchButton");
const loanTopNInput = document.getElementById("loanTopNInput");
const loanScreenButton = document.getElementById("loanScreenButton");
const loanYieldScreenButton = document.getElementById("loanYieldScreenButton");
const downloadLoansButton = document.getElementById("downloadLoansButton");
const loanFilterChips = document.getElementById("loanFilterChips");
const detailCard = document.getElementById("detailCard");
const detailTicker = document.getElementById("detailTicker");
const detailTitle = document.getElementById("detailTitle");
const detailHead = document.getElementById("detailHead");
const detailBody = document.getElementById("detailBody");
const nonQualifyingCard = document.getElementById("nonQualifyingCard");
const nonQualifyingHead = document.getElementById("nonQualifyingHead");
const nonQualifyingBody = document.getElementById("nonQualifyingBody");
const defaultedCard = document.getElementById("defaultedCard");
const defaultedHead = document.getElementById("defaultedHead");
const defaultedBody = document.getElementById("defaultedBody");
const statusBar = document.getElementById("statusBar");
const searchInput = document.getElementById("searchInput");
const clearSearchButton = document.getElementById("clearSearchButton");
const sectorFilter = document.getElementById("sectorFilter");
const reloadDataButton = document.getElementById("reloadDataButton");
const downloadIssuersButton = document.getElementById("downloadIssuersButton");
const downloadDetailButton = document.getElementById("downloadDetailButton");
const hideDetailButton = document.getElementById("hideDetailButton");
const detailToggleBar = document.getElementById("detailToggleBar");

function setLoading(isLoading, message = "Loading...") {
  if (!loadingBanner) return;
  loadingBanner.classList.toggle("hidden", !isLoading);
  loadingBanner.textContent = message;
}

const issuerColumns = [
  "Issuer",
  "Sector",
  "Secured Face ($BN)",
  "Unsecured Face ($BN)",
  "Preferred Face ($BN)",
  "Price",
  "Yield",
  "3M MV Change ($MM)",
  "7D MV Change ($MM)",
  "52W PEAK UPSIDE SECURED ($MM)",
  "52W PEAK UPSIDE UNSECURED ($MM)",
  "52W PEAK UPSIDE PREFERRED ($MM)",
  "RETURN TO PAR SECURED ($MM)",
  "RETURN TO PAR UNSECURED ($MM)",
  "RETURN TO PAR PREFERRED ($MM)",
  "COVERAGE PRIMARY",
  "COVERAGE SECONDARY",
];

const detailColumns = [
  { key: "NAME", label: "Name" },
  { key: "PAYMENT_RANK", label: "Lien" },
  { key: "AMT_OUTSTANDING_MM", label: "Amt Out ($MM)" },
  { key: "PX_MID", label: "Current PX" },
  { key: "PRICE_MOVE_7D", label: "7D Price Move" },
  { key: "PRICE_MOVE_3M", label: "3M Price Move" },
  { key: "YIELD", label: "Yield" },
  { key: "LQA_LIQUIDITY_SCORE", label: "Liquidity Score" },
  { key: "LQA_EXPECTED_DAILY_VOLUME_MM", label: "Exp. Daily Volume ($MM)" },
  { key: "PRICE_RANGE", label: "Price Range" },
];

const moversColumns = [
  { key: "Issuer Name", label: "Issuer Name" },
  { key: "Security Name", label: "Security Name" },
  { key: "Sector", label: "Sector" },
  { key: "Lien", label: "Lien" },
  { key: "Amount Out ($MM)", label: "Amount Out ($MM)" },
  { key: "Current Px", label: "Current Px" },
  { key: "ACTIVE_MOVE", label: "Price Move" },
  { key: "Price Range", label: "Price Range" },
];

const abnormalPriceColumns = [
  { key: "Issuer", label: "Issuer" },
  { key: "Security", label: "Security" },
  { key: "Sector", label: "Sector" },
  { key: "Rank", label: "Rank" },
  { key: "Mat", label: "Mat" },
  { key: "Amount Out ($MM)", label: "Amount Out ($MM)" },
  { key: "Current Px", label: "Current Px" },
  { key: "Prior Px", label: "Prior Px" },
  { key: "3M Price Move", label: "3M Price Move" },
  { key: "Yield", label: "Yield" },
  { key: "52W Low", label: "52W Low" },
  { key: "52W High", label: "52W High" },
];

function fmt(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") {
    return value.toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits });
  }
  return value;
}

const COVERAGE_ANALYSTS = [
  "David", "Bill", "Bronte", "Jake", "Basso", "Evan", "Brie", "Manish",
  "Allen", "Sherman", "Nicoll", "Zhe", "Dante", "Olivia", "Bergstein",
  "Casey", "Bobby", "Patrick", "Brooks", "Nick", "Sam", "Reed",
  "Nick M", "Kendrick", "Ryan", "Mark", "Samson", "Jared",
];

const zeroAsDashColumns = new Set([
  "Secured Face ($MM)", "Unsecured Face ($MM)", "Preferred Face ($MM)", "Face Strike Zone ($MM)",
  "Total Secured Face ($MM)", "Total Unsecured Face ($MM)", "Total Preferred Face ($MM)", "Face Total All ($MM)",
  "52W PEAK UPSIDE SECURED TARGET ($MM)", "52W PEAK UPSIDE UNSECURED TARGET ($MM)", "52W PEAK UPSIDE PREFERRED TARGET ($MM)", "52W PEAK UPSIDE TOTAL TARGET ($MM)",
  "RETURN TO PAR SECURED TARGET ($MM)", "RETURN TO PAR UNSECURED TARGET ($MM)", "RETURN TO PAR PREFERRED TARGET ($MM)", "RETURN TO PAR TOTAL TARGET ($MM)",
]);

function fmtIssuer(column, value) {
  if (zeroAsDashColumns.has(column) && (value === 0 || value === null || value === undefined)) return "-";
  return fmt(value, issuerDigits(column));
}

function issuerDigits(column) {
  if (["Price", "Yield", "Price All", "Yield All"].includes(column)) return 2;
  if (column.endsWith("($BN)")) return 2;
  if (column.endsWith("($MM)")) return 0;
  return 2;
}

function renderCoverageCell(ticker, slot, names) {
  if (!ticker) return "";
  const t = String(ticker).replace(/&/g, "&amp;").replace(/"/g, "&quot;");
  const spans = names.map((name, idx) =>
    `<span class="cov-name" data-ticker="${t}" data-slot="${slot}" data-idx="${idx}" title="Click to remove">${name}</span>`
  ).join('<span class="cov-sep"> / </span>');
  const addBtn = names.length < 2
    ? `<button class="cov-add" data-ticker="${t}" data-slot="${slot}" title="Add analyst">+</button>`
    : "";
  return spans + addBtn;
}

function updateCoverageCells(ticker) {
  const btn = issuerBody.querySelector(`[data-issuer="${CSS.escape(ticker)}"]`);
  if (!btn) return;
  const tds = btn.closest("tr").querySelectorAll("td");
  const entry = state.coverageMap[ticker] || { primary: [], secondary: [] };
  if (tds[11]) tds[11].innerHTML = renderCoverageCell(ticker, "primary", entry.primary || []);
  if (tds[12]) tds[12].innerHTML = renderCoverageCell(ticker, "secondary", entry.secondary || []);
}

function detailDigits(column) {
  if (column === "AMT_OUTSTANDING_MM") return 0;
  if (["PX_MID", "PRICE_MOVE_3M", "PRICE_MOVE_7D", "YIELD", "LQA_EXPECTED_DAILY_VOLUME_MM", "PX_HIGH_52W", "PX_LOW_52W"].includes(column)) return 1;
  if (["OAS", "OAS_DELTA"].includes(column)) return 0;
  return 2;
}

function moversDigits(column) {
  if (column === "Amount Out ($MM)") return 0;
  if (["Current Px", "3M Price Move", "7D Price Move", "ACTIVE_MOVE"].includes(column)) return 1;
  return 2;
}

function abnormalPriceDigits(column) {
  if (column === "Amount Out ($MM)") return 0;
  if (["Current Px", "Prior Px", "3M Price Move", "Yield", "52W Low", "52W High"].includes(column)) return 1;
  return 2;
}

function pillClass(value) {
  if (!value) return "pill";
  if (["Deep Distressed", "Distressed", "Severe", "Recovery Cliff"].includes(value)) return "pill bad";
  if (["Stressed", "Elevated"].includes(value)) return "pill warn";
  if (["Premium", "Normal", "Par-Adjacent"].includes(value)) return "pill good";
  return "pill";
}

function seniorityClass(value) {
  const numeric = Number(value);
  if (Number.isNaN(numeric)) return "";
  if (numeric < 10) return "basis-low";
  if (numeric < 20) return "basis-mid";
  return "basis-high";
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, { credentials: "same-origin", ...options });
  if (!response.ok) {
    let message = "Request failed";
    try {
      const payload = await response.json();
      message = payload.detail || message;
    } catch {}
    throw new Error(message);
  }
  return response.json();
}

async function downloadExcel(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    credentials: "same-origin",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    let message = "Download failed";
    try {
      const data = await response.json();
      message = data.detail || message;
    } catch {}
    throw new Error(message);
  }
  const blob = await response.blob();
  const contentDisposition = response.headers.get("Content-Disposition") || "";
  const match = contentDisposition.match(/filename=\"?([^\";]+)\"?/i);
  const filename = match ? match[1] : "export.xlsx";
  const objectUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = objectUrl;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(objectUrl);
}

function sortIndicator(column) {
  if (state.sortField !== column) return "";
  return state.sortDirection === "asc" ? " &uarr;" : " &darr;";
}

function fmtPriceMove(moveVal) {
  if (moveVal === null || moveVal === undefined || moveVal === "") return "-";
  const move = Number(moveVal);
  if (Number.isNaN(move)) return "-";
  const cssClass = move < 0 ? "negative" : move > 0 ? "positive" : "neutral";
  const prefix = move > 0 ? "+" : "";
  return `<span class="price-move ${cssClass}">${prefix}${fmt(move, 1)}</span>`;
}

function renderPriceMove(row) {
  return fmtPriceMove(row.PRICE_MOVE_3M);
}

function renderPriceMove7D(row) {
  return fmtPriceMove(row.PRICE_MOVE_7D);
}

function renderPriceRange(row) {
  const low = Number(row.PX_LOW_52W);
  const high = Number(row.PX_HIGH_52W);
  const current = Number(row.PX_MID);
  if (Number.isNaN(low) || Number.isNaN(high) || Number.isNaN(current) || high <= low) return "-";
  const clamped = Math.min(Math.max(current, low), high);
  const pct = ((clamped - low) / (high - low)) * 100;
  return `
    <div class="price-range-cell">
      <span class="range-value low">${fmt(low, 1)}</span>
      <div class="range-track"><span class="range-dot" style="left:${pct}%"></span></div>
      <span class="range-value high">${fmt(high, 1)}</span>
    </div>
  `;
}

function renderMoverPriceRange(row) {
  const low = Number(row.PX_LOW_52W);
  const high = Number(row.PX_HIGH_52W);
  const current = Number(row["Current Px"]);
  if (Number.isNaN(low) || Number.isNaN(high) || Number.isNaN(current) || high <= low) return "-";
  const clamped = Math.min(Math.max(current, low), high);
  const pct = ((clamped - low) / (high - low)) * 100;
  return `
    <div class="price-range-cell">
      <span class="range-value low">${fmt(low, 1)}</span>
      <div class="range-track"><span class="range-dot" style="left:${pct}%"></span></div>
      <span class="range-value high">${fmt(high, 1)}</span>
    </div>
  `;
}

const glossaryEntries = [
  { sup: "1", label: "Face ($MM)", def: "<strong>Strike Zone:</strong> face value of instruments meeting screening criteria (see below). <strong>Total:</strong> face value of all non-defaulted instruments regardless of screening criteria." },
  { sup: "2", label: "52W Peak Upside / Return To Par Upside", def: "Dollar upside for Strike Zone securities, assuming prices revert to 52-week highs (52W) or par (Return to Par)." },
  { sup: "3", label: "3M / 7D MV Change", def: "Aggregate market value change ($MM) for Strike Zone securities over the selected time period." },
  { sup: "4", label: "Price", def: "Face-weighted average mid price across Strike Zone securities." },
  { sup: "5", label: "Yield", def: "Face-weighted average yield to maturity across Strike Zone securities." },
  { sup: "6", label: "Exp. Daily Volume ($MM)", def: "Bloomberg LQA Expected Daily Volume: estimated notional a single firm could sell in one day within the current bid-ask spread. Modeled using reference data (amount outstanding, maturity, coupon), market data (trade count, dealer quotes), and pricing data (spread, volatility)." },
  { sup: "7", label: "Liquidity Score", def: "Bloomberg LQA score: a 1–100 percentile ranking of estimated liquidation cost within an asset class. 100 = most liquid (lowest cost); 1 = least liquid. Scores are relative to peers — there is no absolute meaning to any specific value." },
];

function renderIssuerTable() {
  // Face — Strike Zone (criteria-eligible)
  const faceStrikeSecKey   = "Secured Face ($MM)";
  const faceStrikeUnsecKey = "Unsecured Face ($MM)";
  const faceStrikePrefKey  = "Preferred Face ($MM)";
  const faceStrikeTotalKey = "Face Strike Zone ($MM)";
  // Face — Total (all non-defaulted)
  const faceTotalSecKey    = "Total Secured Face ($MM)";
  const faceTotalUnsecKey  = "Total Unsecured Face ($MM)";
  const faceTotalPrefKey   = "Total Preferred Face ($MM)";
  const faceTotalTotalKey  = "Face Total All ($MM)";

  // Upside — always Strike Zone
  const is52w = state.upsideMode === "52w";
  const upsideLabel = is52w ? "52W Peak Upside" : "Return to Par Upside";
  const upsideSup = "<sup>2</sup>";
  const upsideSecKey   = is52w ? "52W PEAK UPSIDE SECURED TARGET ($MM)"  : "RETURN TO PAR SECURED TARGET ($MM)";
  const upsideUnsecKey = is52w ? "52W PEAK UPSIDE UNSECURED TARGET ($MM)" : "RETURN TO PAR UNSECURED TARGET ($MM)";
  const upsidePrefKey  = is52w ? "52W PEAK UPSIDE PREFERRED TARGET ($MM)" : "RETURN TO PAR PREFERRED TARGET ($MM)";
  const upsideTotalKey = is52w ? "52W PEAK UPSIDE TOTAL TARGET ($MM)"     : "RETURN TO PAR TOTAL TARGET ($MM)";

  // MV Change — always Strike Zone
  const is3m = state.moveMode === "3m";
  const mvChangeKey   = is3m ? "3M MV Change TARGET ($MM)" : "7D MV Change TARGET ($MM)";
  const mvPctKey      = is3m ? "3M MV Pct Change"          : "7D MV Pct Change";
  const mvChangeLabel = is3m ? "3M MV △" : "7D MV △";

  // Price & Yield — always Strike Zone
  const priceKey = "Price";
  const yieldKey = "Yield";

  const tableColumnOrder = [
    "Issuer", "Sector",
    faceStrikeSecKey, faceStrikeUnsecKey, faceStrikePrefKey, faceStrikeTotalKey,
    faceTotalSecKey,  faceTotalUnsecKey,  faceTotalPrefKey,  faceTotalTotalKey,
    upsideSecKey, upsideUnsecKey, upsidePrefKey, upsideTotalKey,
    mvChangeKey, mvPctKey,
    priceKey, yieldKey,
    "COVERAGE PRIMARY", "COVERAGE SECONDARY",
  ];

  issuerHead.innerHTML = `
    <tr class="header-group-row">
      <th class="col-group-r"></th>
      <th></th>
      <th colspan="4" class="group-header col-group-l">FACE STRIKE ZONE <em>($MM)</em><sup>1</sup> <button type="button" id="faceStrikePctToggle" class="mv-abs-toggle${state.faceStrikeZonePct ? " active" : ""}">%</button></th>
      <th colspan="4" class="group-header col-group-l">FACE TOTAL <em>($MM)</em></th>
      <th colspan="4" class="group-header upside-toggle col-group-l" id="upsideToggle" title="Click to toggle">${upsideLabel} <em>(Strike Zone, $MM)</em>${upsideSup} &#x21c4;</th>
      <th colspan="2" class="group-header upside-toggle col-group-l" id="moveToggle" title="Click to toggle" style="min-width:155px">${mvChangeLabel}<br><em>(Strike Zone)</em><sup>3</sup> &#x21c4;</th>
      <th colspan="2" class="group-header col-group-l">Price &amp; Yield <em>(Strike Zone)</em></th>
      <th colspan="2" class="group-header col-group-l">Coverage (WIP)</th>
    </tr>
    <tr>
      <th class="col-fit col-group-r">Issuer</th>
      <th class="col-fit">Sector</th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header" data-sort="${faceStrikeSecKey}">Secured${sortIndicator(faceStrikeSecKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${faceStrikeUnsecKey}">Unsecured${sortIndicator(faceStrikeUnsecKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${faceStrikePrefKey}">Preferred${sortIndicator(faceStrikePrefKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${faceStrikeTotalKey}">Total${sortIndicator(faceStrikeTotalKey)}</button></th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header" data-sort="${faceTotalSecKey}">Secured${sortIndicator(faceTotalSecKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${faceTotalUnsecKey}">Unsecured${sortIndicator(faceTotalUnsecKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${faceTotalPrefKey}">Preferred${sortIndicator(faceTotalPrefKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${faceTotalTotalKey}">Total${sortIndicator(faceTotalTotalKey)}</button></th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header" data-sort="${upsideSecKey}">Secured${sortIndicator(upsideSecKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${upsideUnsecKey}">Unsecured${sortIndicator(upsideUnsecKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${upsidePrefKey}">Preferred${sortIndicator(upsidePrefKey)}</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${upsideTotalKey}">Total${sortIndicator(upsideTotalKey)}</button></th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header" data-sort="${mvChangeKey}">$MM${sortIndicator(mvChangeKey) || "&#x21c5;"}</button><button type="button" id="mvAbsToggle" class="mv-abs-toggle${state.mvAbsSort ? " active" : ""}">ABS</button></th>
      <th class="col-tight"><button type="button" class="sort-header" data-sort="${mvPctKey}">%${sortIndicator(mvPctKey)}</button></th>
      <th class="col-group-l"><button type="button" class="sort-header" data-sort="${priceKey}">Price<sup>4</sup>${sortIndicator(priceKey)}</button></th>
      <th><button type="button" class="sort-header" data-sort="${yieldKey}">Yield<sup>5</sup>${sortIndicator(yieldKey)}</button></th>
      <th class="col-group-l">Primary</th>
      <th>Secondary</th>
    </tr>
  `;

  document.getElementById("upsideToggle").addEventListener("click", () => {
    state.upsideMode = state.upsideMode === "52w" ? "par" : "52w";
    renderIssuerTable();
  });

  document.getElementById("moveToggle").addEventListener("click", () => {
    state.moveMode = state.moveMode === "3m" ? "7d" : "3m";
    renderIssuerTable();
  });

  document.getElementById("faceStrikePctToggle").addEventListener("click", () => {
    state.faceStrikeZonePct = !state.faceStrikeZonePct;
    renderIssuerTable();
  });

  document.getElementById("mvAbsToggle").addEventListener("click", (e) => {
    e.stopPropagation();
    state.mvAbsSort = !state.mvAbsSort;
    const mvCols = new Set(["3M MV Change TARGET ($MM)", "7D MV Change TARGET ($MM)", "3M MV Pct Change", "7D MV Pct Change"]);
    if (mvCols.has(state.sortField)) state.sortDirection = "desc";
    applyFilters();
    renderIssuerTable();
  });

  const bodyBorderClass = {
    [faceStrikeSecKey]: "col-group-l",
    [faceTotalSecKey]:  "col-group-l",
    [upsideSecKey]:     "col-group-l",
    [priceKey]:         "col-group-l",
    "COVERAGE PRIMARY": "col-group-l",
  };

  const faceStrikePctMap = {
    [faceStrikeSecKey]:   faceTotalSecKey,
    [faceStrikeUnsecKey]: faceTotalUnsecKey,
    [faceStrikePrefKey]:  faceTotalPrefKey,
    [faceStrikeTotalKey]: faceTotalTotalKey,
  };

  issuerBody.innerHTML = state.filteredIssuers.map((row) => {
    const selected = state.selectedIssuer === row.PARENT_TICKER ? "selected" : "";
    const cells = tableColumnOrder.map((column) => {
      const value = row[column];
      if (column === "Issuer") {
        const marker = row.REPORT_SENTIMENT_COLOR
          ? `<sup class="sentiment-marker ${row.REPORT_SENTIMENT_COLOR}" title="${row.REPORT_SENTIMENT_LABEL || "Report sentiment"}"></sup>`
          : "";
        const hMarker = row._HAS_HOLDING ? `<sup class="holding-marker" title="Portfolio holding">H</sup>` : "";
        return `<td><button class="table-button" data-issuer="${row.PARENT_TICKER}">${fmt(value, 2)}${marker}${hMarker}</button></td>`;
      }
      if (["3M MV Change ($MM)", "7D MV Change ($MM)", "3M MV Change TARGET ($MM)", "7D MV Change TARGET ($MM)"].includes(column)) {
        if (value === null || value === undefined || value === 0) return `<td class="col-tight col-group-l">-</td>`;
        const v = Number(value);
        const cls = v > 0 ? "positive" : v < 0 ? "negative" : "neutral";
        const arrow = v > 0 ? "▲" : "▼";
        return `<td class="col-tight col-group-l price-move ${cls}">${arrow} ${fmt(Math.abs(v), 0)}</td>`;
      }
      if (column === "3M MV Pct Change" || column === "7D MV Pct Change") {
        if (value === null || value === undefined || value === 0) return `<td class="col-tight">-</td>`;
        const v = Number(value);
        const cls = v > 0 ? "positive" : v < 0 ? "negative" : "neutral";
        const arrow = v > 0 ? "▲" : "▼";
        return `<td class="col-tight price-move ${cls}">${arrow} ${fmt(Math.abs(v), 1)}%</td>`;
      }
      if (column === "COVERAGE PRIMARY" || column === "COVERAGE SECONDARY") {
        const slot = column === "COVERAGE PRIMARY" ? "primary" : "secondary";
        const ticker = row.PARENT_TICKER;
        const names = (state.coverageMap[ticker] || {})[slot] || [];
        const covClass = column === "COVERAGE PRIMARY" ? "cov-cell col-group-l" : "cov-cell";
        return `<td class="${covClass}">${renderCoverageCell(ticker, slot, names)}</td>`;
      }
      if (state.faceStrikeZonePct && column in faceStrikePctMap) {
        const bc = bodyBorderClass[column];
        const cls = bc ? ` class="${bc}"` : "";
        const v = Number(value);
        if (value == null || isNaN(v) || v === 0) return `<td${cls}>-</td>`;
        const denom = Number(row[faceStrikePctMap[column]]);
        const pctStr = denom > 0 ? ` <em class="face-pct">(${Math.round(v / denom * 100)}%)</em>` : "";
        return `<td${cls}>${fmtIssuer(column, value)}${pctStr}</td>`;
      }
      const bc = bodyBorderClass[column];
      return `<td${bc ? ` class="${bc}"` : ""}>${fmtIssuer(column, value)}</td>`;
    }).join("");
    const classes = [selected, row._HAS_HOLDING ? "issuer-has-holdings" : ""].filter(Boolean).join(" ");
    return `<tr class="${classes}">${cells}</tr>`;
  }).join("");

  document.querySelectorAll("[data-issuer]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await openIssuerDetail(button.dataset.issuer);
      } catch (error) {
        statusBar.textContent = error.message;
      }
    });
  });

  document.querySelectorAll("[data-sort]").forEach((button) => {
    button.addEventListener("click", () => {
      const column = button.dataset.sort;
      const isMvCol = ["3M MV Change TARGET ($MM)", "7D MV Change TARGET ($MM)", "3M MV Pct Change", "7D MV Pct Change"].includes(column);
      if (state.sortField === column && !(isMvCol && state.mvAbsSort)) {
        state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
      } else {
        state.sortField = column;
        state.sortDirection = "desc";
      }
      applyFilters();
    });
  });
}

function renderMoversChart() {
  const metric = moversChartMetric.value;
  const metricLabel = metric === "count" ? "Security Count by Sector" : "Amount Outstanding ($MM) by Sector";
  moversChartTitle.textContent = metricLabel;
  if (!state.moversSectorSummary.length) {
    moversChart.innerHTML = `<div class="status-bar">No securities matching the current filter criteria.</div>`;
    return;
  }
  const metricKey = metric === "count" ? "Security Count" : "Amount Out ($MM)";
  const maxValue = Math.max(...state.moversSectorSummary.map((row) => Number(row[metricKey]) || 0), 0);
  moversChart.innerHTML = state.moversSectorSummary.map((row) => {
    const value = Number(row[metricKey]) || 0;
    const widthPct = maxValue > 0 ? (value / maxValue) * 100 : 0;
    return `
      <div class="sector-bar-row">
        <div class="sector-bar-label">${row.Sector || "Unknown"}</div>
        <div class="sector-bar-track"><div class="sector-bar-fill" style="width:${widthPct}%"></div></div>
        <div class="sector-bar-value">${fmt(value, 0)}</div>
      </div>
    `;
  }).join("");
}

function renderMoversTable() {
  const activeMoveKey = moversSortField.value === "7D Price Move" ? "7D Price Move" : "3M Price Move";
  const activeMoveLabel = activeMoveKey === "7D Price Move" ? "7D Price Move" : "3M Price Move";
  moversHead.innerHTML = `<tr>${moversColumns.map((column) => {
    if (column.key === "ACTIVE_MOVE") {
      const marker = moversSortDirection.value === "asc" ? " &uarr;" : " &darr;";
      return `<th><button type="button" class="sort-header mover-sort-header" data-mover-sort="${activeMoveKey}">${activeMoveLabel}${marker}</button></th>`;
    }
    if (column.key === "Current Px") {
      const marker = moversSortField.value === "Current Px" ? (moversSortDirection.value === "asc" ? " &uarr;" : " &darr;") : "";
      return `<th><button type="button" class="sort-header mover-sort-header" data-mover-sort="Current Px">Current Px${marker}</button></th>`;
    }
    return `<th>${column.label}</th>`;
  }).join("")}</tr>`;
  moversBody.innerHTML = state.moversRows.map((row) => {
    const cells = moversColumns.map((column) => {
      if (column.key === "ACTIVE_MOVE") {
        return `<td>${fmtPriceMove(row[activeMoveKey])}</td>`;
      }
      if (column.key === "Price Range") {
        return `<td>${renderMoverPriceRange(row)}</td>`;
      }
      return `<td>${fmt(row[column.key], moversDigits(column.key))}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  document.querySelectorAll("[data-mover-sort]").forEach((button) => {
    button.addEventListener("click", () => {
      const column = button.dataset.moverSort;
      if (moversSortField.value === column) {
        moversSortDirection.value = moversSortDirection.value === "asc" ? "desc" : "asc";
      } else {
        moversSortField.value = column;
        moversSortDirection.value = "desc";
      }
      applyMoversFilters();
    });
  });
}

function renderAbnormalPriceTable() {
  abnormalPriceHead.innerHTML = `<tr>${abnormalPriceColumns.map((column) => `<th>${column.label}</th>`).join("")}</tr>`;
  abnormalPriceBody.innerHTML = state.abnormalPriceRows.map((row) => {
    const cells = abnormalPriceColumns.map((column) => {
      if (column.key === "3M Price Move") {
        const move = Number(row[column.key]);
        if (Number.isNaN(move)) return "<td>-</td>";
        const cssClass = move < 0 ? "negative" : move > 0 ? "positive" : "neutral";
        const prefix = move > 0 ? "+" : "";
        return `<td><span class="price-move ${cssClass}">${prefix}${fmt(move, 1)}</span></td>`;
      }
      return `<td>${fmt(row[column.key], abnormalPriceDigits(column.key))}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
  abnormalPriceStatus.textContent = `${state.abnormalPriceRows.length} securities with current price above 120 excluded from issuer weighted averages`;
}

const exclusionColumns = [
  { key: "Issuer", label: "Issuer" },
  { key: "NAME", label: "Security" },
  { key: "SECTOR", label: "Sector" },
  { key: "PAYMENT_RANK", label: "Lien" },
  { key: "MATURITY", label: "Mat" },
  { key: "AMT_OUTSTANDING_MM", label: "Amt Out ($MM)" },
  { key: "PX_MID", label: "Price" },
  { key: "YIELD", label: "Yield" },
  { key: "Exclusion Reason", label: "Exclusion Reason" },
];

function renderExclusionsTable() {
  exclusionsHead.innerHTML = `<tr>${exclusionColumns.map((c) => `<th>${c.label}</th>`).join("")}</tr>`;
  exclusionsBody.innerHTML = state.excludedRows.map((row) => {
    const cells = exclusionColumns.map((c) => `<td>${fmt(row[c.key], 2)}</td>`).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
  exclusionsStatus.textContent = `${state.excludedRows.length} securities excluded from screening`;
}

const loanColumns = [
  { key: "Issuer",             label: "Issuer" },
  { key: "NAME",               label: "Security" },
  { key: "SECTOR",             label: "Sector",       tdClass: "col-tight" },
  { key: "AMT_OUTSTANDING_MM", label: "Amt Out",     sortable: true, tdClass: "col-tight col-group-l" },
  { key: "PX_MID",             label: "Current Px",  tdClass: "col-tight col-group-l" },
  { key: "YIELD",              label: "Yield",        tdClass: "col-tight" },
  { key: "COUPON_DISPLAY",      label: "Coupon",       tdClass: "col-tight" },
  { key: "PRICE_MOVE_3M",      label: "3M Px Move",  sortable: true, tdClass: "col-tight col-group-l" },
  { key: "PRICE_MOVE_7D",      label: "7D Px Move",  sortable: true, tdClass: "col-tight" },
  { key: "MV_CHANGE_3M_MM",    label: "3M MV ($MM)", sortable: true, tdClass: "col-tight col-group-l" },
  { key: "MV_CHANGE_7D_MM",    label: "7D MV ($MM)", sortable: true, tdClass: "col-tight" },
  { key: "PRICE_RANGE",        label: "52W Range",   tdClass: "col-group-l" },
  { key: "COVERAGE_PRIMARY",   label: "Primary",     tdClass: "col-tight col-group-l" },
  { key: "COVERAGE_SECONDARY", label: "Secondary",   tdClass: "col-tight" },
];

const LOAN_FILTER_DEFS = [
  { id: "yieldMin", label: "Yield >", unit: "%", test: (row, v) => Number(row.YIELD) > v },
  { id: "yieldMax", label: "Yield <", unit: "%", test: (row, v) => Number(row.YIELD) < v },
  { id: "priceMax", label: "Price <", unit: "",  test: (row, v) => Number(row.PX_MID) < v },
];

const LOAN_ABS_SORT_FIELDS = new Set(["PRICE_MOVE_3M", "PRICE_MOVE_7D", "MV_CHANGE_3M_MM", "MV_CHANGE_7D_MM"]);

function renderLoanPriceRange(row) {
  const low = Number(row.PX_LOW_52W);
  const high = Number(row.PX_HIGH_52W);
  const current = Number(row.PX_MID);
  if (Number.isNaN(low) || Number.isNaN(high) || Number.isNaN(current) || high <= low) return "-";
  const clamped = Math.min(Math.max(current, low), high);
  const pct = ((clamped - low) / (high - low)) * 100;
  return `
    <div class="price-range-cell">
      <span class="range-value low">${fmt(low, 1)}</span>
      <div class="range-track"><span class="range-dot" style="left:${pct}%"></span></div>
      <span class="range-value high">${fmt(high, 1)}</span>
    </div>
  `;
}

function renderLoanFilterChips() {
  loanFilterChips.innerHTML = LOAN_FILTER_DEFS.map((def) => {
    const active = state.loanActiveFilters[def.id];
    const val = state.loanFilterValues[def.id];
    const unit = def.unit ? `<span class="chip-unit">${def.unit}</span>` : "";
    if (active) {
      return `<span class="filter-chip active">
        ${def.label}
        <input class="chip-val-input" type="number" data-filter-id="${def.id}" value="${val}" step="any" />
        ${unit}
        <button class="chip-remove" data-filter-id="${def.id}" type="button" title="Remove">×</button>
      </span>`;
    }
    return `<span class="filter-chip inactive-chip">
      ${def.label}
      <input class="chip-val-input chip-val-inactive" type="number" data-filter-id="${def.id}" value="${val}" step="any" />
      ${unit}
      <button class="chip-add" data-filter-id="${def.id}" type="button" title="Add filter">+</button>
    </span>`;
  }).join("");

  loanFilterChips.querySelectorAll(".chip-val-input:not(.chip-val-inactive)").forEach((input) => {
    input.addEventListener("change", () => {
      const v = parseFloat(input.value);
      if (!Number.isNaN(v)) { state.loanFilterValues[input.dataset.filterId] = v; applyLoansFilter(); }
    });
  });

  loanFilterChips.querySelectorAll(".chip-remove").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      state.loanActiveFilters[btn.dataset.filterId] = false;
      applyLoansFilter();
    });
  });

  loanFilterChips.querySelectorAll(".chip-add").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.dataset.filterId;
      const input = loanFilterChips.querySelector(`.chip-val-input[data-filter-id="${id}"]`);
      const v = parseFloat(input?.value);
      if (!Number.isNaN(v)) state.loanFilterValues[id] = v;
      state.loanActiveFilters[id] = true;
      applyLoansFilter();
    });
  });
}

function updateLoanScreenButton() {
  const topN = parseInt(loanTopNInput.value, 10) || 50;
  const isSizeOn = state.loanScreenMode === "size";
  const isYieldOn = state.loanScreenMode === "yield";
  loanScreenButton.textContent = isSizeOn ? `Screened: Top ${topN} by Size ✓` : `Screen by Size: Top ${topN}`;
  loanScreenButton.classList.toggle("btn-on", isSizeOn);
  loanScreenButton.classList.toggle("secondary", !isSizeOn);
  loanYieldScreenButton.textContent = isYieldOn ? `Screened: Top ${topN} by Yield ✓` : `Screen by Yield: Top ${topN}`;
  loanYieldScreenButton.classList.toggle("btn-on", isYieldOn);
  loanYieldScreenButton.classList.toggle("secondary", !isYieldOn);
}

function applyLoansFilter() {
  const search = loansSearchInput.value.trim().toLowerCase();
  let rows = search
    ? state.loanRows.filter((row) =>
        String(row.Issuer || "").toLowerCase().includes(search) ||
        String(row.PARENT_TICKER || "").toLowerCase().includes(search) ||
        String(row.NAME || "").toLowerCase().includes(search)
      )
    : [...state.loanRows];

  LOAN_FILTER_DEFS.forEach((def) => {
    if (state.loanActiveFilters[def.id]) {
      const v = state.loanFilterValues[def.id];
      rows = rows.filter((row) => def.test(row, v));
    }
  });

  if (state.loanScreenMode === "size") {
    rows = rows.filter((row) => Number(row.YIELD) >= 9 && Number(row.AMT_OUTSTANDING_MM) >= 700);
  } else if (state.loanScreenMode === "yield") {
    rows = rows.filter((row) => Number(row.AMT_OUTSTANDING_MM) >= 700);
  }

  // Yield screen forces sort by yield desc; otherwise use current sort state
  const sf = state.loanScreenMode === "yield" ? "YIELD" : state.loanSortField;
  const asc = state.loanScreenMode === "yield" ? false : state.loanSortDirection === "asc";
  rows.sort((a, b) => {
    let av = Number(a[sf]);
    let bv = Number(b[sf]);
    if (LOAN_ABS_SORT_FIELDS.has(sf)) {
      av = Number.isNaN(av) ? -Infinity : Math.abs(av);
      bv = Number.isNaN(bv) ? -Infinity : Math.abs(bv);
    } else {
      av = Number.isNaN(av) ? -Infinity : av;
      bv = Number.isNaN(bv) ? -Infinity : bv;
    }
    return asc ? av - bv : bv - av;
  });

  const topN = parseInt(loanTopNInput.value, 10);
  if (!Number.isNaN(topN) && topN > 0) rows = rows.slice(0, topN);

  state.filteredLoanRows = rows;
  clearLoansSearchButton.classList.toggle("hidden", !loansSearchInput.value);
  const screenNote = state.loanScreenMode === "size"
    ? " — Yield ≥9%, Tranche ≥$700M"
    : state.loanScreenMode === "yield"
    ? " — Face ≥$700M, sorted by Yield"
    : "";
  loansStatus.textContent = `${rows.length} loan instrument${rows.length !== 1 ? "s" : ""}${screenNote}`;
  renderLoanFilterChips();
  updateLoanScreenButton();
  renderLoansTable();
}

function loanSortIndicator(key) {
  if (state.loanSortField !== key) return "";
  return state.loanSortDirection === "asc" ? " &uarr;" : " &darr;";
}

function renderLoansTable() {
  loansHead.innerHTML = `
    <tr class="header-group-row">
      <th></th>
      <th></th>
      <th></th>
      <th class="group-header col-group-l">Face<br><em style="white-space:nowrap">($MM)</em></th>
      <th colspan="3" class="group-header col-group-l">Price &amp; Yield</th>
      <th colspan="2" class="group-header col-group-l">Px Move</th>
      <th colspan="2" class="group-header col-group-l">MV Change<br><em style="white-space:nowrap">($MM)</em></th>
      <th class="col-group-l"></th>
      <th colspan="2" class="group-header col-group-l">Coverage</th>
    </tr>
    <tr>
      <th class="col-fit">Issuer</th>
      <th class="col-fit">Security</th>
      <th class="col-tight">Sector</th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header loan-sort-header" data-loan-sort="AMT_OUTSTANDING_MM">Amt Out${loanSortIndicator("AMT_OUTSTANDING_MM")}</button></th>
      <th class="col-tight col-group-l">Current Px</th>
      <th class="col-tight">Yield</th>
      <th class="col-tight">Coupon</th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header loan-sort-header" data-loan-sort="PRICE_MOVE_3M">3M${loanSortIndicator("PRICE_MOVE_3M")}</button></th>
      <th class="col-tight"><button type="button" class="sort-header loan-sort-header" data-loan-sort="PRICE_MOVE_7D">7D${loanSortIndicator("PRICE_MOVE_7D")}</button></th>
      <th class="col-tight col-group-l"><button type="button" class="sort-header loan-sort-header" data-loan-sort="MV_CHANGE_3M_MM">3M${loanSortIndicator("MV_CHANGE_3M_MM")}</button></th>
      <th class="col-tight"><button type="button" class="sort-header loan-sort-header" data-loan-sort="MV_CHANGE_7D_MM">7D${loanSortIndicator("MV_CHANGE_7D_MM")}</button></th>
      <th class="col-group-l">52W Range</th>
      <th class="col-tight col-group-l">Primary</th>
      <th class="col-tight">Secondary</th>
    </tr>
  `;

  loansBody.innerHTML = state.filteredLoanRows.map((row) => {
    const cells = loanColumns.map((col) => {
      const cls = col.tdClass ? ` class="${col.tdClass}"` : "";
      if (col.key === "PRICE_RANGE") return `<td${cls}>${renderLoanPriceRange(row)}</td>`;
      if (col.key === "PRICE_MOVE_3M") return `<td${cls}>${fmtPriceMove(row.PRICE_MOVE_3M)}</td>`;
      if (col.key === "PRICE_MOVE_7D") return `<td${cls}>${fmtPriceMove(row.PRICE_MOVE_7D)}</td>`;
      if (col.key === "MV_CHANGE_3M_MM" || col.key === "MV_CHANGE_7D_MM") {
        const v = Number(row[col.key]);
        if (!row[col.key] || Number.isNaN(v) || v === 0) return `<td${cls}>-</td>`;
        const mvCls = v > 0 ? "positive" : "negative";
        const arrow = v > 0 ? "▲" : "▼";
        return `<td class="${(col.tdClass || "")} price-move ${mvCls}">${arrow} ${fmt(Math.abs(v), 0)}</td>`;
      }
      if (col.key === "AMT_OUTSTANDING_MM") return `<td${cls}>${fmt(row[col.key], 0)}</td>`;
      if (col.key === "PX_MID") return `<td${cls}>${fmt(row[col.key], 2)}</td>`;
      if (col.key === "YIELD") return `<td${cls}>${fmt(row[col.key], 2)}</td>`;
      if (col.key === "COUPON_DISPLAY") {
        const ri = String(row.RESET_INDEX || "").trim().replace(/\s*index\s*/i, "");
        const margin = row.LN_CURRENT_MARGIN;
        if (ri) {
          const marginStr = margin != null && margin !== "" ? ` + ${Math.round(Number(margin))}` : "";
          return `<td${cls}>${ri}${marginStr}</td>`;
        }
        const cpn = row.CPN_VALUE;
        return `<td${cls}>${cpn != null && cpn !== "" ? fmt(cpn, 2) : "-"}</td>`;
      }
      if (col.key === "NAME") {
        const text = row[col.key] != null ? row[col.key] : "-";
        return `<td${cls}>${text}</td>`;
      }
      if (col.key === "COVERAGE_PRIMARY" || col.key === "COVERAGE_SECONDARY") {
        const slot = col.key === "COVERAGE_PRIMARY" ? "primary" : "secondary";
        const ticker = row.PARENT_TICKER;
        const names = (state.coverageMap[ticker] || {})[slot] || [];
        return `<td${cls}>${renderCoverageCell(ticker, slot, names)}</td>`;
      }
      return `<td${cls}>${row[col.key] != null ? row[col.key] : "-"}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  document.querySelectorAll("[data-loan-sort]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const col = btn.dataset.loanSort;
      if (state.loanSortField === col) {
        state.loanSortDirection = state.loanSortDirection === "asc" ? "desc" : "asc";
      } else {
        state.loanSortField = col;
        state.loanSortDirection = "desc";
      }
      applyLoansFilter();
    });
  });
}

function setExceptionsSubtab(tabName) {
  abnormalPriceSubtabButton.classList.toggle("active", tabName === "abnormal-price");
  exclusionsSubtabButton.classList.toggle("active", tabName === "exclusions");
  abnormalBbgSubtabButton.classList.toggle("active", tabName === "abnormal-bbg");
  abnormalPricePanel.classList.toggle("hidden", tabName !== "abnormal-price");
  exclusionsPanel.classList.toggle("hidden", tabName !== "exclusions");
  abnormalBbgPanel.classList.toggle("hidden", tabName !== "abnormal-bbg");
}

function applyMoversFilters() {
  const direction = moversDirectionFilter.value;
  const sortBy = moversSortField.value;
  const ascending = moversSortDirection.value === "asc";
  const is7d = sortBy === "7D Price Move";
  const moveKey = is7d ? "7D Price Move" : "3M Price Move";
  const threshold = is7d ? 3 : 10;

  let rows = [...state.rawMoversRows];
  rows = rows.filter((row) => {
    const currentPx = Number(row["Current Px"]);
    return !Number.isNaN(currentPx) && currentPx <= 105;
  });
  rows = rows.filter((row) => {
    const move = Number(row[moveKey]);
    if (Number.isNaN(move)) return false;
    if (direction === "down") return move <= -threshold;
    if (direction === "up") return move >= threshold;
    return Math.abs(move) > threshold;
  });

  const effectiveSortKey = sortBy === "Current Px" ? "Current Px" : moveKey;
  rows.sort((left, right) => {
    const a = Number(left[effectiveSortKey]);
    const b = Number(right[effectiveSortKey]);
    const result = (Number.isNaN(a) ? 0 : a) - (Number.isNaN(b) ? 0 : b);
    return ascending ? result : -result;
  });

  state.moversRows = rows;
  const sectorMap = new Map();
  const sectorCountMap = new Map();
  rows.forEach((row) => {
    const sector = row.Sector || "Unknown";
    const amount = Number(row["Amount Out ($MM)"]) || 0;
    sectorMap.set(sector, (sectorMap.get(sector) || 0) + amount);
    sectorCountMap.set(sector, (sectorCountMap.get(sector) || 0) + 1);
  });
  state.moversSectorSummary = [...sectorMap.entries()]
    .map(([Sector, amount]) => ({
      Sector,
      "Amount Out ($MM)": amount,
      "Security Count": sectorCountMap.get(Sector) || 0,
    }))
    .sort((a, b) => {
      const metricKey = moversChartMetric.value === "count" ? "Security Count" : "Amount Out ($MM)";
      return (b[metricKey] || 0) - (a[metricKey] || 0);
    });

  const periodLabel = is7d ? "7D" : "3M";
  const directionLabel = direction === "down" ? `${periodLabel} price move down below -${threshold}` : direction === "up" ? `${periodLabel} price move up above ${threshold}` : `absolute ${periodLabel} price move above ${threshold}`;
  moversStatus.textContent = `${state.moversRows.length} securities with ${directionLabel}`;
  renderMoversChart();
  renderMoversTable();
}

function setActiveTab(tabName) {
  const issuerActive = tabName === "issuer";
  const loansActive = tabName === "loans";
  const moversActive = tabName === "movers";
  const exceptionsActive = tabName === "exceptions";
  const documentationActive = tabName === "documentation";
  issuerTabButton.classList.toggle("active", issuerActive);
  loansTabButton.classList.toggle("active", loansActive);
  moversTabButton.classList.toggle("active", moversActive);
  exceptionsTabButton.classList.toggle("active", exceptionsActive);
  documentationTabButton.classList.toggle("active", documentationActive);
  issuerTabPanel.classList.toggle("hidden", !issuerActive);
  loansTabPanel.classList.toggle("hidden", !loansActive);
  moversTabPanel.classList.toggle("hidden", !moversActive);
  exceptionsTabPanel.classList.toggle("hidden", !exceptionsActive);
  documentationTabPanel.classList.toggle("hidden", !documentationActive);
}

function applyFilters() {
  const search = searchInput.value.trim().toLowerCase();
  const sector = sectorFilter.value;
  const sortBy = state.sortField;
  const ascending = state.sortDirection === "asc";

  let rows = [...state.issuers];
  if (sector !== "All") rows = rows.filter((row) => row.Sector === sector);
  if (search) {
    rows = rows.filter((row) =>
      String(row.Issuer || "").toLowerCase().includes(search) ||
      String(row.PARENT_TICKER || "").toLowerCase().includes(search)
    );
  }

  const mvColumns = new Set(["3M MV Change TARGET ($MM)", "7D MV Change TARGET ($MM)"]);
  const isMvSort = mvColumns.has(sortBy);
  rows.sort((left, right) => {
    const a = left[sortBy];
    const b = right[sortBy];
    const aNum = typeof a === "number" ? a : Number(a);
    const bNum = typeof b === "number" ? b : Number(b);
    const numeric = !Number.isNaN(aNum) && !Number.isNaN(bNum);
    const useAbs = isMvSort && state.mvAbsSort;
    const av = numeric && useAbs ? Math.abs(aNum) : aNum;
    const bv = numeric && useAbs ? Math.abs(bNum) : bNum;
    const result = numeric ? av - bv : String(a ?? "").localeCompare(String(b ?? ""));
    return ascending ? result : -result;
  });

  state.filteredIssuers = rows;
  const anchorDateEl = document.getElementById("anchorDate");
  if (anchorDateEl) anchorDateEl.textContent = `As of ${state.metadata.anchor_date || "n/a"}`;
  statusBar.textContent = `${state.issuers.length} total issuers | Showing ${rows.length}`;
  clearSearchButton.classList.toggle("hidden", !searchInput.value);
  if (state.selectedIssuer && !rows.some((row) => row.PARENT_TICKER === state.selectedIssuer)) {
    state.selectedIssuer = null;
    detailCard.classList.add("hidden");
  } else if (state.selectedIssuer) {
    renderIssuerDetail(state.selectedIssuer);
    return;
  }
  renderIssuerTable();
}

function isQualifyingRow(row, minMaturity) {
  if (row._IS_DEFAULTED) return false;
  const rowYield = Number(row.YIELD);
  const rowPx = Number(row.PX_MID);
  const rowAmount = Number(row.AMT_OUTSTANDING_MM);
  const isLoan = row.LOAN_TYPE != null && row.LOAN_TYPE !== "";
  const maturity = row.MATURITY ? new Date(row.MATURITY) : null;
  const maturityEligible = !minMaturity || !maturity || (maturity instanceof Date && !Number.isNaN(maturity.getTime()) && maturity > minMaturity);
  if (isLoan) return maturityEligible && rowAmount >= 700 && rowYield >= 9 && rowYield <= 50 && rowPx <= 95;
  const rankText = String(row.PAYMENT_RANK || "").toLowerCase();
  const isPreferred = rankText === "preferred";
  if (!maturityEligible || rowAmount < 400) return false;
  if (isPreferred) return rowYield >= 10 && rowYield <= 50 && rowPx <= 100;
  return rowYield >= 10 && rowYield <= 50 && rowPx <= 100 && !rankText.includes("subordinated");
}

function getExclusionReasons(row, minMaturity) {
  const reasons = [];
  const currency = String(row.ISS_CURRENCY || "").toUpperCase().trim();
  if (currency && currency !== "USD") reasons.push(`Non-USD (${currency})`);
  const rawYield = row.YIELD;
  const rawPx = row.PX_MID;
  const yieldMissing = rawYield === null || rawYield === undefined || rawYield === "";
  const pxMissing = rawPx === null || rawPx === undefined || rawPx === "";
  const rowYield = Number(rawYield);
  const rowPx = Number(rawPx);
  const rowAmount = Number(row.AMT_OUTSTANDING_MM);
  const isLoan = row.LOAN_TYPE != null && row.LOAN_TYPE !== "";
  if (isLoan) {
    const maturity = row.MATURITY ? new Date(row.MATURITY) : null;
    const maturityEligible = !minMaturity || !maturity || (maturity instanceof Date && !Number.isNaN(maturity.getTime()) && maturity > minMaturity);
    if (!maturityEligible) reasons.push("Maturity ≤2M");
    if (rowAmount < 700) reasons.push("Size < $700M");
    if (yieldMissing) reasons.push("No yield data");
    else if (rowYield < 9) reasons.push("Yield < 9%");
    else if (rowYield > 50) reasons.push("Yield > 50%");
    if (!pxMissing && rowPx > 95) reasons.push("Price > 95");
    return reasons.join("; ") || "Other";
  }
  const rankText = String(row.PAYMENT_RANK || "").toLowerCase();
  const maturity = row.MATURITY ? new Date(row.MATURITY) : null;
  const maturityEligible = !minMaturity || !maturity || (maturity instanceof Date && !Number.isNaN(maturity.getTime()) && maturity > minMaturity);
  if (!maturityEligible) reasons.push("Maturity ≤2M");
  if (rowAmount < 400) reasons.push("Size < $400M");
  if (rankText.includes("subordinated")) reasons.push("Subordinated");
  if (yieldMissing) reasons.push("No yield data");
  else if (rowYield < 10) reasons.push("Yield < 10%");
  else if (rowYield > 50) reasons.push("Yield > 50%");
  if (!pxMissing && rowPx > 100) reasons.push("Price > 100");
  return reasons.join("; ") || "Other";
}

function renderNonQualifyingHead() {
  return `<tr>${detailColumns.map((column) => {
    const sup = column.key === "LQA_EXPECTED_DAILY_VOLUME_MM" ? "<sup>6</sup>" : column.key === "LQA_LIQUIDITY_SCORE" ? "<sup>7</sup>" : "";
    return `<th>${column.label}${sup}</th>`;
  }).join("")}<th>Exclusion Reason</th></tr>`;
}

function renderNonQualifyingRows(rows, minMaturity) {
  return rows.map((row) => {
    const cells = detailColumns.map((column) => {
      const value = row[column.key];
      if (column.key === "NAME") {
        const hMarker = row._IS_HOLDING ? `<sup class="holding-marker" title="Portfolio holding">H</sup>` : "";
        const bbgId = row.ID || "";
        return `<td title="${bbgId}">${fmt(value, 2)}${hMarker}</td>`;
      }
      if (column.key === "PRICE_MOVE_3M") return `<td>${renderPriceMove(row)}</td>`;
      if (column.key === "PRICE_MOVE_7D") return `<td>${renderPriceMove7D(row)}</td>`;
      if (column.key === "PRICE_RANGE") return `<td>${renderPriceRange(row)}</td>`;
      if (column.key === "LQA_EXPECTED_DAILY_VOLUME_MM") {
        const v = Number(value);
        return `<td>${value == null || isNaN(v) || v < 0 ? "-" : fmt(v, 1)}</td>`;
      }
      const className = column.key === "AMT_OUTSTANDING_MM" ? "detail-narrow" : "";
      return `<td class="${className}">${fmt(value, detailDigits(column.key))}</td>`;
    }).join("");
    const reason = getExclusionReasons(row, minMaturity);
    const rowClass = row._IS_HOLDING ? ` class="holding-row"` : "";
    return `<tr${rowClass}>${cells}<td class="exclusion-reason-cell">${reason}</td></tr>`;
  }).join("");
}

function detailSortFn(a, b) {
  const aSecured = String(a.PAYMENT_RANK || "").toLowerCase().includes("secured") && !String(a.PAYMENT_RANK || "").toLowerCase().includes("unsecured");
  const bSecured = String(b.PAYMENT_RANK || "").toLowerCase().includes("secured") && !String(b.PAYMENT_RANK || "").toLowerCase().includes("unsecured");
  if (aSecured !== bSecured) return aSecured ? -1 : 1;
  return (Number(b.AMT_OUTSTANDING_MM) || 0) - (Number(a.AMT_OUTSTANDING_MM) || 0);
}

function renderDetailRows(rows) {
  return rows.map((row) => {
    const cells = detailColumns.map((column) => {
      const value = row[column.key];
      if (column.key === "NAME") {
        const dMarker = row._IS_DEFAULTED ? `<sup class="defaulted-marker" title="Defaulted">D</sup>` : "";
        const hMarker = row._IS_HOLDING ? `<sup class="holding-marker" title="Portfolio holding">H</sup>` : "";
        const bbgId = row.ID || "";
        return `<td title="${bbgId}">${fmt(value, 2)}${dMarker}${hMarker}</td>`;
      }
      if (column.key === "PRICE_MOVE_3M") return `<td>${renderPriceMove(row)}</td>`;
      if (column.key === "PRICE_MOVE_7D") return `<td>${renderPriceMove7D(row)}</td>`;
      if (column.key === "PRICE_RANGE") return `<td>${renderPriceRange(row)}</td>`;
      if (column.key === "LQA_EXPECTED_DAILY_VOLUME_MM") {
        const v = Number(value);
        return `<td>${value == null || isNaN(v) || v < 0 ? "-" : fmt(v, 1)}</td>`;
      }
      const className = column.key === "AMT_OUTSTANDING_MM" ? "detail-narrow" : "";
      return `<td class="${className}">${fmt(value, detailDigits(column.key))}</td>`;
    }).join("");
    const rowClass = row._IS_HOLDING ? ` class="holding-row"` : "";
    return `<tr${rowClass}>${cells}</tr>`;
  }).join("");
}

function renderDetailHead() {
  return `<tr>${detailColumns.map((column) => {
    const sup = column.key === "LQA_EXPECTED_DAILY_VOLUME_MM" ? "<sup>6</sup>" : column.key === "LQA_LIQUIDITY_SCORE" ? "<sup>7</sup>" : "";
    return `<th>${column.label}${sup}</th>`;
  }).join("")}</tr>`;
}

function renderIssuerDetail(parentTicker) {
  const anchorDate = state.metadata.anchor_date ? new Date(state.metadata.anchor_date) : null;
  const minMaturity = anchorDate ? new Date(anchorDate) : null;
  if (minMaturity) minMaturity.setMonth(minMaturity.getMonth() + 2);

  const allRows = state.instrumentMap.get(parentTicker) || [];
  const rows = allRows.filter((row) => isQualifyingRow(row, minMaturity)).sort(detailSortFn);
  const nonQualRows = allRows.filter((row) => !row._IS_DEFAULTED && !isQualifyingRow(row, minMaturity)).sort(detailSortFn);
  const defaultedRows = allRows.filter((row) => row._IS_DEFAULTED);

  state.detailRows = rows;
  const issuer = state.issuers.find((row) => row.PARENT_TICKER === parentTicker);
  const hasHoldings = !!issuer?._HAS_HOLDING;
  state.selectedIssuer = parentTicker;
  detailCard.classList.remove("hidden");
  detailCard.classList.toggle("issuer-has-holdings", hasHoldings);
  detailToggleBar.classList.remove("hidden");
  hideDetailButton.innerHTML = "&#x25BC; Hide Detail";
  detailTicker.textContent = parentTicker;
  const holdingBadge = hasHoldings ? ` <span class="holding-badge">Holding</span>` : "";
  detailTitle.innerHTML = (issuer?.Issuer || parentTicker) + holdingBadge;
  detailHead.innerHTML = renderDetailHead();
  detailBody.innerHTML = renderDetailRows(rows);

  // Non-qualifying securities section
  if (nonQualRows.length > 0) {
    nonQualifyingCard.classList.remove("hidden");
    nonQualifyingHead.innerHTML = renderNonQualifyingHead();
    nonQualifyingBody.innerHTML = renderNonQualifyingRows(nonQualRows, minMaturity);
  } else {
    nonQualifyingCard.classList.add("hidden");
  }

  // Defaulted bonds section
  if (defaultedRows.length > 0) {
    defaultedCard.classList.remove("hidden");
    defaultedHead.innerHTML = renderDetailHead();
    defaultedBody.innerHTML = renderDetailRows(defaultedRows);
  } else {
    defaultedCard.classList.add("hidden");
  }

  renderIssuerTable();
}

async function openIssuerDetail(parentTicker) {
  if (!parentTicker || parentTicker === "null" || parentTicker === "undefined") return;
  if (!state.instrumentMap.has(parentTicker)) {
    const payload = await fetchJson(`/api/issuer-detail?parent_ticker=${encodeURIComponent(parentTicker)}`);
    state.instrumentMap.set(payload.issuer, payload.rows);
  }
  renderIssuerDetail(parentTicker);
}

async function loadDashboard() {
  setLoading(true, "Loading...");
  const [dashboardPayload, moversPayload, abnormalPricesPayload, excludedPayload, loansPayload, coveragePayload] = await Promise.all([
    fetchJson("/api/dashboard"),
    fetchJson("/api/price-movers"),
    fetchJson("/api/abnormal-prices"),
    fetchJson("/api/excluded"),
    fetchJson("/api/loans"),
    fetchJson("/api/coverage"),
  ]);
  state.issuers = dashboardPayload.issuers;
  state.filters = dashboardPayload.filters;
  state.metadata = dashboardPayload.metadata;
  state.selectedIssuer = null;
  state.instrumentMap = new Map();
  state.detailRows = [];
  detailCard.classList.add("hidden");
  sectorFilter.innerHTML = `<option value="All">All</option>${dashboardPayload.filters.sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("")}`;
  state.rawMoversRows = moversPayload.rows.map((row) => ({ ...row, "Price Range": "" }));
  state.abnormalPriceRows = abnormalPricesPayload.rows.map((row) => ({
    Issuer: row.Issuer,
    Security: row.NAME,
    Sector: row.SECTOR,
    Rank: row.PAYMENT_RANK,
    Mat: row.MATURITY,
    "Amount Out ($MM)": row.AMT_OUTSTANDING_MM,
    "Current Px": row.PX_MID,
    "Prior Px": row.PX_MID_T90,
    "3M Price Move": row.PRICE_MOVE_3M,
    Yield: row.YIELD,
    "52W Low": row.PX_LOW_52W,
    "52W High": row.PX_HIGH_52W,
  }));
  state.excludedRows = excludedPayload.rows;
  state.loanRows = loansPayload.rows;
  state.coverageMap = coveragePayload.coverages || {};
  applyFilters();
  applyMoversFilters();
  applyLoansFilter();
  renderAbnormalPriceTable();
  setLoading(false);
}

async function checkSession() {
  const payload = await fetchJson("/api/session");
  if (!payload.authenticated) {
    loginCard.classList.remove("hidden");
    app.classList.add("hidden");
    logoutButton.classList.add("hidden");
    setLoading(false);
    return;
  }
  loginCard.classList.add("hidden");
  app.classList.remove("hidden");
  logoutButton.classList.remove("hidden");
  try {
    await loadDashboard();
  } catch (error) {
    setLoading(false);
    statusBar.textContent = error.message;
    throw error;
  }
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginError.classList.add("hidden");
  const formData = new FormData(loginForm);
  try {
    await fetchJson("/login", { method: "POST", body: formData });
    await checkSession();
  } catch (error) {
    setLoading(false);
    loginError.textContent = error.message;
    loginError.classList.remove("hidden");
  }
});

logoutButton.addEventListener("click", async () => {
  await fetchJson("/logout", { method: "POST" });
  window.location.reload();
});

issuerTabButton.addEventListener("click", () => setActiveTab("issuer"));
loansTabButton.addEventListener("click", () => setActiveTab("loans"));
moversTabButton.addEventListener("click", () => setActiveTab("movers"));
exceptionsTabButton.addEventListener("click", () => setActiveTab("exceptions"));
documentationTabButton.addEventListener("click", () => setActiveTab("documentation"));
abnormalPriceSubtabButton.addEventListener("click", () => {
  setExceptionsSubtab("abnormal-price");
  renderAbnormalPriceTable();
});
exclusionsSubtabButton.addEventListener("click", () => {
  setExceptionsSubtab("exclusions");
  renderExclusionsTable();
});
abnormalBbgSubtabButton.addEventListener("click", () => setExceptionsSubtab("abnormal-bbg"));

[searchInput, sectorFilter].forEach((element) => {
  element.addEventListener("input", applyFilters);
  element.addEventListener("change", applyFilters);
});

[moversDirectionFilter, moversChartMetric, moversSortField, moversSortDirection].forEach((element) => {
  element.addEventListener("input", applyMoversFilters);
  element.addEventListener("change", applyMoversFilters);
});

clearSearchButton.addEventListener("click", () => {
  searchInput.value = "";
  applyFilters();
  searchInput.focus();
});

loansSearchInput.addEventListener("input", applyLoansFilter);
clearLoansSearchButton.addEventListener("click", () => {
  loansSearchInput.value = "";
  applyLoansFilter();
  loansSearchInput.focus();
});

loanTopNInput.addEventListener("input", () => {
  updateLoanScreenButton();
  applyLoansFilter();
});

loanScreenButton.addEventListener("click", () => {
  state.loanScreenMode = state.loanScreenMode === "size" ? null : "size";
  applyLoansFilter();
});

loanYieldScreenButton.addEventListener("click", () => {
  state.loanScreenMode = state.loanScreenMode === "yield" ? null : "yield";
  applyLoansFilter();
});

reloadDataButton.addEventListener("click", async () => {
  reloadDataButton.disabled = true;
  reloadDataButton.textContent = "Reloading...";
  try {
    await fetchJson("/api/admin/reload", { method: "POST" });
    await loadDashboard();
  } catch (error) {
    statusBar.textContent = error.message;
  } finally {
    reloadDataButton.disabled = false;
    reloadDataButton.textContent = "Reload Data";
  }
});

downloadIssuersButton.addEventListener("click", async () => {
  await downloadExcel("/api/export/issuers", {
    rows: state.filteredIssuers.slice(0, 50),
    upsideMode: state.upsideMode,
  });
});

hideDetailButton.addEventListener("click", () => {
  const hiding = !detailCard.classList.contains("hidden");
  detailCard.classList.toggle("hidden", hiding);
  nonQualifyingCard.classList.toggle("hidden", hiding);
  defaultedCard.classList.toggle("hidden", hiding);
  hideDetailButton.innerHTML = hiding ? "&#x25B2; Show Detail" : "&#x25BC; Hide Detail";
});

downloadDetailButton.addEventListener("click", async () => {
  if (!state.selectedIssuer) return;
  await downloadExcel("/api/export/detail", { issuer: state.selectedIssuer, rows: state.detailRows });
});

downloadLoansButton.addEventListener("click", async () => {
  await downloadExcel("/api/export/loans", { rows: state.filteredLoanRows });
});

document.getElementById("glossaryButton").addEventListener("click", () => {
  const existing = document.querySelector(".glossary-popup");
  if (existing) { existing.remove(); return; }
  const popup = document.createElement("div");
  popup.className = "glossary-popup";
  popup.innerHTML = `
    <table>${glossaryEntries.map((e) =>
      `<tr><td><sup>${e.sup}</sup></td><td><strong>${e.label}</strong></td><td>${e.def}</td></tr>`
    ).join("")}</table>
    <div class="glossary-criteria">
      <strong>Screening Criteria</strong>
      <table>
        <tr>
          <td><strong>Loans</strong></td>
          <td>Yield 9–50% &nbsp;|&nbsp; Price &le; 95 &nbsp;|&nbsp; Tranche &ge; $700M &nbsp;|&nbsp; Maturity &gt; 2M</td>
        </tr>
        <tr>
          <td><strong>Bonds</strong></td>
          <td>Yield 10–50% &nbsp;|&nbsp; Price &le; 100 &nbsp;|&nbsp; Tranche &ge; $400M &nbsp;|&nbsp; Maturity &gt; 2M &nbsp;|&nbsp; Non-subordinated</td>
        </tr>
      </table>
    </div>`;
  document.getElementById("glossaryButton").parentElement.style.position = "relative";
  document.getElementById("glossaryButton").parentElement.appendChild(popup);
  setTimeout(() => document.addEventListener("click", (ev) => {
    if (!popup.contains(ev.target) && ev.target.id !== "glossaryButton") popup.remove();
  }, { once: true }), 0);
});

// ── Coverage dropdown ──────────────────────────────────────────────────────
const coverageDropdown = document.createElement("div");
coverageDropdown.className = "coverage-dropdown hidden";
document.body.appendChild(coverageDropdown);

let _covTicker = null;
let _covSlot = null;

function showCoverageDropdown(ticker, slot, anchorEl) {
  _covTicker = ticker;
  _covSlot = slot;
  const entry = state.coverageMap[ticker] || { primary: [], secondary: [] };
  const current = entry[slot] || [];
  const available = COVERAGE_ANALYSTS.filter((n) => !current.includes(n));
  if (!available.length) { hideCoverageDropdown(); return; }
  coverageDropdown.innerHTML = available.map((name) =>
    `<div class="cov-option" data-name="${name.replace(/"/g, "&quot;")}">${name}</div>`
  ).join("");
  const rect = anchorEl.getBoundingClientRect();
  coverageDropdown.style.top = `${rect.bottom + window.scrollY + 2}px`;
  coverageDropdown.style.left = `${rect.left + window.scrollX}px`;
  coverageDropdown.classList.remove("hidden");
}

function hideCoverageDropdown() {
  coverageDropdown.classList.add("hidden");
  _covTicker = null;
  _covSlot = null;
}

function saveCoverage(ticker) {
  const entry = state.coverageMap[ticker] || { primary: [], secondary: [] };
  fetch("/api/coverage", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ticker, primary: entry.primary || [], secondary: entry.secondary || [] }),
  }).catch(() => {});
}

coverageDropdown.addEventListener("click", (e) => {
  const opt = e.target.closest(".cov-option");
  if (!opt || !_covTicker) return;
  const name = opt.dataset.name;
  if (!state.coverageMap[_covTicker]) state.coverageMap[_covTicker] = { primary: [], secondary: [] };
  const arr = state.coverageMap[_covTicker][_covSlot];
  if (arr.length < 2 && !arr.includes(name)) {
    arr.push(name);
    updateCoverageCells(_covTicker);
    renderLoansTable();
    saveCoverage(_covTicker);
  }
  hideCoverageDropdown();
});

issuerBody.addEventListener("click", (e) => {
  const addBtn = e.target.closest(".cov-add");
  if (addBtn) {
    e.stopPropagation();
    const { ticker, slot } = addBtn.dataset;
    if (!coverageDropdown.classList.contains("hidden") && _covTicker === ticker && _covSlot === slot) {
      hideCoverageDropdown();
    } else {
      showCoverageDropdown(ticker, slot, addBtn);
    }
    return;
  }
  const nameSpan = e.target.closest(".cov-name");
  if (nameSpan) {
    e.stopPropagation();
    const { ticker, slot, idx } = nameSpan.dataset;
    if (state.coverageMap[ticker]) {
      state.coverageMap[ticker][slot].splice(Number(idx), 1);
      updateCoverageCells(ticker);
      renderLoansTable();
      saveCoverage(ticker);
    }
    hideCoverageDropdown();
  }
});

loansBody.addEventListener("click", (e) => {
  const addBtn = e.target.closest(".cov-add");
  if (addBtn) {
    e.stopPropagation();
    const { ticker, slot } = addBtn.dataset;
    if (!coverageDropdown.classList.contains("hidden") && _covTicker === ticker && _covSlot === slot) {
      hideCoverageDropdown();
    } else {
      showCoverageDropdown(ticker, slot, addBtn);
    }
    return;
  }
  const nameSpan = e.target.closest(".cov-name");
  if (nameSpan) {
    e.stopPropagation();
    const { ticker, slot, idx } = nameSpan.dataset;
    if (state.coverageMap[ticker]) {
      state.coverageMap[ticker][slot].splice(Number(idx), 1);
      updateCoverageCells(ticker);
      renderLoansTable();
      saveCoverage(ticker);
    }
    hideCoverageDropdown();
  }
});

document.addEventListener("click", (e) => {
  if (!coverageDropdown.classList.contains("hidden") &&
      !coverageDropdown.contains(e.target) &&
      !e.target.closest(".cov-add")) {
    hideCoverageDropdown();
  }
});
// ────────────────────────────────────────────────────────────────────────────

checkSession();
