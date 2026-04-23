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
  upsideMode: "52w",
  moveMode: "3m",
  sortField: "52W PEAK UPSIDE SECURED ($MM)",
  sortDirection: "desc",
};

const loginCard = document.getElementById("loginCard");
const app = document.getElementById("app");
const loginForm = document.getElementById("loginForm");
const loginError = document.getElementById("loginError");
const logoutButton = document.getElementById("logoutButton");
const loadingBanner = document.getElementById("loadingBanner");
const issuerTabButton = document.getElementById("issuerTabButton");
const moversTabButton = document.getElementById("moversTabButton");
const exceptionsTabButton = document.getElementById("exceptionsTabButton");
const issuerTabPanel = document.getElementById("issuerTabPanel");
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
const detailCard = document.getElementById("detailCard");
const detailTicker = document.getElementById("detailTicker");
const detailTitle = document.getElementById("detailTitle");
const detailHead = document.getElementById("detailHead");
const detailBody = document.getElementById("detailBody");
const defaultedCard = document.getElementById("defaultedCard");
const defaultedHead = document.getElementById("defaultedHead");
const defaultedBody = document.getElementById("defaultedBody");
const reportSummaryCard = document.getElementById("reportSummaryCard");
const reportMeta = document.getElementById("reportMeta");
const reportBullets = document.getElementById("reportBullets");
const statusBar = document.getElementById("statusBar");
const searchInput = document.getElementById("searchInput");
const clearSearchButton = document.getElementById("clearSearchButton");
const sectorFilter = document.getElementById("sectorFilter");
const downloadIssuersButton = document.getElementById("downloadIssuersButton");
const downloadDetailButton = document.getElementById("downloadDetailButton");

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
  "3M Price Move",
  "7D Price Move",
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
  { key: "LAST_30D_VOLUME_MM", label: "Last Month Traded Volume ($MM)" },
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

const zeroAsDashColumns = new Set([
  "Secured Face ($BN)", "Unsecured Face ($BN)", "Preferred Face ($BN)",
  "52W PEAK UPSIDE SECURED ($MM)", "52W PEAK UPSIDE UNSECURED ($MM)", "52W PEAK UPSIDE PREFERRED ($MM)",
  "RETURN TO PAR SECURED ($MM)", "RETURN TO PAR UNSECURED ($MM)", "RETURN TO PAR PREFERRED ($MM)",
]);

function fmtIssuer(column, value) {
  if (zeroAsDashColumns.has(column) && (value === 0 || value === null || value === undefined)) return "-";
  return fmt(value, issuerDigits(column));
}

function issuerDigits(column) {
  if (["Secured Face ($BN)", "Unsecured Face ($BN)", "Preferred Face ($BN)", "Price", "Yield", "3M Price Move", "7D Price Move"].includes(column)) return 1;
  if (["52W PEAK UPSIDE SECURED ($MM)", "52W PEAK UPSIDE UNSECURED ($MM)", "52W PEAK UPSIDE PREFERRED ($MM)", "RETURN TO PAR SECURED ($MM)", "RETURN TO PAR UNSECURED ($MM)", "RETURN TO PAR PREFERRED ($MM)", "OAS Delta (bps)"].includes(column)) return 0;
  return 2;
}

function detailDigits(column) {
  if (column === "AMT_OUTSTANDING_MM") return 0;
  if (["PX_MID", "PRICE_MOVE_3M", "PRICE_MOVE_7D", "YIELD", "LAST_30D_VOLUME_MM", "PX_HIGH_52W", "PX_LOW_52W"].includes(column)) return 1;
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
  { sup: "1", label: "Face ($BN)", def: "Face value of instruments that fit the screening criteria" },
  { sup: "2", label: "Price", def: "Price weighted by face amounts outstanding" },
  { sup: "3", label: "Yield", def: "Yield weighted by face amounts outstanding" },
  { sup: "4", label: "3M / 7D Price Move", def: "Largest price movement among qualifying instruments over the time period" },
  { sup: "5", label: "52W Peak Upside", def: "The dollar amount to be gained assuming all securities' prices revert to their 52-week highest prices" },
  { sup: "6", label: "Return To Par Upside", def: "The dollar amount to be gained assuming all securities' prices revert to par (100)" },
  { sup: "7", label: "Last Month Traded Volume ($MM)", def: "Traded volume (in thousands of USD) for the last month based on all trades reported to FINRA TRACE" },
];

function renderIssuerTable() {
  const faceColumns = ["Secured Face ($BN)", "Unsecured Face ($BN)", "Preferred Face ($BN)"];

  const is52w = state.upsideMode === "52w";
  const upsideLabel = is52w ? "52W Peak Upside" : "Return to Par Upside";
  const upsideSup = is52w ? "<sup>5</sup>" : "<sup>6</sup>";
  const upsideSecKey = is52w ? "52W PEAK UPSIDE SECURED ($MM)" : "RETURN TO PAR SECURED ($MM)";
  const upsideUnsecKey = is52w ? "52W PEAK UPSIDE UNSECURED ($MM)" : "RETURN TO PAR UNSECURED ($MM)";
  const upsidePrefKey = is52w ? "52W PEAK UPSIDE PREFERRED ($MM)" : "RETURN TO PAR PREFERRED ($MM)";

  const is3m = state.moveMode === "3m";
  const moveLabel = is3m ? "3M Price Move" : "7D Price Move";
  const moveKey = is3m ? "3M Price Move" : "7D Price Move";

  const tableColumnOrder = [
    "Issuer", "Sector",
    ...faceColumns,
    upsideSecKey,
    upsideUnsecKey,
    upsidePrefKey,
    "Price", "Yield",
    moveKey,
    "COVERAGE PRIMARY",
    "COVERAGE SECONDARY",
  ];

  issuerHead.innerHTML = `
    <tr class="header-group-row">
      <th></th>
      <th></th>
      <th colspan="3" class="group-header">Face ($BN)<sup>1</sup></th>
      <th colspan="3" class="group-header upside-toggle" id="upsideToggle" title="Click to toggle">${upsideLabel}${upsideSup} &#x21c4;</th>
      <th></th>
      <th></th>
      <th class="group-header upside-toggle" id="moveToggle" title="Click to toggle">${moveLabel}<sup>4</sup> &#x21c4;</th>
      <th colspan="2" class="group-header">Coverage (WIP)</th>
    </tr>
    <tr>
      <th class="col-fit">Issuer</th>
      <th class="col-fit">Sector</th>
      <th><button type="button" class="sort-header" data-sort="Secured Face ($BN)">Secured${sortIndicator("Secured Face ($BN)")}</button></th>
      <th><button type="button" class="sort-header" data-sort="Unsecured Face ($BN)">Unsecured${sortIndicator("Unsecured Face ($BN)")}</button></th>
      <th><button type="button" class="sort-header" data-sort="Preferred Face ($BN)">Preferred${sortIndicator("Preferred Face ($BN)")}</button></th>
      <th><button type="button" class="sort-header" data-sort="${upsideSecKey}">Secured ($MM)${sortIndicator(upsideSecKey)}</button></th>
      <th><button type="button" class="sort-header" data-sort="${upsideUnsecKey}">Unsecured ($MM)${sortIndicator(upsideUnsecKey)}</button></th>
      <th><button type="button" class="sort-header" data-sort="${upsidePrefKey}">Preferred ($MM)${sortIndicator(upsidePrefKey)}</button></th>
      <th><button type="button" class="sort-header" data-sort="Price">Price<sup>2</sup>${sortIndicator("Price")}</button></th>
      <th><button type="button" class="sort-header" data-sort="Yield">Yield<sup>3</sup>${sortIndicator("Yield")}</button></th>
      <th><button type="button" class="sort-header" data-sort="${moveKey}">${moveLabel}${sortIndicator(moveKey)}</button></th>
      <th>Primary</th>
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

  issuerBody.innerHTML = state.filteredIssuers.map((row) => {
    const selected = state.selectedIssuer === row.PARENT_TICKER ? "selected" : "";
    const cells = tableColumnOrder.map((column) => {
      const value = row[column];
      if (column === "Issuer") {
        const marker = row.REPORT_SENTIMENT_COLOR
          ? `<sup class="sentiment-marker ${row.REPORT_SENTIMENT_COLOR}" title="${row.REPORT_SENTIMENT_LABEL || "Report sentiment"}"></sup>`
          : "";
        const defaultedMarker = row.HAS_DEFAULTED ? `<sup class="defaulted-marker" title="Defaulted">D</sup>` : "";
        return `<td><button class="table-button" data-issuer="${row.PARENT_TICKER}">${fmt(value, 2)}${defaultedMarker}${marker}</button></td>`;
      }
      if (column === "3M Price Move" || column === "7D Price Move") return `<td>${fmtPriceMove(value)}</td>`;
      return `<td>${fmtIssuer(column, value)}</td>`;
    }).join("");
    return `<tr class="${selected}">${cells}</tr>`;
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
      if (state.sortField === column) {
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
  const moversActive = tabName === "movers";
  const exceptionsActive = tabName === "exceptions";
  const documentationActive = tabName === "documentation";
  issuerTabButton.classList.toggle("active", issuerActive);
  moversTabButton.classList.toggle("active", moversActive);
  exceptionsTabButton.classList.toggle("active", exceptionsActive);
  documentationTabButton.classList.toggle("active", documentationActive);
  issuerTabPanel.classList.toggle("hidden", !issuerActive);
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

  rows.sort((left, right) => {
    const a = left[sortBy];
    const b = right[sortBy];
    const aNum = typeof a === "number" ? a : Number(a);
    const bNum = typeof b === "number" ? b : Number(b);
    const numeric = !Number.isNaN(aNum) && !Number.isNaN(bNum);
    const result = numeric ? aNum - bNum : String(a ?? "").localeCompare(String(b ?? ""));
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
    reportSummaryCard.classList.add("hidden");
  } else if (state.selectedIssuer) {
    renderIssuerDetail(state.selectedIssuer);
    return;
  }
  renderIssuerTable();
}

function renderIssuerDetail(parentTicker) {
  const anchorDate = state.metadata.anchor_date ? new Date(state.metadata.anchor_date) : null;
  const minMaturity = anchorDate ? new Date(anchorDate) : null;
  if (minMaturity) minMaturity.setFullYear(minMaturity.getFullYear() + 1);
  const rows = (state.instrumentMap.get(parentTicker) || []).filter((row) => {
    if (row._IS_DEFAULTED) return false;
    const rowYield = Number(row.YIELD);
    const rowAmount = Number(row.AMT_OUTSTANDING_MM);
    const rankText = String(row.PAYMENT_RANK || "").toLowerCase();
    const isPreferred = rankText === "preferred";
    const rowPx = Number(row.PX_MID);
    const maturity = row.MATURITY ? new Date(row.MATURITY) : null;
    const maturityEligible = !minMaturity || (maturity instanceof Date && !Number.isNaN(maturity.getTime()) && maturity > minMaturity);
    if (isPreferred) return rowYield >= 7 && rowAmount >= 200 && rowPx < 90 && maturityEligible;
    return rowYield >= 7 && rowAmount >= 200 && rowPx < 90 && maturityEligible && !rankText.includes("subordinated");
  });
  rows.sort((a, b) => {
    const aSecured = String(a.PAYMENT_RANK || "").toLowerCase().includes("secured") && !String(a.PAYMENT_RANK || "").toLowerCase().includes("unsecured");
    const bSecured = String(b.PAYMENT_RANK || "").toLowerCase().includes("secured") && !String(b.PAYMENT_RANK || "").toLowerCase().includes("unsecured");
    if (aSecured !== bSecured) return aSecured ? -1 : 1;
    return (Number(b.AMT_OUTSTANDING_MM) || 0) - (Number(a.AMT_OUTSTANDING_MM) || 0);
  });
  state.detailRows = rows;
  const issuer = state.issuers.find((row) => row.PARENT_TICKER === parentTicker);
  state.selectedIssuer = parentTicker;
  detailCard.classList.remove("hidden");
  detailTicker.textContent = parentTicker;
  detailTitle.textContent = issuer?.Issuer || parentTicker;
  if (issuer?.REPORT_SUMMARY_BULLETS?.length) {
    reportSummaryCard.classList.remove("hidden");
    reportMeta.textContent = issuer.REPORT_DATE ? `Report date: ${issuer.REPORT_DATE}` : "";
    reportBullets.innerHTML = issuer.REPORT_SUMMARY_BULLETS.map((bullet) => `<li>${bullet}</li>`).join("");
  } else {
    reportSummaryCard.classList.add("hidden");
    reportMeta.textContent = "";
    reportBullets.innerHTML = "";
  }
  detailHead.innerHTML = `<tr>${detailColumns.map((column) => {
    const sup = column.key === "LAST_30D_VOLUME_MM" ? "<sup>7</sup>" : "";
    return `<th>${column.label}${sup}</th>`;
  }).join("")}</tr>`;
  detailBody.innerHTML = rows.map((row) => {
    const cells = detailColumns.map((column) => {
      const value = row[column.key];
      if (column.key === "NAME") {
        const dMarker = row._IS_DEFAULTED ? `<sup class="defaulted-marker" title="Defaulted">D</sup>` : "";
        const bbgId = row.ID || "";
        return `<td title="${bbgId}">${fmt(value, 2)}${dMarker}</td>`;
      }
      if (column.key === "PRICE_MOVE_3M") return `<td>${renderPriceMove(row)}</td>`;
      if (column.key === "PRICE_MOVE_7D") return `<td>${renderPriceMove7D(row)}</td>`;
      if (column.key === "PRICE_RANGE") return `<td>${renderPriceRange(row)}</td>`;
      const className = column.key === "AMT_OUTSTANDING_MM" ? "detail-narrow" : "";
      return `<td class="${className}">${fmt(value, detailDigits(column.key))}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");

  // Defaulted bonds table
  const defaultedRows = (state.instrumentMap.get(parentTicker) || []).filter((row) => row._IS_DEFAULTED);
  if (defaultedRows.length > 0) {
    defaultedCard.classList.remove("hidden");
    defaultedHead.innerHTML = `<tr>${detailColumns.map((column) => {
      const sup = column.key === "LAST_30D_VOLUME_MM" ? "<sup>7</sup>" : "";
      return `<th>${column.label}${sup}</th>`;
    }).join("")}</tr>`;
    defaultedBody.innerHTML = defaultedRows.map((row) => {
      const cells = detailColumns.map((column) => {
        const value = row[column.key];
        if (column.key === "NAME") {
          const bbgId = row.ID || "";
          return `<td title="${bbgId}">${fmt(value, 2)}<sup class="defaulted-marker" title="Defaulted">D</sup></td>`;
        }
        if (column.key === "PRICE_MOVE_3M") return `<td>${renderPriceMove(row)}</td>`;
        if (column.key === "PRICE_MOVE_7D") return `<td>${renderPriceMove7D(row)}</td>`;
        if (column.key === "PRICE_RANGE") return `<td>${renderPriceRange(row)}</td>`;
        const className = column.key === "AMT_OUTSTANDING_MM" ? "detail-narrow" : "";
        return `<td class="${className}">${fmt(value, detailDigits(column.key))}</td>`;
      }).join("");
      return `<tr>${cells}</tr>`;
    }).join("");
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
  const [dashboardPayload, moversPayload, abnormalPricesPayload, excludedPayload] = await Promise.all([
    fetchJson("/api/dashboard"),
    fetchJson("/api/price-movers"),
    fetchJson("/api/abnormal-prices"),
    fetchJson("/api/excluded"),
  ]);
  state.issuers = dashboardPayload.issuers;
  state.filters = dashboardPayload.filters;
  state.metadata = dashboardPayload.metadata;
  state.selectedIssuer = null;
  state.instrumentMap = new Map();
  state.detailRows = [];
  detailCard.classList.add("hidden");
  reportSummaryCard.classList.add("hidden");
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
  applyFilters();
  applyMoversFilters();
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

downloadIssuersButton.addEventListener("click", async () => {
  await downloadExcel("/api/export/issuers", { rows: state.filteredIssuers });
});

downloadDetailButton.addEventListener("click", async () => {
  if (!state.selectedIssuer) return;
  await downloadExcel("/api/export/detail", { issuer: state.selectedIssuer, rows: state.detailRows });
});

document.getElementById("glossaryButton").addEventListener("click", () => {
  const existing = document.querySelector(".glossary-popup");
  if (existing) { existing.remove(); return; }
  const popup = document.createElement("div");
  popup.className = "glossary-popup";
  popup.innerHTML = `<table>${glossaryEntries.map((e) =>
    `<tr><td><sup>${e.sup}</sup></td><td><strong>${e.label}</strong></td><td>${e.def}</td></tr>`
  ).join("")}</table>`;
  document.getElementById("glossaryButton").parentElement.style.position = "relative";
  document.getElementById("glossaryButton").parentElement.appendChild(popup);
  setTimeout(() => document.addEventListener("click", (ev) => {
    if (!popup.contains(ev.target) && ev.target.id !== "glossaryButton") popup.remove();
  }, { once: true }), 0);
});

checkSession();
