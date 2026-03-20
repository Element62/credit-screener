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
};

const loginCard = document.getElementById("loginCard");
const app = document.getElementById("app");
const loginForm = document.getElementById("loginForm");
const loginError = document.getElementById("loginError");
const logoutButton = document.getElementById("logoutButton");
const loadingBanner = document.getElementById("loadingBanner");
const issuerTabButton = document.getElementById("issuerTabButton");
const moversTabButton = document.getElementById("moversTabButton");
const issuerTabPanel = document.getElementById("issuerTabPanel");
const moversTabPanel = document.getElementById("moversTabPanel");
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
const detailCard = document.getElementById("detailCard");
const detailTicker = document.getElementById("detailTicker");
const detailTitle = document.getElementById("detailTitle");
const detailHead = document.getElementById("detailHead");
const detailBody = document.getElementById("detailBody");
const reportSummaryCard = document.getElementById("reportSummaryCard");
const reportMeta = document.getElementById("reportMeta");
const reportBullets = document.getElementById("reportBullets");
const statusBar = document.getElementById("statusBar");
const searchInput = document.getElementById("searchInput");
const clearSearchButton = document.getElementById("clearSearchButton");
const sectorFilter = document.getElementById("sectorFilter");
const issuerMinYieldInput = document.getElementById("issuerMinYieldInput");
const sortField = document.getElementById("sortField");
const sortDirection = document.getElementById("sortDirection");
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
  "Face ($BN)",
  "WA Price",
  "WA Yield",
  "52W PEAK UPSIDE SECURED ($MM)",
  "52W PEAK UPSIDE UNSECURED ($MM)",
  "Secured (%)",
  "Seniority Basis (pts)",
  "<1Y",
  "1-3Y",
  "3-5Y",
  "Nearest Maturity",
  "# Tranches",
  "COVERAGE PRIMARY",
  "COVERAGE SECONDARY",
];

const detailColumns = [
  { key: "ID", label: "ID" },
  { key: "NAME", label: "Name" },
  { key: "PAYMENT_RANK", label: "Rank" },
  { key: "MATURITY", label: "Mat" },
  { key: "AMT_OUTSTANDING_MM", label: "Amt Out ($MM)" },
  { key: "PX_MID", label: "Current PX" },
  { key: "PRICE_MOVE_3M", label: "3M Price Move" },
  { key: "YIELD", label: "Yield" },
  { key: "PRICE_RANGE", label: "Price Range" },
];

const moversColumns = [
  { key: "Issuer Name", label: "Issuer Name" },
  { key: "Security Name", label: "Security Name" },
  { key: "Sector", label: "Sector" },
  { key: "Rank", label: "Rank" },
  { key: "Mat", label: "Mat" },
  { key: "Amount Out ($MM)", label: "Amount Out ($MM)" },
  { key: "Current Px", label: "Current Px" },
  { key: "3M Price Move", label: "3M Price Move" },
  { key: "Price Range", label: "Price Range" },
];

function fmt(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "number") {
    return value.toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits });
  }
  return value;
}

function issuerDigits(column) {
  if (["Face ($BN)", "WA Price", "WA Yield", "Secured (%)"].includes(column)) return 1;
  if (["52W PEAK UPSIDE SECURED ($MM)", "52W PEAK UPSIDE UNSECURED ($MM)", "OAS Delta (bps)", "# Tranches"].includes(column)) return 0;
  return 2;
}

function detailDigits(column) {
  if (column === "AMT_OUTSTANDING_MM") return 0;
  if (["PX_MID", "PRICE_MOVE_3M", "YIELD", "PX_HIGH_52W", "PX_LOW_52W"].includes(column)) return 1;
  if (["OAS", "OAS_DELTA"].includes(column)) return 0;
  return 2;
}

function moversDigits(column) {
  if (column === "Amount Out ($MM)") return 0;
  if (["Current Px", "3M Price Move"].includes(column)) return 1;
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
  if (sortField.value !== column) return "";
  return sortDirection.value === "asc" ? " &uarr;" : " &darr;";
}

function renderPriceMove(row) {
  const currentRaw = row.PX_MID;
  const priorRaw = row.PX_MID_T90;
  if (currentRaw === null || currentRaw === undefined || currentRaw === "") return "N/A";
  if (priorRaw === null || priorRaw === undefined || priorRaw === "") return "N/A";
  const current = Number(currentRaw);
  const prior = Number(priorRaw);
  if (Number.isNaN(current) || Number.isNaN(prior)) return "N/A";
  const move = current - prior;
  const cssClass = move < 0 ? "negative" : move > 0 ? "positive" : "neutral";
  const prefix = move > 0 ? "+" : "";
  return `<span class="price-move ${cssClass}">${prefix}${fmt(move, 1)}</span>`;
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

function renderIssuerTable() {
  const leadingColumns = ["Issuer", "Sector", "Face ($BN)", "WA Price", "WA Yield"];
  const trailingColumns = ["Secured (%)", "Seniority Basis (pts)", "<1Y", "1-3Y", "3-5Y", "Nearest Maturity", "# Tranches"];
  const tableColumnOrder = [
    ...leadingColumns,
    "52W PEAK UPSIDE SECURED ($MM)",
    "52W PEAK UPSIDE UNSECURED ($MM)",
    ...trailingColumns,
    "COVERAGE PRIMARY",
    "COVERAGE SECONDARY",
  ];

  issuerHead.innerHTML = `
    <tr>
      ${leadingColumns.map((column) => `<th rowspan="2"><button type="button" class="sort-header" data-sort="${column}">${column}${sortIndicator(column)}</button></th>`).join("")}
      <th colspan="2" class="group-header">52W Peak Upside</th>
      ${trailingColumns.map((column) => `<th rowspan="2"><button type="button" class="sort-header" data-sort="${column}">${column}${sortIndicator(column)}</button></th>`).join("")}
      <th colspan="2" class="group-header">Coverage</th>
    </tr>
    <tr>
      <th><button type="button" class="sort-header" data-sort="52W PEAK UPSIDE SECURED ($MM)">Secured ($MM)${sortIndicator("52W PEAK UPSIDE SECURED ($MM)")}</button></th>
      <th><button type="button" class="sort-header" data-sort="52W PEAK UPSIDE UNSECURED ($MM)">Unsecured ($MM)${sortIndicator("52W PEAK UPSIDE UNSECURED ($MM)")}</button></th>
      <th><button type="button" class="sort-header" data-sort="COVERAGE PRIMARY">Primary${sortIndicator("COVERAGE PRIMARY")}</button></th>
      <th><button type="button" class="sort-header" data-sort="COVERAGE SECONDARY">Secondary${sortIndicator("COVERAGE SECONDARY")}</button></th>
    </tr>
  `;

  issuerBody.innerHTML = state.filteredIssuers.map((row) => {
    const selected = state.selectedIssuer === row.PARENT_TICKER ? "selected" : "";
    const cells = tableColumnOrder.map((column) => {
      const value = row[column];
      if (column === "Issuer") {
        const marker = row.REPORT_SENTIMENT_COLOR
          ? `<sup class="sentiment-marker ${row.REPORT_SENTIMENT_COLOR}" title="${row.REPORT_SENTIMENT_LABEL || "Report sentiment"}"></sup>`
          : "";
        return `<td><button class="table-button" data-issuer="${row.PARENT_TICKER}">${fmt(value, 2)}${marker}</button></td>`;
      }
      if (column === "Seniority Basis (pts)") {
        return `<td><span class="basis-value ${seniorityClass(value)}">${fmt(value, 1)}</span></td>`;
      }
      return `<td>${fmt(value, issuerDigits(column))}</td>`;
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
      if (sortField.value === column) {
        sortDirection.value = sortDirection.value === "asc" ? "desc" : "asc";
      } else {
        sortField.value = column;
        sortDirection.value = "desc";
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
    moversChart.innerHTML = `<div class="status-bar">No securities with absolute 3M price move above 10.</div>`;
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
  moversHead.innerHTML = `<tr>${moversColumns.map((column) => {
    if (["Current Px", "3M Price Move"].includes(column.key)) {
      const marker = moversSortField.value === column.key ? (moversSortDirection.value === "asc" ? " &uarr;" : " &darr;") : "";
      return `<th><button type="button" class="sort-header mover-sort-header" data-mover-sort="${column.key}">${column.label}${marker}</button></th>`;
    }
    return `<th>${column.label}</th>`;
  }).join("")}</tr>`;
  moversBody.innerHTML = state.moversRows.map((row) => {
    const cells = moversColumns.map((column) => {
      if (column.key === "3M Price Move") {
        const move = Number(row[column.key]);
        const cssClass = move < 0 ? "negative" : move > 0 ? "positive" : "neutral";
        const prefix = move > 0 ? "+" : "";
        return `<td><span class="price-move ${cssClass}">${prefix}${fmt(move, 1)}</span></td>`;
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

function applyMoversFilters() {
  const direction = moversDirectionFilter.value;
  const sortBy = moversSortField.value;
  const ascending = moversSortDirection.value === "asc";

  let rows = [...state.rawMoversRows];
  rows = rows.filter((row) => {
    const currentPx = Number(row["Current Px"]);
    return !Number.isNaN(currentPx) && currentPx <= 105;
  });
  if (direction === "down") rows = rows.filter((row) => Number(row["3M Price Move"]) <= -10);
  if (direction === "up") rows = rows.filter((row) => Number(row["3M Price Move"]) >= 10);

  rows.sort((left, right) => {
    const a = Number(left[sortBy]);
    const b = Number(right[sortBy]);
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

  const directionLabel = direction === "down" ? "price move down below -10" : direction === "up" ? "price move up above 10" : "absolute 3M price move above 10";
  moversStatus.textContent = `${state.moversRows.length} securities with ${directionLabel}`;
  renderMoversChart();
  renderMoversTable();
}

function setActiveTab(tabName) {
  const issuerActive = tabName === "issuer";
  issuerTabButton.classList.toggle("active", issuerActive);
  moversTabButton.classList.toggle("active", !issuerActive);
  issuerTabPanel.classList.toggle("hidden", !issuerActive);
  moversTabPanel.classList.toggle("hidden", issuerActive);
}

function applyFilters() {
  const search = searchInput.value.trim().toLowerCase();
  const sector = sectorFilter.value;
  const minIssuerYield = Number(issuerMinYieldInput.value || 0);
  const sortBy = sortField.value;
  const ascending = sortDirection.value === "asc";

  let rows = [...state.issuers];
  rows = rows.filter((row) => {
    const maxYield = Number(row["Max Yield"]);
    return !Number.isNaN(maxYield) && maxYield >= minIssuerYield;
  });
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
  statusBar.textContent = `Anchor: ${state.metadata.anchor_date || "n/a"} | T-90: ${state.metadata.t90_date || "n/a"} | ${state.issuers.length} total issuers | Showing ${rows.length}`;
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
  const minYield = Number(issuerMinYieldInput.value || 0);
  const anchorDate = state.metadata.anchor_date ? new Date(state.metadata.anchor_date) : null;
  const minMaturity = anchorDate ? new Date(anchorDate) : null;
  if (minMaturity) minMaturity.setFullYear(minMaturity.getFullYear() + 1);
  const rows = (state.instrumentMap.get(parentTicker) || []).filter((row) => {
    const rowYield = Number(row.YIELD);
    const rowAmount = Number(row.AMT_OUTSTANDING_MM);
    const maturity = row.MATURITY ? new Date(row.MATURITY) : null;
    const rankText = String(row.PAYMENT_RANK || "").toLowerCase();
    const maturityEligible = !minMaturity || (maturity instanceof Date && !Number.isNaN(maturity.getTime()) && maturity > minMaturity);
    return rowYield >= minYield && rowAmount >= 200 && maturityEligible && !rankText.includes("subordinated");
  });
  state.detailRows = rows;
  const issuer = state.issuers.find((row) => row.PARENT_TICKER === parentTicker);
  state.selectedIssuer = parentTicker;
  detailCard.classList.remove("hidden");
  detailTicker.textContent = parentTicker;
  detailTitle.textContent = `${issuer?.Issuer || parentTicker} capital stack`;
  if (issuer?.REPORT_SUMMARY_BULLETS?.length) {
    reportSummaryCard.classList.remove("hidden");
    reportMeta.textContent = issuer.REPORT_DATE ? `Report date: ${issuer.REPORT_DATE}` : "";
    reportBullets.innerHTML = issuer.REPORT_SUMMARY_BULLETS.map((bullet) => `<li>${bullet}</li>`).join("");
  } else {
    reportSummaryCard.classList.add("hidden");
    reportMeta.textContent = "";
    reportBullets.innerHTML = "";
  }
  detailHead.innerHTML = `<tr>${detailColumns.map((column) => `<th>${column.label}</th>`).join("")}</tr>`;
  detailBody.innerHTML = rows.map((row) => {
    const cells = detailColumns.map((column) => {
      const value = row[column.key];
      if (column.key === "PRICE_MOVE_3M") return `<td>${renderPriceMove(row)}</td>`;
      if (column.key === "PRICE_RANGE") return `<td>${renderPriceRange(row)}</td>`;
      const className = column.key === "AMT_OUTSTANDING_MM" ? "detail-narrow" : "";
      return `<td class="${className}">${fmt(value, detailDigits(column.key))}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
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
  const [dashboardPayload, moversPayload] = await Promise.all([
    fetchJson("/api/dashboard"),
    fetchJson("/api/price-movers"),
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
  applyFilters();
  applyMoversFilters();
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

[searchInput, sectorFilter, issuerMinYieldInput, sortField, sortDirection].forEach((element) => {
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

checkSession();
