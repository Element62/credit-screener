const state = { issuers: [], filteredIssuers: [], filters: {}, metadata: {}, selectedIssuer: null, instrumentMap: new Map(), detailRows: [] };

const loginCard = document.getElementById("loginCard");
const app = document.getElementById("app");
const loginForm = document.getElementById("loginForm");
const loginError = document.getElementById("loginError");
const logoutButton = document.getElementById("logoutButton");
const issuerHead = document.getElementById("issuerHead");
const issuerBody = document.getElementById("issuerBody");
const detailCard = document.getElementById("detailCard");
const detailTicker = document.getElementById("detailTicker");
const detailTitle = document.getElementById("detailTitle");
const detailHead = document.getElementById("detailHead");
const detailBody = document.getElementById("detailBody");
const reportSummaryCard = document.getElementById("reportSummaryCard");
const reportMeta = document.getElementById("reportMeta");
const reportBullets = document.getElementById("reportBullets");
const detailMinYieldInput = document.getElementById("detailMinYieldInput");
const statusBar = document.getElementById("statusBar");
const searchInput = document.getElementById("searchInput");
const clearSearchButton = document.getElementById("clearSearchButton");
const sectorFilter = document.getElementById("sectorFilter");
const issuerMinYieldInput = document.getElementById("issuerMinYieldInput");
const sortField = document.getElementById("sortField");
const sortDirection = document.getElementById("sortDirection");
const downloadIssuersButton = document.getElementById("downloadIssuersButton");
const downloadDetailButton = document.getElementById("downloadDetailButton");

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
  const current = Number(row.PX_MID);
  const prior = Number(row.PX_MID_T90);
  if (Number.isNaN(current) || Number.isNaN(prior)) return "-";
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

function renderIssuerTable() {
  const leadingColumns = ["Issuer", "Sector", "Face ($BN)", "WA Price", "WA Yield"];
  const trailingColumns = ["Secured (%)", "Seniority Basis (pts)", "<1Y", "1-3Y", "3-5Y", "Nearest Maturity", "# Tranches"];
  const tableColumnOrder = [...leadingColumns, "52W PEAK UPSIDE SECURED ($MM)", "52W PEAK UPSIDE UNSECURED ($MM)", ...trailingColumns];

  issuerHead.innerHTML = `
    <tr>
      ${leadingColumns.map((column) => `<th rowspan="2"><button type="button" class="sort-header" data-sort="${column}">${column}${sortIndicator(column)}</button></th>`).join("")}
      <th colspan="2" class="group-header">52W Peak Upside</th>
      ${trailingColumns.map((column) => `<th rowspan="2"><button type="button" class="sort-header" data-sort="${column}">${column}${sortIndicator(column)}</button></th>`).join("")}
    </tr>
    <tr>
      <th><button type="button" class="sort-header" data-sort="52W PEAK UPSIDE SECURED ($MM)">Secured ($MM)${sortIndicator("52W PEAK UPSIDE SECURED ($MM)")}</button></th>
      <th><button type="button" class="sort-header" data-sort="52W PEAK UPSIDE UNSECURED ($MM)">Unsecured ($MM)${sortIndicator("52W PEAK UPSIDE UNSECURED ($MM)")}</button></th>
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
        uploadStatus.textContent = error.message;
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
  renderIssuerTable();
}

function renderIssuerDetail(parentTicker) {
  const minYield = Number(detailMinYieldInput.value || 0);
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
  const payload = await fetchJson("/api/dashboard");
  state.issuers = payload.issuers;
  state.filters = payload.filters;
  state.metadata = payload.metadata;
  state.selectedIssuer = null;
  state.instrumentMap = new Map();
  state.detailRows = [];
  detailCard.classList.add("hidden");
  reportSummaryCard.classList.add("hidden");

  sectorFilter.innerHTML = `<option value="All">All</option>${payload.filters.sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("")}`;

  applyFilters();
}

async function checkSession() {
  const payload = await fetchJson("/api/session");
  if (!payload.authenticated) {
    loginCard.classList.remove("hidden");
    app.classList.add("hidden");
    logoutButton.classList.add("hidden");
    return;
  }
  loginCard.classList.add("hidden");
  app.classList.remove("hidden");
  logoutButton.classList.remove("hidden");
  await loadDashboard();
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginError.classList.add("hidden");
  const formData = new FormData(loginForm);
  try {
    await fetchJson("/login", { method: "POST", body: formData });
    await checkSession();
  } catch (error) {
    loginError.textContent = error.message;
    loginError.classList.remove("hidden");
  }
});

logoutButton.addEventListener("click", async () => {
  await fetchJson("/logout", { method: "POST" });
  window.location.reload();
});

[searchInput, sectorFilter, issuerMinYieldInput, sortField, sortDirection].forEach((element) => {
  element.addEventListener("input", applyFilters);
  element.addEventListener("change", applyFilters);
});

clearSearchButton.addEventListener("click", () => {
  searchInput.value = "";
  applyFilters();
  searchInput.focus();
});

[detailMinYieldInput].forEach((element) => {
  element.addEventListener("input", () => {
    if (state.selectedIssuer) renderIssuerDetail(state.selectedIssuer);
  });
  element.addEventListener("change", () => {
    if (state.selectedIssuer) renderIssuerDetail(state.selectedIssuer);
  });
});

downloadIssuersButton.addEventListener("click", async () => {
  await downloadExcel("/api/export/issuers", { rows: state.filteredIssuers });
});

downloadDetailButton.addEventListener("click", async () => {
  if (!state.selectedIssuer) return;
  await downloadExcel("/api/export/detail", { issuer: state.selectedIssuer, rows: state.detailRows });
});

checkSession();
