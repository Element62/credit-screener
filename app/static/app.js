const state = { issuers: [], filteredIssuers: [], filters: {}, metadata: {}, selectedIssuer: null, instrumentMap: new Map() };

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
const detailMinYieldInput = document.getElementById("detailMinYieldInput");
const statusBar = document.getElementById("statusBar");
const searchInput = document.getElementById("searchInput");
const sectorFilter = document.getElementById("sectorFilter");
const distressFilter = document.getElementById("distressFilter");
const minFaceInput = document.getElementById("minFaceInput");
const sortField = document.getElementById("sortField");
const sortDirection = document.getElementById("sortDirection");
const uploadForm = document.getElementById("uploadForm");
const uploadStatus = document.getElementById("uploadStatus");
const processReportsButton = document.getElementById("processReportsButton");

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
    button.addEventListener("click", () => renderIssuerDetail(button.dataset.issuer));
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
  const distress = distressFilter.value;
  const minFace = Number(minFaceInput.value || 0);
  const sortBy = sortField.value;
  const ascending = sortDirection.value === "asc";

  let rows = [...state.issuers];
  rows = rows.filter((row) => Number(row["Face ($MM)"] || 0) >= minFace);
  if (sector !== "All") rows = rows.filter((row) => row.Sector === sector);
  if (search) {
    rows = rows.filter((row) =>
      String(row.Issuer || "").toLowerCase().includes(search) ||
      String(row.PARENT_TICKER || "").toLowerCase().includes(search)
    );
  }
  if (distress !== "All") {
    rows = rows.filter((row) => (row.__distressTiers || []).includes(distress));
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
  renderIssuerTable();
}

function renderIssuerDetail(parentTicker) {
  const minYield = Number(detailMinYieldInput.value || 0);
  const rows = (state.instrumentMap.get(parentTicker) || []).filter((row) => Number(row.YIELD) >= minYield);
  const issuer = state.issuers.find((row) => row.PARENT_TICKER === parentTicker);
  state.selectedIssuer = parentTicker;
  detailCard.classList.remove("hidden");
  detailTicker.textContent = parentTicker;
  detailTitle.textContent = `${issuer?.Issuer || parentTicker} capital stack`;
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

async function loadDashboard() {
  const payload = await fetchJson("/api/dashboard");
  state.issuers = payload.issuers;
  state.filters = payload.filters;
  state.metadata = payload.metadata;
  state.selectedIssuer = null;
  state.instrumentMap = new Map();
  detailCard.classList.add("hidden");

  sectorFilter.innerHTML = `<option value="All">All</option>${payload.filters.sectors.map((sector) => `<option value="${sector}">${sector}</option>`).join("")}`;
  distressFilter.innerHTML = `<option value="All">All</option>${payload.filters.distress_tiers.map((tier) => `<option value="${tier}">${tier}</option>`).join("")}`;

  const detailPayloads = await Promise.all(state.issuers.map((issuer) => fetchJson(`/api/issuers/${encodeURIComponent(issuer.PARENT_TICKER)}`)));
  detailPayloads.forEach((payload) => state.instrumentMap.set(payload.issuer, payload.rows));

  state.issuers = state.issuers.map((issuer) => ({
    ...issuer,
    __distressTiers: Array.from(new Set((state.instrumentMap.get(issuer.PARENT_TICKER) || []).map((row) => row.DISTRESS_TIER).filter(Boolean))),
  }));

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

[searchInput, sectorFilter, distressFilter, minFaceInput, sortField, sortDirection].forEach((element) => {
  element.addEventListener("input", applyFilters);
  element.addEventListener("change", applyFilters);
});

[detailMinYieldInput].forEach((element) => {
  element.addEventListener("input", () => {
    if (state.selectedIssuer) renderIssuerDetail(state.selectedIssuer);
  });
  element.addEventListener("change", () => {
    if (state.selectedIssuer) renderIssuerDetail(state.selectedIssuer);
  });
});

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  uploadStatus.textContent = "Uploading workbook...";
  const formData = new FormData(uploadForm);
  try {
    const payload = await fetchJson("/api/admin/upload", { method: "POST", body: formData });
    uploadStatus.textContent = `Workbook refreshed. Anchor date: ${payload.metadata.anchor_date || "n/a"}.`;
    await loadDashboard();
  } catch (error) {
    uploadStatus.textContent = error.message;
  }
});

processReportsButton.addEventListener("click", async () => {
  uploadStatus.textContent = "Processing latest reports...";
  try {
    const payload = await fetchJson("/api/admin/process-reports", { method: "POST" });
    uploadStatus.textContent = `Processed report cache refreshed. ${payload.processed_report_count} report record(s) available.`;
    await loadDashboard();
  } catch (error) {
    uploadStatus.textContent = error.message;
  }
});

checkSession();
