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
const reportSummaryCard = document.getElementById("reportSummaryCard");
const reportMeta = document.getElementById("reportMeta");
const reportBullets = document.getElementById("reportBullets");
const reportSentimentBadge = document.getElementById("reportSentimentBadge");
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

const issuerColumns = ["Issuer", "Sector", "Report Sentiment", "Face ($BN)", "WA Price", "Dist to Par (pts)", "Par Upside ($MM)", "52W Peak Upside ($MM)", "Secured (%)", "Seniority Basis (pts)", "Gap Signal", "<1Y", "1-3Y", "3-5Y", "Nearest Maturity", "OAS Delta (bps)", "# Tranches"];
const detailColumns = ["ID", "NAME", "PAYMENT_RANK", "MATURITY", "AMT_OUTSTANDING_MM", "PX_MID", "PX_MID_T90", "DIST_TO_PAR", "DISLOCATION_MM", "PX_HIGH_52W", "PX_LOW_52W", "OAS", "OAS_DELTA", "DISTRESS_TIER"];

function fmt(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "number") return value.toLocaleString(undefined, { maximumFractionDigits: digits, minimumFractionDigits: digits });
  return value;
}

function pillClass(value) {
  if (!value) return "pill";
  if (["Deep Distressed", "Distressed", "Severe", "Recovery Cliff"].includes(value)) return "pill bad";
  if (["Stressed", "Elevated"].includes(value)) return "pill warn";
  if (["Premium", "Normal", "Par-Adjacent"].includes(value)) return "pill good";
  return "pill";
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
  return sortDirection.value === "asc" ? " ▲" : " ▼";
}

function renderIssuerTable() {
  issuerHead.innerHTML = `<tr>${issuerColumns.map((column) => `<th><button type="button" class="sort-header" data-sort="${column}">${column}${sortIndicator(column)}</button></th>`).join("")}</tr>`;
  issuerBody.innerHTML = state.filteredIssuers.map((row) => {
    const selected = state.selectedIssuer === row.PARENT_TICKER ? "selected" : "";
    const cells = issuerColumns.map((column) => {
      const value = row[column];
      if (column === "Issuer") {
        return `<td><button class="table-button" data-issuer="${row.PARENT_TICKER}">${fmt(value, 2)}</button></td>`;
      }
      if (column === "Report Sentiment") {
        if (row.REPORT_SENTIMENT_SCORE === null || row.REPORT_SENTIMENT_SCORE === undefined) return `<td>—</td>`;
        return `<td><span class="pill ${row.REPORT_SENTIMENT_COLOR || ""}">${row.REPORT_SENTIMENT_SCORE} ${row.REPORT_SENTIMENT_LABEL || ""}</span></td>`;
      }
      if (column === "Gap Signal") {
        return `<td><span class="${pillClass(value)}">${fmt(value, 0)}</span></td>`;
      }
      return `<td>${fmt(value)}</td>`;
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
  if (sector !== "All") rows = rows.filter((row) => row["Sector"] === sector);
  if (search) {
    rows = rows.filter((row) =>
      String(row["Issuer"] || "").toLowerCase().includes(search) ||
      String(row["PARENT_TICKER"] || "").toLowerCase().includes(search)
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
  const rows = state.instrumentMap.get(parentTicker) || [];
  const issuer = state.issuers.find((row) => row.PARENT_TICKER === parentTicker);
  state.selectedIssuer = parentTicker;
  detailCard.classList.remove("hidden");
  detailTicker.textContent = parentTicker;
  detailTitle.textContent = `${issuer?.Issuer || parentTicker} capital stack`;
  if (issuer?.REPORT_SUMMARY_BULLETS?.length) {
    reportSummaryCard.classList.remove("hidden");
    reportMeta.textContent = `Report date: ${issuer.REPORT_DATE || "n/a"}${issuer.REPORT_SOURCE_FILE ? ` | ${issuer.REPORT_SOURCE_FILE.split(/[\\\\/]/).pop()}` : ""}`;
    reportSentimentBadge.innerHTML = `<span class="pill ${issuer.REPORT_SENTIMENT_COLOR || ""}">${issuer.REPORT_SENTIMENT_SCORE} ${issuer.REPORT_SENTIMENT_LABEL || ""}</span>`;
    reportBullets.innerHTML = issuer.REPORT_SUMMARY_BULLETS.map((bullet) => `<li>${bullet}</li>`).join("");
  } else {
    reportSummaryCard.classList.add("hidden");
    reportMeta.textContent = "";
    reportSentimentBadge.innerHTML = "";
    reportBullets.innerHTML = "";
  }
  detailHead.innerHTML = `<tr>${detailColumns.map((column) => `<th>${column}</th>`).join("")}</tr>`;
  detailBody.innerHTML = rows.map((row) => {
    const cells = detailColumns.map((column) => {
      const value = row[column];
      if (column === "DISTRESS_TIER") return `<td><span class="${pillClass(value)}">${fmt(value, 0)}</span></td>`;
      return `<td>${fmt(value)}</td>`;
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
  reportSummaryCard.classList.add("hidden");

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
