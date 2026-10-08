// report.js — front end for grouped reports (e.g. Food Pipeline's daily delivery
// report). Generic: the report is picked by <main data-report-id="...">, and the
// section names/colours come from GET /api/reports, so nothing here is specific
// to Food Pipeline.
//
// Flow: POST /api/upload -> POST /api/reports/{id}/generate -> show the summary.
// Everything that came from the server or the spreadsheet (client names,
// warnings) is inserted with textContent, never innerHTML, because client names
// are free text and could contain "<" or "&".

const API_BASE = "";
const REPORT_ID = document.querySelector("[data-report-id]").dataset.reportId;

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const dropzoneFilename = document.getElementById("dropzone-filename");
const createBtn = document.getElementById("create-btn");
const statusEl = document.getElementById("status");
const panelUpload = document.getElementById("panel-upload");
const panelResults = document.getElementById("panel-results");

let selectedFile = null;
let sections = [];            // [{key, title, title_color}] in report order

// ---- status helpers (same look as the rest of the app) ----

function setLoading(msg) { statusEl.className = "status is-loading"; statusEl.textContent = msg; }
function showError(msg) { statusEl.className = "status is-error"; statusEl.textContent = msg; }
function clearStatus() { statusEl.className = "status"; statusEl.textContent = ""; }

// FastAPI sends detail as a string for our errors, but as a list for request-validation errors.
function errorText(data, fallback) {
  return data && typeof data.detail === "string" ? data.detail : fallback;
}

// ---- section names/colours, so the table headings match the PDF ----

async function loadReportInfo() {
  try {
    const res = await fetch(`${API_BASE}/api/reports`);
    const list = await res.json();
    const info = list.find((r) => r.id === REPORT_ID);
    if (info) sections = info.sections;
  } catch (err) {
    // Not fatal: the results table falls back to the section keys the server returns.
  }
}
loadReportInfo();

// ---- file selection (same behaviour as the veggie upload page) ----

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fileInput.click(); }
});
["dragenter", "dragover"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.add("is-dragover"); })
);
["dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => { e.preventDefault(); dropzone.classList.remove("is-dragover"); })
);
dropzone.addEventListener("drop", (e) => {
  const file = e.dataTransfer.files[0];
  if (file) handleFileSelected(file);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) handleFileSelected(fileInput.files[0]);
});

function handleFileSelected(file) {
  clearStatus();
  // Quick name check only. The server inspects the real contents (the TMS export
  // is named .xls but is really .xlsx inside).
  if (!/\.xlsx?$/i.test(file.name)) {
    showError(`"${file.name}" isn't an Excel file. Please choose the .xlsx / .xls export from the TMS.`);
    selectedFile = null;
    createBtn.disabled = true;
    dropzoneFilename.textContent = "";
    return;
  }
  selectedFile = file;
  dropzoneFilename.textContent = file.name;
  createBtn.disabled = false;
}

// ---- create the reports ----

createBtn.addEventListener("click", () => createReports());

async function createReports() {
  if (!selectedFile) return;
  createBtn.disabled = true;
  try {
    setLoading("Reading your file…");
    const upForm = new FormData();
    upForm.append("file", selectedFile);
    const upRes = await fetch(`${API_BASE}/api/upload`, { method: "POST", body: upForm });
    const upData = await upRes.json();
    if (!upRes.ok) { showError(errorText(upData, "Something went wrong reading that file.")); return; }

    setLoading("Building one report per client…");
    const genForm = new FormData();
    genForm.append("session_id", upData.session_id);
    const genRes = await fetch(`${API_BASE}/api/reports/${encodeURIComponent(REPORT_ID)}/generate`,
                               { method: "POST", body: genForm });
    const summary = await genRes.json();
    if (!genRes.ok) { showError(errorText(summary, "Something went wrong building the reports.")); return; }

    clearStatus();
    renderResults(summary);
  } catch (err) {
    showError("Couldn't reach the server. Is the backend running?");
  } finally {
    createBtn.disabled = !selectedFile;
  }
}

// ---- results ----

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function renderResults(summary) {
  document.getElementById("results-title").textContent =
    summary.report_date ? `Reports for ${summary.report_date}` : "Reports";
  document.getElementById("results-subtitle").textContent =
    summary.client_count === 0
      ? "No reports were created."
      : `${summary.client_count} client${summary.client_count === 1 ? "" : "s"}. Check the details below, then download.`;

  // Warnings: things worth a look before anything is sent to a client.
  const warnBox = document.getElementById("warnings");
  warnBox.replaceChildren();
  warnBox.hidden = summary.warnings.length === 0;
  if (summary.warnings.length) {
    warnBox.append(el("p", "warning-box__title",
      summary.client_count === 0 ? "Why nothing was created" : "Please review before sending"));
    const ul = el("ul");
    summary.warnings.forEach((w) => ul.appendChild(el("li", null, w)));
    warnBox.appendChild(ul);
  }

  // Table: one row per client, one count column per report section.
  const keys = sections.length ? sections.map((s) => s.key)
                               : Object.keys((summary.clients[0] || {}).counts || {});
  const meta = (key) => sections.find((s) => s.key === key) || { key, title: key, title_color: null };

  const table = document.getElementById("results-table");
  table.replaceChildren();
  document.getElementById("table-wrap").hidden = summary.client_count === 0;
  if (summary.client_count) {
    const head = el("tr");
    head.appendChild(el("th", null, "Client"));
    keys.forEach((k) => head.appendChild(el("th", "num", meta(k).title)));
    head.appendChild(el("th", "num", "PDF"));
    table.appendChild(el("thead")).appendChild(head);

    const body = el("tbody");
    summary.clients.forEach((c) => {
      const tr = el("tr");
      tr.appendChild(el("td", null, c.name));
      keys.forEach((k) => {
        const n = c.counts[k] || 0;
        const td = el("td", "num " + (n === 0 ? "count--zero" : "count--alert"), String(n));
        // A non-zero count takes the section's own colour (e.g. red for failed), so problems stand out.
        if (n > 0 && meta(k).title_color) td.style.color = meta(k).title_color;
        else if (n > 0) td.className = "num";
        tr.appendChild(td);
      });
      const link = el("a", "report-table__link", "Download");
      link.href = c.download_url;
      link.setAttribute("download", c.filename);
      tr.appendChild(el("td", "num")).appendChild(link);
      body.appendChild(tr);
    });
    table.appendChild(body);
  }

  const zip = document.getElementById("zip-link");
  zip.hidden = !summary.zip_url;
  if (summary.zip_url) zip.href = summary.zip_url;

  panelUpload.hidden = true;
  panelResults.hidden = false;
  panelResults.scrollIntoView({ behavior: "smooth", block: "start" });
}

document.getElementById("again-btn").addEventListener("click", () => {
  selectedFile = null;
  fileInput.value = "";
  dropzoneFilename.textContent = "";
  createBtn.disabled = true;
  clearStatus();
  panelResults.hidden = true;
  panelUpload.hidden = false;
});
