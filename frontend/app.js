const API_BASE = window.CODOCS_API_BASE || "http://127.0.0.1:8000";

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const dropzoneFilename = document.getElementById("dropzone-filename");
const headerRowInput = document.getElementById("header-row");
const detectBtn = document.getElementById("detect-btn");
const statusEl = document.getElementById("status");

const panelHeaders = document.getElementById("panel-headers");
const headersSummary = document.getElementById("headers-summary");
const headerChips = document.getElementById("header-chips");
const previewWrap = document.getElementById("preview-table-wrap");
const previewTable = document.getElementById("preview-table");
const wrongRowBtn = document.getElementById("wrong-row-btn");
const continueBtn = document.getElementById("continue-btn");

let selectedFile = null;
let currentSession = null; // { session_id, filename, headers }

// ---- File selection ----

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});

["dragenter", "dragover"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add("is-dragover");
  })
);
["dragleave", "drop"].forEach((evt) =>
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove("is-dragover");
  })
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
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showError(`"${file.name}" isn't an .xlsx file. Please choose an Excel file.`);
    selectedFile = null;
    detectBtn.disabled = true;
    dropzoneFilename.textContent = "";
    return;
  }
  selectedFile = file;
  dropzoneFilename.textContent = file.name;
  detectBtn.disabled = false;
  panelHeaders.hidden = true;
}

// ---- Header detection ----

detectBtn.addEventListener("click", () => detectHeaders());
wrongRowBtn.addEventListener("click", () => {
  headerRowInput.focus();
  headerRowInput.select();
  detectHeaders();
});

async function detectHeaders() {
  if (!selectedFile) return;

  const headerRow = parseInt(headerRowInput.value, 10) || 2;
  setLoading("Reading your file…");
  detectBtn.disabled = true;
  panelHeaders.hidden = true;

  try {
    const form = new FormData();
    form.append("file", selectedFile);
    form.append("header_row", headerRow);

    const res = await fetch(`${API_BASE}/api/upload`, { method: "POST", body: form });
    const data = await res.json();

    if (!res.ok) {
      showError(data.detail || "Something went wrong reading that file.");
      return;
    }

    currentSession = data;
    clearStatus();
    showHeaders(data);
  } catch (err) {
    showError("Couldn't reach the server. Is the backend running?");
  } finally {
    detectBtn.disabled = false;
  }
}

function showHeaders(data) {
  headersSummary.textContent =
    `Found ${data.headers.length} column${data.headers.length === 1 ? "" : "s"} in "${data.filename}".` +
    (data.headers.length <= 1
      ? " That seems low for a real spreadsheet — try a different header row below if this looks wrong."
      : " Check these match what you expect before continuing.");

  headerChips.innerHTML = "";
  data.headers.forEach((h) => {
    const li = document.createElement("li");
    li.textContent = h;
    headerChips.appendChild(li);
  });

  renderPreviewTable(data.headers, data.preview_rows);

  // Document-type selection (step 5) isn't built yet — leave Continue disabled
  // rather than dead-ending the user silently.
  continueBtn.disabled = true;

  panelHeaders.hidden = false;
  markStepDone(1);
  panelHeaders.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderPreviewTable(headers, rows) {
  if (!rows || rows.length === 0) {
    previewWrap.hidden = true;
    return;
  }
  const thead = `<thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows
    .map(
      (row) =>
        `<tr>${headers.map((h) => `<td>${escapeHtml(row[h] ?? "")}</td>`).join("")}</tr>`
    )
    .join("")}</tbody>`;
  previewTable.innerHTML = thead + tbody;
  previewWrap.hidden = false;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = String(str);
  return div.innerHTML;
}

// ---- Status / step helpers ----

function setLoading(msg) {
  statusEl.className = "status is-loading";
  statusEl.textContent = msg;
}
function showError(msg) {
  statusEl.className = "status is-error";
  statusEl.textContent = msg;
}
function clearStatus() {
  statusEl.className = "status";
  statusEl.textContent = "";
}
function markStepDone(stepNumber) {
  const item = document.querySelector(`.steps__item[data-step="${stepNumber}"]`);
  if (item) {
    item.classList.remove("is-active");
    item.classList.add("is-done");
  }
}
