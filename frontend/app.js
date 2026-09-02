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

const panelPreset = document.getElementById("panel-preset");
const presetList = document.getElementById("preset-list");
const presetStatus = document.getElementById("preset-status");

const panelMapping = document.getElementById("panel-mapping");
const mappingHint = document.getElementById("mapping-hint");
const mappingList = document.getElementById("mapping-list");
const generateBtn = document.getElementById("generate-btn");
const generateStatus = document.getElementById("generate-status");
const resultPanel = document.getElementById("result-panel");
const resultText = document.getElementById("result-text");
const downloadLink = document.getElementById("download-link");

let selectedFile = null;
let currentSession = null; // { session_id, filename, headers, preview_rows }
let selectedConfig = null; // full DocumentConfig of the chosen preset
let columnMapping = {};    // { columnKey: excelHeader }

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

// ---- Step 2: preset picker ----

continueBtn.addEventListener("click", () => loadPresets());

async function loadPresets() {
  markStepActive(2);
  presetStatus.className = "status is-loading";
  presetStatus.textContent = "Loading document types…";
  presetList.innerHTML = "";
  panelPreset.hidden = false;
  panelMapping.hidden = true;
  resultPanel.hidden = true;
  panelPreset.scrollIntoView({ behavior: "smooth", block: "start" });

  try {
    const res = await fetch(`${API_BASE}/api/presets`);
    const presets = await res.json();

    if (!res.ok) {
      presetStatus.className = "status is-error";
      presetStatus.textContent = "Couldn't load document types.";
      return;
    }
    if (presets.length === 0) {
      presetStatus.className = "status is-error";
      presetStatus.textContent = "No document types are set up yet.";
      return;
    }

    presetStatus.className = "status";
    presetStatus.textContent = "";
    presetList.innerHTML = "";
    presets.forEach((p) => {
      const li = document.createElement("li");
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "preset-card";
      btn.innerHTML = `
        <p class="preset-card__name">${escapeHtml(p.name)}</p>
        <p class="preset-card__meta">${escapeHtml(p.document_title)} · ${p.column_labels.map(escapeHtml).join(", ")}</p>
      `;
      btn.addEventListener("click", () => selectPreset(p.id, btn));
      li.appendChild(btn);
      presetList.appendChild(li);
    });
  } catch (err) {
    presetStatus.className = "status is-error";
    presetStatus.textContent = "Couldn't reach the server.";
  }
}

async function selectPreset(presetId, clickedBtn) {
  document.querySelectorAll(".preset-card").forEach((el) => el.classList.remove("is-selected"));
  clickedBtn.classList.add("is-selected");

  presetStatus.className = "status is-loading";
  presetStatus.textContent = "Loading template…";

  try {
    const res = await fetch(`${API_BASE}/api/presets/${presetId}`);
    const config = await res.json();
    if (!res.ok) {
      presetStatus.className = "status is-error";
      presetStatus.textContent = "Couldn't load that template.";
      return;
    }
    presetStatus.className = "status";
    presetStatus.textContent = "";
    selectedConfig = config;
    markStepDone(2);
    showMappingForm(config);
  } catch (err) {
    presetStatus.className = "status is-error";
    presetStatus.textContent = "Couldn't reach the server.";
  }
}

// ---- Step 3: column mapping ----

function showMappingForm(config) {
  markStepActive(3);
  mappingHint.textContent =
    `"${config.name}" — match each field to a column from your file. Fields matching your file's headers are pre-filled; check them and adjust anything that's wrong.`;

  columnMapping = {};
  mappingList.innerHTML = "";

  config.columns.forEach((col) => {
    if (col.type === "image") {
      const row = document.createElement("div");
      row.className = "mapping-row mapping-row--image";
      row.innerHTML = `
        <span class="mapping-row__label">${escapeHtml(col.label)}</span>
        <span class="mapping-row__note">Matched automatically from the file's embedded photos</span>
      `;
      mappingList.appendChild(row);
      return;
    }

    const bestMatch = findBestHeaderMatch(col.source_header, currentSession.headers);
    columnMapping[col.key] = bestMatch || "";

    const row = document.createElement("div");
    row.className = "mapping-row";

    const label = document.createElement("span");
    label.className = "mapping-row__label";
    label.innerHTML = escapeHtml(col.label) + (col.optional ? ` <span class="mapping-row__optional-tag">optional</span>` : "");

    const select = document.createElement("select");
    select.className = "mapping-row__select";
    select.dataset.columnKey = col.key;
    select.dataset.optional = col.optional ? "1" : "0";

    const blankOpt = document.createElement("option");
    blankOpt.value = "";
    blankOpt.textContent = "— not mapped —";
    select.appendChild(blankOpt);

    currentSession.headers.forEach((h) => {
      const opt = document.createElement("option");
      opt.value = h;
      opt.textContent = h;
      if (h === bestMatch) opt.selected = true;
      select.appendChild(opt);
    });

    select.addEventListener("change", () => {
      columnMapping[col.key] = select.value;
      updateSelectValidity(select);
      updateGenerateAvailability();
    });

    updateSelectValidity(select);

    row.appendChild(label);
    row.appendChild(select);
    mappingList.appendChild(row);
  });

  panelMapping.hidden = false;
  resultPanel.hidden = true;
  generateStatus.className = "status";
  generateStatus.textContent = "";
  updateGenerateAvailability();
  panelMapping.scrollIntoView({ behavior: "smooth", block: "start" });
}

function findBestHeaderMatch(suggestedHeader, actualHeaders) {
  if (!suggestedHeader) return null;
  const exact = actualHeaders.find((h) => h === suggestedHeader);
  if (exact) return exact;
  const loose = actualHeaders.find((h) => h.trim().toLowerCase() === suggestedHeader.trim().toLowerCase());
  return loose || null;
}

function updateSelectValidity(select) {
  const isOptional = select.dataset.optional === "1";
  select.classList.toggle("is-unmapped", !isOptional && !select.value);
}

function updateGenerateAvailability() {
  const selects = mappingList.querySelectorAll(".mapping-row__select");
  const allRequiredMapped = Array.from(selects).every(
    (s) => s.dataset.optional === "1" || s.value
  );
  generateBtn.disabled = !allRequiredMapped;
}

// ---- Generate + download ----

generateBtn.addEventListener("click", () => generatePdf());

async function generatePdf() {
  if (!selectedConfig || !currentSession) return;

  const finalConfig = JSON.parse(JSON.stringify(selectedConfig));
  finalConfig.columns = finalConfig.columns.map((col) => {
    if (col.type === "image") return col;
    return { ...col, source_header: columnMapping[col.key] || null };
  });

  generateBtn.disabled = true;
  generateStatus.className = "status is-loading";
  generateStatus.textContent = "Generating your PDF…";
  resultPanel.hidden = true;

  try {
    const form = new FormData();
    form.append("session_id", currentSession.session_id);
    form.append("config_json", JSON.stringify(finalConfig));

    const res = await fetch(`${API_BASE}/api/generate`, { method: "POST", body: form });
    const data = await res.json();

    if (!res.ok) {
      generateStatus.className = "status is-error";
      generateStatus.textContent = data.detail || "Something went wrong generating the PDF.";
      return;
    }

    generateStatus.className = "status is-success";
    generateStatus.textContent = "Done.";
    resultText.textContent = `${data.item_count} item${data.item_count === 1 ? "" : "s"} included.`;
    downloadLink.href = `${API_BASE}${data.download_url}`;
    resultPanel.hidden = false;
    markStepDone(3);
  } catch (err) {
    generateStatus.className = "status is-error";
    generateStatus.textContent = "Couldn't reach the server.";
  } finally {
    generateBtn.disabled = false;
  }
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
function markStepActive(stepNumber) {
  document.querySelectorAll(".steps__item").forEach((el) => el.classList.remove("is-active"));
  const item = document.querySelector(`.steps__item[data-step="${stepNumber}"]`);
  if (item) item.classList.add("is-active");
}
