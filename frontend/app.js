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

const customizeToggle = document.getElementById("customize-toggle");
const customizeBody = document.getElementById("customize-body");
const docTitleInput = document.getElementById("doc-title-input");
const customizeList = document.getElementById("customize-list");
const previewFrame = document.getElementById("preview-frame");
const previewFrameStatus = document.getElementById("preview-frame-status");

const WIDTH_PRESETS = { narrow: 0.6, medium: 1.0, wide: 1.5, xwide: 2.2 };
const WIDTH_PRESETS_MM = { narrow: 20, medium: 30, wide: 45, xwide: 60 };
const WIDTH_PRESET_LABELS = { narrow: "Narrow", medium: "Medium", wide: "Wide", xwide: "Extra wide" };
const WIDTH_PRESET_ORDER = ["narrow", "medium", "wide", "xwide"];

let selectedFile = null;
let currentSession = null; // { session_id, filename, headers, preview_rows }
let selectedConfig = null; // full DocumentConfig of the chosen preset
let columnMapping = {};    // { columnKey: excelHeader }
let customization = {};    // { columnKey: { label, widthPreset } }
let previewDebounceTimer = null;
let previewObjectUrl = null;
let previewRequestSeq = 0; // guards against an in-flight preview response arriving out of order

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
      schedulePreview();
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
  buildCustomizeSection(config);
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

// ---- Step 3b: customize headers & widths, with live preview ----

customizeToggle.addEventListener("click", () => {
  const isOpen = customizeToggle.getAttribute("aria-expanded") === "true";
  customizeToggle.setAttribute("aria-expanded", String(!isOpen));
  customizeBody.hidden = isOpen;
  if (!isOpen) schedulePreview(true); // opening the panel: show a preview right away
});

function closestWidthPreset(col) {
  const widthMode = col.width_mode || "fixed"; // schema default, in case an unnormalized config slips through
  const map = widthMode === "fixed" ? WIDTH_PRESETS_MM : WIDTH_PRESETS;
  const currentValue = widthMode === "fixed" ? col.width_mm : col.flex_weight;
  if (currentValue == null) return "medium";
  let best = "medium";
  let bestDiff = Infinity;
  for (const key of WIDTH_PRESET_ORDER) {
    const diff = Math.abs(map[key] - currentValue);
    if (diff < bestDiff) {
      bestDiff = diff;
      best = key;
    }
  }
  return best;
}

function buildCustomizeSection(config) {
  docTitleInput.value = config.document_title;
  docTitleInput.oninput = () => schedulePreview();

  customization = {};
  customizeList.innerHTML = "";

  config.columns.forEach((col) => {
    customization[col.key] = {
      label: col.label,
      widthPreset: col.type === "text" ? closestWidthPreset(col) : null,
    };

    const row = document.createElement("div");
    row.className = "customize-row";

    const labelInput = document.createElement("input");
    labelInput.type = "text";
    labelInput.className = "customize-row__label-input";
    labelInput.value = col.label;
    labelInput.setAttribute("aria-label", `Header text for ${col.label}`);
    labelInput.addEventListener("input", () => {
      customization[col.key].label = labelInput.value;
      schedulePreview();
    });
    row.appendChild(labelInput);

    if (col.type === "text") {
      const picker = document.createElement("div");
      picker.className = "width-picker";
      WIDTH_PRESET_ORDER.forEach((presetKey) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.textContent = WIDTH_PRESET_LABELS[presetKey];
        btn.classList.toggle("is-selected", customization[col.key].widthPreset === presetKey);
        btn.addEventListener("click", () => {
          customization[col.key].widthPreset = presetKey;
          picker.querySelectorAll("button").forEach((b) => b.classList.remove("is-selected"));
          btn.classList.add("is-selected");
          schedulePreview();
        });
        picker.appendChild(btn);
      });
      row.appendChild(picker);
    } else {
      const note = document.createElement("span");
      note.className = "customize-row__original";
      note.textContent = "Photo size is fixed";
      row.appendChild(note);
    }

    customizeList.appendChild(row);
  });
}

function buildFinalConfig(rowLimit) {
  const finalConfig = JSON.parse(JSON.stringify(selectedConfig));
  finalConfig.document_title = docTitleInput.value || selectedConfig.document_title;
  finalConfig.columns = finalConfig.columns.map((col) => {
    const custom = customization[col.key] || {};
    const updated = { ...col, label: custom.label || col.label };
    if (col.type === "image") return updated;

    updated.source_header = columnMapping[col.key] || null;
    if (custom.widthPreset) {
      const widthMode = col.width_mode || "fixed";
      if (widthMode === "fixed") {
        updated.width_mm = WIDTH_PRESETS_MM[custom.widthPreset];
      } else {
        updated.flex_weight = WIDTH_PRESETS[custom.widthPreset];
      }
    }
    return updated;
  });
  return finalConfig;
}

function schedulePreview(immediate) {
  if (customizeBody.hidden) return; // no point rendering a preview nobody can see yet
  clearTimeout(previewDebounceTimer);
  const requiredOk = !generateBtn.disabled;
  if (!requiredOk) {
    previewFrameStatus.textContent = "Map required fields to see a preview";
    return;
  }
  previewFrameStatus.textContent = "Updating…";
  previewDebounceTimer = setTimeout(runPreview, immediate ? 0 : 600);
}

async function runPreview() {
  if (!currentSession || !selectedConfig) return;
  const mySeq = ++previewRequestSeq;

  try {
    const form = new FormData();
    form.append("session_id", currentSession.session_id);
    form.append("config_json", JSON.stringify(buildFinalConfig()));
    form.append("row_limit", "3");

    const res = await fetch(`${API_BASE}/api/preview`, { method: "POST", body: form });

    if (mySeq !== previewRequestSeq) return; // a newer request superseded this one

    if (!res.ok) {
      let message = "Couldn't generate a preview.";
      try {
        const data = await res.json();
        message = data.detail || message;
      } catch (_) {}
      previewFrameStatus.textContent = message;
      return;
    }

    const blob = await res.blob();
    if (mySeq !== previewRequestSeq) return;

    if (previewObjectUrl) URL.revokeObjectURL(previewObjectUrl);
    previewObjectUrl = URL.createObjectURL(blob);
    previewFrame.src = previewObjectUrl;
    previewFrameStatus.textContent = "";
  } catch (err) {
    if (mySeq === previewRequestSeq) previewFrameStatus.textContent = "Couldn't reach the server.";
  }
}

// ---- Generate + download ----

generateBtn.addEventListener("click", () => generatePdf());

async function generatePdf() {
  if (!selectedConfig || !currentSession) return;

  const finalConfig = buildFinalConfig();

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
