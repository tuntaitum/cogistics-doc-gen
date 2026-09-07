const API_BASE = window.CODOCS_API_BASE || "http://127.0.0.1:8000";

// ---- Element refs ----

// Page 1: upload
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

// Page 2: document type
const backToUploadBtn = document.getElementById("back-to-upload");
const presetList = document.getElementById("preset-list");
const presetStatus = document.getElementById("preset-status");
const panelPresetPreview = document.getElementById("panel-preset-preview");
const presetPreviewFrame = document.getElementById("preset-preview-frame");
const presetPreviewStatus = document.getElementById("preset-preview-status");
const continueToMappingBtn = document.getElementById("continue-to-mapping-btn");

// Page 3: map, customize, generate
const backToPresetBtn = document.getElementById("back-to-preset");
const mappingReferenceTable = document.getElementById("mapping-reference-table");
const mappingHint = document.getElementById("mapping-hint");
const mappingList = document.getElementById("mapping-list");
const docTitleInput = document.getElementById("doc-title-input");
const customizeList = document.getElementById("customize-list");
const previewFrame = document.getElementById("preview-frame");
const previewFrameStatus = document.getElementById("preview-frame-status");
const filenameInput = document.getElementById("filename-input");
const generateBtn = document.getElementById("generate-btn");
const generateStatus = document.getElementById("generate-status");
const resultPanel = document.getElementById("result-panel");
const resultText = document.getElementById("result-text");
const downloadLink = document.getElementById("download-link");

const pages = {
  1: document.getElementById("page-upload"),
  2: document.getElementById("page-preset"),
  3: document.getElementById("page-mapping"),
};

const WIDTH_PRESETS = { narrow: 0.6, medium: 1.0, wide: 1.5, xwide: 2.2 };
const WIDTH_PRESETS_MM = { narrow: 20, medium: 30, wide: 45, xwide: 60 };
const WIDTH_PRESET_LABELS = { narrow: "Narrow", medium: "Medium", wide: "Wide", xwide: "Extra wide" };
const WIDTH_PRESET_ORDER = ["narrow", "medium", "wide", "xwide"];

let selectedFile = null;
let currentSession = null; // { session_id, filename, headers, preview_rows }
let selectedConfig = null; // full DocumentConfig of the chosen preset
let columnMapping = {};    // { columnKey: excelHeader } — computed once on preset select, edited on page 3
let customization = {};    // { columnKey: { label, widthPreset } }
let previewDebounceTimer = null;
let previewObjectUrl = null;
let previewRequestSeq = 0;
let presetPreviewObjectUrl = null;
let presetPreviewSeq = 0;

// ---- Page navigation ----

function goToPage(n) {
  Object.entries(pages).forEach(([num, el]) => {
    el.hidden = Number(num) !== n;
  });
  document.querySelectorAll(".steps__item").forEach((el) => {
    const step = Number(el.dataset.step);
    el.classList.remove("is-active", "is-done");
    if (step < n) el.classList.add("is-done");
    if (step === n) el.classList.add("is-active");
  });
  window.scrollTo({ top: 0, behavior: "smooth" });
}

backToUploadBtn.addEventListener("click", () => goToPage(1));
backToPresetBtn.addEventListener("click", () => goToPage(2));

// ---- Page 1: file selection ----

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

// ---- Page 1: header detection ----

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

  renderReferenceTable(previewTable, previewWrap, data.headers, data.preview_rows);

  panelHeaders.hidden = false;
  panelHeaders.scrollIntoView({ behavior: "smooth", block: "start" });
}

function renderReferenceTable(tableEl, wrapEl, headers, rows) {
  if (!rows || rows.length === 0) {
    if (wrapEl) wrapEl.hidden = true;
    tableEl.innerHTML = "";
    return;
  }
  const thead = `<thead><tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${rows
    .map(
      (row) =>
        `<tr>${headers.map((h) => `<td>${escapeHtml(row[h] ?? "")}</td>`).join("")}</tr>`
    )
    .join("")}</tbody>`;
  tableEl.innerHTML = thead + tbody;
  if (wrapEl) wrapEl.hidden = false;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = String(str);
  return div.innerHTML;
}

// ---- Page 2: preset picker + inline preview ----

continueBtn.addEventListener("click", () => {
  goToPage(2);
  loadPresets();
});

async function loadPresets() {
  presetStatus.className = "status is-loading";
  presetStatus.textContent = "Loading document types…";
  presetList.innerHTML = "";
  panelPresetPreview.hidden = true;

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
  panelPresetPreview.hidden = true;

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
    computeAutoMapping(config);
    panelPresetPreview.hidden = false;
    runPresetPreview();
  } catch (err) {
    presetStatus.className = "status is-error";
    presetStatus.textContent = "Couldn't reach the server.";
  }
}

function computeAutoMapping(config) {
  columnMapping = {};
  config.columns.forEach((col) => {
    if (col.type === "image") return;
    columnMapping[col.key] = findBestHeaderMatch(col.source_header, currentSession.headers) || "";
  });
}

function findBestHeaderMatch(suggestedHeader, actualHeaders) {
  if (!suggestedHeader) return null;
  const exact = actualHeaders.find((h) => h === suggestedHeader);
  if (exact) return exact;
  const loose = actualHeaders.find((h) => h.trim().toLowerCase() === suggestedHeader.trim().toLowerCase());
  return loose || null;
}

async function runPresetPreview() {
  if (!selectedConfig || !currentSession) return;
  const mySeq = ++presetPreviewSeq;
  presetPreviewStatus.textContent = "Loading preview…";

  const previewConfig = JSON.parse(JSON.stringify(selectedConfig));
  previewConfig.columns = previewConfig.columns.map((col) => {
    if (col.type === "image") return col;
    return { ...col, source_header: columnMapping[col.key] || null };
  });

  try {
    const form = new FormData();
    form.append("session_id", currentSession.session_id);
    form.append("config_json", JSON.stringify(previewConfig));
    form.append("row_limit", "3");

    const res = await fetch(`${API_BASE}/api/preview`, { method: "POST", body: form });
    if (mySeq !== presetPreviewSeq) return;

    if (!res.ok) {
      let message = "Couldn't generate a preview — you can still continue and map columns manually.";
      try {
        const data = await res.json();
        message = data.detail || message;
      } catch (_) {}
      presetPreviewStatus.textContent = message;
      return;
    }

    const blob = await res.blob();
    if (mySeq !== presetPreviewSeq) return;

    if (presetPreviewObjectUrl) URL.revokeObjectURL(presetPreviewObjectUrl);
    presetPreviewObjectUrl = URL.createObjectURL(blob);
    presetPreviewFrame.src = presetPreviewObjectUrl;
    presetPreviewStatus.textContent = "";
  } catch (err) {
    if (mySeq === presetPreviewSeq) {
      presetPreviewStatus.textContent = "Couldn't reach the server — you can still continue.";
    }
  }
}

continueToMappingBtn.addEventListener("click", () => {
  goToPage(3);
  showMappingForm(selectedConfig);
});

// ---- Page 3: column mapping ----

function showMappingForm(config) {
  mappingHint.textContent =
    `"${config.name}" — match each field to a column from your file. Fields matching your file's headers are pre-filled; check them and adjust anything that's wrong.`;

  renderReferenceTable(mappingReferenceTable, null, currentSession.headers, currentSession.preview_rows);

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

    const currentValue = columnMapping[col.key] || "";
    currentSession.headers.forEach((h) => {
      const opt = document.createElement("option");
      opt.value = h;
      opt.textContent = h;
      if (h === currentValue) opt.selected = true;
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

  buildCustomizeSection(config);
  buildFilenameDefault(config);
  resultPanel.hidden = true;
  generateStatus.className = "status";
  generateStatus.textContent = "";
  updateGenerateAvailability();
  schedulePreview(true);
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

// ---- Page 3: customize headers & widths, with live preview ----

function closestWidthPreset(col) {
  const widthMode = col.width_mode || "fixed";
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

function buildFinalConfig() {
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

    if (mySeq !== previewRequestSeq) return;

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

// ---- Page 3: file name ----

function buildFilenameDefault(config) {
  filenameInput.value = sanitizeFilenameInput(config.document_title || config.name);
}

function sanitizeFilenameInput(raw) {
  return String(raw || "document")
    .replace(/[\\/:*?"<>|\r\n\t]/g, "")
    .trim()
    .slice(0, 150) || "document";
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

    const finalName = sanitizeFilenameInput(filenameInput.value);
    downloadLink.href = `${API_BASE}${data.download_url}?filename=${encodeURIComponent(finalName)}`;
    downloadLink.download = finalName.toLowerCase().endsWith(".pdf") ? finalName : `${finalName}.pdf`;
    resultPanel.hidden = false;
  } catch (err) {
    generateStatus.className = "status is-error";
    generateStatus.textContent = "Couldn't reach the server.";
  } finally {
    generateBtn.disabled = false;
  }
}

// ---- Status helpers ----

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
