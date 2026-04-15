const INVENTORY_CSV_PATH = "resources/84 Inventory - 12.30.25 - W0889757.csv";
const MAX_SUGGESTIONS = 8;
const DRAFT_STORAGE_KEY = "material-request-draft-v1";
const HISTORY_STORAGE_KEY = "material-request-history-v1";
const REDO_STORAGE_KEY = "material-request-redo-v1";
const SUBDIVISION_STORAGE_KEY = "material-request-subdivisions-v1";
const LOT_STORAGE_KEY = "material-request-lots-v1";
const HISTORY_LIMIT = 60;
const MAX_SUBDIVISION_SUGGESTIONS = 10;
const MAX_LOT_SUGGESTIONS = 10;

const state = {
  inventory: [],
  lotCount: 0,
  isRestoring: false,
  saveTimer: null,
  activeSubdivisionInput: null,
  lastEmailSubject: "",
};

const lotsContainer = document.getElementById("lotsContainer");
const addLotBtn = document.getElementById("addLotBtn");
const undoBtn = document.getElementById("undoBtn");
const redoBtn = document.getElementById("redoBtn");
const sendBtn = document.getElementById("sendBtn");
const outputSection = document.getElementById("outputSection");
const emailOutput = document.getElementById("emailOutput");
const copyBtn = document.getElementById("copyBtn");
const emailBtn = document.getElementById("emailBtn");
const lotCardTemplate = document.getElementById("lotCardTemplate");
const itemRowTemplate = document.getElementById("itemRowTemplate");

addLotBtn.addEventListener("click", () => addLotCard());
undoBtn.addEventListener("click", undoLastChange);
redoBtn.addEventListener("click", redoLastUndo);
sendBtn.addEventListener("click", handleSend);
copyBtn.addEventListener("click", copyEmailText);
emailBtn.addEventListener("click", openEmailDraft);

lotsContainer.addEventListener("input", handleDraftInput);
lotsContainer.addEventListener("change", () => saveDraftSnapshot(true));

document.addEventListener("click", (event) => {
  const clickedInsideItemSuggestions = event.target.closest(".field-item-search");
  if (!clickedInsideItemSuggestions) {
    closeAllSuggestions();
  }

  const clickedInsideSubdivisionSuggestions = event.target.closest(".field-subdivision-search");
  if (!clickedInsideSubdivisionSuggestions) {
    closeAllSubdivisionSuggestions();
  }

  const clickedInsideLotSuggestions = event.target.closest(".field-lot-search");
  if (!clickedInsideLotSuggestions) {
    closeAllLotSuggestions();
  }
});

boot();

async function boot() {
  try {
    state.inventory = await loadInventory(INVENTORY_CSV_PATH);
  } catch (error) {
    console.error(error);
    alert("Could not load inventory CSV. Please run this from a local web server.");
  }

  const restored = restoreDraftState();
  if (!restored) {
    addLotCard(null, { skipSave: true, skipFocus: true });
    saveDraftSnapshot(true);
  }

  syncHistoryButtons();
}

async function loadInventory(csvPath) {
  const response = await fetch(encodeURI(csvPath));
  if (!response.ok) {
    throw new Error(`Inventory CSV failed to load: ${response.status}`);
  }

  const csvText = await response.text();
  const rows = parseCsv(csvText);

  return rows
    .filter((row) => row["Pos#"] && row.Description)
    .map((row) => ({
      pos: String(row["Pos#"]).trim(),
      description: String(row.Description).trim(),
    }))
    .filter((row) => row.pos && row.description);
}

function parseCsv(csvText) {
  const lines = csvText.split(/\r?\n/).filter((line) => line.length > 0);
  if (!lines.length) {
    return [];
  }

  const headers = splitCsvLine(lines[0]).map((h) => h.trim());
  const records = [];

  for (let i = 1; i < lines.length; i += 1) {
    const values = splitCsvLine(lines[i]);
    const record = {};

    for (let j = 0; j < headers.length; j += 1) {
      record[headers[j]] = values[j] ?? "";
    }

    records.push(record);
  }

  return records;
}

function splitCsvLine(line) {
  const parts = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      parts.push(current.trim());
      current = "";
      continue;
    }

    current += char;
  }

  parts.push(current.trim());
  return parts;
}

function createEmptyItemData() {
  return { qty: "1", description: "", pos: "" };
}

function createEmptyLotData() {
  return { subdivision: "", lot: "", items: [createEmptyItemData()] };
}

function addLotCard(initialLotData = null, options = {}) {
  const { skipSave = false, skipFocus = false } = options;
  const fragment = lotCardTemplate.content.cloneNode(true);
  const lotCard = fragment.querySelector(".lot-card");
  const removeLotBtn = lotCard.querySelector(".remove-lot-btn");
  const addItemBtn = lotCard.querySelector(".add-item-btn");
  const itemsContainer = lotCard.querySelector(".items-container");
  const subdivisionInput = lotCard.querySelector(".subdivision-input");
  const subdivisionSuggestionsBox = lotCard.querySelector(".subdivision-suggestions");
  const lotInput = lotCard.querySelector(".lot-input");
  const lotSuggestionsBox = lotCard.querySelector(".lot-suggestions");

  state.lotCount += 1;

  const lotData = initialLotData || createEmptyLotData();
  subdivisionInput.value = lotData.subdivision || "";
  lotInput.value = lotData.lot || "";

  subdivisionInput.addEventListener("focus", () => {
    state.activeSubdivisionInput = subdivisionInput;
    renderSubdivisionSuggestions(subdivisionInput, subdivisionSuggestionsBox);
  });

  subdivisionInput.addEventListener("input", () => {
    state.activeSubdivisionInput = subdivisionInput;
    renderSubdivisionSuggestions(subdivisionInput, subdivisionSuggestionsBox);
  });

  subdivisionInput.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      hideSubdivisionSuggestions(subdivisionSuggestionsBox);
    }
  });

  lotInput.addEventListener("focus", () => {
    renderLotSuggestions(lotInput, lotSuggestionsBox);
  });

  lotInput.addEventListener("input", () => {
    renderLotSuggestions(lotInput, lotSuggestionsBox);
  });

  lotInput.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      hideLotSuggestions(lotSuggestionsBox);
    }
  });

  addItemBtn.addEventListener("click", () => {
    addItemRow(itemsContainer);
  });

  removeLotBtn.addEventListener("click", () => {
    lotCard.remove();
    renumberLots();
    if (!lotsContainer.querySelector(".lot-card")) {
      addLotCard(null, { skipFocus: true });
      return;
    }

    saveDraftSnapshot(true);
  });

  lotsContainer.appendChild(fragment);

  const lotItems = Array.isArray(lotData.items) && lotData.items.length ? lotData.items : [createEmptyItemData()];
  lotItems.forEach((itemData) => {
    addItemRow(itemsContainer, itemData, { skipSave: true, skipFocus: true });
  });

  if (!skipSave) {
    saveDraftSnapshot(true);
  }

  if (!skipFocus) {
    subdivisionInput.focus();
  }
}

function renumberLots() {
  const lotCards = Array.from(lotsContainer.querySelectorAll(".lot-card"));
  state.lotCount = lotCards.length;
}

function addItemRow(itemsContainer, initialItemData = null, options = {}) {
  const { skipSave = false, skipFocus = false } = options;
  const fragment = itemRowTemplate.content.cloneNode(true);
  const row = fragment.querySelector(".item-row");
  const qtyInput = row.querySelector(".qty-input");
  const itemInput = row.querySelector(".item-input");
  const posInput = row.querySelector(".pos-input");
  const suggestionsBox = row.querySelector(".suggestions");
  const removeBtn = row.querySelector(".remove-item-btn");

  const itemData = initialItemData || createEmptyItemData();
  qtyInput.value = itemData.qty || "1";
  itemInput.value = itemData.description || "";
  posInput.value = itemData.pos || "";

  itemInput.addEventListener("input", () => {
    posInput.value = "";
    renderSuggestions(itemInput, suggestionsBox, posInput);
  });

  itemInput.addEventListener("focus", () => {
    renderSuggestions(itemInput, suggestionsBox, posInput);
  });

  itemInput.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      hideSuggestions(suggestionsBox);
    }
  });

  removeBtn.addEventListener("click", () => {
    row.remove();

    if (!itemsContainer.querySelector(".item-row")) {
      addItemRow(itemsContainer, createEmptyItemData(), { skipFocus: true });
      return;
    }

    saveDraftSnapshot(true);
  });

  itemsContainer.appendChild(fragment);

  if (!skipSave) {
    saveDraftSnapshot(true);
  }

  if (!skipFocus) {
    itemInput.focus();
  }
}

function renderSuggestions(itemInput, suggestionsBox, posInput) {
  closeAllSuggestions(suggestionsBox);

  const query = itemInput.value.trim().toLowerCase();
  if (!query) {
    hideSuggestions(suggestionsBox);
    return;
  }

  const matches = state.inventory
    .filter((item) => item.description.toLowerCase().includes(query))
    .slice(0, MAX_SUGGESTIONS);

  if (!matches.length) {
    suggestionsBox.innerHTML = "";
    hideSuggestions(suggestionsBox);
    return;
  }

  suggestionsBox.innerHTML = "";

  matches.forEach((item) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "suggestion-btn";
    btn.innerHTML = `
      <div class="suggestion-top">${escapeHtml(item.description)}</div>
      <div class="suggestion-sub">POS# ${escapeHtml(item.pos)}</div>
    `;

    btn.addEventListener("click", () => {
      itemInput.value = item.description;
      posInput.value = item.pos;
      hideSuggestions(suggestionsBox);
      saveDraftSnapshot(true);
    });

    suggestionsBox.appendChild(btn);
  });

  suggestionsBox.classList.remove("hidden");
}

function hideSuggestions(suggestionsBox) {
  suggestionsBox.classList.add("hidden");
}

function closeAllSuggestions(exceptBox = null) {
  const all = document.querySelectorAll(".suggestions");
  all.forEach((box) => {
    if (box !== exceptBox) {
      box.classList.add("hidden");
    }
  });
}

function renderSubdivisionSuggestions(subdivisionInput, suggestionsBox) {
  closeAllSubdivisionSuggestions(suggestionsBox);

  const values = getSubdivisionMemory();
  if (!values.length) {
    suggestionsBox.innerHTML = "";
    hideSubdivisionSuggestions(suggestionsBox);
    return;
  }

  const query = subdivisionInput.value.trim().toLowerCase();
  const matches = values
    .filter((value) => value.toLowerCase().includes(query))
    .slice(0, MAX_SUBDIVISION_SUGGESTIONS);

  if (!matches.length) {
    suggestionsBox.innerHTML = "";
    hideSubdivisionSuggestions(suggestionsBox);
    return;
  }

  suggestionsBox.innerHTML = "";

  matches.forEach((value) => {
    const row = document.createElement("div");
    row.className = "subdivision-option";

    const selectBtn = document.createElement("button");
    selectBtn.type = "button";
    selectBtn.className = "subdivision-select-btn";
    selectBtn.textContent = value;
    selectBtn.addEventListener("click", () => {
      subdivisionInput.value = value;
      hideSubdivisionSuggestions(suggestionsBox);
      saveDraftSnapshot(true);
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "subdivision-remove-btn";
    removeBtn.setAttribute("aria-label", `Remove subdivision ${value}`);
    removeBtn.textContent = "x";
    removeBtn.addEventListener("click", () => {
      removeSubdivisionFromMemory(value);
      renderSubdivisionSuggestions(subdivisionInput, suggestionsBox);
    });

    row.appendChild(selectBtn);
    row.appendChild(removeBtn);
    suggestionsBox.appendChild(row);
  });

  suggestionsBox.classList.remove("hidden");
}

function hideSubdivisionSuggestions(suggestionsBox) {
  suggestionsBox.classList.add("hidden");
}

function closeAllSubdivisionSuggestions(exceptBox = null) {
  const all = document.querySelectorAll(".subdivision-suggestions");
  all.forEach((box) => {
    if (box !== exceptBox) {
      box.classList.add("hidden");
    }
  });
}

function renderLotSuggestions(lotInput, suggestionsBox) {
  closeAllLotSuggestions(suggestionsBox);

  const values = getLotMemory();
  if (!values.length) {
    suggestionsBox.innerHTML = "";
    hideLotSuggestions(suggestionsBox);
    return;
  }

  const query = lotInput.value.trim().toLowerCase();
  const matches = values
    .filter((value) => value.toLowerCase().includes(query))
    .slice(0, MAX_LOT_SUGGESTIONS);

  if (!matches.length) {
    suggestionsBox.innerHTML = "";
    hideLotSuggestions(suggestionsBox);
    return;
  }

  suggestionsBox.innerHTML = "";

  matches.forEach((value) => {
    const row = document.createElement("div");
    row.className = "lot-option";

    const selectBtn = document.createElement("button");
    selectBtn.type = "button";
    selectBtn.className = "lot-select-btn";
    selectBtn.textContent = value;
    selectBtn.addEventListener("click", () => {
      lotInput.value = value;
      hideLotSuggestions(suggestionsBox);
      saveDraftSnapshot(true);
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "lot-remove-btn";
    removeBtn.setAttribute("aria-label", `Remove lot ${value}`);
    removeBtn.textContent = "x";
    removeBtn.addEventListener("click", () => {
      removeLotFromMemory(value);
      renderLotSuggestions(lotInput, suggestionsBox);
    });

    row.appendChild(selectBtn);
    row.appendChild(removeBtn);
    suggestionsBox.appendChild(row);
  });

  suggestionsBox.classList.remove("hidden");
}

function hideLotSuggestions(suggestionsBox) {
  suggestionsBox.classList.add("hidden");
}

function closeAllLotSuggestions(exceptBox = null) {
  const all = document.querySelectorAll(".lot-suggestions");
  all.forEach((box) => {
    if (box !== exceptBox) {
      box.classList.add("hidden");
    }
  });
}

function getEditorState() {
  const lotCards = Array.from(lotsContainer.querySelectorAll(".lot-card"));

  const lots = lotCards.map((lotCard) => {
    const subdivision = lotCard.querySelector(".subdivision-input").value.trim();
    const lot = lotCard.querySelector(".lot-input").value.trim();
    const rows = Array.from(lotCard.querySelectorAll(".item-row"));
    const items = rows.map((row) => {
      const qty = row.querySelector(".qty-input").value.trim();
      const description = row.querySelector(".item-input").value.trim();
      const pos = row.querySelector(".pos-input").value.trim();
      return { qty, description, pos };
    });

    return { subdivision, lot, items };
  });

  return { lots };
}

function getRequestLots() {
  const snapshot = getEditorState();

  return snapshot.lots.map((entry) => {
    const filteredItems = entry.items.filter((item) => item.qty || item.description);
    return {
      subdivision: entry.subdivision,
      lot: entry.lot,
      items: filteredItems,
    };
  });
}

function restoreSnapshot(snapshot) {
  const lots = Array.isArray(snapshot?.lots) && snapshot.lots.length ? snapshot.lots : [createEmptyLotData()];

  state.isRestoring = true;
  closeAllSuggestions();
  closeAllSubdivisionSuggestions();
  closeAllLotSuggestions();
  lotsContainer.innerHTML = "";
  state.lotCount = 0;

  lots.forEach((lotData) => {
    addLotCard(lotData, { skipSave: true, skipFocus: true });
  });

  renumberLots();
  state.isRestoring = false;
  state.activeSubdivisionInput = lotsContainer.querySelector(".subdivision-input");
}

function readJsonObject(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return null;
    }

    return JSON.parse(raw);
  } catch (error) {
    console.warn(`Could not read local storage key: ${key}`, error);
    return null;
  }
}

function readJsonArray(key) {
  const parsed = readJsonObject(key);
  return Array.isArray(parsed) ? parsed : [];
}

function writeHistory(history) {
  localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(history));
  syncHistoryButtons();
}

function getHistory() {
  return readJsonArray(HISTORY_STORAGE_KEY);
}

function writeRedoStack(redoStack) {
  localStorage.setItem(REDO_STORAGE_KEY, JSON.stringify(redoStack));
  syncHistoryButtons();
}

function getRedoStack() {
  return readJsonArray(REDO_STORAGE_KEY);
}

function restoreDraftState() {
  const draft = readJsonObject(DRAFT_STORAGE_KEY);
  if (!draft || !Array.isArray(draft.lots)) {
    return false;
  }

  restoreSnapshot(draft);

  const history = getHistory();
  if (!history.length) {
    writeHistory([draft]);
  }

  const redoStack = getRedoStack();
  if (!Array.isArray(redoStack)) {
    writeRedoStack([]);
  }

  return true;
}

function pushHistorySnapshot(snapshot) {
  const history = getHistory();
  const nextSerialized = JSON.stringify(snapshot);
  const lastSerialized = history.length ? JSON.stringify(history[history.length - 1]) : "";

  if (nextSerialized === lastSerialized) {
    return;
  }

  history.push(snapshot);

  if (history.length > HISTORY_LIMIT) {
    history.splice(0, history.length - HISTORY_LIMIT);
  }

  writeHistory(history);
}

function saveDraftSnapshot(pushHistory = true) {
  if (state.isRestoring) {
    return;
  }

  const snapshot = getEditorState();
  localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(snapshot));
  updateSubdivisionMemory(snapshot);
  updateLotMemory(snapshot);

  if (pushHistory) {
    pushHistorySnapshot(snapshot);
    writeRedoStack([]);
  }
}

function handleDraftInput() {
  if (state.isRestoring) {
    return;
  }

  if (state.saveTimer) {
    clearTimeout(state.saveTimer);
  }

  state.saveTimer = setTimeout(() => {
    saveDraftSnapshot(true);
  }, 280);
}

function syncHistoryButtons() {
  const history = getHistory();
  const redoStack = getRedoStack();
  undoBtn.disabled = history.length < 2;
  redoBtn.disabled = redoStack.length === 0;
}

function undoLastChange() {
  const history = getHistory();
  if (history.length < 2) {
    alert("Nothing to undo yet.");
    return;
  }

  const redoStack = getRedoStack();
  const current = history[history.length - 1];
  redoStack.push(current);

  history.pop();
  const previous = history[history.length - 1];
  writeHistory(history);
  writeRedoStack(redoStack);
  localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(previous));
  restoreSnapshot(previous);
}

function redoLastUndo() {
  const redoStack = getRedoStack();
  if (!redoStack.length) {
    alert("Nothing to redo yet.");
    return;
  }

  const history = getHistory();
  const restored = redoStack.pop();
  history.push(restored);

  if (history.length > HISTORY_LIMIT) {
    history.splice(0, history.length - HISTORY_LIMIT);
  }

  writeHistory(history);
  writeRedoStack(redoStack);
  localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(restored));
  restoreSnapshot(restored);
}

function getSubdivisionMemory() {
  const values = readJsonArray(SUBDIVISION_STORAGE_KEY);
  return values
    .map((value) => String(value || "").trim())
    .filter(Boolean);
}

function getLotMemory() {
  const values = readJsonArray(LOT_STORAGE_KEY);
  return values
    .map((value) => String(value || "").trim())
    .filter(Boolean);
}

function writeSubdivisionMemory(values) {
  localStorage.setItem(SUBDIVISION_STORAGE_KEY, JSON.stringify(values));
}

function writeLotMemory(values) {
  localStorage.setItem(LOT_STORAGE_KEY, JSON.stringify(values));
}

function updateSubdivisionMemory(snapshot) {
  const existing = getSubdivisionMemory();
  const map = new Map(existing.map((value) => [value.toLowerCase(), value]));

  snapshot.lots.forEach((entry) => {
    const value = String(entry.subdivision || "").trim();
    if (!value) {
      return;
    }

    map.set(value.toLowerCase(), value);
  });

  const merged = Array.from(map.values()).sort((a, b) => a.localeCompare(b));
  if (JSON.stringify(merged) !== JSON.stringify(existing)) {
    writeSubdivisionMemory(merged);
  }
}

function updateLotMemory(snapshot) {
  const existing = getLotMemory();
  const map = new Map(existing.map((value) => [value.toLowerCase(), value]));

  snapshot.lots.forEach((entry) => {
    const value = String(entry.lot || "").trim();
    if (!value) {
      return;
    }

    map.set(value.toLowerCase(), value);
  });

  const merged = Array.from(map.values()).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  if (JSON.stringify(merged) !== JSON.stringify(existing)) {
    writeLotMemory(merged);
  }
}

function removeSubdivisionFromMemory(valueToRemove) {
  const values = getSubdivisionMemory();
  const filtered = values.filter((value) => value.toLowerCase() !== valueToRemove.toLowerCase());
  writeSubdivisionMemory(filtered);

  const subdivisionInputs = lotsContainer.querySelectorAll(".subdivision-input");
  subdivisionInputs.forEach((input) => {
    if (input.value.trim().toLowerCase() === valueToRemove.toLowerCase()) {
      input.value = "";
    }
  });

  saveDraftSnapshot(true);
}

function removeLotFromMemory(valueToRemove) {
  const values = getLotMemory();
  const filtered = values.filter((value) => value.toLowerCase() !== valueToRemove.toLowerCase());
  writeLotMemory(filtered);

  const lotInputs = lotsContainer.querySelectorAll(".lot-input");
  lotInputs.forEach((input) => {
    if (input.value.trim().toLowerCase() === valueToRemove.toLowerCase()) {
      input.value = "";
    }
  });

  saveDraftSnapshot(true);
}

function handleSend() {
  const lots = getRequestLots();
  const activeLots = lots.filter((entry) => entry.subdivision || entry.lot || entry.items.length);

  if (!activeLots.length) {
    alert("Please add at least one lot with items.");
    return;
  }

  const lotMissingDetails = activeLots.find((entry) => !entry.subdivision || !entry.lot);
  if (lotMissingDetails) {
    alert("Each active lot needs both subdivision and lot number.");
    return;
  }

  const lotMissingItems = activeLots.find((entry) => entry.items.length === 0);
  if (lotMissingItems) {
    alert("Each active lot needs at least one item.");
    return;
  }

  const invalidItem = activeLots
    .flatMap((entry) => entry.items)
    .find((item) => !item.qty || !item.description || !item.pos);
  if (invalidItem) {
    alert("Each item needs quantity and a selected inventory item from suggestions.");
    return;
  }

  const emailSubject = buildEmailSubject(activeLots);
  const emailText = buildEmailText(activeLots);
  state.lastEmailSubject = emailSubject;
  emailOutput.value = emailText;
  outputSection.classList.remove("hidden");

  const pslText = buildPslText(activeLots);
  downloadTextFile(buildFileName(activeLots), pslText);
  openEmailDraft(emailSubject, emailText);
}

function buildEmailSubject(lots) {
  return lots
    .map((entry) => `${entry.subdivision} Lot ${entry.lot}`)
    .join(" | ");
}

function buildEmailText(lots) {
  const lines = [];

  lots.forEach((entry) => {
    lines.push(`Subdivision: ${entry.subdivision}`);
    lines.push(`Lot: ${entry.lot}`);

    entry.items.forEach((item) => {
      lines.push(`(${item.qty}) ${item.description}`);
    });

    lines.push("");
  });

  return lines.join("\n");
}

function buildPslText(lots) {
  return lots
    .flatMap((entry) => entry.items)
    .map((item) => `send "${item.pos}<cr>${item.qty}<cr>"`)
    .join("\n");
}

function buildFileName(lots) {
  const grouped = [];
  const groupedMap = new Map();

  lots.forEach((entry) => {
    const subdivisionToken = normalizeToken(entry.subdivision) || "SUBDIVISION";

    if (!groupedMap.has(subdivisionToken)) {
      groupedMap.set(subdivisionToken, []);
      grouped.push(subdivisionToken);
    }

    groupedMap.get(subdivisionToken).push(entry.lot);
  });

  const segments = grouped.map((subdivisionToken) => {
    const lotLabel = formatLotLabel(groupedMap.get(subdivisionToken));
    return `${subdivisionToken}_${lotLabel}`;
  });

  if (!segments.length) {
    return "MATERIAL_REQUEST.psl";
  }

  return `${segments.join("_")}.psl`;
}

function normalizeToken(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "");
}

function formatLotLabel(lotValues) {
  const cleaned = lotValues.map((value) => String(value || "").trim()).filter(Boolean);
  if (!cleaned.length) {
    return "LOT";
  }

  const numericLots = cleaned.map((value) => (/^\d+$/.test(value) ? Number(value) : NaN));
  const allNumeric = numericLots.every((value) => Number.isFinite(value));

  if (allNumeric) {
    const sortedUnique = Array.from(new Set(numericLots)).sort((a, b) => a - b);
    const ranges = compressNumberRanges(sortedUnique);

    if (ranges.length === 1 && !ranges[0].includes("-")) {
      return `LOT${ranges[0]}`;
    }

    return `LOTS${ranges.join("_")}`;
  }

  const lotTokens = cleaned.map((value) => normalizeToken(value)).filter(Boolean);

  if (!lotTokens.length) {
    return "LOT";
  }

  if (lotTokens.length === 1) {
    return `LOT${lotTokens[0]}`;
  }

  return `LOTS${lotTokens.join("_")}`;
}

function compressNumberRanges(numbers) {
  const ranges = [];
  let start = numbers[0];
  let previous = numbers[0];

  for (let i = 1; i <= numbers.length; i += 1) {
    const current = numbers[i];
    const isConsecutive = current === previous + 1;

    if (isConsecutive) {
      previous = current;
      continue;
    }

    if (start === previous) {
      ranges.push(String(start));
    } else {
      ranges.push(`${start}-${previous}`);
    }

    start = current;
    previous = current;
  }

  return ranges;
}

function downloadTextFile(fileName, text) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function copyEmailText() {
  const text = emailOutput.value;
  if (!text) {
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    copyBtn.textContent = "Copied";
    setTimeout(() => {
      copyBtn.textContent = "Copy Text";
    }, 1500);
  } catch (error) {
    console.error(error);
    alert("Clipboard copy failed. You can still select and copy manually.");
  }
}

function openEmailDraft(subjectText = null, bodyText = null) {
  const text = bodyText ?? emailOutput.value;
  if (!text) {
    alert("Click Send first to generate the email text.");
    return;
  }

  const rawSubject = (subjectText ?? state.lastEmailSubject) || "Fill-In App Request";
  const subject = encodeURIComponent(rawSubject);
  const body = encodeURIComponent(text);
  window.location.href = `mailto:?subject=${subject}&body=${body}`;
}

function escapeHtml(value) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
