const SHEET_NAME = "Links";
const SPIN_SPEED = 110;

/*
  Achtung:
  Dieser Passwortschutz ist für GitHub Pages nur ein Frontend-Schutz.
  Für echte Sicherheit wäre ein Backend nötig.
*/
const ADMIN_PASSWORD = "ChangeThis123!";

const CATEGORY_CONFIG = {
  "Digitalization": {
    image: "assets/digitalization.png"
  },
  "Finance Transformation": {
    image: "assets/finance-transformation.png"
  },
  "Smart Work": {
    image: "assets/smart-work.png"
  },
  "International": {
    image: "assets/international.png"
  },
  "Mindset": {
    image: "assets/mindset.png"
  },
  "(Automotive) Future": {
    image: "assets/automotive-future.png"
  }
};

const CATEGORY_ORDER = [
  "Digitalization",
  "Finance Transformation",
  "Smart Work",
  "International",
  "Mindset",
  "(Automotive) Future"
];

const FALLBACK_DATA = [
  {
    category: "Digitalization",
    author: "Demo Source",
    title: "Digital Innovation Overview",
    url: "https://example.com/digitalization",
    type: "text"
  },
  {
    category: "Finance Transformation",
    author: "Demo Source",
    title: "Future of Finance",
    url: "https://example.com/finance",
    type: "video"
  },
  {
    category: "Smart Work",
    author: "Demo Source",
    title: "Smart Collaboration Models",
    url: "https://example.com/smartwork",
    type: "podcast"
  },
  {
    category: "International",
    author: "Demo Source",
    title: "Global Market Signals",
    url: "https://example.com/international",
    type: "text"
  },
  {
    category: "Mindset",
    author: "Demo Source",
    title: "Growth Mindset Basics",
    url: "https://example.com/mindset",
    type: "video"
  },
  {
    category: "(Automotive) Future",
    author: "Demo Source",
    title: "Mobility 2030",
    url: "https://example.com/future",
    type: "text"
  }
];

const state = {
  grouped: {},
  timers: new Map(),
  selected: new Map(),
  isUnlocked: false
};

const categoriesGrid = document.getElementById("categoriesGrid");
const finalSelection = document.getElementById("finalSelection");
const categoryTemplate = document.getElementById("categoryTemplate");

const startAllBtn = document.getElementById("startAllBtn");
const stopAllBtn = document.getElementById("stopAllBtn");
const continueAllBtn = document.getElementById("continueAllBtn");

const toggleAdminBtn = document.getElementById("toggleAdminBtn");
const adminPanel = document.getElementById("adminPanel");
const authPanel = document.getElementById("authPanel");
const uploadPanel = document.getElementById("uploadPanel");
const passwordInput = document.getElementById("passwordInput");
const unlockBtn = document.getElementById("unlockBtn");
const lockBtn = document.getElementById("lockBtn");
const excelFile = document.getElementById("excelFile");
const authMessage = document.getElementById("authMessage");
const statusBar = document.getElementById("statusBar");

startAllBtn.addEventListener("click", startAll);
stopAllBtn.addEventListener("click", stopAll);
continueAllBtn.addEventListener("click", startAll);

toggleAdminBtn.addEventListener("click", toggleAdminPanel);
unlockBtn.addEventListener("click", unlockAdmin);
lockBtn.addEventListener("click", lockAdmin);
passwordInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    unlockAdmin();
  }
});
excelFile.addEventListener("change", handleExcelUpload);

init();

function init() {
  state.grouped = groupByCategory(FALLBACK_DATA);
  renderApp();
}

function toggleAdminPanel() {
  adminPanel.classList.toggle("hidden");
}

function unlockAdmin() {
  const entered = passwordInput.value.trim();

  if (entered === ADMIN_PASSWORD) {
    state.isUnlocked = true;
    authPanel.classList.add("hidden");
    uploadPanel.classList.remove("hidden");
    authMessage.textContent = "";
    passwordInput.value = "";
  } else {
    authMessage.textContent = "Passwort nicht korrekt.";
  }
}

function lockAdmin() {
  state.isUnlocked = false;
  uploadPanel.classList.add("hidden");
  authPanel.classList.remove("hidden");
  passwordInput.value = "";
  authMessage.textContent = "";
  statusBar.textContent = "";
}

async function handleExcelUpload(event) {
  const file = event.target.files?.[0];
  if (!file || !state.isUnlocked) return;

  try {
    setStatus(`Lade ${file.name} ...`);

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const worksheet = workbook.Sheets[SHEET_NAME];

    if (!worksheet) {
      setStatus(`Kein Blatt mit dem Namen "${SHEET_NAME}" gefunden.`);
      return;
    }

    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    const normalizedRows = normalizeRows(rows);
    const preparedRows = normalizedRows.filter(isValidRow);

    if (!preparedRows.length) {
      setStatus("Keine gültigen aktiven Zeilen gefunden.");
      return;
    }

    state.grouped = groupByCategory(preparedRows);
    renderApp();
    setStatus(`${file.name} erfolgreich geladen.`);
  } catch (error) {
    console.error(error);
    setStatus("Die Excel-Datei konnte nicht verarbeitet werden.");
  }
}

function setStatus(message) {
  statusBar.textContent = message;
}

function normalizeRows(rows) {
  return rows.map((row) => {
    const normalized = {};
    Object.keys(row).forEach((key) => {
      normalized[String(key).trim().toLowerCase()] =
        typeof row[key] === "string" ? row[key].trim() : row[key];
    });
    return normalized;
  });
}

function isValidRow(row) {
  const category = row.category;
  const author = row.author;
  const title = row.title;
  const url = row.url;
  const type = String(row.type || "").toLowerCase();
  const active = parseActive(row.active);

  const validCategory = CATEGORY_ORDER.includes(category);
  const validType = ["text", "video", "podcast"].includes(type);
  const validUrl = /^https?:\/\//i.test(String(url || "").trim());

  return Boolean(category && author && title && validCategory && validType && validUrl && active);
}

function parseActive(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["true", "1", "yes", "ja", "active"].includes(normalized);
}

function groupByCategory(rows) {
  const grouped = {};

  CATEGORY_ORDER.forEach((category) => {
    grouped[category] = rows.filter((row) => row.category === category);
  });

  return grouped;
}

function renderApp() {
  stopAll();
  categoriesGrid.innerHTML = "";
  state.selected.clear();

  CATEGORY_ORDER.forEach((category) => {
    const items = state.grouped[category] || [];
    if (!items.length) return;

    const node = categoryTemplate.content.firstElementChild.cloneNode(true);
    node.dataset.category = category;

    const image = node.querySelector(".category-image");
    const title = node.querySelector(".category-title");
    const resultCard = node.querySelector(".result-card");
    const resultAuthor = node.querySelector(".result-author");
    const resultTitle = node.querySelector(".result-title");
    const typeIcon = node.querySelector(".type-icon");
    const typeLabel = node.querySelector(".type-label");

    const startBtn = node.querySelector(".start-btn");
    const stopBtn = node.querySelector(".stop-btn");
    const continueBtn = node.querySelector(".continue-btn");

    title.textContent = category;
    image.src = CATEGORY_CONFIG[category].image;
    image.alt = category;

    const renderRow = (row) => {
      resultCard.href = row.url;
      resultAuthor.textContent = row.author;
      resultTitle.textContent = row.title;
      typeIcon.textContent = getTypeIcon(row.type);
      typeLabel.textContent = formatType(row.type);
      state.selected.set(category, row);
      renderFinalSelection();
    };

    node._renderRow = renderRow;

    const firstRow = items[Math.floor(Math.random() * items.length)];
    renderRow(firstRow);

    startBtn.addEventListener("click", () => startSpin(category));
    stopBtn.addEventListener("click", () => stopSpin(category));
    continueBtn.addEventListener("click", () => startSpin(category));

    categoriesGrid.appendChild(node);
  });

  renderFinalSelection();
}

function startSpin(category) {
  const items = state.grouped[category] || [];
  if (!items.length) return;

  stopSpin(category);

  const timer = setInterval(() => {
    const nextItem = items[Math.floor(Math.random() * items.length)];
    paintCategory(category, nextItem);
  }, SPIN_SPEED);

  state.timers.set(category, timer);
}

function stopSpin(category) {
  const timer = state.timers.get(category);
  if (timer) {
    clearInterval(timer);
    state.timers.delete(category);
  }
}

function startAll() {
  CATEGORY_ORDER.forEach((category) => {
    if ((state.grouped[category] || []).length) {
      startSpin(category);
    }
  });
}

function stopAll() {
  for (const timer of state.timers.values()) {
    clearInterval(timer);
  }
  state.timers.clear();
}

function paintCategory(category, row) {
  const card = [...document.querySelectorAll(".category-card")]
    .find((el) => el.dataset.category === category);

  if (!card || !card._renderRow) return;
  card._renderRow(row);
}

function renderFinalSelection() {
  finalSelection.innerHTML = "";

  CATEGORY_ORDER.forEach((category) => {
    const row = state.selected.get(category);
    if (!row) return;

    const item = document.createElement("div");
    item.className = "final-item";

    item.innerHTML = `
      <h4>${escapeHtml(category)}</h4>
      <a href="${row.url}" target="_blank" rel="noopener noreferrer">${escapeHtml(row.title)}</a>
      <div class="meta">${escapeHtml(row.author)} · ${formatType(row.type)}</div>
    `;

    finalSelection.appendChild(item);
  });
}

function getTypeIcon(type) {
  const value = String(type || "").toLowerCase();
  if (value === "video") return "🎬";
  if (value === "podcast") return "🎧";
  return "📄";
}

function formatType(type) {
  const value = String(type || "").toLowerCase();
  if (value === "video") return "Video";
  if (value === "podcast") return "Podcast";
  return "Text";
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
