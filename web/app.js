// ─── Translations ────────────────────────────────────────────────────────────

const TRANSLATIONS = {
  cs: {
    title: "Počítání replik a obsazování",
    lede: 'Vlož vstupy po epizodách, nech si spočítat <strong>VSTUPY</strong> a <strong>REPLIKY</strong> podle pravidla <code>ceil(slov / 8)</code>, doplň dabéry a stáhni hotový <code>.xlsx</code> export.',
    projectNameLabel: "Název projektu",
    projectNameDefault: "Obsazení projektu",
    recalculate: "Přepočítat",
    recalculating: "Počítám…",
    exportXlsx: "Export XLSX",
    exporting: "Exportuji…",
    inputKicker: "Vstup",
    episodesHeading: "Epizody",
    clearActive: "Vyčistit aktivní",
    loadFile: "Načíst soubor",
    episodeLabel: "Epizoda {label}",
    formatsSupported: "Podporované formáty:",
    formatRaw: "<code>POSTAVA / TC / TEXT</code> pro surový dialogový výpis",
    formatSummary: "<code>POSTAVA / VSTUPY / REPLIKY</code> pro už spočítaný souhrn",
    contentPlaceholder: "Sem vlož TSV/CSV s hlavičkami POSTAVA, TC, TEXT nebo POSTAVA, VSTUPY, REPLIKY.",
    castingKicker: "Obsazení",
    fullCastHeading: "Komplet postav",
    summaryKicker: "Souhrn",
    actorsHeading: "Dabéři",
    // badges
    badgeCharacters: "Postavy",
    badgeInputs: "Vstupy",
    badgeReplicas: "Repliky",
    badgeMissing: "Neobsazeno",
    // table headers
    colCharacter: "Postava",
    colInputs: "Vstupy",
    colReplicas: "Repliky",
    colActor: "Dabér",
    colNote: "Poznámka",
    // actor summary
    actorInputs: "{n} vstupů",
    actorReplicas: "{n} replik",
    // status
    statusMissing: "Chybí obsadit {chars} postav, {inputs} vstupů a {replicas} replik.",
    statusOk: "Všechny postavy mají přiřazeného dabéra.",
    // error
    unknownError: "Neznámá chyba.",
  },
  en: {
    title: "Replica counting and casting",
    lede: 'Paste episode inputs, let it calculate <strong>INPUTS</strong> and <strong>REPLICAS</strong> using <code>ceil(words / 8)</code>, fill in voice actors and download the <code>.xlsx</code> export.',
    projectNameLabel: "Project name",
    projectNameDefault: "Project casting",
    recalculate: "Recalculate",
    recalculating: "Calculating…",
    exportXlsx: "Export XLSX",
    exporting: "Exporting…",
    inputKicker: "Input",
    episodesHeading: "Episodes",
    clearActive: "Clear active",
    loadFile: "Load file",
    episodeLabel: "Episode {label}",
    formatsSupported: "Supported formats:",
    formatRaw: "<code>CHARACTER / TC / TEXT</code> for raw dialogue export",
    formatSummary: "<code>CHARACTER / INPUTS / REPLICAS</code> for pre-calculated summary",
    contentPlaceholder: "Paste TSV/CSV with headers CHARACTER, TC, TEXT or CHARACTER, INPUTS, REPLICAS.",
    castingKicker: "Casting",
    fullCastHeading: "Full cast",
    summaryKicker: "Summary",
    actorsHeading: "Voice actors",
    // badges
    badgeCharacters: "Characters",
    badgeInputs: "Inputs",
    badgeReplicas: "Replicas",
    badgeMissing: "Unassigned",
    // table headers
    colCharacter: "Character",
    colInputs: "Inputs",
    colReplicas: "Replicas",
    colActor: "Voice actor",
    colNote: "Note",
    // actor summary
    actorInputs: "{n} inputs",
    actorReplicas: "{n} replicas",
    // status
    statusMissing: "Missing {chars} characters, {inputs} inputs and {replicas} replicas.",
    statusOk: "All characters have an assigned voice actor.",
    // error
    unknownError: "Unknown error.",
  },
};

let currentLang = localStorage.getItem("obsazovani-lang") || "cs";

function t(key, vars = {}) {
  const dict = TRANSLATIONS[currentLang] || TRANSLATIONS.cs;
  let str = dict[key] || TRANSLATIONS.cs[key] || key;
  for (const [k, v] of Object.entries(vars)) {
    str = str.replaceAll(`{${k}}`, v);
  }
  return str;
}

function applyI18n() {
  document.documentElement.lang = currentLang;
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    el.textContent = t(el.dataset.i18n);
  });
  document.querySelectorAll("[data-i18n-html]").forEach((el) => {
    el.innerHTML = t(el.dataset.i18nHtml);
  });
  document.getElementById("project-title").placeholder = t("projectNameDefault");
  document.getElementById("episode-content").placeholder = t("contentPlaceholder");

  document.querySelectorAll(".lang-switcher button").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.lang === currentLang);
  });

  // Re-render episode label
  const episode = state.episodes[state.currentEpisode];
  if (episode) {
    elements.episodeLabel.textContent = t("episodeLabel", { label: episode.label });
  }
}

// ─── State ───────────────────────────────────────────────────────────────────

const state = {
  title: "",
  currentEpisode: 0,
  episodes: Array.from({ length: 6 }, (_, index) => ({
    label: String(index + 1).padStart(2, "0"),
    content: "",
  })),
  assignments: {},
  analysis: null,
};

const elements = {
  title: document.getElementById("project-title"),
  analyzeButton: document.getElementById("analyze-button"),
  exportButton: document.getElementById("export-button"),
  clearCurrentButton: document.getElementById("clear-current-button"),
  episodeFileInput: document.getElementById("episode-file-input"),
  episodeTabs: document.getElementById("episode-tabs"),
  episodeLabel: document.getElementById("episode-label"),
  episodeContent: document.getElementById("episode-content"),
  castingHead: document.querySelector("#casting-table thead"),
  castingBody: document.querySelector("#casting-table tbody"),
  actorSummary: document.getElementById("actor-summary"),
  summaryBadges: document.getElementById("summary-badges"),
  statusBox: document.getElementById("status-box"),
  badgeTemplate: document.getElementById("badge-template"),
  langCsButton: document.getElementById("lang-cs-button"),
  langEnButton: document.getElementById("lang-en-button"),
};

function payload() {
  return {
    title: state.title,
    episodes: state.episodes,
    assignments: state.assignments,
  };
}

function saveDraft() {
  localStorage.setItem("obsazovani-draft", JSON.stringify(payload()));
}

function loadDraft() {
  const draft = localStorage.getItem("obsazovani-draft");
  if (!draft) return;
  try {
    const parsed = JSON.parse(draft);
    state.title = parsed.title || state.title;
    if (Array.isArray(parsed.episodes)) {
      parsed.episodes.slice(0, 6).forEach((episode, index) => {
        state.episodes[index] = {
          label: episode.label || state.episodes[index].label,
          content: episode.content || "",
        };
      });
    }
    state.assignments = parsed.assignments || {};
  } catch (error) {
    console.error(error);
  }
}

async function postJson(url, body, responseType = "json") {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const problem = await response.json().catch(() => ({ error: t("unknownError") }));
    throw new Error(problem.error || t("unknownError"));
  }
  if (responseType === "blob") {
    return response.blob();
  }
  return response.json();
}

async function analyze() {
  try {
    elements.analyzeButton.disabled = true;
    elements.analyzeButton.textContent = t("recalculating");
    state.analysis = await postJson("/api/analyze", payload());
    renderAnalysis();
  } catch (error) {
    elements.statusBox.textContent = error.message;
    elements.statusBox.className = "status-box is-error";
  } finally {
    elements.analyzeButton.disabled = false;
    elements.analyzeButton.textContent = t("recalculate");
  }
}

async function exportWorkbook() {
  try {
    elements.exportButton.disabled = true;
    elements.exportButton.textContent = t("exporting");
    const blob = await postJson("/api/export", payload(), "blob");
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${(state.title || "casting").replace(/\s+/g, "-")}.xlsx`;
    anchor.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    elements.statusBox.textContent = error.message;
    elements.statusBox.className = "status-box is-error";
  } finally {
    elements.exportButton.disabled = false;
    elements.exportButton.textContent = t("exportXlsx");
  }
}

function scheduleAnalyze() {
  saveDraft();
  window.clearTimeout(scheduleAnalyze.timer);
  scheduleAnalyze.timer = window.setTimeout(analyze, 180);
}

function updateEpisodeEditor() {
  const episode = state.episodes[state.currentEpisode];
  elements.episodeLabel.textContent = t("episodeLabel", { label: episode.label });
  elements.episodeContent.value = episode.content;
}

function renderEpisodeTabs() {
  elements.episodeTabs.innerHTML = "";
  state.episodes.forEach((episode, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `episode-tab${state.currentEpisode === index ? " is-active" : ""}`;
    button.textContent = episode.label;
    button.addEventListener("click", () => {
      state.currentEpisode = index;
      renderEpisodeTabs();
      updateEpisodeEditor();
    });
    elements.episodeTabs.appendChild(button);
  });
}

function renderBadges() {
  elements.summaryBadges.innerHTML = "";
  if (!state.analysis) return;

  const items = [
    [t("badgeCharacters"), state.analysis.stats.characterCount],
    [t("badgeInputs"), state.analysis.stats.inputs],
    [t("badgeReplicas"), state.analysis.stats.replicas],
    [t("badgeMissing"), state.analysis.missing.characters],
  ];

  for (const [label, value] of items) {
    const badge = elements.badgeTemplate.content.firstElementChild.cloneNode(true);
    badge.querySelector(".badge-label").textContent = label;
    badge.querySelector(".badge-value").textContent = value;
    elements.summaryBadges.appendChild(badge);
  }
}

function renderCastingTable() {
  if (!state.analysis) return;

  const episodeHeaders = state.analysis.episodes
    .map((episode) => `<th>${episode.label}</th>`)
    .join("");

  elements.castingHead.innerHTML = `
    <tr>
      <th>${t("colCharacter")}</th>
      ${episodeHeaders}
      <th>${t("colInputs")}</th>
      <th>${t("colReplicas")}</th>
      <th>${t("colActor")}</th>
      <th>${t("colNote")}</th>
    </tr>
  `;

  elements.castingBody.innerHTML = "";
  state.analysis.complete.forEach((row) => {
    const tr = document.createElement("tr");
    if (!row.actor) tr.classList.add("is-missing");
    tr.innerHTML = `
      <td class="cell-character">${row.character}</td>
      ${row.episodes.map((episode) => `<td class="cell-episode">${episode.display}</td>`).join("")}
      <td>${row.totalInputs}</td>
      <td>${row.totalReplicas}</td>
      <td><input data-field="actor" data-character="${escapeAttribute(row.character)}" value="${escapeAttribute(row.actor || "")}" /></td>
      <td><input data-field="note" data-character="${escapeAttribute(row.character)}" value="${escapeAttribute(row.note || "")}" /></td>
    `;
    elements.castingBody.appendChild(tr);
  });

  elements.castingBody.querySelectorAll("input").forEach((input) => {
    input.addEventListener("change", (event) => {
      const character = event.target.dataset.character;
      const field = event.target.dataset.field;
      state.assignments[character] = state.assignments[character] || { actor: "", note: "" };
      state.assignments[character][field] = event.target.value;
      scheduleAnalyze();
    });
  });
}

function renderActorSummary() {
  if (!state.analysis) return;
  elements.actorSummary.innerHTML = "";

  state.analysis.actors.forEach((actor) => {
    const article = document.createElement("article");
    article.className = "actor-row";
    article.innerHTML = `
      <div>
        <strong>${actor.actor}</strong>
        <p>${t("actorInputs", { n: actor.totalInputs })}</p>
      </div>
      <span>${t("actorReplicas", { n: actor.totalReplicas })}</span>
    `;
    elements.actorSummary.appendChild(article);
  });
}

function renderStatus() {
  if (!state.analysis) return;
  const missing = state.analysis.missing;
  elements.statusBox.className = `status-box${missing.characters ? " is-warning" : " is-ok"}`;
  elements.statusBox.textContent = missing.characters
    ? t("statusMissing", { chars: missing.characters, inputs: missing.inputs, replicas: missing.replicas })
    : t("statusOk");
}

function renderAnalysis() {
  renderBadges();
  renderCastingTable();
  renderActorSummary();
  renderStatus();
}

function escapeAttribute(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

// ─── Language switching ───────────────────────────────────────────────────────

function switchLanguage(lang) {
  if (currentLang === lang) return;
  currentLang = lang;
  localStorage.setItem("obsazovani-lang", lang);
  applyI18n();
  renderAnalysis();
}

elements.langCsButton.addEventListener("click", () => switchLanguage("cs"));
elements.langEnButton.addEventListener("click", () => switchLanguage("en"));

// ─── Event listeners ─────────────────────────────────────────────────────────

elements.title.addEventListener("input", (event) => {
  state.title = event.target.value;
  saveDraft();
});

elements.episodeContent.addEventListener("input", (event) => {
  state.episodes[state.currentEpisode].content = event.target.value;
  saveDraft();
});

elements.clearCurrentButton.addEventListener("click", () => {
  state.episodes[state.currentEpisode].content = "";
  updateEpisodeEditor();
  scheduleAnalyze();
});

elements.episodeFileInput.addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  const content = await file.text();
  state.episodes[state.currentEpisode].content = content;
  updateEpisodeEditor();
  scheduleAnalyze();
  event.target.value = "";
});

elements.analyzeButton.addEventListener("click", analyze);
elements.exportButton.addEventListener("click", exportWorkbook);

// ─── Init ─────────────────────────────────────────────────────────────────────

loadDraft();
currentLang = localStorage.getItem("obsazovani-lang") || "cs";
elements.title.value = state.title;
applyI18n();
renderEpisodeTabs();
updateEpisodeEditor();
analyze();
