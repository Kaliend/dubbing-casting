const state = {
  title: "Obsazení projektu",
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
    const problem = await response.json().catch(() => ({ error: "Neznámá chyba." }));
    throw new Error(problem.error || "Neznámá chyba.");
  }
  if (responseType === "blob") {
    return response.blob();
  }
  return response.json();
}

async function analyze() {
  try {
    elements.analyzeButton.disabled = true;
    elements.analyzeButton.textContent = "Počítám…";
    state.analysis = await postJson("/api/analyze", payload());
    renderAnalysis();
  } catch (error) {
    elements.statusBox.textContent = error.message;
    elements.statusBox.className = "status-box is-error";
  } finally {
    elements.analyzeButton.disabled = false;
    elements.analyzeButton.textContent = "Přepočítat";
  }
}

async function exportWorkbook() {
  try {
    elements.exportButton.disabled = true;
    elements.exportButton.textContent = "Exportuji…";
    const blob = await postJson("/api/export", payload(), "blob");
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${(state.title || "obsazeni").replace(/\s+/g, "-")}.xlsx`;
    anchor.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    elements.statusBox.textContent = error.message;
    elements.statusBox.className = "status-box is-error";
  } finally {
    elements.exportButton.disabled = false;
    elements.exportButton.textContent = "Export XLSX";
  }
}

function scheduleAnalyze() {
  saveDraft();
  window.clearTimeout(scheduleAnalyze.timer);
  scheduleAnalyze.timer = window.setTimeout(analyze, 180);
}

function updateEpisodeEditor() {
  const episode = state.episodes[state.currentEpisode];
  elements.episodeLabel.textContent = `Epizoda ${episode.label}`;
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
    ["Postavy", state.analysis.stats.characterCount],
    ["Vstupy", state.analysis.stats.inputs],
    ["Repliky", state.analysis.stats.replicas],
    ["Neobsazeno", state.analysis.missing.characters],
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
      <th>Postava</th>
      ${episodeHeaders}
      <th>Vstupy</th>
      <th>Repliky</th>
      <th>Dabér</th>
      <th>Poznámka</th>
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
        <p>${actor.totalInputs} vstupů</p>
      </div>
      <span>${actor.totalReplicas} replik</span>
    `;
    elements.actorSummary.appendChild(article);
  });
}

function renderStatus() {
  if (!state.analysis) return;
  const missing = state.analysis.missing;
  elements.statusBox.className = `status-box${missing.characters ? " is-warning" : " is-ok"}`;
  elements.statusBox.textContent = missing.characters
    ? `Chybí obsadit ${missing.characters} postav, ${missing.inputs} vstupů a ${missing.replicas} replik.`
    : "Všechny postavy mají přiřazeného dabéra.";
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

loadDraft();
elements.title.value = state.title;
renderEpisodeTabs();
updateEpisodeEditor();
analyze();
