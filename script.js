"use strict";

const DATABASE_URL = "database.xlsx";

const state = {
  items: [],
  order: [],
  currentPosition: 0,
  currentItem: null,
  revealedIndexes: new Set(),
  correct: 0,
  wrong: 0,
  streak: 0,
  answeredCurrent: false,
  hintShown: false,
};

const el = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  loadDatabase();
});

function cacheElements() {
  [
    "gameLayout", "progressText", "progressBar", "gameTitle", "wordLengthText",
    "secondaryHint", "secondaryHintText", "maskedWord", "guessForm", "guessInput",
    "submitButton", "feedbackText", "hintButton", "skipButton", "correctStat",
    "streakStat", "wrongStat", "remainingStat", "accuracyStat", "accuracyBar",
    "databaseStatus", "restartButton", "reloadDatabaseButton", "helpButton",
    "helpDialog", "closeHelpButton", "understoodButton", "endDialog", "endSummary",
    "endScore", "playAgainButton", "loadingLayer", "loadingTitle", "loadingMessage",
    "databaseDialog", "databaseErrorText", "databaseFileInput"
  ].forEach((id) => {
    el[id] = document.getElementById(id);
  });
}

function bindEvents() {
  el.guessForm.addEventListener("submit", handleGuess);
  el.skipButton.addEventListener("click", skipCurrent);
  el.hintButton.addEventListener("click", showSecondaryHint);
  el.restartButton.addEventListener("click", () => startNewGame(true));
  el.reloadDatabaseButton.addEventListener("click", () => loadDatabase(true));

  el.helpButton.addEventListener("click", () => el.helpDialog.showModal());
  el.closeHelpButton.addEventListener("click", () => el.helpDialog.close());
  el.understoodButton.addEventListener("click", () => el.helpDialog.close());

  el.playAgainButton.addEventListener("click", () => {
    el.endDialog.close();
    startNewGame(true);
  });

  el.databaseFileInput.addEventListener("change", handleManualDatabaseFile);

  [el.helpDialog, el.endDialog, el.databaseDialog].forEach((dialog) => {
    dialog.addEventListener("click", (event) => {
      if (event.target === dialog) dialog.close();
    });
  });
}

async function loadDatabase(isReload = false) {
  showLoading(
    isReload ? "Atualizando database.xlsx" : "Carregando database.xlsx",
    isReload ? "Buscando a versão mais recente da planilha." : "Preparando as palavras da partida."
  );

  try {
    if (typeof XLSX === "undefined") {
      throw new Error("A biblioteca de leitura da planilha não foi carregada.");
    }

    const separator = DATABASE_URL.includes("?") ? "&" : "?";
    const response = await fetch(`${DATABASE_URL}${separator}v=${Date.now()}`, {
      cache: "no-store",
    });

    if (!response.ok) {
      throw new Error(`Falha ao carregar a planilha (HTTP ${response.status}).`);
    }

    const buffer = await response.arrayBuffer();
    const items = parseWorkbook(buffer);
    setDatabase(items, "database.xlsx");
  } catch (error) {
    console.error(error);
    hideLoading();
    el.gameLayout.setAttribute("aria-busy", "false");
    el.databaseErrorText.textContent =
      `${error.message} Confirme se database.xlsx está na mesma pasta do site. ` +
      "Se estiver abrindo o arquivo HTML diretamente no computador, selecione a planilha manualmente.";
    if (!el.databaseDialog.open) el.databaseDialog.showModal();
  }
}

async function handleManualDatabaseFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  showLoading("Lendo planilha", `Abrindo ${file.name}.`);

  try {
    const buffer = await file.arrayBuffer();
    const items = parseWorkbook(buffer);
    setDatabase(items, file.name);
    if (el.databaseDialog.open) el.databaseDialog.close();
  } catch (error) {
    console.error(error);
    hideLoading();
    el.databaseErrorText.textContent = `Não foi possível ler a planilha: ${error.message}`;
  } finally {
    event.target.value = "";
  }
}

function parseWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error("A planilha não possui abas legíveis.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });

  if (rows.length < 2) {
    throw new Error("A planilha está vazia ou não possui linhas de dados.");
  }

  return normalizeRows(rows);
}

/**
 * Aceita automaticamente dois formatos:
 *
 * FORMATO A — igual ao exemplo enviado:
 * | (sem título) | exibicao | dica |
 * | Era          | Grande divisão... | ... |
 *
 * FORMATO B — duas colunas convencionais:
 * | exibicao | dica |
 * | Era      | Grande divisão... |
 */
function normalizeRows(rows) {
  const header = rows[0].map(normalizeHeader);
  const dataRows = rows.slice(1);

  const exibicaoIndex = header.indexOf("exibicao");
  const dicaIndex = header.indexOf("dica");
  const firstHeaderIsBlank = !header[0];

  let items = [];

  // Formato A: 1ª coluna sem cabeçalho = resposta; "exibicao" = contexto; "dica" = dica opcional.
  if (firstHeaderIsBlank && exibicaoIndex >= 0) {
    items = dataRows.map((row, index) => ({
      answer: cleanCell(row[0]),
      clue: cleanCell(row[exibicaoIndex]),
      extraHint: dicaIndex >= 0 ? cleanCell(row[dicaIndex]) : "",
      sourceRow: index + 2,
    }));
  }
  // Formato B: "exibicao" = resposta; "dica" = contexto.
  else if (exibicaoIndex >= 0 && dicaIndex >= 0) {
    items = dataRows.map((row, index) => ({
      answer: cleanCell(row[exibicaoIndex]),
      clue: cleanCell(row[dicaIndex]),
      extraHint: "",
      sourceRow: index + 2,
    }));
  }
  // Fallback: primeiras duas colunas não vazias = resposta e contexto.
  else {
    items = dataRows.map((row, index) => {
      const values = row.map(cleanCell).filter(Boolean);
      return {
        answer: values[0] || "",
        clue: values[1] || "",
        extraHint: values[2] || "",
        sourceRow: index + 2,
      };
    });
  }

  items = items.filter((item) => item.answer && item.clue);

  // Remove duplicatas exatas de resposta + contexto, sem alterar a primeira ocorrência.
  const seen = new Set();
  items = items.filter((item) => {
    const key = `${normalizeForComparison(item.answer)}::${normalizeForComparison(item.clue)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  if (!items.length) {
    throw new Error(
      "Nenhuma linha válida foi encontrada. É necessário haver uma resposta e um texto de exibição/dica em cada linha."
    );
  }

  return items;
}

function setDatabase(items, sourceName) {
  state.items = items;
  el.databaseStatus.textContent = `${items.length} ${items.length === 1 ? "item" : "itens"}`;
  el.databaseStatus.title = `Fonte: ${sourceName}`;
  el.gameLayout.setAttribute("aria-busy", "false");
  hideLoading();
  startNewGame(false);
}

function startNewGame(announce = false) {
  if (!state.items.length) return;

  state.order = shuffle([...Array(state.items.length).keys()]);
  state.currentPosition = 0;
  state.correct = 0;
  state.wrong = 0;
  state.streak = 0;
  state.answeredCurrent = false;
  state.hintShown = false;

  if (el.endDialog.open) el.endDialog.close();
  loadCurrentItem();
  updateStats();

  if (announce) setFeedback("Nova partida iniciada.", "success");
}

function loadCurrentItem() {
  if (state.currentPosition >= state.order.length) {
    finishGame();
    return;
  }

  const itemIndex = state.order[state.currentPosition];
  state.currentItem = state.items[itemIndex];
  state.answeredCurrent = false;
  state.hintShown = false;

  const answerLength = countLetters(state.currentItem.answer);
  state.revealedIndexes = chooseRevealedIndexes(state.currentItem.answer, getRevealCount(answerLength));

  el.gameTitle.textContent = state.currentItem.clue;
  el.wordLengthText.textContent = `${answerLength} ${answerLength === 1 ? "letra" : "letras"}`;
  el.secondaryHint.hidden = true;
  el.secondaryHintText.textContent = state.currentItem.extraHint || "";
  el.hintButton.hidden = !state.currentItem.extraHint;
  el.hintButton.disabled = false;
  el.hintButton.textContent = "Mostrar dica";
  el.guessInput.value = "";
  el.guessInput.disabled = false;
  el.submitButton.disabled = false;
  el.skipButton.disabled = false;
  clearFeedback();

  renderMaskedWord(false);
  updateStats();
  updateProgress();

  requestAnimationFrame(() => el.guessInput.focus({ preventScroll: true }));
}

function getRevealCount(letterCount) {
  if (letterCount <= 4) return Math.min(1, letterCount);
  if (letterCount <= 7) return 2;
  if (letterCount <= 10) return 3;
  return 4;
}

function chooseRevealedIndexes(answer, revealCount) {
  const characters = Array.from(answer);
  const letterIndexes = [];

  characters.forEach((char, index) => {
    if (isLetterOrNumber(char)) letterIndexes.push(index);
  });

  if (revealCount >= letterIndexes.length) return new Set(letterIndexes);

  // Distribui as revelações pela palavra em vez de concentrá-las no início.
  const chosen = new Set();
  const available = [...letterIndexes];

  while (chosen.size < revealCount && available.length) {
    const randomPosition = Math.floor(Math.random() * available.length);
    chosen.add(available.splice(randomPosition, 1)[0]);
  }

  return chosen;
}

function renderMaskedWord(revealAll = false) {
  el.maskedWord.innerHTML = "";
  const fragment = document.createDocumentFragment();

  Array.from(state.currentItem.answer).forEach((char, index) => {
    if (/\s/u.test(char)) {
      const space = document.createElement("span");
      space.className = "character-space";
      space.setAttribute("aria-hidden", "true");
      fragment.appendChild(space);
      return;
    }

    if (!isLetterOrNumber(char)) {
      const visible = document.createElement("span");
      visible.className = "character-visible";
      visible.textContent = char;
      visible.setAttribute("aria-hidden", "true");
      fragment.appendChild(visible);
      return;
    }

    const slot = document.createElement("span");
    const shouldReveal = revealAll || state.revealedIndexes.has(index);
    slot.className = `letter-slot${shouldReveal ? " revealed" : ""}`;
    slot.textContent = shouldReveal ? char : "";
    slot.setAttribute("aria-hidden", "true");
    fragment.appendChild(slot);
  });

  el.maskedWord.appendChild(fragment);

  const accessibleText = revealAll
    ? `Resposta: ${state.currentItem.answer}`
    : "Palavra parcialmente revelada. Use o contexto para descobrir a resposta.";
  el.maskedWord.setAttribute("aria-label", accessibleText);
}

function handleGuess(event) {
  event.preventDefault();
  if (!state.currentItem || state.answeredCurrent) return;

  const guess = el.guessInput.value.trim();
  if (!guess) {
    setFeedback("Digite uma resposta antes de confirmar.", "error");
    el.guessInput.focus();
    return;
  }

  const isCorrect = normalizeForComparison(guess) === normalizeForComparison(state.currentItem.answer);

  if (isCorrect) {
    state.correct += 1;
    state.streak += 1;
    state.answeredCurrent = true;
    renderMaskedWord(true);
    setFeedback("Correto. Avançando para a próxima palavra…", "success");
    lockCurrentRound();
    updateStats();
    window.setTimeout(goToNextItem, 850);
  } else {
    state.wrong += 1;
    state.streak = 0;
    setFeedback("Ainda não. Tente outra vez.", "error");
    updateStats();
    selectInputText();
  }
}

function skipCurrent() {
  if (!state.currentItem || state.answeredCurrent) return;

  state.wrong += 1;
  state.streak = 0;
  state.answeredCurrent = true;
  renderMaskedWord(true);
  setFeedback(`Resposta: ${state.currentItem.answer}`, "error");
  lockCurrentRound();
  updateStats();
  window.setTimeout(goToNextItem, 1250);
}

function showSecondaryHint() {
  if (!state.currentItem?.extraHint) return;
  state.hintShown = true;
  el.secondaryHint.hidden = false;
  el.hintButton.disabled = true;
  el.hintButton.textContent = "Dica exibida";
}

function lockCurrentRound() {
  el.guessInput.disabled = true;
  el.submitButton.disabled = true;
  el.skipButton.disabled = true;
  el.hintButton.disabled = true;
}

function goToNextItem() {
  state.currentPosition += 1;
  loadCurrentItem();
}

function finishGame() {
  const attempts = state.correct + state.wrong;
  const accuracy = attempts ? Math.round((state.correct / attempts) * 100) : 0;

  el.endSummary.textContent =
    `Você acertou ${state.correct} ${state.correct === 1 ? "resposta" : "respostas"} ` +
    `e registrou ${state.wrong} ${state.wrong === 1 ? "erro" : "erros"}.`;
  el.endScore.textContent = `${accuracy}%`;

  if (!el.endDialog.open) el.endDialog.showModal();
}

function updateProgress() {
  const total = state.order.length;
  const current = Math.min(state.currentPosition + 1, total);
  const percent = total ? (state.currentPosition / total) * 100 : 0;

  el.progressText.textContent = `${current} de ${total}`;
  el.progressBar.style.width = `${percent}%`;
}

function updateStats() {
  const total = state.order.length || state.items.length;
  const completed = Math.min(state.currentPosition, total);
  const remaining = Math.max(total - completed, 0);
  const attempts = state.correct + state.wrong;
  const accuracy = attempts ? Math.round((state.correct / attempts) * 100) : 0;

  el.correctStat.textContent = String(state.correct);
  el.streakStat.textContent = String(state.streak);
  el.wrongStat.textContent = String(state.wrong);
  el.remainingStat.textContent = String(remaining);
  el.accuracyStat.textContent = `${accuracy}%`;
  el.accuracyBar.style.width = `${accuracy}%`;
}

function normalizeForComparison(value) {
  return String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("pt-BR")
    .replace(/[’‘´`]/g, "'")
    .replace(/[^\p{L}\p{N}]+/gu, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function countLetters(value) {
  return Array.from(String(value)).filter(isLetterOrNumber).length;
}

function isLetterOrNumber(char) {
  return /[\p{L}\p{N}]/u.test(char);
}

function normalizeHeader(value) {
  return String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function cleanCell(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function shuffle(array) {
  for (let i = array.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}

function setFeedback(message, type = "") {
  el.feedbackText.textContent = message;
  el.feedbackText.className = `feedback${type ? ` ${type}` : ""}`;
}

function clearFeedback() {
  setFeedback("");
}

function selectInputText() {
  requestAnimationFrame(() => {
    el.guessInput.focus();
    el.guessInput.select();
  });
}

function showLoading(title, message) {
  el.loadingTitle.textContent = title;
  el.loadingMessage.textContent = message;
  el.loadingLayer.hidden = false;
  el.gameLayout.setAttribute("aria-busy", "true");
}

function hideLoading() {
  el.loadingLayer.hidden = true;
}
