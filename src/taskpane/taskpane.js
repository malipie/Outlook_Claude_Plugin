import "./taskpane.css";
import { getProvider, getAllProviders } from "./providers.js";

// --- Constants ---
const STORAGE_KEYS = {
  provider: "ai_provider",
  system: "ai_system_prompt",
  style: "ai_writing_style",
};

const DEFAULT_SYSTEM_PROMPT = `Jesteś asystentem do pisania odpowiedzi na maile.

Zasady:
- Odpowiadaj W TYM SAMYM JĘZYKU co oryginalny mail
- Pisz TYLKO treść odpowiedzi — bez tematu, bez nagłówków
- Nie dodawaj stopki, podpisu ani pozdrowień na końcu (chyba że styl pisania mówi inaczej)
- Bądź zwięzły — pisz krótko i na temat, bez zbędnych uprzejmości
- Pisz naturalnie, jak człowiek — nie jak korporacyjny bot
- Nie zaczynaj od "Dziękuję za maila" ani podobnych formułek`;

const LOADING_LABELS = {
  anthropic: "Claude myśli...",
  openai: "GPT generuje...",
  gemini: "Gemini pracuje...",
};

/** @param {string} providerId */
function providerApiKeyKey(providerId) {
  return `${providerId}_api_key`;
}

/** @param {string} providerId */
function providerModelKey(providerId) {
  return `${providerId}_model`;
}

// --- DOM References ---
const els = {};

function cacheDom() {
  els.settingsPanel = document.getElementById("settings-panel");
  els.mainPanel = document.getElementById("main-panel");
  els.noKeyWarning = document.getElementById("no-key-warning");
  els.errorBar = document.getElementById("error-bar");
  els.responseArea = document.getElementById("response-area");
  els.responseText = document.getElementById("response-text");

  els.btnSettings = document.getElementById("btn-settings");
  els.btnSaveSettings = document.getElementById("btn-save-settings");
  els.btnOpenSettings = document.getElementById("btn-open-settings");
  els.btnToggleKey = document.getElementById("btn-toggle-key");
  els.btnGenerate = document.getElementById("btn-generate");
  els.btnGenerateText = document.getElementById("btn-generate-text");
  els.btnGenerateSpinner = document.getElementById("btn-generate-spinner");
  els.btnCopy = document.getElementById("btn-copy");
  els.btnInsert = document.getElementById("btn-insert");

  els.selectProvider = document.getElementById("select-provider");
  els.inputApiKey = document.getElementById("input-api-key");
  els.selectModel = document.getElementById("select-model");
  els.textareaSystem = document.getElementById("textarea-system");
  els.btnResetSystem = document.getElementById("btn-reset-system");
  els.textareaStyle = document.getElementById("textarea-style");
  els.textareaContext = document.getElementById("textarea-context");
}

// --- State ---
let isGenerating = false;
let isOutlook = false;
let isComposeMode = false;

// --- Initialization ---

Office.onReady((info) => {
  isOutlook = info.host === Office.HostType.Outlook;
  cacheDom();
  migrateOldSettings();
  loadSettings();
  bindEvents();
  detectMode();
  updateView();
});

/**
 * Detect whether we are in read mode or compose mode.
 * In compose mode, item.subject has getAsync method (async property).
 * In read mode, item.subject is a plain string.
 */
function detectMode() {
  if (!isOutlook) return;
  const item = Office.context.mailbox.item;
  if (!item) return;
  isComposeMode = typeof item.subject !== "string";
}

// --- Migration from old single-provider localStorage keys ---

function migrateOldSettings() {
  // Migrate from old claude-* keys
  const oldKey = localStorage.getItem("claude-api-key");
  if (oldKey) {
    if (!localStorage.getItem(providerApiKeyKey("anthropic"))) {
      localStorage.setItem(providerApiKeyKey("anthropic"), oldKey);
    }

    const oldModel = localStorage.getItem("claude-model");
    if (oldModel && !localStorage.getItem(providerModelKey("anthropic"))) {
      localStorage.setItem(providerModelKey("anthropic"), oldModel);
    }

    const oldStyle = localStorage.getItem("claude-writing-style");
    if (oldStyle && !localStorage.getItem(STORAGE_KEYS.style)) {
      localStorage.setItem(STORAGE_KEYS.style, oldStyle);
    }

    if (!localStorage.getItem(STORAGE_KEYS.provider)) {
      localStorage.setItem(STORAGE_KEYS.provider, "anthropic");
    }

    localStorage.removeItem("claude-api-key");
    localStorage.removeItem("claude-model");
    localStorage.removeItem("claude-writing-style");
  }

  // Migrate: old ai_system_prompt held writing style, move it to ai_writing_style
  const oldSystemValue = localStorage.getItem(STORAGE_KEYS.system);
  if (oldSystemValue && !localStorage.getItem(STORAGE_KEYS.style)) {
    localStorage.setItem(STORAGE_KEYS.style, oldSystemValue);
    localStorage.removeItem(STORAGE_KEYS.system);
  }
}

// --- Settings ---

function loadSettings() {
  const providerId = localStorage.getItem(STORAGE_KEYS.provider) || "anthropic";
  const systemPrompt = localStorage.getItem(STORAGE_KEYS.system);
  const style = localStorage.getItem(STORAGE_KEYS.style) || "";

  els.selectProvider.value = providerId;
  els.textareaSystem.value = systemPrompt !== null ? systemPrompt : DEFAULT_SYSTEM_PROMPT;
  els.textareaStyle.value = style;

  populateProviderFields(providerId);
}

/**
 * Load API key, placeholder, and model list for the selected provider.
 * @param {string} providerId
 */
function populateProviderFields(providerId) {
  const provider = getProvider(providerId);

  const savedKey = localStorage.getItem(providerApiKeyKey(providerId)) || "";
  els.inputApiKey.value = savedKey;
  els.inputApiKey.placeholder = provider.placeholder;

  const savedModel = localStorage.getItem(providerModelKey(providerId));
  els.selectModel.innerHTML = "";
  provider.models.forEach((m) => {
    const opt = document.createElement("option");
    opt.value = m.id;
    opt.textContent = m.label;
    if (m.id === savedModel) opt.selected = true;
    els.selectModel.appendChild(opt);
  });
}

function saveSettings() {
  const providerId = els.selectProvider.value;
  const apiKey = els.inputApiKey.value.trim();
  const model = els.selectModel.value;
  const systemPrompt = els.textareaSystem.value.trim();
  const style = els.textareaStyle.value.trim();

  localStorage.setItem(STORAGE_KEYS.provider, providerId);
  localStorage.setItem(providerApiKeyKey(providerId), apiKey);
  localStorage.setItem(providerModelKey(providerId), model);
  localStorage.setItem(STORAGE_KEYS.system, systemPrompt);
  localStorage.setItem(STORAGE_KEYS.style, style);

  els.settingsPanel.classList.add("hidden");
  updateView();
}

function hasApiKey() {
  const providerId = localStorage.getItem(STORAGE_KEYS.provider) || "anthropic";
  return Boolean(localStorage.getItem(providerApiKeyKey(providerId)));
}

function updateView() {
  if (hasApiKey()) {
    els.noKeyWarning.classList.add("hidden");
    els.mainPanel.classList.remove("hidden");
  } else {
    els.noKeyWarning.classList.remove("hidden");
    els.mainPanel.classList.add("hidden");
  }

  // Adjust insert button label based on mode
  if (isComposeMode) {
    els.btnInsert.textContent = "Wstaw do treści";
  }
}

// --- Events ---

function bindEvents() {
  els.btnSettings.addEventListener("click", () => {
    els.settingsPanel.classList.toggle("hidden");
  });

  els.btnOpenSettings.addEventListener("click", () => {
    els.settingsPanel.classList.remove("hidden");
  });

  els.btnSaveSettings.addEventListener("click", saveSettings);

  els.btnToggleKey.addEventListener("click", () => {
    els.inputApiKey.type = els.inputApiKey.type === "password" ? "text" : "password";
  });

  els.selectProvider.addEventListener("change", (e) => {
    populateProviderFields(e.target.value);
  });

  els.btnResetSystem.addEventListener("click", () => {
    els.textareaSystem.value = DEFAULT_SYSTEM_PROMPT;
  });

  els.btnGenerate.addEventListener("click", handleGenerate);
  els.btnCopy.addEventListener("click", handleCopy);
  els.btnInsert.addEventListener("click", handleInsert);
}

// --- Email Reading ---

/**
 * Read email in READ mode (viewing a received message).
 * @returns {Promise<{subject: string, from: string, fromEmail: string, body: string}>}
 */
function readEmailReadMode() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("Nie wybrano żadnego maila."));
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({
          subject: item.subject || "(brak tematu)",
          from: item.from ? item.from.displayName : "Nieznany",
          fromEmail: item.from ? item.from.emailAddress : "",
          body: result.value,
        });
      } else {
        reject(new Error("Nie udało się odczytać treści maila."));
      }
    });
  });
}

/**
 * Read email in COMPOSE mode (replying/composing).
 * The body contains the quoted original message.
 * @returns {Promise<{subject: string, from: string, fromEmail: string, body: string}>}
 */
function readEmailComposeMode() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("Nie wybrano żadnego maila."));
      return;
    }

    // In compose mode, subject is async
    item.subject.getAsync((subjectResult) => {
      const subject = subjectResult.status === Office.AsyncResultStatus.Succeeded
        ? subjectResult.value
        : "(brak tematu)";

      item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
        if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve({
            subject,
            from: "(z cytowanej wiadomości)",
            fromEmail: "",
            body: bodyResult.value,
          });
        } else {
          reject(new Error("Nie udało się odczytać treści maila."));
        }
      });
    });
  });
}

/**
 * Read email content — dispatches to read or compose mode handler.
 * @returns {Promise<{subject: string, from: string, fromEmail: string, body: string}>}
 */
function readEmail() {
  if (!isOutlook) {
    return Promise.reject(new Error("Otwórz ten panel w Outlooku, aby odczytać maila."));
  }
  return isComposeMode ? readEmailComposeMode() : readEmailReadMode();
}

// --- LLM API Call ---

/**
 * Build the full system prompt from saved system instructions, style, and per-reply context.
 * @param {string} systemPrompt - base system instructions from settings
 * @param {string} style - personal writing style
 * @param {string} extraContext - per-reply additional instructions
 * @returns {string}
 */
function buildSystemPrompt(systemPrompt, style, extraContext) {
  let prompt = systemPrompt || DEFAULT_SYSTEM_PROMPT;

  if (style) {
    prompt += `\n\nStyl pisania użytkownika:\n${style}`;
  }
  if (extraContext) {
    prompt += `\n\nDodatkowe instrukcje do tej odpowiedzi:\n${extraContext}`;
  }

  return prompt;
}

/**
 * Call the selected LLM provider to generate an email reply.
 * @param {{subject: string, from: string, fromEmail: string, body: string}} email
 * @param {string} extraContext
 * @returns {Promise<string>}
 */
async function callLLM(email, extraContext) {
  const providerId = localStorage.getItem(STORAGE_KEYS.provider) || "anthropic";
  const provider = getProvider(providerId);

  const apiKey = localStorage.getItem(providerApiKeyKey(providerId));
  if (!apiKey) {
    throw new Error(`Brak klucza API dla ${provider.name} — wpisz go w ustawieniach.`);
  }

  const model = localStorage.getItem(providerModelKey(providerId)) || provider.models[0].id;
  const savedSystem = localStorage.getItem(STORAGE_KEYS.system);
  const style = localStorage.getItem(STORAGE_KEYS.style) || "";

  const systemPrompt = buildSystemPrompt(savedSystem, style, extraContext);
  const userMessage = `From: ${email.from} <${email.fromEmail}>\nSubject: ${email.subject}\n\n${email.body}`;

  const { url, headers, body } = provider.buildRequest(apiKey, model, systemPrompt, userMessage);

  const response = await fetch(url, { method: "POST", headers, body });
  const data = await response.json();

  if (!response.ok) {
    const msg = data?.error?.message || `HTTP ${response.status}`;
    throw new Error(`Błąd ${provider.name}: ${msg}`);
  }

  return provider.parseResponse(data);
}

// --- Action Handlers ---

async function handleGenerate() {
  if (isGenerating) return;

  hideError();
  els.responseArea.classList.add("hidden");

  try {
    setLoading(true);

    const email = await readEmail();
    const extraContext = els.textareaContext.value.trim();
    const reply = await callLLM(email, extraContext);

    els.responseText.textContent = reply;
    els.responseArea.classList.remove("hidden");
  } catch (err) {
    showError(err.message);
  } finally {
    setLoading(false);
  }
}

async function handleCopy() {
  const text = els.responseText.textContent;
  if (!text) return;

  try {
    await navigator.clipboard.writeText(text);
    els.btnCopy.textContent = "Skopiowano!";
    els.btnCopy.classList.add("btn-copied");
    setTimeout(() => {
      els.btnCopy.textContent = "Kopiuj";
      els.btnCopy.classList.remove("btn-copied");
    }, 1500);
  } catch {
    showError("Nie udało się skopiować do schowka.");
  }
}

function handleInsert() {
  const text = els.responseText.textContent;
  if (!text) return;

  if (!isOutlook) {
    showError("Wstawianie działa tylko w Outlooku.");
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) {
    showError("Nie wybrano maila.");
    return;
  }

  if (isComposeMode) {
    // Compose mode: insert text at cursor position (top of body)
    item.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          showInsertSuccess();
        } else {
          copyFallback(text);
        }
      }
    );
  } else {
    // Read mode: open Reply All form with generated text as body
    item.displayReplyAllForm(text);
  }
}

/** Show brief success feedback on the insert button. */
function showInsertSuccess() {
  els.btnInsert.textContent = "Wstawiono!";
  els.btnInsert.classList.add("btn-copied");
  setTimeout(() => {
    els.btnInsert.textContent = isComposeMode ? "Wstaw do treści" : "Wstaw do odpowiedzi";
    els.btnInsert.classList.remove("btn-copied");
  }, 1500);
}

/**
 * Fallback: copy to clipboard when direct insertion fails.
 * @param {string} text
 */
function copyFallback(text) {
  navigator.clipboard.writeText(text).then(() => {
    showError("Wstawienie nie powiodło się — tekst skopiowany do schowka.");
  });
}

// --- UI Helpers ---

/** @param {boolean} loading */
function setLoading(loading) {
  isGenerating = loading;
  els.btnGenerate.disabled = loading;

  if (loading) {
    const providerId = localStorage.getItem(STORAGE_KEYS.provider) || "anthropic";
    els.btnGenerateText.textContent = LOADING_LABELS[providerId] || "Generuję...";
  } else {
    els.btnGenerateText.textContent = "Wygeneruj odpowiedź";
  }

  els.btnGenerateSpinner.classList.toggle("hidden", !loading);
}

/** @param {string} msg */
function showError(msg) {
  els.errorBar.textContent = msg;
  els.errorBar.classList.remove("hidden");
}

function hideError() {
  els.errorBar.textContent = "";
  els.errorBar.classList.add("hidden");
}
