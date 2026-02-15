/**
 * Provider Adapter Pattern for LLM APIs.
 *
 * Each provider implements a unified interface:
 *   - name (string)          — display name
 *   - models (array)         — [{id, label}]
 *   - placeholder (string)   — API key input placeholder
 *   - buildRequest(apiKey, model, systemPrompt, userMessage) → {url, headers, body}
 *   - parseResponse(jsonData) → string (extracted reply text)
 *
 * Adding a new provider requires ONLY:
 *   1. A new adapter object with the methods above
 *   2. An entry in the PROVIDERS registry
 * No UI or taskpane.js changes needed.
 */

// ============================================================
// ANTHROPIC CLAUDE
// Requires `anthropic-dangerous-direct-browser-access: true`
// header for CORS to work from browser.
// ============================================================
const AnthropicProvider = {
  name: "Anthropic Claude",
  placeholder: "sk-ant-... (z console.anthropic.com)",
  models: [
    { id: "claude-sonnet-4-20250514", label: "Claude Sonnet 4 (rekomendowany)" },
    { id: "claude-opus-4-0-20250115", label: "Claude Opus 4" },
    { id: "claude-haiku-4-5-20251001", label: "Claude Haiku 4.5 (najtańszy)" },
  ],
  buildRequest(apiKey, model, systemPrompt, userMessage) {
    return {
      url: "https://api.anthropic.com/v1/messages",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model,
        max_tokens: 2048,
        system: systemPrompt,
        messages: [{ role: "user", content: userMessage }],
      }),
    };
  },
  parseResponse(data) {
    if (data.error) throw new Error(data.error.message);
    return data.content[0].text;
  },
};

// ============================================================
// OPENAI GPT
// Works from browser with standard Bearer token, no extra
// CORS headers needed.
// ============================================================
const OpenAIProvider = {
  name: "OpenAI GPT",
  placeholder: "sk-... (z platform.openai.com)",
  models: [
    { id: "gpt-4o", label: "GPT-4o (rekomendowany)" },
    { id: "gpt-4o-mini", label: "GPT-4o Mini (najtańszy)" },
    { id: "gpt-4.1", label: "GPT-4.1" },
    { id: "o3-mini", label: "o3-mini (reasoning)" },
  ],
  buildRequest(apiKey, model, systemPrompt, userMessage) {
    return {
      url: "https://api.openai.com/v1/chat/completions",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        max_tokens: 2048,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userMessage },
        ],
      }),
    };
  },
  parseResponse(data) {
    if (data.error) throw new Error(data.error.message);
    return data.choices[0].message.content;
  },
};

// ============================================================
// GOOGLE GEMINI
// API key is passed as a URL parameter (?key=), not a header.
// Works from browser without CORS issues.
// ============================================================
const GeminiProvider = {
  name: "Google Gemini",
  placeholder: "AIza... (z aistudio.google.com)",
  models: [
    { id: "gemini-2.5-flash", label: "Gemini 2.5 Flash (rekomendowany)" },
    { id: "gemini-2.5-pro", label: "Gemini 2.5 Pro" },
    { id: "gemini-2.0-flash-lite", label: "Gemini 2.0 Flash Lite (najtańszy)" },
  ],
  buildRequest(apiKey, model, systemPrompt, userMessage) {
    return {
      url: `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        systemInstruction: {
          parts: [{ text: systemPrompt }],
        },
        contents: [
          {
            parts: [{ text: userMessage }],
          },
        ],
        generationConfig: {
          maxOutputTokens: 2048,
        },
      }),
    };
  },
  parseResponse(data) {
    if (data.error) throw new Error(data.error.message);
    return data.candidates[0].content.parts[0].text;
  },
};

// ============================================================
// PROVIDER REGISTRY
// ============================================================
const PROVIDERS = {
  anthropic: AnthropicProvider,
  openai: OpenAIProvider,
  gemini: GeminiProvider,
};

/** @param {string} id */
function getProvider(id) {
  return PROVIDERS[id] || AnthropicProvider;
}

function getAllProviders() {
  return Object.entries(PROVIDERS).map(([id, p]) => ({
    id,
    name: p.name,
    models: p.models,
  }));
}

export { PROVIDERS, getProvider, getAllProviders };
