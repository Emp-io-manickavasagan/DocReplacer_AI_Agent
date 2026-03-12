/**
 * aiProviders.js
 * ─────────────────────────────────────────────────────────────────────────
 * Unified AI provider service for DocForge / DocAgent.
 *
 * Supported providers (free tier):
 *   • Cohere  — command-r-plus-08-2024  (4096 max output tokens)
 *   • Gemini  — gemini-2.0-flash        (generous free quota)
 *
 * Usage:
 *   import { askAI, PROVIDERS } from "./aiProviders";
 *
 *   const result = await askAI({
 *     provider  : "cohere",          // "cohere" | "gemini"
 *     systemMsg : "You are ...",
 *     userMsg   : "Modify this JSON chunk...",
 *   });
 *
 *   if (result.ok) console.log(result.text);
 *   else           console.error(result.error);
 * ─────────────────────────────────────────────────────────────────────────
 */

/* ── Provider registry ─────────────────────────────────────────────────── */

export const PROVIDERS = {
  cohere: {
    id      : "cohere",
    label   : "Cohere",
    model   : "command-r-plus-08-2024",
    maxTokens: 4096,
    color   : "#8B5CF6",   // violet
    badge   : "command-r-plus",
    envKey  : "VITE_COHERE_API_KEY",
    docsUrl : "https://dashboard.cohere.com/api-keys",
  },
  gemini: {
    id      : "gemini",
    label   : "Gemini",
    model   : "gemini-2.0-flash",
    maxTokens: 8192,
    color   : "#0EA5E9",   // sky blue
    badge   : "gemini-2.0-flash",
    envKey  : "VITE_GEMINI_API_KEY",
    docsUrl : "https://aistudio.google.com/app/apikey",
  },
};

/* ── Key helpers ───────────────────────────────────────────────────────── */

/** Returns the API key for a provider from import.meta.env */
function getKey(providerId) {
  const keys = {
    cohere : import.meta.env.VITE_COHERE_API_KEY,
    gemini : import.meta.env.VITE_GEMINI_API_KEY,
  };
  return keys[providerId] || "";
}

/** Checks whether a key looks valid (non-empty, not placeholder) */
export function isKeyConfigured(providerId) {
  const k = getKey(providerId);
  return !!k && k !== "your_cohere_api_key_here" && k !== "your_gemini_api_key_here";
}

/* ── System prompt ─────────────────────────────────────────────────────── */

const DEFAULT_SYSTEM = `You are a precise DOCX XML editor assistant.
You will receive a JSON object representing a Word document paragraph (w:p node).
Your job:
  1. Read the "instruction" field and apply ONLY the requested change.
  2. Modify ONLY _text values — never change _tag or _attrs fields.
  3. Return ONLY the updated paragraph JSON object.
  4. Output must be valid JSON — no markdown fences, no explanation text.
  5. Preserve every _tag, _attrs, and _children structure exactly.
  6. If no change is needed, return the currentParagraph unchanged.`;

/* ── Cohere call ───────────────────────────────────────────────────────── */

async function callCohere({ userMsg, systemMsg }) {
  const key = getKey("cohere");
  if (!key || key === "your_cohere_api_key_here") {
    return { ok: false, error: "Cohere API key not set. Add VITE_COHERE_API_KEY to your .env file." };
  }

  const body = {
    model     : PROVIDERS.cohere.model,
    max_tokens: PROVIDERS.cohere.maxTokens,
    messages  : [
      { role: "user", content: userMsg },
    ],
    system    : systemMsg || DEFAULT_SYSTEM,
  };

  try {
    const res = await fetch("https://api.cohere.com/v2/chat", {
      method : "POST",
      headers: {
        "Content-Type" : "application/json",
        "Authorization": `Bearer ${key}`,
        "X-Client-Name": "docforge",
      },
      body: JSON.stringify(body),
    });

    const data = await res.json();

    if (!res.ok) {
      const msg = data?.message || data?.error || `HTTP ${res.status}`;
      return { ok: false, error: `Cohere error: ${msg}` };
    }

    // v2 response: data.message.content[0].text
    const text = data?.message?.content?.[0]?.text
      || data?.text
      || "";

    if (!text) return { ok: false, error: "Cohere returned an empty response." };
    return { ok: true, text: text.trim(), provider: "cohere", model: PROVIDERS.cohere.model };

  } catch (e) {
    return { ok: false, error: `Cohere fetch failed: ${e.message}` };
  }
}

/* ── Gemini call ───────────────────────────────────────────────────────── */

async function callGemini({ userMsg, systemMsg }) {
  const key = getKey("gemini");
  if (!key || key === "your_gemini_api_key_here") {
    return { ok: false, error: "Gemini API key not set. Add VITE_GEMINI_API_KEY to your .env file." };
  }

  const model   = PROVIDERS.gemini.model;
  const url     = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;

  const body = {
    system_instruction: {
      parts: [{ text: systemMsg || DEFAULT_SYSTEM }],
    },
    contents: [
      {
        role : "user",
        parts: [{ text: userMsg }],
      },
    ],
    generationConfig: {
      maxOutputTokens : PROVIDERS.gemini.maxTokens,
      temperature     : 0.1,   // low temp for deterministic JSON edits
      responseMimeType: "application/json",
    },
  };

  try {
    const res  = await fetch(url, {
      method : "POST",
      headers: { "Content-Type": "application/json" },
      body   : JSON.stringify(body),
    });

    const data = await res.json();

    if (!res.ok) {
      const msg = data?.error?.message || `HTTP ${res.status}`;
      return { ok: false, error: `Gemini error: ${msg}` };
    }

    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    if (!text) return { ok: false, error: "Gemini returned an empty response." };

    return { ok: true, text: text.trim(), provider: "gemini", model };

  } catch (e) {
    return { ok: false, error: `Gemini fetch failed: ${e.message}` };
  }
}

/* ── Main unified call ─────────────────────────────────────────────────── */

/**
 * Calls the selected AI provider and returns { ok, text, provider, model }
 * or { ok: false, error }.
 *
 * @param {object} opts
 * @param {"cohere"|"gemini"} opts.provider
 * @param {string} opts.userMsg     — the full user message / JSON chunk
 * @param {string} [opts.systemMsg] — optional system override
 */
export async function askAI({ provider, userMsg, systemMsg }) {
  if (!provider || !PROVIDERS[provider]) {
    return { ok: false, error: `Unknown provider: "${provider}". Use "cohere" or "gemini".` };
  }
  if (!userMsg?.trim()) {
    return { ok: false, error: "userMsg is empty." };
  }

  switch (provider) {
    case "cohere": return callCohere({ userMsg, systemMsg });
    case "gemini": return callGemini({ userMsg, systemMsg });
    default:       return { ok: false, error: `Provider "${provider}" not implemented.` };
  }
}

/**
 * Strips markdown code fences if the model wrapped its response in them.
 * Safe to call even if the text is already clean JSON.
 */
export function stripFences(text) {
  return text
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();
}