/**
 * aiProviders.js
 * Gemini + Cohere provider config and API call implementations.
 * API keys: read from .env (VITE_GEMINI_API_KEY / VITE_COHERE_API_KEY)
 *           OR passed in directly (fallback for when env is missing)
 */

/* ─── .env keys (requires `vite` restart after editing .env) ─── */
export const ENV_KEYS = {
  gemini: import.meta.env.VITE_GEMINI_API_KEY || "",
  cohere: import.meta.env.VITE_COHERE_API_KEY || "",
};

/* ─── Provider definitions ─── */
export const PROVIDERS = {
  gemini: {
    label   : "Google Gemini",
    icon    : "✦",
    color   : "#1a73e8",
    gradient: "linear-gradient(135deg,#1a73e8,#0f4c8a)",
    models  : [
      "gemini-2.0-flash",
      "gemini-2.0-flash-lite",
      "gemini-1.5-flash",
      "gemini-1.5-pro",
    ],
    envVar  : "VITE_GEMINI_API_KEY",
  },
  cohere: {
    label   : "Cohere",
    icon    : "◈",
    color   : "#39594d",
    gradient: "linear-gradient(135deg,#39594d,#1a3a30)",
    models  : [
      "command-a-03-2025",
      "command-r-plus",
      "command-r",
      "command-r7b-12-2024",
    ],
    envVar  : "VITE_COHERE_API_KEY",
  },
};

/* ─── Tool schemas ─── */
export const SEARCH_SCHEMA = {
  type: "object",
  properties: {
    keywords: {
      type: "array",
      items: { type: "string" },
      description: "Keywords or phrases to find in the document paragraphs",
    },
  },
  required: ["keywords"],
};

export const PATCH_SCHEMA = {
  type: "object",
  properties: {
    patches: {
      type: "array",
      items: {
        type: "object",
        properties: {
          nodeId      : { type: "string", description: "nodeId from the search result — copy exactly" },
          paraPath    : { type: "array",  items: {}, description: "paragraphPath from the search result — copy exactly" },
          newParagraph: { type: "object", description: "Modified paragraph JSON. CRITICAL: keep ALL _tag and _attrs exactly as-is. Only change _text values." },
        },
        required: ["nodeId", "paraPath", "newParagraph"],
      },
    },
  },
  required: ["patches"],
};

/* ─── System prompt ─── */
export const SYSTEM_PROMPT =
`You are a precise DOCX document editor. The user will give you editing instructions.

Workflow:
1. Call search_paragraphs with ALL relevant keywords in ONE call.
2. Review the returned paragraph JSON carefully.
3. Call patch_paragraphs with ALL modifications in ONE call.
   RULES: Copy nodeId and paragraphPath exactly. Keep every _tag/_attrs exactly as-is. Only change _text values.
4. Briefly confirm what you changed.

Never ask the user to search manually — use the tools yourself.`;

/* ═══════════════════════════════════════════════════════════
   GEMINI — generateContent with functionDeclarations
═══════════════════════════════════════════════════════════ */
export async function callGemini(model, apiKey, history, systemPrompt = SYSTEM_PROMPT) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const tools = [{
    functionDeclarations: [
      {
        name       : "search_paragraphs",
        description: "Search the DOCX paragraph index. Call first to find paragraphs to edit. Pass all keywords in one call.",
        parameters : SEARCH_SCHEMA,
      },
      {
        name       : "patch_paragraphs",
        description: "Apply edits to document paragraphs. Keep _tag/_attrs exactly as-is, only change _text values.",
        parameters : PATCH_SCHEMA,
      },
    ],
  }];

  const res = await fetch(url, {
    method : "POST",
    headers: { "Content-Type": "application/json" },
    body   : JSON.stringify({
      systemInstruction: { parts: [{ text: systemPrompt }] },
      tools,
      contents        : history,
      generationConfig: { maxOutputTokens: 8192, temperature: 0.1 },
    }),
  });

  if (!res.ok) {
    const t = await res.text().catch(() => res.statusText);
    let msg = t;
    try { msg = JSON.parse(t)?.error?.message || msg; } catch {}
    throw new Error(`Gemini ${res.status}: ${msg}`);
  }

  const data  = await res.json();
  const parts = data.candidates?.[0]?.content?.parts || [];

  const toolCalls = parts
    .filter(p => p.functionCall)
    .map(p => ({ id: p.functionCall.name + "_" + Date.now(), name: p.functionCall.name, input: p.functionCall.args || {} }));

  const text = parts.filter(p => p.text).map(p => p.text).join("\n").trim();

  return { toolCalls, text, done: toolCalls.length === 0, rawParts: parts };
}

/* ═══════════════════════════════════════════════════════════
   COHERE — /v2/chat with OpenAI-style tools
═══════════════════════════════════════════════════════════ */
export async function callCohere(model, apiKey, messages) {
  const tools = [
    {
      type    : "function",
      function: {
        name       : "search_paragraphs",
        description: "Search the DOCX paragraph index to find paragraphs to edit.",
        parameters : SEARCH_SCHEMA,
      },
    },
    {
      type    : "function",
      function: {
        name       : "patch_paragraphs",
        description: "Apply edits to document paragraphs. Keep _tag/_attrs exactly as-is.",
        parameters : PATCH_SCHEMA,
      },
    },
  ];

  const res = await fetch("https://api.cohere.com/v2/chat", {
    method : "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization : `Bearer ${apiKey}`,
    },
    body: JSON.stringify({ model, messages, tools, max_tokens: 8192, temperature: 0.1 }),
  });

  if (!res.ok) {
    const t = await res.text().catch(() => res.statusText);
    let msg = t;
    try { msg = JSON.parse(t)?.message || msg; } catch {}
    throw new Error(`Cohere ${res.status}: ${msg}`);
  }

  const data    = await res.json();
  const content = Array.isArray(data.message?.content) ? data.message.content : [];

  const toolCalls = content
    .filter(b => b.type === "tool_call")
    .map(b => ({
      id   : b.id || b.tool_call?.id || (b.tool_call?.function?.name + "_" + Date.now()),
      name : b.tool_call?.function?.name || "",
      input: (() => { try { return JSON.parse(b.tool_call?.function?.arguments || "{}"); } catch { return {}; } })(),
    }));

  const text = content.filter(b => b.type === "text").map(b => b.text).join("\n").trim();
  const done = data.finish_reason !== "TOOL_CALL" && toolCalls.length === 0;

  return { toolCalls, text, done, rawMsg: data.message };
}