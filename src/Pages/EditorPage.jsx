/**
 * EditorPage.jsx — Agentic DOCX Editor
 *
 * Providers  : Google Gemini · Cohere
 * API Keys   : Read from .env (VITE_GEMINI_API_KEY / VITE_COHERE_API_KEY)
 *
 * Flow (zero manual steps):
 *   1. User picks provider + model
 *   2. User types a natural-language editing prompt → Run Agent
 *   3. LLM calls search_paragraphs({ keywords[] }) → App searches paragraph index
 *   4. LLM calls patch_paragraphs({ patches[] })  → App applies all edits at once
 *   5. Loop until done → Export .docx
 */

import { useEffect, useState, useCallback, useRef } from "react";
import { renderAsync } from "docx-preview";
import * as JSZipModule from "jszip";
import { zipToNodeMap, nodeMapToDocxBuffer } from "../utils/xmlToJson";

const JSZip = JSZipModule.default || JSZipModule;

/* ═══════════════════════════════════════════════════════════
   PROVIDER CONFIG — keys from Vite env
═══════════════════════════════════════════════════════════ */
const ENV_GEMINI = import.meta.env.VITE_GEMINI_API_KEY || "";
const ENV_COHERE = import.meta.env.VITE_COHERE_API_KEY || "";

const PROVIDERS = {
  gemini: {
    label: "Google Gemini",
    icon: "✦",
    color: "#4285f4",
    models: [
      "gemini-2.0-flash",
      "gemini-2.0-flash-lite",
      "gemini-1.5-flash",
      "gemini-1.5-pro",
    ],
    envKey: ENV_GEMINI,
    envVar: "VITE_GEMINI_API_KEY",
  },
  cohere: {
    label: "Cohere",
    icon: "◈",
    color: "#39594d",
    models: [
      "command-a-03-2025",
      "command-r-plus",
      "command-r",
      "command-r7b-12-2024",
    ],
    envKey: ENV_COHERE,
    envVar: "VITE_COHERE_API_KEY",
  },
};

/* ═══════════════════════════════════════════════════════════
   TOOL SCHEMAS
═══════════════════════════════════════════════════════════ */
const SEARCH_SCHEMA = {
  type: "object",
  properties: {
    keywords: {
      type: "array",
      items: { type: "string" },
      description: "List of keywords or phrases to find in the document paragraphs",
    },
  },
  required: ["keywords"],
};

const PATCH_SCHEMA = {
  type: "object",
  properties: {
    patches: {
      type: "array",
      items: {
        type: "object",
        properties: {
          nodeId:       { type: "string", description: "nodeId from search result" },
          paraPath:     { type: "array",  items: {}, description: "paragraphPath from search result — copy exactly" },
          newParagraph: { type: "object", description: "Modified paragraph JSON. CRITICAL: keep ALL _tag and _attrs exactly as-is. Only change _text values." },
        },
        required: ["nodeId", "paraPath", "newParagraph"],
      },
    },
  },
  required: ["patches"],
};

/* ═══════════════════════════════════════════════════════════
   SYSTEM PROMPT
═══════════════════════════════════════════════════════════ */
const SYSTEM_PROMPT = `You are a precise DOCX document editor.
The user will give you editing instructions for a Word document.

Your workflow:
1. Call search_paragraphs with ALL relevant keywords in ONE call to find the paragraphs you need to edit.
2. Review the returned paragraph JSON carefully.
3. Call patch_paragraphs with ALL your modifications in ONE call.
   CRITICAL RULES for patching:
   - Copy nodeId and paragraphPath EXACTLY from the search result.
   - Keep every _tag and _attrs field EXACTLY as-is — never remove or change them.
   - Only modify _text string values.
   - Include ALL fields from the original paragraph.
4. Briefly confirm what you changed.

Never ask the user to search manually. Use the tools yourself.`;

/* ═══════════════════════════════════════════════════════════
   GEMINI API CALL
   Uses Gemini's function calling (generateContent endpoint)
═══════════════════════════════════════════════════════════ */
async function callGemini(model, apiKey, history, systemPrompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const tools = [{
    functionDeclarations: [
      { name: "search_paragraphs", description: "Search the DOCX paragraph index. Call this first to find paragraphs you need to edit. Pass all keywords in one call.", parameters: SEARCH_SCHEMA },
      { name: "patch_paragraphs",  description: "Apply edits to document paragraphs. Pass ALL patches in one call. Keep _tag/_attrs exactly as-is, only change _text values.", parameters: PATCH_SCHEMA },
    ],
  }];

  const body = {
    systemInstruction: { parts: [{ text: systemPrompt }] },
    tools,
    contents: history,
    generationConfig: { maxOutputTokens: 8192, temperature: 0.1 },
  };

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const t = await res.text().catch(() => res.statusText);
    let msg = t;
    try { msg = JSON.parse(t)?.error?.message || msg; } catch {}
    throw new Error(`Gemini ${res.status}: ${msg}`);
  }

  const data = await res.json();
  const parts = data.candidates?.[0]?.content?.parts || [];

  const toolCalls = parts
    .filter(p => p.functionCall)
    .map(p => ({ id: p.functionCall.name + "_" + Date.now(), name: p.functionCall.name, input: p.functionCall.args || {} }));

  const text = parts.filter(p => p.text).map(p => p.text).join("\n").trim();
  const done = toolCalls.length === 0;

  return { toolCalls, text, done, rawParts: parts, finishReason: data.candidates?.[0]?.finishReason };
}

/* ═══════════════════════════════════════════════════════════
   COHERE API CALL
   Uses Cohere's chat endpoint with tool use
═══════════════════════════════════════════════════════════ */
async function callCohere(model, apiKey, messages) {
  const url = "https://api.cohere.com/v2/chat";

  const tools = [
    {
      type: "function",
      function: {
        name: "search_paragraphs",
        description: "Search the DOCX paragraph index. Call this first to find paragraphs you need to edit.",
        parameters: SEARCH_SCHEMA,
      },
    },
    {
      type: "function",
      function: {
        name: "patch_paragraphs",
        description: "Apply edits to document paragraphs. Keep _tag/_attrs exactly as-is.",
        parameters: PATCH_SCHEMA,
      },
    },
  ];

  const body = { model, messages, tools, max_tokens: 8192, temperature: 0.1 };

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const t = await res.text().catch(() => res.statusText);
    let msg = t;
    try { msg = JSON.parse(t)?.message || msg; } catch {}
    throw new Error(`Cohere ${res.status}: ${msg}`);
  }

  const data = await res.json();
  const msg  = data.message || {};
  const content = Array.isArray(msg.content) ? msg.content : [];

  const toolCalls = content
    .filter(b => b.type === "tool_call")
    .map(b => ({
      id    : b.id || b.tool_call?.id || (b.tool_call?.function?.name + "_" + Date.now()),
      name  : b.tool_call?.function?.name || "",
      input : (() => { try { return JSON.parse(b.tool_call?.function?.arguments || "{}"); } catch { return {}; } })(),
    }));

  const text = content.filter(b => b.type === "text").map(b => b.text).join("\n").trim();
  const done = data.finish_reason !== "TOOL_CALL" && toolCalls.length === 0;

  return { toolCalls, text, done, rawMsg: msg, finishReason: data.finish_reason };
}

/* ═══════════════════════════════════════════════════════════
   PURE HELPERS
═══════════════════════════════════════════════════════════ */
const getAtPath = (obj, path) =>
  path.reduce((cur, k) => (cur != null ? cur[k] : undefined), obj);

const safeClone = (obj) =>
  JSON.parse(JSON.stringify(obj, (k, v) => (k === "_raw" ? undefined : v)));

function collectText(node, out = []) {
  if (!node || typeof node !== "object") return out;
  if (node._text) out.push(node._text);
  (node._children || []).forEach(c => collectText(c, out));
  return out;
}

/* ═══════════════════════════════════════════════════════════
   PARAGRAPH INDEX
═══════════════════════════════════════════════════════════ */
function buildParagraphIndex(masterJson) {
  const index = [];
  function walk(node, nodeId, nodePath, nodeFile, path) {
    if (!node || typeof node !== "object") return;
    if (node._tag === "w:p") {
      const fullText = collectText(node).join("").trim();
      if (fullText) index.push({ nodeId, nodePath, nodeFile, paraPath: path, fullText, paraNode: node });
      return;
    }
    (node._children || []).forEach((c, i) =>
      walk(c, nodeId, nodePath, nodeFile, [...path, "_children", i])
    );
  }
  for (const [nodeId, node] of Object.entries(masterJson.nodes)) {
    if (node.type !== "xml" || !node.content) continue;
    walk(node.content, nodeId, node.path, node.fileName, []);
  }
  return index;
}

function searchMultiple(index, keywords) {
  const byKeyword = {}, seen = new Set(), allFlat = [];
  for (const kw of keywords) {
    const q = kw.toLowerCase().trim();
    const hits = q
      ? index.filter(e => e.fullText.toLowerCase().includes(q))
              .map(e => ({ ...e, keyword: kw, key: `${e.nodeId}|${e.paraPath.join(".")}` }))
      : [];
    byKeyword[kw] = hits;
    for (const h of hits) { if (!seen.has(h.key)) { seen.add(h.key); allFlat.push(h); } }
  }
  return { byKeyword, allFlat };
}

/* ═══════════════════════════════════════════════════════════
   PATCH
═══════════════════════════════════════════════════════════ */
function applyPatch(masterJson, nodeId, paraPath, newPara) {
  const next = safeClone(masterJson);
  for (const [id, n] of Object.entries(next.nodes))
    if (n.type === "binary") n.content._raw = masterJson.nodes[id]?.content?._raw;
  const content = next.nodes[nodeId]?.content;
  if (!content) throw new Error("Node not found: " + nodeId);
  const parent = getAtPath(content, paraPath.slice(0, -1));
  if (parent == null) throw new Error("Bad paragraph path: " + JSON.stringify(paraPath));
  parent[paraPath[paraPath.length - 1]] = newPara;
  return next;
}

function applyMultiPatch(masterJson, patches) {
  let cur = masterJson;
  for (const p of patches) cur = applyPatch(cur, p.nodeId, p.paraPath, p.newPara);
  return cur;
}

/* ═══════════════════════════════════════════════════════════
   JSON TREE
═══════════════════════════════════════════════════════════ */
function JNode({ label, value, depth = 0, highlight = null }) {
  const [open, setOpen] = useState(depth < 2);
  const isObj  = value !== null && typeof value === "object" && !Array.isArray(value);
  const isArr  = Array.isArray(value);
  const isLeaf = !isObj && !isArr;
  const indent = depth * 16;
  const keys   = isObj ? Object.keys(value).filter(k => k !== "_raw") : [];
  const count  = isArr ? value.length : keys.length;

  if (isObj && value._binary) return (
    <div style={{ paddingLeft: indent, marginBottom: 2, display: "flex", alignItems: "center", gap: 8, fontFamily: "monospace", fontSize: 12 }}>
      {label && <span style={{ color: "#64748b", fontWeight: 700 }}>{label}: </span>}
      <span style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 4, padding: "1px 8px", fontSize: 11, color: "#92400e", fontWeight: 700 }}>{value._type?.toUpperCase()}</span>
      {value._size && <span style={{ color: "#94a3b8", fontSize: 11 }}>{(value._size / 1024).toFixed(1)} KB</span>}
    </div>
  );

  if (isLeaf) {
    const col   = typeof value === "string" ? "#16a34a" : typeof value === "number" ? "#0369a1" : "#7c3aed";
    const isHit = highlight && typeof value === "string" && value.toLowerCase().includes(highlight.toLowerCase());
    return (
      <div style={{ paddingLeft: indent, marginBottom: 1, fontFamily: "monospace", fontSize: 12, lineHeight: 1.8 }}>
        {label && <span style={{ color: "#64748b", fontWeight: 600 }}>{label}: </span>}
        <span style={{ color: col, background: isHit ? "#fef08a" : "transparent", borderRadius: 3, padding: isHit ? "0 3px" : 0 }}>
          {typeof value === "string" ? `"${value}"` : String(value)}
        </span>
      </div>
    );
  }

  const [o, c] = isArr ? ["[", "]"] : ["{", "}"];
  const isNodeEntry = depth === 1 && label && /^\d+$/.test(label);
  const nodeInfo    = isNodeEntry && isObj ? value : null;

  return (
    <div style={{ paddingLeft: depth > 0 ? indent : 0, marginBottom: 2 }}>
      <div onClick={() => setOpen(x => !x)}
        style={{ display: "flex", alignItems: "center", gap: 5, cursor: "pointer", padding: "2px 5px", borderRadius: 6, marginLeft: -5, userSelect: "none", background: isNodeEntry && open ? "#f0f4ff" : "transparent", transition: "background .12s" }}
        onMouseEnter={e => { if (!(isNodeEntry && open)) e.currentTarget.style.background = "#f8faff"; }}
        onMouseLeave={e => { e.currentTarget.style.background = isNodeEntry && open ? "#f0f4ff" : "transparent"; }}>
        <svg viewBox="0 0 10 10" width="9" height="9" style={{ color: "#94a3b8", flexShrink: 0, transform: open ? "rotate(90deg)" : "none", transition: "transform .15s" }}>
          <path d="M3 1l4 4-4 4" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
        {isNodeEntry && <span style={{ background: "#1e3a8a", color: "#fff", fontSize: 10, fontWeight: 800, fontFamily: "monospace", padding: "1px 7px", borderRadius: 5, flexShrink: 0 }}>ID {label}</span>}
        {label && !isNodeEntry && <span style={{ fontFamily: "monospace", fontSize: 12, fontWeight: 700, color: "#475569" }}>{label}:</span>}
        <span style={{ fontFamily: "monospace", fontSize: 12, color: "#94a3b8" }}>{o}</span>
        {!open && <>
          <span style={{ background: "#e0e7ff", color: "#3730a3", fontSize: 10, fontWeight: 700, padding: "1px 7px", borderRadius: 8 }}>{count} {isArr ? "items" : "keys"}</span>
          {isNodeEntry && nodeInfo?.fileName && <span style={{ fontSize: 11, color: "#64748b", fontFamily: "monospace" }}>— {nodeInfo.fileName}</span>}
          <span style={{ fontFamily: "monospace", fontSize: 12, color: "#94a3b8" }}>{c}</span>
        </>}
        {isNodeEntry && open && nodeInfo?.path && <span style={{ fontSize: 10.5, color: "#6366f1", fontFamily: "monospace", background: "#eef2ff", padding: "1px 7px", borderRadius: 4 }}>{nodeInfo.path}</span>}
        {isNodeEntry && open && nodeInfo?.type && <span style={{ fontSize: 10, fontWeight: 700, padding: "1px 6px", borderRadius: 4, background: nodeInfo.type === "xml" ? "#eff6ff" : "#fff7ed", color: nodeInfo.type === "xml" ? "#1d4ed8" : "#c2410c", border: `1px solid ${nodeInfo.type === "xml" ? "#bfdbfe" : "#fed7aa"}` }}>{nodeInfo.type?.toUpperCase()}</span>}
      </div>
      {open && (
        <div style={{ paddingLeft: 14, borderLeft: "2px solid #e0e7ff", marginLeft: 4, marginTop: 2, marginBottom: 2 }}>
          {isArr
            ? value.map((item, i) => <JNode key={i} label={String(i)} value={item} depth={depth + 1} highlight={highlight} />)
            : keys.map(k => <JNode key={k} label={k} value={value[k]} depth={depth + 1} highlight={highlight} />)
          }
        </div>
      )}
      {open && <div style={{ paddingLeft: indent, fontFamily: "monospace", fontSize: 12, color: "#94a3b8" }}>{c}</div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════
   MAIN PAGE
═══════════════════════════════════════════════════════════ */
export default function EditorPage({ arrayBuffer, fileName, onBack }) {

  /* master JSON + index */
  const [masterJson, setMasterJson] = useState(null);
  const [nodeArray,  setNodeArray]  = useState([]);
  const [isLoading,  setIsLoading]  = useState(true);
  const [loadErr,    setLoadErr]    = useState("");
  const indexRef   = useRef([]);
  const masterRef  = useRef(null);

  /* docx preview */
  const previewRef  = useRef(null);
  const [rendering, setRendering] = useState(true);
  const [renderErr, setRenderErr] = useState("");

  /* model */
  const [providerKey, setProviderKey] = useState("gemini");
  const [model,       setModel]       = useState(PROVIDERS.gemini.models[0]);

  /* agent */
  const [prompt,     setPrompt]     = useState("");
  const [running,    setRunning]    = useState(false);
  const [agentLog,   setAgentLog]   = useState([]);
  const [agentError, setAgentError] = useState("");
  const [patchLog,   setPatchLog]   = useState([]);
  const [hlKw,       setHlKw]       = useState("");
  const logEndRef  = useRef(null);

  /* export */
  const [exporting,  setExporting]  = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [exportErr,  setExportErr]  = useState("");

  const zipRef = useRef(null);

  /* keep masterRef current */
  useEffect(() => { masterRef.current = masterJson; }, [masterJson]);

  /* ── LOAD ── */
  useEffect(() => {
    if (!arrayBuffer) return;
    (async () => {
      try {
        const zip = await JSZip.loadAsync(arrayBuffer);
        zipRef.current = zip;
        const nodes = await zipToNodeMap(zip);
        setNodeArray(nodes);
        const nodesObj = {};
        nodes.forEach(n => { nodesObj[String(n.nodeId)] = { path: n.path, fileName: n.fileName, type: n.type, content: n.content }; });
        const master = { docId: crypto.randomUUID?.() || Date.now().toString(36), fileName: fileName || "document.docx", builtAt: new Date().toISOString(), nodeCount: nodes.length, nodes: nodesObj };
        setMasterJson(master);
        masterRef.current = master;
        indexRef.current = buildParagraphIndex(master);
      } catch (e) { setLoadErr(e.message); }
      finally { setIsLoading(false); }
    })();
  }, [arrayBuffer]);

  /* ── DOCX PREVIEW ── */
  useEffect(() => {
    if (!arrayBuffer || !previewRef.current) return;
    renderAsync(arrayBuffer, previewRef.current, null, { inWrapper: true, breakPages: true, experimental: true, useBase64URL: true, renderHeaders: true, renderFooters: true })
      .then(() => setRendering(false))
      .catch(e => { setRenderErr(e.message); setRendering(false); });
  }, [arrayBuffer]);

  /* provider change */
  const handleProviderChange = (pk) => {
    setProviderKey(pk);
    setModel(PROVIDERS[pk].models[0]);
  };

  /* log helpers */
  const addLog = (type, text) => {
    setAgentLog(l => [...l, { id: Date.now() + Math.random(), type, text, ts: new Date().toLocaleTimeString() }]);
  };
  useEffect(() => { logEndRef.current?.scrollIntoView({ behavior: "smooth" }); }, [agentLog]);

  /* ═══════════════════════════════════════════════
     TOOL HANDLER — shared between Gemini & Cohere
  ═══════════════════════════════════════════════ */
  const handleToolCall = (tc, localIndex, localMaster, newPatchList) => {
    /* ── search_paragraphs ── */
    if (tc.name === "search_paragraphs") {
      const keywords = tc.input?.keywords || [];
      addLog("search", `🔍 Searching ${keywords.length} keyword${keywords.length !== 1 ? "s" : ""}: ${keywords.map(k => `"${k}"`).join(", ")}`);
      setHlKw(keywords[0] || "");

      const { allFlat } = searchMultiple(localIndex, keywords);
      addLog("result", `   Found ${allFlat.length} matching paragraph${allFlat.length !== 1 ? "s" : ""}`);

      const chunks = allFlat.map(hit => ({
        nodeId:        hit.nodeId,
        filePath:      hit.nodePath,
        paragraphPath: hit.paraPath,
        fullText:      hit.fullText,
        paragraph:     safeClone(hit.paraNode),
      }));

      return { ok: true, output: JSON.stringify({ found: chunks.length, paragraphs: chunks }) };
    }

    /* ── patch_paragraphs ── */
    if (tc.name === "patch_paragraphs") {
      const patches = tc.input?.patches || [];
      addLog("patch", `✏️  Patching ${patches.length} paragraph${patches.length !== 1 ? "s" : ""}…`);

      try {
        const norm = patches.map(p => ({ nodeId: String(p.nodeId), paraPath: p.paraPath, newPara: p.newParagraph }));
        const updated = applyMultiPatch(localMaster.current, norm);

        // Commit to ref immediately so next tool call sees fresh state
        localMaster.current = updated;
        localIndex.current  = buildParagraphIndex(updated);

        patches.forEach(p => {
          const hit = indexRef.current.find(e => e.nodeId === String(p.nodeId) && e.paraPath.join(".") === p.paraPath?.join?.("."));
          newPatchList.push({
            id: Date.now() + Math.random(), time: new Date().toLocaleTimeString(),
            nodeId: String(p.nodeId),
            file: updated.nodes[String(p.nodeId)]?.fileName || "unknown",
            prevText: hit?.fullText?.slice(0, 60) || "(previous)",
            newText: collectText(p.newParagraph).join("").slice(0, 60),
          });
        });

        addLog("done", `   ✓ Patched ${patches.length} paragraph${patches.length !== 1 ? "s" : ""}`);
        return { ok: true, output: JSON.stringify({ success: true, patched: patches.length }) };
      } catch (e) {
        addLog("error", `   ✗ Patch error: ${e.message}`);
        return { ok: false, output: JSON.stringify({ success: false, error: e.message }) };
      }
    }

    return { ok: false, output: JSON.stringify({ error: "Unknown tool: " + tc.name }) };
  };

  /* ═══════════════════════════════════════════════
     AGENT LOOP
  ═══════════════════════════════════════════════ */
  const runAgent = useCallback(async () => {
    if (!prompt.trim() || !masterRef.current) return;

    const prov   = PROVIDERS[providerKey];
    const apiKey = prov.envKey;

    if (!apiKey) {
      setAgentError(`API key not found. Add ${prov.envVar}=your_key to your .env file and restart the dev server.`);
      return;
    }

    setRunning(true);
    setAgentError("");
    setAgentLog([]);
    setHlKw("");

    addLog("user", prompt.trim());

    /* mutable refs for the loop */
    const loopMaster = { current: masterRef.current };
    const loopIndex  = { current: indexRef.current };
    const newPatches = [];

    const MAX_TURNS = 12;
    let turn = 0;

    try {
      if (providerKey === "gemini") {
        /* ── GEMINI LOOP ── */
        // Gemini uses its own history format
        const history = [{ role: "user", parts: [{ text: prompt.trim() }] }];

        while (turn < MAX_TURNS) {
          turn++;
          addLog("thinking", `Turn ${turn} — calling ${prov.label} (${model})…`);

          const resp = await callGemini(model, apiKey, history, SYSTEM_PROMPT);

          if (resp.toolCalls.length === 0) {
            if (resp.text) addLog("ai", resp.text);
            break;
          }

          // Add model response to history
          history.push({ role: "model", parts: resp.rawParts });

          // Execute tool calls and build function response parts
          const funcResponseParts = [];
          for (const tc of resp.toolCalls) {
            const { output } = handleToolCall(tc, loopIndex, loopMaster, newPatches);
            funcResponseParts.push({
              functionResponse: {
                name: tc.name,
                response: { output },
              },
            });
          }

          history.push({ role: "user", parts: funcResponseParts });
        }

      } else {
        /* ── COHERE LOOP ── */
        const messages = [
          { role: "system",  content: SYSTEM_PROMPT },
          { role: "user",    content: prompt.trim() },
        ];

        while (turn < MAX_TURNS) {
          turn++;
          addLog("thinking", `Turn ${turn} — calling ${prov.label} (${model})…`);

          const resp = await callCohere(model, apiKey, messages);

          if (resp.toolCalls.length === 0) {
            if (resp.text) addLog("ai", resp.text);
            break;
          }

          // Add assistant turn to history
          messages.push({ role: "assistant", ...resp.rawMsg });

          // Execute tools and build tool results
          for (const tc of resp.toolCalls) {
            const { output } = handleToolCall(tc, loopIndex, loopMaster, newPatches);
            messages.push({ role: "tool", tool_call_id: tc.id, content: output });
          }
        }
      }

      if (turn >= MAX_TURNS) addLog("warn", "⚠️ Reached max turns limit.");

      /* ── Commit all changes to React state ── */
      setMasterJson(loopMaster.current);
      masterRef.current = loopMaster.current;
      indexRef.current  = loopIndex.current;
      setNodeArray(prev =>
        prev.map(n => { const u = loopMaster.current.nodes[String(n.nodeId)]; return u ? { ...n, content: u.content } : n; })
      );
      if (newPatches.length) setPatchLog(log => [...newPatches, ...log]);

    } catch (e) {
      setAgentError(e.message);
      addLog("error", `❌ ${e.message}`);
    } finally {
      setRunning(false);
    }
  }, [prompt, providerKey, model]);

  /* ── DOWNLOAD master.json ── */
  const handleDownloadMaster = () => {
    if (!masterJson) return;
    const blob = new Blob([JSON.stringify(safeClone(masterJson), null, 2)], { type: "application/json" });
    const url  = URL.createObjectURL(blob);
    Object.assign(document.createElement("a"), { href: url, download: (fileName || "doc").replace(/\.docx$/i, "") + "_master.json" }).click();
    URL.revokeObjectURL(url);
  };

  /* ── EXPORT .docx ── */
  const handleExport = async () => {
    if (!nodeArray.length || !zipRef.current) return;
    setExporting(true); setExportErr("");
    try {
      const buf  = await nodeMapToDocxBuffer(nodeArray, zipRef.current);
      const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const url  = URL.createObjectURL(blob);
      Object.assign(document.createElement("a"), { href: url, download: (fileName || "doc").replace(/\.docx$/i, "") + "_modified.docx" }).click();
      URL.revokeObjectURL(url);
      setExportDone(true); setTimeout(() => setExportDone(false), 3000);
    } catch (e) { setExportErr(e.message); }
    finally { setExporting(false); }
  };

  /* ── DERIVED ── */
  const safeDisplayJson = masterJson ? safeClone(masterJson) : null;
  const xmlCount        = nodeArray.filter(n => n.type === "xml").length;
  const binCount        = nodeArray.filter(n => n.type === "binary").length;
  const patchCount      = patchLog.length;
  const prov            = PROVIDERS[providerKey];
  const keyMissing      = !prov.envKey;

  /* log colours */
  const LOG_STYLE = {
    user    : { bg: "#eff6ff", border: "#bfdbfe", color: "#1e40af" },
    thinking: { bg: "#fafafa", border: "#e2e8f0", color: "#64748b" },
    search  : { bg: "#f0f9ff", border: "#bae6fd", color: "#0369a1" },
    result  : { bg: "#f8fafc", border: "#e2e8f0", color: "#475569" },
    patch   : { bg: "#fdf4ff", border: "#e9d5ff", color: "#7c3aed" },
    done    : { bg: "#f0fdf4", border: "#bbf7d0", color: "#065f46" },
    ai      : { bg: "#fff",    border: "#e2e8f0", color: "#1e293b" },
    warn    : { bg: "#fffbeb", border: "#fde68a", color: "#92400e" },
    error   : { bg: "#fef2f2", border: "#fecaca", color: "#dc2626" },
  };

  /* ═══════════════════════════════════════════════
     RENDER
  ═══════════════════════════════════════════════ */
  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Serif+Display&display=swap');
        *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
        html,body,#root{height:100%;}
        .jv{font-family:'DM Sans',sans-serif;height:100vh;display:flex;flex-direction:column;background:#f1f5f9;overflow:hidden;}

        /* TOP BAR */
        .jv-bar{height:52px;background:#1e3a8a;display:flex;align-items:center;padding:0 16px;gap:10px;flex-shrink:0;box-shadow:0 2px 14px rgba(15,23,42,.25);z-index:20;}
        .jv-back{display:flex;align-items:center;gap:5px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);border-radius:7px;padding:5px 12px;color:#fff;font-size:12.5px;font-weight:500;cursor:pointer;transition:background .15s;white-space:nowrap;}
        .jv-back:hover{background:rgba(255,255,255,.2);}
        .jv-bar-title{font-family:'DM Serif Display',serif;font-size:15px;color:rgba(255,255,255,.92);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
        .jv-bar-title em{color:#93c5fd;font-style:normal;}
        .jv-chip{border-radius:6px;padding:3px 9px;font-size:11px;font-weight:600;white-space:nowrap;}
        .jv-chip.dim{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.12);color:rgba(255,255,255,.65);}
        .jv-chip.grn{background:#4ade80;color:#14532d;}
        .jv-btn{display:flex;align-items:center;gap:5px;border:none;border-radius:7px;padding:6px 13px;font-family:'DM Sans',sans-serif;font-size:12.5px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap;flex-shrink:0;}
        .jv-btn:disabled{opacity:.4;cursor:not-allowed;}
        .jv-btn.ghost{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.85);}
        .jv-btn.ghost:not(:disabled):hover{background:rgba(255,255,255,.2);}
        .jv-btn.solid{background:#fff;color:#1e3a8a;}
        .jv-btn.solid:not(:disabled):hover{background:#eff6ff;}
        .jv-btn.solid.done{background:#d1fae5;color:#065f46;}

        /* BODY */
        .jv-body{flex:1;display:flex;overflow:hidden;}

        /* LEFT — JSON */
        .jv-json-panel{width:400px;flex-shrink:0;display:flex;flex-direction:column;overflow:hidden;background:#fff;border-right:1px solid #e2e8f0;}
        .jv-json-head{padding:10px 16px;background:#fff;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;justify-content:space-between;gap:10px;flex-shrink:0;}
        .jv-file-tab{display:flex;align-items:center;gap:8px;background:#f0f4ff;border:1.5px solid #c7d2fe;border-radius:8px;padding:5px 12px;}
        .jv-file-name{font-size:13px;font-weight:700;color:#1e3a8a;font-family:monospace;}
        .jv-file-size{font-size:11px;color:#6366f1;background:#eef2ff;border-radius:4px;padding:1px 7px;font-family:monospace;}
        .jv-json-scroll{flex:1;overflow:auto;padding:16px 20px;background:#fff;}
        .jv-json-scroll::-webkit-scrollbar{width:6px;}
        .jv-json-scroll::-webkit-scrollbar-thumb{background:#e2e8f0;border-radius:99px;}

        /* CENTER — DOCX preview */
        .jv-preview-panel{flex:1;display:flex;flex-direction:column;overflow:hidden;background:#e5eaf0;position:relative;}
        .jv-preview-label{position:absolute;top:12px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#94a3b8;letter-spacing:.14em;text-transform:uppercase;background:rgba(255,255,255,.85);border-radius:20px;padding:3px 12px;z-index:5;pointer-events:none;backdrop-filter:blur(4px);}
        .jv-preview-scroll{flex:1;overflow-y:auto;padding:44px 24px 40px;display:flex;flex-direction:column;align-items:center;}
        .jv-preview-scroll::-webkit-scrollbar{width:6px;}
        .jv-preview-scroll::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:99px;}
        .jv-preview-scroll .docx-wrapper{background:transparent !important;padding:0 !important;display:flex !important;flex-direction:column !important;align-items:center !important;gap:20px !important;width:100% !important;}
        .jv-preview-scroll .docx-wrapper > section.docx{box-shadow:0 2px 8px rgba(15,23,42,.1),0 12px 40px rgba(15,23,42,.15) !important;border-radius:4px !important;margin:0 auto !important;}
        .jv-preview-overlay{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:12px;background:#e5eaf0;z-index:4;}

        /* RIGHT — AGENT */
        .ag-panel{width:380px;flex-shrink:0;background:#fff;border-left:1px solid #e2e8f0;display:flex;flex-direction:column;overflow:hidden;}

        /* agent header */
        .ag-hd{padding:0 14px;height:52px;display:flex;align-items:center;gap:10px;flex-shrink:0;}
        .ag-hd.gemini{background:linear-gradient(135deg,#1a73e8,#0f4c8a);}
        .ag-hd.cohere{background:linear-gradient(135deg,#39594d,#1a3a30);}
        .ag-hd-title{font-size:13px;font-weight:700;color:#fff;flex:1;}
        .ag-hd-badge{font-size:10.5px;font-weight:700;padding:3px 9px;border-radius:20px;background:rgba(255,255,255,.18);color:#fff;border:1px solid rgba(255,255,255,.2);}

        /* provider tabs */
        .ag-provider-tabs{display:flex;border-bottom:1px solid #f1f5f9;flex-shrink:0;}
        .ag-tab{flex:1;padding:10px 8px;border:none;background:transparent;cursor:pointer;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;color:#94a3b8;display:flex;align-items:center;justify-content:center;gap:6px;border-bottom:2px solid transparent;transition:all .15s;margin-bottom:-1px;}
        .ag-tab:hover{color:#475569;background:#f8fafc;}
        .ag-tab.act.gemini{color:#1a73e8;border-bottom-color:#1a73e8;background:#eff6ff;}
        .ag-tab.act.cohere{color:#39594d;border-bottom-color:#39594d;background:#f0fdf4;}
        .ag-tab-icon{font-size:16px;}

        /* model row */
        .ag-model-row{padding:10px 14px;border-bottom:1px solid #f1f5f9;display:flex;align-items:center;gap:8px;flex-shrink:0;background:#fafafa;}
        .ag-model-label{font-size:10.5px;font-weight:700;color:#64748b;white-space:nowrap;}
        .ag-model-select{flex:1;border:1.5px solid #e2e8f0;border-radius:8px;padding:5px 10px;font-family:'DM Sans',sans-serif;font-size:12px;color:#1e293b;outline:none;background:#fff;cursor:pointer;transition:border-color .15s;}
        .ag-model-select:focus{border-color:#1e3a8a;}
        .ag-key-ok{font-size:11px;font-weight:700;padding:3px 8px;border-radius:5px;background:#d1fae5;color:#065f46;white-space:nowrap;}
        .ag-key-missing{font-size:11px;font-weight:700;padding:3px 8px;border-radius:5px;background:#fef2f2;color:#dc2626;white-space:nowrap;}

        /* prompt area */
        .ag-prompt{padding:10px 14px;border-bottom:1px solid #f1f5f9;flex-shrink:0;}
        .ag-ta{width:100%;border:1.5px solid #e2e8f0;border-radius:10px;outline:none;resize:none;font-family:'DM Sans',sans-serif;font-size:13px;color:#1e293b;padding:10px 12px;background:#fff;line-height:1.6;transition:border-color .15s;}
        .ag-ta:focus{border-color:#1e3a8a;box-shadow:0 0 0 3px rgba(30,58,138,.07);}
        .ag-ta::placeholder{color:#94a3b8;}
        .ag-ta:disabled{background:#f8fafc;cursor:not-allowed;}
        .ag-run{margin-top:8px;width:100%;border:none;border-radius:10px;padding:10px 16px;color:#fff;font-family:'DM Sans',sans-serif;font-size:13.5px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;transition:opacity .15s;}
        .ag-run.gemini{background:linear-gradient(135deg,#1a73e8,#0f4c8a);box-shadow:0 2px 8px rgba(26,115,232,.3);}
        .ag-run.cohere{background:linear-gradient(135deg,#39594d,#1a3a30);box-shadow:0 2px 8px rgba(57,89,77,.3);}
        .ag-run:hover{opacity:.9;}
        .ag-run:disabled{opacity:.4;cursor:not-allowed;background:#94a3b8;box-shadow:none;}
        .ag-err{margin-top:6px;font-size:11px;color:#dc2626;background:#fef2f2;border:1px solid #fecaca;border-radius:7px;padding:6px 10px;line-height:1.5;}
        .ag-hint{font-size:10px;color:#94a3b8;margin-top:5px;text-align:center;}

        /* agent log */
        .ag-log-section{flex:1;overflow:hidden;display:flex;flex-direction:column;}
        .ag-log-hd{padding:6px 14px;font-size:10px;font-weight:800;color:#94a3b8;letter-spacing:.14em;text-transform:uppercase;border-bottom:1px solid #f1f5f9;flex-shrink:0;display:flex;align-items:center;gap:6px;}
        .ag-log-scroll{flex:1;overflow-y:auto;padding:10px 12px;display:flex;flex-direction:column;gap:5px;}
        .ag-log-scroll::-webkit-scrollbar{width:4px;}
        .ag-log-scroll::-webkit-scrollbar-thumb{background:#e2e8f0;border-radius:99px;}
        .ag-log-entry{border-radius:8px;padding:7px 10px;border-width:1px;border-style:solid;font-size:12px;line-height:1.5;}
        .ag-log-meta{display:flex;align-items:center;gap:5px;margin-bottom:3px;}
        .ag-log-ts{font-size:9.5px;color:#94a3b8;margin-left:auto;}
        .ag-log-text{white-space:pre-wrap;word-break:break-word;}

        /* patch list */
        .ag-patches{flex-shrink:0;border-top:1px solid #e2e8f0;max-height:140px;overflow-y:auto;}
        .ag-patches-hd{padding:6px 14px;font-size:10.5px;font-weight:700;color:#065f46;background:#f0fdf4;display:flex;align-items:center;gap:5px;position:sticky;top:0;}
        .ag-patches-dot{width:7px;height:7px;background:#4ade80;border-radius:50%;flex-shrink:0;}
        .ag-patch-item{padding:5px 14px;border-bottom:1px solid #f0fdf4;}
        .ag-patch-file{font-size:10px;font-weight:700;color:#1e3a8a;}
        .ag-patch-prev{font-size:10px;color:#dc2626;text-decoration:line-through;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .ag-patch-new{font-size:10px;color:#059669;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}

        /* export */
        .ag-export{padding:10px 14px;border-top:1px solid #e2e8f0;flex-shrink:0;}
        .ag-export-btn{width:100%;border:none;border-radius:9px;padding:9px;background:#059669;color:#fff;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:7px;transition:background .15s;}
        .ag-export-btn:hover{background:#047857;}
        .ag-export-btn:disabled{opacity:.4;cursor:not-allowed;}
        .ag-export-note{font-size:10px;color:#94a3b8;text-align:center;margin-top:4px;}

        /* misc */
        .ag-empty{display:flex;flex-direction:column;align-items:center;padding:28px 16px;gap:9px;text-align:center;}
        .ag-empty-t{font-size:13px;color:#64748b;font-weight:600;}
        .ag-empty-s{font-size:11.5px;color:#94a3b8;line-height:1.6;}
        .jv-spin{width:28px;height:28px;border:3px solid #e2e8f0;border-top-color:#1e3a8a;border-radius:50%;}
        @keyframes spin{to{transform:rotate(360deg)}} .jv-spin{animation:spin .7s linear infinite;}
        @keyframes spin2{to{transform:rotate(360deg)}} .spinning{animation:spin2 .75s linear infinite;display:inline-block;}
      `}</style>

      <div className="jv">

        {/* ═══ TOP BAR ═══ */}
        <header className="jv-bar">
          <button className="jv-back" onClick={onBack}>
            <svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">
              <path fillRule="evenodd" d="M17 10a.75.75 0 01-.75.75H5.612l4.158 3.96a.75.75 0 11-1.04 1.08l-5.5-5.25a.75.75 0 010-1.08l5.5-5.25a.75.75 0 111.04 1.08L5.612 9.25H16.25A.75.75 0 0117 10z" clipRule="evenodd" />
            </svg>
            Back
          </button>
          <div className="jv-bar-title">
            {(fileName || "document").replace(/\.docx$/i, "")} <em>/ Editor</em>
          </div>
          {!isLoading && masterJson && (
            <div className="jv-chip dim">{xmlCount} XML · {binCount} bin · {indexRef.current.length} ¶</div>
          )}
          {patchCount > 0 && <div className="jv-chip grn">✓ {patchCount} patch{patchCount > 1 ? "es" : ""}</div>}
          <button className="jv-btn ghost" onClick={handleDownloadMaster} disabled={!masterJson}>
            <svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z" /><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z" /></svg>
            master.json
          </button>
          <button className={`jv-btn solid${exportDone ? " done" : ""}`} onClick={handleExport} disabled={exporting || !masterJson}>
            {exporting
              ? <><svg className="spinning" viewBox="0 0 24 24" fill="none" width="12" height="12"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity=".25" /><path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75" /></svg>Exporting…</>
              : exportDone ? "✓ Exported!"
              : <><svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z" /><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z" /></svg>Export .docx</>
            }
          </button>
        </header>

        {/* ═══ BODY ═══ */}
        <div className="jv-body">

          {/* ─── LEFT: JSON ─── */}
          <div className="jv-json-panel">
            <div className="jv-json-head">
              <div className="jv-file-tab">
                <span style={{ fontSize: 14 }}>📄</span>
                <span className="jv-file-name">master.json</span>
                {masterJson && <span className="jv-file-size">{masterJson.nodeCount} nodes</span>}
              </div>
              {masterJson && (
                <button onClick={() => navigator.clipboard.writeText(JSON.stringify(safeDisplayJson, null, 2))}
                  style={{ display: "flex", alignItems: "center", gap: 4, padding: "4px 10px", borderRadius: 7, border: "1px solid #e2e8f0", background: "#fff", color: "#475569", fontFamily: "'DM Sans',sans-serif", fontSize: 11.5, fontWeight: 500, cursor: "pointer" }}>
                  <svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M0 6.75C0 5.784.784 5 1.75 5h1.5a.75.75 0 010 1.5h-1.5a.25.25 0 00-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 00.25-.25v-1.5a.75.75 0 011.5 0v1.5A1.75 1.75 0 019.25 16h-7.5A1.75 1.75 0 010 14.25z" /><path d="M5 1.75C5 .784 5.784 0 6.75 0h7.5C15.216 0 16 .784 16 1.75v7.5A1.75 1.75 0 0114.25 11h-7.5A1.75 1.75 0 015 9.25zm1.75-.25a.25.25 0 00-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 00.25-.25v-7.5a.25.25 0 00-.25-.25z" /></svg>
                  Copy
                </button>
              )}
            </div>
            <div className="jv-json-scroll">
              {isLoading ? (
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", padding: 28, gap: 10, textAlign: "center" }}>
                  <div className="jv-spin" />
                  <div style={{ fontSize: 12.5, color: "#64748b", fontWeight: 600 }}>Building master.json…</div>
                  <div style={{ fontSize: 11, color: "#94a3b8" }}>Parsing XML · indexing paragraphs</div>
                </div>
              ) : loadErr ? (
                <div style={{ color: "#dc2626", fontSize: 12, padding: 16 }}>{loadErr}</div>
              ) : safeDisplayJson ? (
                <JNode label={null} value={safeDisplayJson} depth={0} highlight={hlKw || null} />
              ) : null}
            </div>
          </div>

          {/* ─── CENTER: DOCX Preview ─── */}
          <div className="jv-preview-panel">
            <div className="jv-preview-label">Document Preview</div>
            {rendering && !renderErr && (
              <div className="jv-preview-overlay">
                <div className="jv-spin" />
                <div style={{ fontSize: 13, color: "#475569", fontWeight: 600 }}>Rendering…</div>
              </div>
            )}
            {renderErr && <div className="jv-preview-overlay"><div style={{ fontSize: 12, color: "#dc2626" }}>{renderErr}</div></div>}
            <div className="jv-preview-scroll" style={{ opacity: rendering ? 0 : 1, transition: "opacity .3s" }}>
              <div ref={previewRef} />
            </div>
          </div>

          {/* ─── RIGHT: Agent ─── */}
          <div className="ag-panel">

            {/* Header */}
            <div className={`ag-hd ${providerKey}`}>
              <span style={{ fontSize: 18 }}>{prov.icon}</span>
              <span className="ag-hd-title">{prov.label}</span>
              <span className="ag-hd-badge">{model}</span>
            </div>

            {/* Provider tabs */}
            <div className="ag-provider-tabs">
              {Object.entries(PROVIDERS).map(([k, p]) => (
                <button key={k} className={`ag-tab${providerKey === k ? ` act ${k}` : ""}`}
                  onClick={() => handleProviderChange(k)}>
                  <span className="ag-tab-icon">{p.icon}</span>
                  {p.label}
                </button>
              ))}
            </div>

            {/* Model + key status row */}
            <div className="ag-model-row">
              <span className="ag-model-label">Model</span>
              <select className="ag-model-select" value={model} onChange={e => setModel(e.target.value)}>
                {prov.models.map(m => <option key={m} value={m}>{m}</option>)}
              </select>
              {keyMissing
                ? <span className="ag-key-missing">No API key</span>
                : <span className="ag-key-ok">✓ Key loaded</span>
              }
            </div>

            {/* Key missing warning */}
            {keyMissing && (
              <div style={{ padding: "8px 14px", background: "#fffbeb", borderBottom: "1px solid #fde68a", fontSize: 11, color: "#92400e", lineHeight: 1.5 }}>
                ⚠️ Add <code style={{ background: "#fef3c7", padding: "1px 4px", borderRadius: 3, fontFamily: "monospace" }}>{prov.envVar}=your_key</code> to your <strong>.env</strong> file and restart <code style={{ fontFamily: "monospace" }}>vite</code>.
              </div>
            )}

            {/* Prompt */}
            <div className="ag-prompt">
              <textarea className="ag-ta" rows={3}
                placeholder={`Describe what to change…\n\ne.g. "Fix typos in the intro and capitalize all section headings"`}
                value={prompt}
                onChange={e => setPrompt(e.target.value)}
                disabled={running}
                onKeyDown={e => { if (e.key === "Enter" && e.ctrlKey) runAgent(); }}
              />
              <button className={`ag-run ${providerKey}`} onClick={runAgent}
                disabled={running || !prompt.trim() || isLoading || keyMissing}>
                {running
                  ? <><svg className="spinning" viewBox="0 0 24 24" fill="none" width="14" height="14"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity=".25" /><path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75" /></svg>Agent running…</>
                  : <><svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path fillRule="evenodd" d="M11.3 1.046A1 1 0 0112 2v5h4a1 1 0 01.82 1.573l-7 10A1 1 0 018 18v-5H4a1 1 0 01-.82-1.573l7-10a1 1 0 011.12-.38z" clipRule="evenodd" /></svg>Run Agent</>
                }
              </button>
              {agentError && <div className="ag-err">{agentError}</div>}
              <div className="ag-hint">Ctrl+Enter to run · Agent searches and patches automatically</div>
            </div>

            {/* Activity Log */}
            <div className="ag-log-section">
              <div className="ag-log-hd">
                Activity Log
                {running && <svg className="spinning" viewBox="0 0 24 24" fill="none" width="11" height="11" style={{ marginLeft: 4 }}><circle cx="12" cy="12" r="10" stroke="#1e3a8a" strokeWidth="3" opacity=".25" /><path fill="#1e3a8a" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75" /></svg>}
              </div>

              <div className="ag-log-scroll">
                {agentLog.length === 0 ? (
                  <div className="ag-empty">
                    <svg viewBox="0 0 24 24" fill="none" stroke="#e2e8f0" strokeWidth="1.5" width="42" height="42">
                      <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    <div className="ag-empty-t">
                      {isLoading ? "Indexing document…" : `Ready · ${indexRef.current.length} paragraphs indexed`}
                    </div>
                    <div className="ag-empty-s">
                      {isLoading
                        ? "Building paragraph index"
                        : "Type your editing instructions above.\nThe agent will search and patch automatically."
                      }
                    </div>
                  </div>
                ) : agentLog.map(entry => {
                  const s = LOG_STYLE[entry.type] || LOG_STYLE.ai;
                  const showHeader = ["user", "ai", "thinking"].includes(entry.type);
                  return (
                    <div key={entry.id} className="ag-log-entry"
                      style={{ background: s.bg, borderColor: s.border, color: s.color }}>
                      {showHeader && (
                        <div className="ag-log-meta">
                          <span style={{ fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: ".06em" }}>
                            {entry.type === "user" ? "👤 You" : entry.type === "ai" ? `${prov.icon} Agent` : "⏳ System"}
                          </span>
                          <span className="ag-log-ts">{entry.ts}</span>
                        </div>
                      )}
                      <div className="ag-log-text">{entry.text}</div>
                    </div>
                  );
                })}
                <div ref={logEndRef} />
              </div>
            </div>

            {/* Patch summary */}
            {patchLog.length > 0 && (
              <div className="ag-patches">
                <div className="ag-patches-hd">
                  <div className="ag-patches-dot" />
                  {patchLog.length} patch{patchLog.length > 1 ? "es" : ""} applied
                </div>
                {patchLog.map(p => (
                  <div key={p.id} className="ag-patch-item">
                    <div className="ag-patch-file">Node {p.nodeId} · {p.file} · {p.time}</div>
                    <div className="ag-patch-prev">{p.prevText || "—"}</div>
                    <div className="ag-patch-new">→ {p.newText || "(modified)"}</div>
                  </div>
                ))}
              </div>
            )}

            {/* Export */}
            {patchLog.length > 0 && (
              <div className="ag-export">
                <button className="ag-export-btn" onClick={handleExport} disabled={exporting}>
                  {exporting ? "Exporting…"
                    : exportDone ? "✓ Exported!"
                    : <><svg viewBox="0 0 16 16" fill="currentColor" width="13" height="13"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z" /><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z" /></svg>Export modified .docx</>
                  }
                </button>
                {exportErr && <div className="ag-err" style={{ marginTop: 6 }}>{exportErr}</div>}
                <div className="ag-export-note">JSON → XML → ZIP → .docx</div>
              </div>
            )}

          </div>
        </div>
      </div>
    </>
  );
}