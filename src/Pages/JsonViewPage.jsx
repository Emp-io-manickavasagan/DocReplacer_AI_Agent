/**
 * JsonViewPage.jsx
 *
 * The entire .docx is ONE master JSON:
 * {
 *   "docId": "...",
 *   "fileName": "report.docx",
 *   "nodes": {
 *     "1": { "path": "[Content_Types].xml", "fileName": "...", "content": {...} },
 *     "2": { "path": "word/document.xml",   "fileName": "...", "content": {...} },
 *     "3": { "path": "word/styles.xml",      "fileName": "...", "content": {...} },
 *     ...
 *   }
 * }
 *
 * The viewer shows this ONE JSON file.
 * Each top-level key under "nodes" is the Node ID.
 * Clicking a node ID in the tree jumps to that node.
 *
 * Search scans ALL nodes' content for matching text → extracts the paragraph chunk.
 * Patch replaces that chunk in-place inside the single master JSON.
 * Export converts master JSON back → XML per node → repacked .docx.
 */

import { useEffect, useState, useCallback, useRef } from "react";
import * as JSZipModule from "jszip";
import { zipToNodeMap, nodeMapToDocxBuffer } from "../utils/xmlToJson";

const JSZip = JSZipModule.default || JSZipModule;

/* ═══════════════════════════════════════════════
   PURE HELPERS
═══════════════════════════════════════════════ */

const getAtPath = (obj, path) =>
  path.reduce((cur, k) => (cur != null ? cur[k] : undefined), obj);

const safeClone = (obj) =>
  JSON.parse(JSON.stringify(obj, (k, v) => (k === "_raw" ? undefined : v)));

function allTexts(node, out = []) {
  if (!node || typeof node !== "object") return out;
  if (node._text) out.push(node._text);
  (node._children || []).forEach(c => allTexts(c, out));
  return out;
}

function walkText(obj, query, path = [], hits = []) {
  if (!obj || typeof obj !== "object") return hits;
  if (obj._text && obj._text.toLowerCase().includes(query.toLowerCase()))
    hits.push({ textPath: [...path], text: obj._text });
  (obj._children || []).forEach((c, i) =>
    walkText(c, query, [...path, "_children", i], hits)
  );
  return hits;
}

function nearestParagraph(content, hitPath) {
  for (let len = hitPath.length; len >= 0; len--) {
    const node = getAtPath(content, hitPath.slice(0, len));
    if (node?._tag === "w:p") return { paraPath: hitPath.slice(0, len), paraNode: node };
  }
  return null;
}

/** Search masterJson.nodes object (keyed by nodeId string) */
function searchMaster(masterJson, query) {
  if (!query.trim() || !masterJson?.nodes) return [];
  const results = [];
  const seen = new Set();

  for (const [nodeId, node] of Object.entries(masterJson.nodes)) {
    if (node.type !== "xml" || !node.content) continue;
    const hits = walkText(node.content, query);
    for (const hit of hits) {
      const para = nearestParagraph(node.content, hit.textPath);
      if (!para) continue;
      const key = `${nodeId}|${para.paraPath.join(".")}`;
      if (seen.has(key)) continue;
      seen.add(key);
      results.push({
        key,
        nodeId,
        nodePath: node.path,
        nodeFile: node.fileName,
        paraPath: para.paraPath,
        paraNode: para.paraNode,
        preview:  allTexts(para.paraNode).join("") || "(empty)",
        matched:  hit.text,
      });
    }
  }
  return results;
}

function applyPatch(masterJson, nodeId, paraPath, newPara) {
  const next = safeClone(masterJson);
  // Restore _raw on binary nodes
  for (const [id, node] of Object.entries(next.nodes)) {
    if (node.type === "binary") node.content._raw = masterJson.nodes[id]?.content?._raw;
  }
  const content = next.nodes[nodeId]?.content;
  if (!content) throw new Error("Node not found: " + nodeId);
  const parentPath = paraPath.slice(0, -1);
  const lastKey = paraPath[paraPath.length - 1];
  const parent = getAtPath(content, parentPath);
  if (parent == null) throw new Error("Bad paragraph path");
  parent[lastKey] = newPara;
  return next;
}

function buildChunk(result) {
  return JSON.stringify({
    instruction:
      "Modify only the _text values inside currentParagraph as needed. " +
      "Keep all _tag and _attrs fields exactly as-is. " +
      "Return ONLY the updated paragraph JSON object — no extra text.",
    nodeId:           result.nodeId,
    filePath:         result.nodePath,
    paragraphPath:    result.paraPath,
    currentParagraph: safeClone(result.paraNode),
  }, null, 2);
}

/* ═══════════════════════════════════════════════
   JSON TREE RENDERER — single unified JSON view
═══════════════════════════════════════════════ */

/** Renders ONE node in the JSON tree */
function JNode({ label, value, depth = 0, highlight = null, onJump }) {
  const [open, setOpen] = useState(depth < 2);
  const isObj  = value !== null && typeof value === "object" && !Array.isArray(value);
  const isArr  = Array.isArray(value);
  const isLeaf = !isObj && !isArr;
  const indent = depth * 16;
  const keys   = isObj ? Object.keys(value).filter(k => k !== "_raw") : [];
  const count  = isArr ? value.length : keys.length;

  // Binary blob — show badge, not tree
  if (isObj && value._binary) return (
    <div style={{ paddingLeft: indent, marginBottom: 2, display: "flex", alignItems: "center", gap: 8, fontFamily: "monospace", fontSize: 12 }}>
      {label && <span style={{ color: "#64748b", fontWeight: 700 }}>{label}: </span>}
      <span style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 4, padding: "1px 8px", fontSize: 11, color: "#92400e", fontWeight: 700 }}>{value._type?.toUpperCase()}</span>
      {value._size && <span style={{ color: "#94a3b8", fontSize: 11 }}>{(value._size / 1024).toFixed(1)} KB</span>}
    </div>
  );

  // Leaf
  if (isLeaf) {
    const col = typeof value === "string" ? "#16a34a" : typeof value === "number" ? "#0369a1" : "#7c3aed";
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

  const [open0, close0] = isArr ? ["[", "]"] : ["{", "}"];

  // If this is a top-level node entry (depth === 1, label is a number string = nodeId)
  const isNodeEntry = depth === 1 && label && /^\d+$/.test(label);
  const nodeInfo = isNodeEntry && isObj ? value : null;

  return (
    <div style={{ paddingLeft: depth > 0 ? indent : 0, marginBottom: 2 }}>
      <div
        onClick={() => setOpen(o => !o)}
        style={{
          display: "flex", alignItems: "center", gap: 5,
          cursor: "pointer", padding: "2px 5px", borderRadius: 6,
          marginLeft: -5, userSelect: "none",
          background: isNodeEntry && open ? "#f0f4ff" : "transparent",
          transition: "background .12s",
        }}
        onMouseEnter={e => { if (!isNodeEntry || !open) e.currentTarget.style.background = "#f8faff"; }}
        onMouseLeave={e => { e.currentTarget.style.background = isNodeEntry && open ? "#f0f4ff" : "transparent"; }}
      >
        {/* Expand arrow */}
        <svg viewBox="0 0 10 10" width="9" height="9"
          style={{ color: "#94a3b8", flexShrink: 0, transform: open ? "rotate(90deg)" : "none", transition: "transform .15s" }}>
          <path d="M3 1l4 4-4 4" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round" />
        </svg>

        {/* Node ID badge for top-level nodes */}
        {isNodeEntry && (
          <span style={{
            background: "#1e3a8a", color: "#fff",
            fontSize: 10, fontWeight: 800, fontFamily: "monospace",
            padding: "1px 7px", borderRadius: 5, flexShrink: 0,
          }}>ID {label}</span>
        )}

        {/* Regular label */}
        {label && !isNodeEntry && (
          <span style={{ fontFamily: "monospace", fontSize: 12, fontWeight: 700, color: "#475569" }}>
            {label}:
          </span>
        )}

        <span style={{ fontFamily: "monospace", fontSize: 12, color: "#94a3b8" }}>{open0}</span>

        {/* Collapsed summary */}
        {!open && (
          <>
            <span style={{ background: "#e0e7ff", color: "#3730a3", fontSize: 10, fontWeight: 700, padding: "1px 7px", borderRadius: 8 }}>
              {count} {isArr ? "items" : "keys"}
            </span>
            {isNodeEntry && nodeInfo?.fileName && (
              <span style={{ fontSize: 11, color: "#64748b", fontFamily: "monospace" }}>— {nodeInfo.fileName}</span>
            )}
            <span style={{ fontFamily: "monospace", fontSize: 12, color: "#94a3b8" }}>{close0}</span>
          </>
        )}

        {/* Node path badge when open */}
        {isNodeEntry && open && nodeInfo?.path && (
          <span style={{ fontSize: 10.5, color: "#6366f1", fontFamily: "monospace", background: "#eef2ff", padding: "1px 7px", borderRadius: 4 }}>
            {nodeInfo.path}
          </span>
        )}
        {isNodeEntry && open && nodeInfo?.type && (
          <span style={{
            fontSize: 10, fontWeight: 700, padding: "1px 6px", borderRadius: 4,
            background: nodeInfo.type === "xml" ? "#eff6ff" : "#fff7ed",
            color: nodeInfo.type === "xml" ? "#1d4ed8" : "#c2410c",
            border: `1px solid ${nodeInfo.type === "xml" ? "#bfdbfe" : "#fed7aa"}`,
          }}>{nodeInfo.type?.toUpperCase()}</span>
        )}
      </div>

      {open && (
        <div style={{ paddingLeft: 14, borderLeft: "2px solid #e0e7ff", marginLeft: 4, marginTop: 2, marginBottom: 2 }}>
          {isArr
            ? value.map((item, i) => <JNode key={i} label={String(i)} value={item} depth={depth + 1} highlight={highlight} onJump={onJump} />)
            : keys.map(k => <JNode key={k} label={k} value={value[k]} depth={depth + 1} highlight={highlight} onJump={onJump} />)
          }
        </div>
      )}

      {open && (
        <div style={{ paddingLeft: indent, fontFamily: "monospace", fontSize: 12, color: "#94a3b8" }}>{close0}</div>
      )}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   COPY BUTTON
═══════════════════════════════════════════════ */
function CopyBtn({ text, label = "Copy", style = {} }) {
  const [done, setDone] = useState(false);
  return (
    <button
      onClick={() => { navigator.clipboard.writeText(text); setDone(true); setTimeout(() => setDone(false), 2000); }}
      style={{
        display: "flex", alignItems: "center", gap: 4, padding: "5px 11px",
        borderRadius: 7, border: `1px solid ${done ? "#6ee7b7" : "#e2e8f0"}`,
        background: done ? "#d1fae5" : "#fff", color: done ? "#065f46" : "#475569",
        fontFamily: "'DM Sans',sans-serif", fontSize: 12, fontWeight: 500,
        cursor: "pointer", transition: "all .15s", whiteSpace: "nowrap", ...style,
      }}
    >
      {done
        ? <>✓ Copied!</>
        : <><svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M0 6.75C0 5.784.784 5 1.75 5h1.5a.75.75 0 010 1.5h-1.5a.25.25 0 00-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 00.25-.25v-1.5a.75.75 0 011.5 0v1.5A1.75 1.75 0 019.25 16h-7.5A1.75 1.75 0 010 14.25z" /><path d="M5 1.75C5 .784 5.784 0 6.75 0h7.5C15.216 0 16 .784 16 1.75v7.5A1.75 1.75 0 0114.25 11h-7.5A1.75 1.75 0 015 9.25zm1.75-.25a.25.25 0 00-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 00.25-.25v-7.5a.25.25 0 00-.25-.25z" /></svg>{label}</>
      }
    </button>
  );
}

/* ═══════════════════════════════════════════════
   MAIN PAGE
═══════════════════════════════════════════════ */
export default function JsonViewPage({ arrayBuffer, fileName, onBack }) {

  const [masterJson,   setMasterJson]   = useState(null); // { docId, fileName, nodes: { "1": {...}, "2": {...} } }
  const [nodeArray,    setNodeArray]    = useState([]);   // original array for export
  const [isLoading,    setIsLoading]    = useState(true);
  const [loadErr,      setLoadErr]      = useState("");

  const [query,        setQuery]        = useState("");
  const [results,      setResults]      = useState([]);
  const [activeResult, setActiveResult] = useState(null);
  const [showSearch,   setShowSearch]   = useState(false);

  const [llmText,      setLlmText]      = useState("");
  const [patchErr,     setPatchErr]     = useState("");
  const [patchLog,     setPatchLog]     = useState([]);

  const [exporting,    setExporting]    = useState(false);
  const [exportDone,   setExportDone]   = useState(false);
  const [exportErr,    setExportErr]    = useState("");

  const zipRef = useRef(null);

  /* ── LOAD ── */
  useEffect(() => {
    if (!arrayBuffer) return;
    (async () => {
      try {
        const zip   = await JSZip.loadAsync(arrayBuffer);
        zipRef.current = zip;
        const nodes = await zipToNodeMap(zip);
        setNodeArray(nodes);

        // Build ONE master JSON — nodes keyed by their ID
        const nodesObj = {};
        nodes.forEach(n => {
          nodesObj[String(n.nodeId)] = {
            path:     n.path,
            fileName: n.fileName,
            type:     n.type,
            content:  n.content,
          };
        });

        const master = {
          docId    : crypto.randomUUID?.() || Date.now().toString(36),
          fileName : fileName || "document.docx",
          builtAt  : new Date().toISOString(),
          nodeCount: nodes.length,
          nodes    : nodesObj,
        };
        setMasterJson(master);
      } catch (e) {
        setLoadErr(e.message);
      } finally {
        setIsLoading(false);
      }
    })();
  }, [arrayBuffer]);

  /* ── SEARCH ── */
  const doSearch = useCallback(() => {
    if (!query.trim() || !masterJson) return;
    const found = searchMaster(masterJson, query);
    setResults(found);
    setActiveResult(null);
    setLlmText("");
    setPatchErr("");
    setShowSearch(true);
  }, [query, masterJson]);

  const clearSearch = () => {
    setQuery("");
    setResults([]);
    setActiveResult(null);
    setLlmText("");
    setPatchErr("");
    setShowSearch(false);
  };

  /* ── PATCH ── */
  const handleApplyPatch = () => {
    if (!activeResult) { setPatchErr("Select a search result first."); return; }
    if (!llmText.trim()) { setPatchErr("Paste the LLM response JSON first."); return; }
    try {
      let parsed = JSON.parse(llmText.trim());
      if (parsed.updatedParagraph) parsed = parsed.updatedParagraph;
      if (!parsed._tag) throw new Error('Response must have a "_tag" field');

      const updated = applyPatch(masterJson, activeResult.nodeId, activeResult.paraPath, parsed);
      setMasterJson(updated);

      // Sync nodeArray (needed for export)
      setNodeArray(prev => prev.map(n => {
        const upd = updated.nodes[String(n.nodeId)];
        return upd ? { ...n, content: upd.content } : n;
      }));

      setPatchLog(log => [{
        id: Date.now(), time: new Date().toLocaleTimeString(),
        nodeId: activeResult.nodeId, file: activeResult.nodeFile,
        preview: activeResult.preview.slice(0, 60),
        newText: allTexts(parsed).join("").slice(0, 60),
      }, ...log]);

      const newNodeContent = updated.nodes[activeResult.nodeId]?.content;
      const newPara = newNodeContent ? getAtPath(newNodeContent, activeResult.paraPath) : activeResult.paraNode;
      setActiveResult({ ...activeResult, paraNode: newPara, preview: allTexts(newPara).join("") });
      setLlmText("");
      setPatchErr("");
    } catch (e) {
      setPatchErr(e.message);
    }
  };

  /* ── DOWNLOAD MASTER JSON ── */
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
      setExportDone(true);
      setTimeout(() => setExportDone(false), 3000);
    } catch (e) {
      setExportErr(e.message);
    } finally {
      setExporting(false);
    }
  };

  const safeDisplayJson = masterJson ? safeClone(masterJson) : null;
  const xmlCount  = nodeArray.filter(n => n.type === "xml").length;
  const binCount  = nodeArray.filter(n => n.type === "binary").length;
  const patchCount = patchLog.length;

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Serif+Display&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
        html, body, #root { height: 100%; }

        .jv { font-family: 'DM Sans', sans-serif; height: 100vh; display: flex; flex-direction: column; background: #f1f5f9; overflow: hidden; }

        /* Top bar */
        .jv-bar { height: 52px; background: #1e3a8a; display: flex; align-items: center; padding: 0 16px; gap: 10px; flex-shrink: 0; box-shadow: 0 2px 14px rgba(15,23,42,.25); z-index: 20; }
        .jv-back { display:flex;align-items:center;gap:5px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);border-radius:7px;padding:5px 12px;color:#fff;font-size:12.5px;font-weight:500;cursor:pointer;transition:background .15s;white-space:nowrap; }
        .jv-back:hover { background:rgba(255,255,255,.2); }
        .jv-bar-title { font-family:'DM Serif Display',serif;font-size:15px;color:rgba(255,255,255,.92);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
        .jv-bar-title em { color:#93c5fd;font-style:normal; }
        .jv-chip { border-radius:6px;padding:3px 9px;font-size:11px;font-weight:600;white-space:nowrap; }
        .jv-chip.dim { background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.12);color:rgba(255,255,255,.65); }
        .jv-chip.green { background:#4ade80;color:#14532d; }

        /* bar buttons */
        .jv-btn { display:flex;align-items:center;gap:5px;border:none;border-radius:7px;padding:6px 13px;font-family:'DM Sans',sans-serif;font-size:12.5px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap;flex-shrink:0; }
        .jv-btn:disabled { opacity:.4;cursor:not-allowed; }
        .jv-btn.ghost { background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.85); }
        .jv-btn.ghost:not(:disabled):hover { background:rgba(255,255,255,.2); }
        .jv-btn.solid { background:#fff;color:#1e3a8a; }
        .jv-btn.solid:not(:disabled):hover { background:#eff6ff; }
        .jv-btn.solid.done { background:#d1fae5;color:#065f46; }
        .jv-btn.green { background:#059669;color:#fff; }
        .jv-btn.green:not(:disabled):hover { background:#047857; }

        /* Search bar */
        .jv-search-row { padding:9px 14px;background:#fff;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;gap:8px;flex-shrink:0; }
        .jv-search-box { flex:1;display:flex;align-items:center;gap:8px;background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:10px;padding:7px 12px;transition:border-color .15s; }
        .jv-search-box:focus-within { border-color:#1e3a8a;background:#fff; }
        .jv-search-box input { flex:1;background:transparent;border:none;outline:none;font-family:'DM Sans',sans-serif;font-size:13.5px;color:#1e293b; }
        .jv-search-box input::placeholder { color:#94a3b8; }
        .jv-search-btn { background:#1e3a8a;color:#fff;border:none;border-radius:8px;padding:7px 16px;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:600;cursor:pointer; }
        .jv-search-btn:hover { background:#1e40af; }
        .jv-clear-btn { background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;border-radius:8px;padding:7px 11px;font-family:'DM Sans',sans-serif;font-size:12.5px;cursor:pointer; }
        .jv-clear-btn:hover { background:#e2e8f0; }

        /* Body: 2 panels */
        .jv-body { flex:1;display:flex;overflow:hidden; }

        /* ── MAIN JSON viewer ── */
        .jv-json-panel { flex:1;display:flex;flex-direction:column;overflow:hidden;background:#fff; }

        /* JSON header bar */
        .jv-json-head {
          padding:10px 18px;background:#fff;border-bottom:1px solid #e2e8f0;
          display:flex;align-items:center;justify-content:space-between;gap:12px;
          flex-shrink:0;
        }
        .jv-file-tab {
          display:flex;align-items:center;gap:8px;
          background:#f0f4ff;border:1.5px solid #c7d2fe;
          border-radius:8px;padding:5px 13px;
        }
        .jv-file-icon { font-size:14px; }
        .jv-file-name { font-size:13px;font-weight:700;color:#1e3a8a;font-family:monospace; }
        .jv-file-size { font-size:11px;color:#6366f1;background:#eef2ff;border-radius:4px;padding:1px 7px;font-family:monospace; }

        /* JSON content area */
        .jv-json-scroll {
          flex:1;overflow:auto;padding:18px 22px;
          background:#fff;font-family:monospace;
        }
        .jv-json-scroll::-webkit-scrollbar { width:6px; }
        .jv-json-scroll::-webkit-scrollbar-thumb { background:#e2e8f0;border-radius:99px; }

        /* ── LLM panel ── */
        .jv-llm-panel { width:340px;flex-shrink:0;background:#fff;border-left:1px solid #e2e8f0;display:flex;flex-direction:column;overflow:hidden; }
        .jv-llm-hd { padding:11px 14px;background:#1e3a8a;display:flex;align-items:center;gap:8px;flex-shrink:0; }
        .jv-llm-hd-title { font-size:12.5px;font-weight:700;color:#fff;flex:1; }
        .jv-llm-scroll { flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:12px; }
        .jv-llm-scroll::-webkit-scrollbar { width:4px; }
        .jv-llm-scroll::-webkit-scrollbar-thumb { background:#e2e8f0;border-radius:99px; }

        /* Search results in LLM panel */
        .jv-results-wrap { display:flex;flex-direction:column;gap:8px; }
        .jv-results-label { font-size:10.5px;font-weight:700;color:#94a3b8;letter-spacing:.12em;text-transform:uppercase; }
        .jv-result-item {
          display:flex;align-items:flex-start;gap:8px;
          padding:10px 11px;border-radius:9px;cursor:pointer;
          border:1.5px solid #e2e8f0;transition:all .13s;background:#fff;
        }
        .jv-result-item:hover { border-color:#93c5fd;background:#f8faff; }
        .jv-result-item.act { border-color:#1e3a8a;background:#eff6ff; }
        .jv-r-nid {
          width:28px;height:28px;border-radius:6px;
          background:#1e3a8a;color:#fff;font-size:10.5px;font-weight:700;
          display:flex;align-items:center;justify-content:center;
          flex-shrink:0;font-family:monospace;
        }
        .jv-result-item.act .jv-r-nid { background:#1e40af; }
        .jv-r-body { flex:1;min-width:0; }
        .jv-r-file { font-size:10.5px;font-weight:700;color:#1e3a8a;margin-bottom:2px; }
        .jv-r-text { font-size:12px;color:#334155;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.5; }
        .jv-r-text mark { background:#fef08a;color:#713f12;border-radius:2px;padding:0 2px; }
        .jv-r-path { font-size:9.5px;color:#94a3b8;font-family:monospace;margin-top:2px; }

        /* Workflow step */
        .jv-step { border-radius:11px;overflow:hidden;border:1px solid #e2e8f0; }
        .jv-step-hd { padding:8px 12px;display:flex;align-items:center;gap:7px; }
        .jv-step-hd.b { background:#eff6ff;border-bottom:1px solid #dbeafe; }
        .jv-step-hd.i { background:#f5f3ff;border-bottom:1px solid #e0e7ff; }
        .jv-step-hd.g { background:#f0fdf4;border-bottom:1px solid #bbf7d0; }
        .jv-step-num { width:20px;height:20px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:800;flex-shrink:0; }
        .jv-step-num.b { background:#1e3a8a;color:#fff; }
        .jv-step-num.i { background:#4f46e5;color:#fff; }
        .jv-step-num.g { background:#059669;color:#fff; }
        .jv-step-tit { font-size:11.5px;font-weight:700;flex:1; }
        .jv-step-tit.b { color:#1e3a8a; }
        .jv-step-tit.i { color:#4338ca; }
        .jv-step-tit.g { color:#065f46; }
        .jv-step-bd { padding:10px 12px;background:#fafbff; }

        .jv-chunk-pre { font-family:monospace;font-size:10.5px;color:#334155;line-height:1.6;max-height:130px;overflow-y:auto;white-space:pre-wrap;word-break:break-all;background:#f1f5f9;border-radius:7px;padding:9px 10px; }
        .jv-chunk-pre::-webkit-scrollbar { width:3px; }
        .jv-chunk-pre::-webkit-scrollbar-thumb { background:#cbd5e1;border-radius:99px; }
        .jv-ta { width:100%;border:1.5px solid #e2e8f0;border-radius:8px;outline:none;resize:vertical;font-family:monospace;font-size:11px;color:#1e293b;padding:9px 10px;background:#fff;line-height:1.6;min-height:100px;transition:border-color .15s; }
        .jv-ta:focus { border-color:#6366f1; }
        .jv-ta::placeholder { color:#94a3b8;font-family:'DM Sans',sans-serif;font-size:12px; }
        .jv-err { font-size:11.5px;color:#dc2626;background:#fef2f2;border:1px solid #fecaca;border-radius:7px;padding:7px 10px; }

        /* Patch history */
        .jv-hist { border:1px solid #e2e8f0;border-radius:11px;overflow:hidden; }
        .jv-hist-hd { padding:8px 12px;background:#f8fafc;border-bottom:1px solid #e2e8f0;font-size:11px;font-weight:700;color:#475569;display:flex;align-items:center;gap:6px; }
        .jv-hist-dot { width:7px;height:7px;border-radius:50%;background:#4ade80;flex-shrink:0; }
        .jv-hist-list { max-height:140px;overflow-y:auto; }
        .jv-hist-item { padding:7px 12px;border-bottom:1px solid #f1f5f9;display:flex;gap:7px; }
        .jv-hist-item:last-child { border-bottom:none; }
        .jv-hist-file { font-size:10.5px;font-weight:700;color:#1e3a8a; }
        .jv-hist-old  { font-size:10.5px;color:#dc2626;text-decoration:line-through;white-space:nowrap;overflow:hidden;text-overflow:ellipsis; }
        .jv-hist-new  { font-size:10.5px;color:#059669;white-space:nowrap;overflow:hidden;text-overflow:ellipsis; }
        .jv-hist-time { font-size:9.5px;color:#94a3b8; }

        /* Empty / spinner */
        .jv-empty { flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:10px;padding:30px;text-align:center; }
        .jv-empty-t { font-size:13px;color:#64748b;font-weight:600; }
        .jv-empty-s { font-size:11.5px;color:#94a3b8;line-height:1.5; }
        .jv-spin { width:30px;height:30px;border:3px solid #e2e8f0;border-top-color:#1e3a8a;border-radius:50%; }
        @keyframes spin  { to { transform:rotate(360deg) } } .jv-spin { animation:spin .7s linear infinite; }
        @keyframes spin2 { to { transform:rotate(360deg) } } .spinning { animation:spin2 .75s linear infinite;display:inline-block; }

        /* Flow strip */
        .jv-flow { padding:6px 14px;background:linear-gradient(90deg,#eff6ff,#f5f3ff);border-bottom:1px solid #e0e7ff;display:flex;align-items:center;gap:5px;flex-shrink:0;flex-wrap:wrap; }
        .fs { font-size:10.5px;color:#4338ca;font-weight:500;white-space:nowrap; }
        .fs.hi { background:#c7d2fe;border-radius:4px;padding:1px 6px;color:#1e40af;font-weight:700; }
        .fa { font-size:10.5px;color:#a5b4fc; }
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
            {(fileName || "document").replace(/\.docx$/i, "")} <em>/ Master JSON</em>
          </div>

          {!isLoading && masterJson && <>
            <div className="jv-chip dim">{xmlCount} XML</div>
            <div className="jv-chip dim">{binCount} binary</div>
            {patchCount > 0 && <div className="jv-chip green">✓ {patchCount} patch{patchCount > 1 ? "es" : ""}</div>}
          </>}

          <button className="jv-btn ghost" onClick={handleDownloadMaster} disabled={!masterJson}>
            <svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z" /><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z" /></svg>
            Download JSON
          </button>

          <button
            className={`jv-btn solid${exportDone ? " done" : ""}`}
            onClick={handleExport}
            disabled={exporting || !masterJson || isLoading}
          >
            {exporting
              ? <><svg className="spinning" viewBox="0 0 24 24" fill="none" width="13" height="13"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity=".25" /><path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75" /></svg>Exporting…</>
              : exportDone ? "✓ Exported!"
              : <><svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z" /><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z" /></svg>Export .docx</>
            }
          </button>
        </header>

        {/* ═══ SEARCH STRIP ═══ */}
        <div className="jv-search-row">
          <div className="jv-search-box">
            <svg viewBox="0 0 20 20" fill="#94a3b8" width="15" height="15" style={{ flexShrink: 0 }}>
              <path fillRule="evenodd" d="M9 3.5a5.5 5.5 0 100 11 5.5 5.5 0 000-11zM2 9a7 7 0 1112.452 4.391l3.328 3.329a.75.75 0 11-1.06 1.06l-3.329-3.328A7 7 0 012 9z" clipRule="evenodd" />
            </svg>
            <input
              placeholder='Search content to edit… e.g. "Problem Definition"'
              value={query}
              onChange={e => setQuery(e.target.value)}
              onKeyDown={e => e.key === "Enter" && doSearch()}
            />
            {query && (
              <svg onClick={() => setQuery("")} viewBox="0 0 16 16" fill="#94a3b8" width="14" height="14" style={{ cursor: "pointer", flexShrink: 0 }}>
                <path d="M3.72 3.72a.75.75 0 011.06 0L8 6.94l3.22-3.22a.75.75 0 111.06 1.06L9.06 8l3.22 3.22a.75.75 0 11-1.06 1.06L8 9.06l-3.22 3.22a.75.75 0 01-1.06-1.06L6.94 8 3.72 4.78a.75.75 0 010-1.06z" />
              </svg>
            )}
          </div>
          <button className="jv-search-btn" onClick={doSearch} disabled={!query.trim() || !masterJson}>Search</button>
          {showSearch && <button className="jv-clear-btn" onClick={clearSearch}>Clear</button>}
        </div>

        {/* ═══ FLOW STRIP ═══ */}
        {!isLoading && masterJson && (
          <div className="jv-flow">
            <span className="fs">📦 .docx → Master JSON</span><span className="fa">→</span>
            <span className={`fs${showSearch ? " hi" : ""}`}>🔍 Search content</span><span className="fa">→</span>
            <span className={`fs${activeResult ? " hi" : ""}`}>✂️ Extract chunk</span><span className="fa">→</span>
            <span className="fs">🤖 Send to LLM</span><span className="fa">→</span>
            <span className={`fs${patchCount > 0 ? " hi" : ""}`}>🔧 Patch JSON</span><span className="fa">→</span>
            <span className="fs">📄 Export .docx</span>
          </div>
        )}

        {/* ═══ BODY ═══ */}
        <div className="jv-body">

          {/* ─── LEFT: Single Master JSON Viewer ─── */}
          <div className="jv-json-panel">

            {/* File tab header */}
            <div className="jv-json-head">
              <div className="jv-file-tab">
                <span className="jv-file-icon">📄</span>
                <span className="jv-file-name">master.json</span>
                {masterJson && (
                  <span className="jv-file-size">{masterJson.nodeCount} nodes</span>
                )}
              </div>
              {masterJson && (
                <div style={{ display: "flex", gap: 7 }}>
                  <CopyBtn
                    text={JSON.stringify(safeDisplayJson, null, 2)}
                    label="Copy All"
                  />
                </div>
              )}
            </div>

            {/* JSON tree */}
            <div className="jv-json-scroll">
              {isLoading ? (
                <div className="jv-empty">
                  <div className="jv-spin" />
                  <div className="jv-empty-t">Building master.json…</div>
                  <div className="jv-empty-s">Parsing all XML files from the .docx</div>
                </div>
              ) : loadErr ? (
                <div className="jv-empty">
                  <div style={{ color: "#dc2626" }}>{loadErr}</div>
                </div>
              ) : safeDisplayJson ? (
                /* Render the entire master JSON as one tree */
                <JNode
                  label={null}
                  value={safeDisplayJson}
                  depth={0}
                  highlight={showSearch ? query : null}
                />
              ) : null}
            </div>
          </div>

          {/* ─── RIGHT: LLM Edit Workflow ─── */}
          <div className="jv-llm-panel">
            <div className="jv-llm-hd">
              <svg viewBox="0 0 20 20" fill="#93c5fd" width="14" height="14" style={{ flexShrink: 0 }}>
                <path fillRule="evenodd" d="M11.3 1.046A1 1 0 0112 2v5h4a1 1 0 01.82 1.573l-7 10A1 1 0 018 18v-5H4a1 1 0 01-.82-1.573l7-10a1 1 0 011.12-.38z" clipRule="evenodd" />
              </svg>
              <span className="jv-llm-hd-title">LLM Edit Workflow</span>
            </div>

            <div className="jv-llm-scroll">

              {/* ── Search results ── */}
              {showSearch && (
                <div className="jv-results-wrap">
                  <div className="jv-results-label">
                    {results.length === 0 ? `No matches for "${query}"` : `${results.length} paragraph${results.length > 1 ? "s" : ""} found`}
                  </div>
                  {results.map(r => {
                    const pL  = r.preview.toLowerCase();
                    const qL  = query.toLowerCase();
                    const idx = pL.indexOf(qL);
                    const hi  = idx === -1
                      ? r.preview
                      : `${r.preview.slice(0, idx)}<mark>${r.preview.slice(idx, idx + query.length)}</mark>${r.preview.slice(idx + query.length)}`;

                    return (
                      <div
                        key={r.key}
                        className={`jv-result-item${activeResult?.key === r.key ? " act" : ""}`}
                        onClick={() => { setActiveResult(r); setLlmText(""); setPatchErr(""); }}
                      >
                        <div className="jv-r-nid">{r.nodeId}</div>
                        <div className="jv-r-body">
                          <div className="jv-r-file">{r.nodeFile}</div>
                          <div className="jv-r-text" dangerouslySetInnerHTML={{ __html: r.preview.length > 100 ? hi.slice(0, 100) + "…" : hi }} />
                          <div className="jv-r-path">nodes.{r.nodeId} › para[{r.paraPath.join("›")}]</div>
                        </div>
                        <svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12" style={{ color: "#94a3b8", flexShrink: 0, alignSelf: "center" }}>
                          <path fillRule="evenodd" d="M6.22 3.22a.75.75 0 011.06 0l4.25 4.25a.75.75 0 010 1.06l-4.25 4.25a.75.75 0 01-1.06-1.06L9.94 8 6.22 4.28a.75.75 0 010-1.06z" clipRule="evenodd" />
                        </svg>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* ── No selection empty state ── */}
              {!activeResult && !showSearch && (
                <div className="jv-empty" style={{ flex: "none", padding: "30px 10px" }}>
                  <svg viewBox="0 0 24 24" fill="none" stroke="#cbd5e1" strokeWidth="1.5" width="36" height="36">
                    <path d="M13 10V3L4 14h7v7l9-11h-7z" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                  <div className="jv-empty-t">No content selected</div>
                  <div className="jv-empty-s">Search for text you want to edit above, then select a result to begin the LLM workflow</div>
                </div>
              )}

              {/* ── Workflow steps (only when a result is selected) ── */}
              {activeResult && (<>

                <div className="jv-step">
                  <div className="jv-step-hd b">
                    <div className="jv-step-num b">1</div>
                    <div className="jv-step-tit b">Copy chunk → Send to LLM</div>
                    <CopyBtn text={buildChunk(activeResult)} label="Copy" style={{ padding: "3px 9px", fontSize: 11 }} />
                  </div>
                  <div className="jv-step-bd">
                    <div style={{ fontSize: 10.5, color: "#64748b", marginBottom: 6 }}>
                      <strong style={{ color: "#1e3a8a" }}>master.json → nodes.{activeResult.nodeId}</strong>
                      <span style={{ color: "#94a3b8" }}> ({activeResult.nodeFile})</span>
                      <br />
                      <span style={{ fontFamily: "monospace", color: "#6366f1" }}>para @ [{activeResult.paraPath.join(" › ")}]</span>
                    </div>
                    <div className="jv-chunk-pre">{buildChunk(activeResult)}</div>
                    <div style={{ marginTop: 7, fontSize: 11, color: "#64748b", lineHeight: 1.5 }}>
                      This is the only part sent to LLM — not the whole JSON. Ask LLM to modify <code style={{ background: "#f1f5f9", padding: "1px 4px", borderRadius: 3 }}>_text</code> values and return the updated paragraph.
                    </div>
                  </div>
                </div>

                <div className="jv-step">
                  <div className="jv-step-hd i">
                    <div className="jv-step-num i">2</div>
                    <div className="jv-step-tit i">Paste LLM response</div>
                  </div>
                  <div className="jv-step-bd" style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                    <textarea
                      className="jv-ta"
                      placeholder={'Paste modified paragraph JSON…\n\n{ "_tag": "w:p", "_children": [...] }'}
                      value={llmText}
                      onChange={e => { setLlmText(e.target.value); setPatchErr(""); }}
                    />
                    {patchErr && <div className="jv-err">⚠ {patchErr}</div>}
                  </div>
                </div>

                <div className="jv-step">
                  <div className="jv-step-hd g">
                    <div className="jv-step-num g">3</div>
                    <div className="jv-step-tit g">Patch master.json</div>
                  </div>
                  <div className="jv-step-bd">
                    <div style={{ fontSize: 11, color: "#64748b", marginBottom: 9, lineHeight: 1.5 }}>
                      Updates <code style={{ background: "#f1f5f9", padding: "1px 4px", borderRadius: 3, fontFamily: "monospace" }}>master.json → nodes.{activeResult.nodeId}</code> in place. Then Export rebuilds the .docx.
                    </div>
                    <button
                      className="jv-btn green"
                      style={{ width: "100%", justifyContent: "center" }}
                      onClick={handleApplyPatch}
                    >
                      <svg viewBox="0 0 16 16" fill="currentColor" width="13" height="13">
                        <path d="M11.013 1.427a1.75 1.75 0 012.474 0l1.086 1.086a1.75 1.75 0 010 2.474l-8.61 8.61c-.21.21-.47.364-.756.445l-3.251.93a.75.75 0 01-.927-.928l.929-3.25c.081-.286.235-.547.445-.758l8.61-8.61z" />
                      </svg>
                      Apply Patch to master.json
                    </button>
                  </div>
                </div>

              </>)}

              {/* ── Patch history ── */}
              {patchLog.length > 0 && (
                <div className="jv-hist">
                  <div className="jv-hist-hd"><div className="jv-hist-dot" /> Patches Applied ({patchLog.length})</div>
                  <div className="jv-hist-list">
                    {patchLog.map(p => (
                      <div key={p.id} className="jv-hist-item">
                        <div className="jv-hist-dot" style={{ marginTop: 4 }} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div className="jv-hist-file">Node {p.nodeId} · {p.file}</div>
                          <div className="jv-hist-old">{p.preview || "—"}</div>
                          <div className="jv-hist-new">→ {p.newText || "(modified)"}</div>
                          <div className="jv-hist-time">{p.time}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* ── Export reminder ── */}
              {patchLog.length > 0 && (
                <div style={{ border: "1.5px solid #bbf7d0", borderRadius: 11, background: "#f0fdf4", padding: "12px 14px" }}>
                  <div style={{ fontSize: 11.5, color: "#065f46", marginBottom: 9, lineHeight: 1.5 }}>
                    <strong>{patchLog.length} patch{patchLog.length > 1 ? "es" : ""}</strong> applied to master.json.
                    Click Export to convert <strong>JSON → XML → ZIP → .docx</strong>
                  </div>
                  <button className="jv-btn green" style={{ width: "100%", justifyContent: "center" }} onClick={handleExport} disabled={exporting}>
                    {exporting
                      ? <><svg className="spinning" viewBox="0 0 24 24" fill="none" width="13" height="13"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity=".25" /><path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75" /></svg>Exporting…</>
                      : exportDone ? "✓ Exported!"
                      : <><svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z" /><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z" /></svg>Export .docx</>
                    }
                  </button>
                  {exportErr && <div className="jv-err" style={{ marginTop: 8 }}>{exportErr}</div>}
                </div>
              )}

            </div>
          </div>

        </div>
      </div>
    </>
  );
}