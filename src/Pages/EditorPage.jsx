/**
 * EditorPage.jsx
 *
 * Layout: [master.json tree] | [DOCX preview] | [Agent Pipeline]
 *
 * Search Algorithm: Paragraph Index
 *   – Built ONCE at load from masterJson
 *   – Each w:p paragraph: all _text runs joined → fullText
 *   – Exact substring match, case-insensitive
 *   – Solves split XML runs: "Problem " + "Defin" + "ition" → found
 *
 * Agentic Pipeline (LLM API placeholder):
 *   User prompt → [LLM API — not connected yet]
 *       ↓  LLM calls search(keywords[])  ← all keywords in ONE call
 *   searchMultiple returns paragraph chunks per keyword
 *       ↓  LLM edits each chunk
 *   User queues edits → applyMultiPatch → all patches in ONE shot
 *       ↓  Export → JSON → XML → ZIP → .docx
 */

import { useEffect, useState, useCallback, useRef } from "react";
import { renderAsync } from "docx-preview";
import * as JSZipModule from "jszip";
import { zipToNodeMap, nodeMapToDocxBuffer } from "../utils/xmlToJson";

const JSZip = JSZipModule.default || JSZipModule;

/* ═══════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════════
   PARAGRAPH INDEX  — built once, never re-walks XML
═══════════════════════════════════════════════ */
function buildParagraphIndex(masterJson) {
  const index = [];

  function walk(node, nodeId, nodePath, nodeFile, path) {
    if (!node || typeof node !== "object") return;
    if (node._tag === "w:p") {
      const fullText = collectText(node).join("").trim();
      if (fullText)
        index.push({ nodeId, nodePath, nodeFile, paraPath: path, fullText, paraNode: node });
      return;
    }
    (node._children || []).forEach((child, i) =>
      walk(child, nodeId, nodePath, nodeFile, [...path, "_children", i])
    );
  }

  for (const [nodeId, node] of Object.entries(masterJson.nodes)) {
    if (node.type !== "xml" || !node.content) continue;
    walk(node.content, nodeId, node.path, node.fileName, []);
  }
  return index;
}

function searchIndex(index, keyword) {
  const q = keyword.toLowerCase().trim();
  if (!q) return [];
  return index
    .filter(e => e.fullText.toLowerCase().includes(q))
    .map(e => ({ ...e, keyword, key: `${e.nodeId}|${e.paraPath.join(".")}` }));
}

/** All keywords in one shot — same as LLM tool_call would do */
function searchMultiple(index, keywords) {
  const byKeyword = {};
  const seen = new Set();
  const allFlat = [];
  for (const kw of keywords) {
    const hits = searchIndex(index, kw);
    byKeyword[kw] = hits;
    for (const h of hits) {
      if (!seen.has(h.key)) { seen.add(h.key); allFlat.push(h); }
    }
  }
  return { byKeyword, allFlat };
}

/* ═══════════════════════════════════════════════
   PATCH
═══════════════════════════════════════════════ */
function applyPatch(masterJson, nodeId, paraPath, newPara) {
  const next = safeClone(masterJson);
  for (const [id, node] of Object.entries(next.nodes))
    if (node.type === "binary") node.content._raw = masterJson.nodes[id]?.content?._raw;
  const content = next.nodes[nodeId]?.content;
  if (!content) throw new Error("Node not found: " + nodeId);
  const parent = getAtPath(content, paraPath.slice(0, -1));
  if (parent == null) throw new Error("Bad paragraph path");
  parent[paraPath[paraPath.length - 1]] = newPara;
  return next;
}

/** All patches in one shot */
function applyMultiPatch(masterJson, patches) {
  let cur = masterJson;
  for (const p of patches) cur = applyPatch(cur, p.nodeId, p.paraPath, p.newPara);
  return cur;
}

function buildChunk(entry) {
  return JSON.stringify({
    instruction:
      "Modify only _text values inside currentParagraph. " +
      "Keep all _tag and _attrs fields exactly as-is. " +
      "Return ONLY the updated paragraph JSON — no extra text.",
    nodeId:           entry.nodeId,
    filePath:         entry.nodePath,
    paragraphPath:    entry.paraPath,
    currentParagraph: safeClone(entry.paraNode),
  }, null, 2);
}

/* ═══════════════════════════════════════════════
   JSON TREE
═══════════════════════════════════════════════ */
function JNode({ label, value, depth = 0, highlight = null }) {
  const [open, setOpen] = useState(depth < 2);
  const isObj  = value !== null && typeof value === "object" && !Array.isArray(value);
  const isArr  = Array.isArray(value);
  const isLeaf = !isObj && !isArr;
  const indent = depth * 16;
  const keys   = isObj ? Object.keys(value).filter(k => k !== "_raw") : [];
  const count  = isArr ? value.length : keys.length;

  if (isObj && value._binary) return (
    <div style={{ paddingLeft:indent, marginBottom:2, display:"flex", alignItems:"center", gap:8, fontFamily:"monospace", fontSize:12 }}>
      {label && <span style={{ color:"#64748b", fontWeight:700 }}>{label}: </span>}
      <span style={{ background:"#fef3c7", border:"1px solid #fde68a", borderRadius:4, padding:"1px 8px", fontSize:11, color:"#92400e", fontWeight:700 }}>{value._type?.toUpperCase()}</span>
      {value._size && <span style={{ color:"#94a3b8", fontSize:11 }}>{(value._size/1024).toFixed(1)} KB</span>}
    </div>
  );

  if (isLeaf) {
    const col   = typeof value==="string" ? "#16a34a" : typeof value==="number" ? "#0369a1" : "#7c3aed";
    const isHit = highlight && typeof value==="string" && value.toLowerCase().includes(highlight.toLowerCase());
    return (
      <div style={{ paddingLeft:indent, marginBottom:1, fontFamily:"monospace", fontSize:12, lineHeight:1.8 }}>
        {label && <span style={{ color:"#64748b", fontWeight:600 }}>{label}: </span>}
        <span style={{ color:col, background:isHit?"#fef08a":"transparent", borderRadius:3, padding:isHit?"0 3px":0 }}>
          {typeof value==="string" ? `"${value}"` : String(value)}
        </span>
      </div>
    );
  }

  const [o, c] = isArr ? ["[","]"] : ["{","}"];
  const isNodeEntry = depth===1 && label && /^\d+$/.test(label);
  const nodeInfo    = isNodeEntry && isObj ? value : null;

  return (
    <div style={{ paddingLeft:depth>0?indent:0, marginBottom:2 }}>
      <div onClick={() => setOpen(x=>!x)} style={{ display:"flex", alignItems:"center", gap:5, cursor:"pointer", padding:"2px 5px", borderRadius:6, marginLeft:-5, userSelect:"none", background:isNodeEntry&&open?"#f0f4ff":"transparent", transition:"background .12s" }}
        onMouseEnter={e=>{ if(!(isNodeEntry&&open)) e.currentTarget.style.background="#f8faff"; }}
        onMouseLeave={e=>{ e.currentTarget.style.background=isNodeEntry&&open?"#f0f4ff":"transparent"; }}>
        <svg viewBox="0 0 10 10" width="9" height="9" style={{ color:"#94a3b8", flexShrink:0, transform:open?"rotate(90deg)":"none", transition:"transform .15s" }}>
          <path d="M3 1l4 4-4 4" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
        </svg>
        {isNodeEntry && <span style={{ background:"#1e3a8a", color:"#fff", fontSize:10, fontWeight:800, fontFamily:"monospace", padding:"1px 7px", borderRadius:5, flexShrink:0 }}>ID {label}</span>}
        {label && !isNodeEntry && <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:700, color:"#475569" }}>{label}:</span>}
        <span style={{ fontFamily:"monospace", fontSize:12, color:"#94a3b8" }}>{o}</span>
        {!open && <>
          <span style={{ background:"#e0e7ff", color:"#3730a3", fontSize:10, fontWeight:700, padding:"1px 7px", borderRadius:8 }}>{count} {isArr?"items":"keys"}</span>
          {isNodeEntry && nodeInfo?.fileName && <span style={{ fontSize:11, color:"#64748b", fontFamily:"monospace" }}>— {nodeInfo.fileName}</span>}
          <span style={{ fontFamily:"monospace", fontSize:12, color:"#94a3b8" }}>{c}</span>
        </>}
        {isNodeEntry && open && nodeInfo?.path && <span style={{ fontSize:10.5, color:"#6366f1", fontFamily:"monospace", background:"#eef2ff", padding:"1px 7px", borderRadius:4 }}>{nodeInfo.path}</span>}
        {isNodeEntry && open && nodeInfo?.type && <span style={{ fontSize:10, fontWeight:700, padding:"1px 6px", borderRadius:4, background:nodeInfo.type==="xml"?"#eff6ff":"#fff7ed", color:nodeInfo.type==="xml"?"#1d4ed8":"#c2410c", border:`1px solid ${nodeInfo.type==="xml"?"#bfdbfe":"#fed7aa"}` }}>{nodeInfo.type?.toUpperCase()}</span>}
      </div>
      {open && (
        <div style={{ paddingLeft:14, borderLeft:"2px solid #e0e7ff", marginLeft:4, marginTop:2, marginBottom:2 }}>
          {isArr
            ? value.map((item,i) => <JNode key={i} label={String(i)} value={item} depth={depth+1} highlight={highlight}/>)
            : keys.map(k => <JNode key={k} label={k} value={value[k]} depth={depth+1} highlight={highlight}/>)
          }
        </div>
      )}
      {open && <div style={{ paddingLeft:indent, fontFamily:"monospace", fontSize:12, color:"#94a3b8" }}>{c}</div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════
   COPY BUTTON
═══════════════════════════════════════════════ */
function CopyBtn({ text, label="Copy", style={} }) {
  const [done, setDone] = useState(false);
  return (
    <button onClick={() => { navigator.clipboard.writeText(text); setDone(true); setTimeout(()=>setDone(false),2000); }}
      style={{ display:"flex", alignItems:"center", gap:4, padding:"5px 11px", borderRadius:7, border:`1px solid ${done?"#6ee7b7":"#e2e8f0"}`, background:done?"#d1fae5":"#fff", color:done?"#065f46":"#475569", fontFamily:"'DM Sans',sans-serif", fontSize:12, fontWeight:500, cursor:"pointer", transition:"all .15s", whiteSpace:"nowrap", ...style }}>
      {done ? "✓ Copied!" : <><svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M0 6.75C0 5.784.784 5 1.75 5h1.5a.75.75 0 010 1.5h-1.5a.25.25 0 00-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 00.25-.25v-1.5a.75.75 0 011.5 0v1.5A1.75 1.75 0 019.25 16h-7.5A1.75 1.75 0 010 14.25z"/><path d="M5 1.75C5 .784 5.784 0 6.75 0h7.5C15.216 0 16 .784 16 1.75v7.5A1.75 1.75 0 0114.25 11h-7.5A1.75 1.75 0 015 9.25zm1.75-.25a.25.25 0 00-.25.25v7.5c0 .138.112.25.25.25h7.5a.25.25 0 00.25-.25v-7.5a.25.25 0 00-.25-.25z"/></svg>{label}</>}
    </button>
  );
}

/* ═══════════════════════════════════════════════
   MAIN PAGE
═══════════════════════════════════════════════ */
export default function EditorPage({ arrayBuffer, fileName, onBack }) {

  /* master JSON + paragraph index */
  const [masterJson, setMasterJson] = useState(null);
  const [nodeArray,  setNodeArray]  = useState([]);
  const [isLoading,  setIsLoading]  = useState(true);
  const [loadErr,    setLoadErr]    = useState("");
  const indexRef = useRef([]);   // flat paragraph index — built once

  /* docx preview */
  const previewRef  = useRef(null);
  const [rendering, setRendering] = useState(true);
  const [renderErr, setRenderErr] = useState("");

  /* search */
  const [kwInput,       setKwInput]       = useState("");
  const [searchResults, setSearchResults] = useState(null);   // { byKeyword, allFlat }
  const [highlightKw,   setHighlightKw]   = useState("");

  /* edit queue  — [{ qid, entry, editJson, editErr }] */
  const [queue,    setQueue]    = useState([]);

  /* apply */
  const [applying, setApplying] = useState(false);
  const [applyErr, setApplyErr] = useState("");
  const [patchLog, setPatchLog] = useState([]);

  /* export */
  const [exporting,  setExporting]  = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [exportErr,  setExportErr]  = useState("");

  const zipRef = useRef(null);

  /* ── LOAD master JSON + build index ── */
  useEffect(() => {
    if (!arrayBuffer) return;
    (async () => {
      try {
        const zip = await JSZip.loadAsync(arrayBuffer);
        zipRef.current = zip;
        const nodes = await zipToNodeMap(zip);
        setNodeArray(nodes);
        const nodesObj = {};
        nodes.forEach(n => {
          nodesObj[String(n.nodeId)] = { path:n.path, fileName:n.fileName, type:n.type, content:n.content };
        });
        const master = {
          docId    : crypto.randomUUID?.() || Date.now().toString(36),
          fileName : fileName || "document.docx",
          builtAt  : new Date().toISOString(),
          nodeCount: nodes.length,
          nodes    : nodesObj,
        };
        setMasterJson(master);
        indexRef.current = buildParagraphIndex(master);
      } catch(e) { setLoadErr(e.message); }
      finally { setIsLoading(false); }
    })();
  }, [arrayBuffer]);

  /* ── LOAD docx preview ── */
  useEffect(() => {
    if (!arrayBuffer || !previewRef.current) return;
    renderAsync(arrayBuffer, previewRef.current, null, {
      inWrapper:true, breakPages:true, experimental:true,
      useBase64URL:true, renderHeaders:true, renderFooters:true,
    })
      .then(() => setRendering(false))
      .catch(e  => { setRenderErr(e.message); setRendering(false); });
  }, [arrayBuffer]);

  /* ── SEARCH — multi-keyword, single call ── */
  const handleSearch = useCallback(() => {
    const kws = kwInput.split("\n").map(k=>k.trim()).filter(Boolean);
    if (!kws.length || !indexRef.current.length) return;
    setSearchResults(searchMultiple(indexRef.current, kws));
    setHighlightKw(kws[0] || "");
    setApplyErr("");
  }, [kwInput]);

  const clearSearch = () => {
    setKwInput(""); setSearchResults(null);
    setHighlightKw(""); setQueue([]);
  };

  /* ── QUEUE ── */
  const addToQueue = (entry) =>
    setQueue(q => q.find(x=>x.qid===entry.key) ? q : [...q, { qid:entry.key, entry, editJson:"", editErr:"" }]);

  const removeFromQueue = (qid) =>
    setQueue(q => q.filter(x=>x.qid!==qid));

  const updateQueueItem = (qid, editJson) =>
    setQueue(q => q.map(x => x.qid===qid ? {...x, editJson, editErr:""} : x));

  /* ── APPLY ALL PATCHES — one shot ── */
  const handleApplyAll = useCallback(() => {
    if (!queue.length) return;
    setApplying(true); setApplyErr("");

    const patches = [];
    const errors  = [];

    queue.forEach(item => {
      if (!item.editJson.trim()) {
        errors.push({ qid:item.qid, msg:"Empty — paste the modified paragraph JSON" });
        return;
      }
      try {
        let p = JSON.parse(item.editJson.trim());
        if (p.updatedParagraph) p = p.updatedParagraph;
        if (!p._tag) throw new Error('Missing "_tag" field');
        patches.push({ nodeId:item.entry.nodeId, paraPath:item.entry.paraPath, newPara:p, item });
      } catch(e) { errors.push({ qid:item.qid, msg:e.message }); }
    });

    if (errors.length) {
      setQueue(q => q.map(x => { const e=errors.find(e=>e.qid===x.qid); return e?{...x,editErr:e.msg}:x; }));
      setApplying(false); return;
    }

    try {
      const updated = applyMultiPatch(masterJson, patches);
      setMasterJson(updated);
      indexRef.current = buildParagraphIndex(updated);   // rebuild index
      setNodeArray(prev => prev.map(n => { const u=updated.nodes[String(n.nodeId)]; return u?{...n,content:u.content}:n; }));

      setPatchLog(log => [
        ...patches.map(p => ({
          id      : Date.now()+Math.random(),
          time    : new Date().toLocaleTimeString(),
          nodeId  : p.nodeId,
          file    : p.item.entry.nodeFile,
          prevText: p.item.entry.fullText.slice(0,60),
          newText : collectText(p.newPara).join("").slice(0,60),
        })),
        ...log,
      ]);

      setQueue([]); setSearchResults(null);
      setKwInput(""); setHighlightKw("");
    } catch(e) { setApplyErr(e.message); }
    finally { setApplying(false); }
  }, [queue, masterJson]);

  /* ── DOWNLOAD JSON ── */
  const handleDownloadMaster = () => {
    if (!masterJson) return;
    const blob = new Blob([JSON.stringify(safeClone(masterJson),null,2)], { type:"application/json" });
    const url  = URL.createObjectURL(blob);
    Object.assign(document.createElement("a"), { href:url, download:(fileName||"doc").replace(/\.docx$/i,"")+"_master.json" }).click();
    URL.revokeObjectURL(url);
  };

  /* ── EXPORT .docx ── */
  const handleExport = async () => {
    if (!nodeArray.length || !zipRef.current) return;
    setExporting(true); setExportErr("");
    try {
      const buf  = await nodeMapToDocxBuffer(nodeArray, zipRef.current);
      const blob = new Blob([buf], { type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      const url  = URL.createObjectURL(blob);
      Object.assign(document.createElement("a"), { href:url, download:(fileName||"doc").replace(/\.docx$/i,"")+"_modified.docx" }).click();
      URL.revokeObjectURL(url);
      setExportDone(true); setTimeout(()=>setExportDone(false),3000);
    } catch(e) { setExportErr(e.message); }
    finally { setExporting(false); }
  };

  const safeDisplayJson = masterJson ? safeClone(masterJson) : null;
  const xmlCount        = nodeArray.filter(n=>n.type==="xml").length;
  const binCount        = nodeArray.filter(n=>n.type==="binary").length;
  const patchCount      = patchLog.length;
  const totalResults    = searchResults?.allFlat?.length ?? 0;

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Serif+Display&display=swap');
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
        .jv-chip.idx{background:#7c3aed;color:#fff;}
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
        .jv-json-panel{width:420px;flex-shrink:0;display:flex;flex-direction:column;overflow:hidden;background:#fff;border-right:1px solid #e2e8f0;}
        .jv-json-head{padding:10px 18px;background:#fff;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-shrink:0;}
        .jv-file-tab{display:flex;align-items:center;gap:8px;background:#f0f4ff;border:1.5px solid #c7d2fe;border-radius:8px;padding:5px 13px;}
        .jv-file-icon{font-size:14px;}
        .jv-file-name{font-size:13px;font-weight:700;color:#1e3a8a;font-family:monospace;}
        .jv-file-size{font-size:11px;color:#6366f1;background:#eef2ff;border-radius:4px;padding:1px 7px;font-family:monospace;}
        .jv-json-scroll{flex:1;overflow:auto;padding:18px 22px;background:#fff;font-family:monospace;}
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

        /* RIGHT — Agent pipeline */
        .jv-agent-panel{width:370px;flex-shrink:0;background:#fff;border-left:1px solid #e2e8f0;display:flex;flex-direction:column;overflow:hidden;}
        .jv-agent-hd{padding:11px 14px;background:#1e3a8a;display:flex;align-items:center;gap:8px;flex-shrink:0;}
        .jv-agent-hd-title{font-size:12.5px;font-weight:700;color:#fff;flex:1;}
        .jv-agent-scroll{flex:1;overflow-y:auto;padding:12px;display:flex;flex-direction:column;gap:10px;}
        .jv-agent-scroll::-webkit-scrollbar{width:4px;}
        .jv-agent-scroll::-webkit-scrollbar-thumb{background:#e2e8f0;border-radius:99px;}

        /* LLM placeholder */
        .jv-llm-ph{border:1.5px dashed #c7d2fe;border-radius:11px;padding:12px 13px;background:#f8faff;}
        .jv-llm-ph-title{font-size:11px;font-weight:700;color:#3730a3;display:flex;align-items:center;gap:6px;margin-bottom:7px;}
        .jv-llm-ph-badge{background:#e0e7ff;color:#3730a3;font-size:9.5px;font-weight:800;padding:1px 7px;border-radius:4px;letter-spacing:.05em;}
        .jv-llm-ph-ta{width:100%;background:#f1f5f9;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-family:'DM Sans',sans-serif;font-size:12.5px;color:#94a3b8;resize:none;outline:none;cursor:not-allowed;}
        .jv-llm-ph-btn{margin-top:7px;width:100%;padding:7px;border:none;border-radius:8px;background:#e0e7ff;color:#6366f1;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;cursor:not-allowed;display:flex;align-items:center;justify-content:center;gap:6px;}
        .jv-llm-ph-note{margin-top:7px;font-size:10.5px;color:#94a3b8;line-height:1.5;text-align:center;}

        /* Section headers */
        .jv-sec{font-size:10px;font-weight:800;color:#64748b;letter-spacing:.14em;text-transform:uppercase;display:flex;align-items:center;gap:6px;padding:2px 0;}
        .jv-sec .dot{width:6px;height:6px;border-radius:50%;flex-shrink:0;}
        .jv-sec .bl{background:#1e3a8a;}
        .jv-sec .pu{background:#7c3aed;}
        .jv-sec .gn{background:#059669;}
        .jv-badge{margin-left:auto;font-size:9.5px;font-weight:800;padding:1px 6px;border-radius:4px;font-family:monospace;}
        .jv-badge.bl{background:#dbeafe;color:#1e3a8a;}
        .jv-badge.pu{background:#ede9fe;color:#7c3aed;}
        .jv-badge.gn{background:#d1fae5;color:#065f46;}

        /* Keyword textarea */
        .jv-kw-ta{width:100%;border:1.5px solid #e2e8f0;border-radius:8px;outline:none;resize:vertical;font-family:monospace;font-size:12px;color:#1e293b;padding:8px 10px;background:#fff;line-height:1.7;min-height:72px;transition:border-color .15s;}
        .jv-kw-ta:focus{border-color:#1e3a8a;}
        .jv-kw-ta::placeholder{color:#94a3b8;font-family:'DM Sans',sans-serif;font-size:12px;}
        .jv-sbtns{display:flex;gap:6px;}
        .jv-sbtn{flex:1;border:none;border-radius:8px;padding:7px 0;font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s;}
        .jv-sbtn.go{background:#1e3a8a;color:#fff;}
        .jv-sbtn.go:hover{background:#1e40af;}
        .jv-sbtn.go:disabled{opacity:.4;cursor:not-allowed;}
        .jv-sbtn.clr{background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;}
        .jv-sbtn.clr:hover{background:#e2e8f0;}

        /* Hit card */
        .jv-hit{border:1.5px solid #e2e8f0;border-radius:9px;padding:9px 11px;background:#fff;transition:border-color .13s;}
        .jv-hit:hover{border-color:#93c5fd;}
        .jv-hit-top{display:flex;align-items:center;gap:6px;margin-bottom:4px;}
        .jv-hit-nid{background:#1e3a8a;color:#fff;font-size:9.5px;font-weight:800;font-family:monospace;padding:2px 6px;border-radius:4px;flex-shrink:0;}
        .jv-hit-file{font-size:10.5px;color:#1e3a8a;font-weight:700;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
        .jv-hit-text{font-size:11.5px;color:#334155;line-height:1.5;margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .jv-hit-text mark{background:#fef08a;color:#713f12;border-radius:2px;padding:0 2px;}
        .jv-hit-foot{display:flex;gap:5px;align-items:center;}
        .jv-add-btn{flex:1;border:none;border-radius:7px;padding:5px 0;background:#eff6ff;color:#1e3a8a;font-family:'DM Sans',sans-serif;font-size:11.5px;font-weight:600;cursor:pointer;transition:background .12s;display:flex;align-items:center;justify-content:center;gap:4px;}
        .jv-add-btn:hover{background:#dbeafe;}
        .jv-add-btn.queued{background:#f0fdf4;color:#059669;cursor:default;}
        .jv-path-tag{font-size:9.5px;color:#94a3b8;font-family:monospace;background:#f8fafc;border-radius:4px;padding:1px 5px;}

        /* Queue item */
        .jv-q-item{border:1.5px solid #e2e8f0;border-radius:10px;overflow:hidden;}
        .jv-q-hd{padding:7px 11px;background:#f8fafc;border-bottom:1px solid #f1f5f9;display:flex;align-items:center;gap:6px;}
        .jv-q-nid{background:#7c3aed;color:#fff;font-size:9.5px;font-weight:800;font-family:monospace;padding:2px 6px;border-radius:4px;flex-shrink:0;}
        .jv-q-info{flex:1;min-width:0;}
        .jv-q-file{font-size:10.5px;font-weight:700;color:#1e293b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .jv-q-prev{font-size:10px;color:#64748b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .jv-q-rm{background:none;border:none;cursor:pointer;color:#94a3b8;padding:2px 4px;border-radius:4px;font-size:16px;line-height:1;flex-shrink:0;}
        .jv-q-rm:hover{color:#dc2626;background:#fef2f2;}
        .jv-q-bd{padding:9px 11px;background:#fff;display:flex;flex-direction:column;gap:6px;}
        .jv-edit-ta{width:100%;border:1.5px solid #e2e8f0;border-radius:7px;outline:none;resize:vertical;font-family:monospace;font-size:10.5px;color:#1e293b;padding:7px 9px;background:#fff;line-height:1.6;min-height:80px;transition:border-color .15s;}
        .jv-edit-ta:focus{border-color:#7c3aed;}
        .jv-edit-ta.err{border-color:#fca5a5;}
        .jv-edit-ta::placeholder{color:#94a3b8;font-family:'DM Sans',sans-serif;font-size:11px;}
        .jv-q-err{font-size:10.5px;color:#dc2626;background:#fef2f2;border:1px solid #fecaca;border-radius:6px;padding:5px 8px;}

        /* Apply button */
        .jv-apply-btn{width:100%;border:none;border-radius:9px;padding:9px;background:#059669;color:#fff;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:7px;transition:background .15s;}
        .jv-apply-btn:hover{background:#047857;}
        .jv-apply-btn:disabled{opacity:.4;cursor:not-allowed;}
        .jv-apply-err{font-size:11px;color:#dc2626;background:#fef2f2;border:1px solid #fecaca;border-radius:7px;padding:7px 10px;}

        /* Patch log */
        .jv-log{border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;}
        .jv-log-hd{padding:7px 11px;background:#f8fafc;border-bottom:1px solid #e2e8f0;font-size:10.5px;font-weight:700;color:#475569;display:flex;align-items:center;gap:5px;}
        .jv-log-dot{width:6px;height:6px;border-radius:50%;background:#4ade80;flex-shrink:0;}
        .jv-log-list{max-height:150px;overflow-y:auto;}
        .jv-log-item{padding:7px 11px;border-bottom:1px solid #f1f5f9;}
        .jv-log-item:last-child{border-bottom:none;}
        .jv-log-file{font-size:10px;font-weight:700;color:#1e3a8a;}
        .jv-log-prev{font-size:10px;color:#dc2626;text-decoration:line-through;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .jv-log-new{font-size:10px;color:#059669;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
        .jv-log-time{font-size:9px;color:#94a3b8;}

        /* Export card */
        .jv-export-card{border:1.5px solid #bbf7d0;border-radius:11px;background:#f0fdf4;padding:11px 13px;display:flex;flex-direction:column;gap:8px;}
        .jv-export-note{font-size:11px;color:#065f46;line-height:1.5;}
        .jv-export-btn{border:none;border-radius:8px;padding:8px;background:#059669;color:#fff;font-family:'DM Sans',sans-serif;font-size:12.5px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;transition:background .15s;}
        .jv-export-btn:hover{background:#047857;}
        .jv-export-btn:disabled{opacity:.4;cursor:not-allowed;}

        /* Empty / spin */
        .jv-empty{display:flex;flex-direction:column;align-items:center;padding:24px 10px;gap:8px;text-align:center;}
        .jv-empty-t{font-size:12.5px;color:#64748b;font-weight:600;}
        .jv-empty-s{font-size:11px;color:#94a3b8;line-height:1.5;}
        .jv-spin{width:30px;height:30px;border:3px solid #e2e8f0;border-top-color:#1e3a8a;border-radius:50%;}
        @keyframes spin{to{transform:rotate(360deg)}} .jv-spin{animation:spin .7s linear infinite;}
        @keyframes spin2{to{transform:rotate(360deg)}} .spinning{animation:spin2 .75s linear infinite;display:inline-block;}

        /* Flow strip */
        .jv-flow{padding:5px 14px;background:linear-gradient(90deg,#eff6ff,#f5f3ff);border-bottom:1px solid #e0e7ff;display:flex;align-items:center;gap:5px;flex-shrink:0;flex-wrap:wrap;}
        .fs{font-size:10.5px;color:#4338ca;font-weight:500;white-space:nowrap;}
        .fs.hi{background:#c7d2fe;border-radius:4px;padding:1px 6px;color:#1e40af;font-weight:700;}
        .fa{font-size:10.5px;color:#a5b4fc;}
      `}</style>

      <div className="jv">

        {/* TOP BAR */}
        <header className="jv-bar">
          <button className="jv-back" onClick={onBack}>
            <svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">
              <path fillRule="evenodd" d="M17 10a.75.75 0 01-.75.75H5.612l4.158 3.96a.75.75 0 11-1.04 1.08l-5.5-5.25a.75.75 0 010-1.08l5.5-5.25a.75.75 0 111.04 1.08L5.612 9.25H16.25A.75.75 0 0117 10z" clipRule="evenodd"/>
            </svg>
            Back
          </button>
          <div className="jv-bar-title">
            {(fileName||"document").replace(/\.docx$/i,"")} <em>/ Editor</em>
          </div>
          {!isLoading && masterJson && <>
            <div className="jv-chip dim">{xmlCount} XML · {binCount} bin</div>
            <div className="jv-chip idx">⚡ {indexRef.current.length} paragraphs</div>
            {patchCount > 0 && <div className="jv-chip grn">✓ {patchCount} patch{patchCount>1?"es":""}</div>}
          </>}
          <button className="jv-btn ghost" onClick={handleDownloadMaster} disabled={!masterJson}>
            <svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z"/><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z"/></svg>
            master.json
          </button>
          <button className={`jv-btn solid${exportDone?" done":""}`} onClick={handleExport} disabled={exporting||!masterJson}>
            {exporting
              ? <><svg className="spinning" viewBox="0 0 24 24" fill="none" width="12" height="12"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity=".25"/><path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75"/></svg>Exporting…</>
              : exportDone ? "✓ Exported!"
              : <><svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z"/><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z"/></svg>Export .docx</>
            }
          </button>
        </header>

        {/* FLOW STRIP */}
        {!isLoading && masterJson && (
          <div className="jv-flow">
            <span className="fs">📦 .docx → Index</span><span className="fa">→</span>
            <span className={`fs${searchResults?" hi":""}`}>🔍 Multi-keyword search</span><span className="fa">→</span>
            <span className={`fs${queue.length>0?" hi":""}`}>✏️ Queue edits</span><span className="fa">→</span>
            <span className="fs">🤖 LLM modifies JSON</span><span className="fa">→</span>
            <span className={`fs${patchCount>0?" hi":""}`}>🔧 Patch master.json</span><span className="fa">→</span>
            <span className="fs">📄 Export .docx</span>
          </div>
        )}

        {/* BODY */}
        <div className="jv-body">

          {/* LEFT — master.json tree */}
          <div className="jv-json-panel">
            <div className="jv-json-head">
              <div className="jv-file-tab">
                <span className="jv-file-icon">📄</span>
                <span className="jv-file-name">master.json</span>
                {masterJson && <span className="jv-file-size">{masterJson.nodeCount} nodes</span>}
              </div>
              {masterJson && <CopyBtn text={JSON.stringify(safeDisplayJson,null,2)} label="Copy All"/>}
            </div>
            <div className="jv-json-scroll">
              {isLoading ? (
                <div className="jv-empty">
                  <div className="jv-spin"/>
                  <div className="jv-empty-t">Building master.json…</div>
                  <div className="jv-empty-s">Parsing XML + indexing paragraphs</div>
                </div>
              ) : loadErr ? (
                <div className="jv-empty"><div style={{color:"#dc2626"}}>{loadErr}</div></div>
              ) : safeDisplayJson ? (
                <JNode label={null} value={safeDisplayJson} depth={0} highlight={highlightKw||null}/>
              ) : null}
            </div>
          </div>

          {/* CENTER — DOCX preview */}
          <div className="jv-preview-panel">
            <div className="jv-preview-label">Document Preview</div>
            {rendering && !renderErr && (
              <div className="jv-preview-overlay">
                <div className="jv-spin"/>
                <div style={{fontSize:13,color:"#475569",fontWeight:600}}>Rendering document…</div>
              </div>
            )}
            {renderErr && (
              <div className="jv-preview-overlay">
                <div style={{fontSize:12,color:"#dc2626"}}>{renderErr}</div>
              </div>
            )}
            <div className="jv-preview-scroll" style={{opacity:rendering?0:1,transition:"opacity .3s"}}>
              <div ref={previewRef}/>
            </div>
          </div>

          {/* RIGHT — Agent Pipeline */}
          <div className="jv-agent-panel">
            <div className="jv-agent-hd">
              <svg viewBox="0 0 20 20" fill="#93c5fd" width="14" height="14" style={{flexShrink:0}}>
                <path fillRule="evenodd" d="M11.3 1.046A1 1 0 0112 2v5h4a1 1 0 01.82 1.573l-7 10A1 1 0 018 18v-5H4a1 1 0 01-.82-1.573l7-10a1 1 0 011.12-.38z" clipRule="evenodd"/>
              </svg>
              <span className="jv-agent-hd-title">Agent Pipeline</span>
              {queue.length > 0 && (
                <span style={{background:"#7c3aed",color:"#fff",fontSize:10,fontWeight:800,padding:"1px 7px",borderRadius:5,fontFamily:"monospace"}}>
                  {queue.length} queued
                </span>
              )}
            </div>

            <div className="jv-agent-scroll">

              {/* LLM prompt placeholder */}
              <div className="jv-llm-ph">
                <div className="jv-llm-ph-title">
                  🤖 LLM Prompt
                  <span className="jv-llm-ph-badge">API NOT CONNECTED</span>
                </div>
                <textarea className="jv-llm-ph-ta" rows={3} disabled
                  placeholder='e.g. "Change all headings to title case and fix the intro paragraph typo"'/>
                <button className="jv-llm-ph-btn" disabled>
                  Run Agent — Connect LLM API to enable
                </button>
                <div className="jv-llm-ph-note">
                  When connected, the LLM will call the search and patch tools below automatically from your single prompt.
                </div>
              </div>

              {/* ── STEP 1: Search ── */}
              <div className="jv-sec">
                <span className="dot bl"/>
                STEP 1 — Search Tool
                {searchResults && (
                  <span className="jv-badge bl">{totalResults} found</span>
                )}
              </div>

              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                <div style={{fontSize:11,color:"#64748b",lineHeight:1.5}}>
                  One keyword per line. All searched in a single call — exactly how the LLM agent will use this tool.
                </div>
                <textarea
                  className="jv-kw-ta"
                  placeholder={"Problem Definition\nExecutive Summary\nConclusion"}
                  value={kwInput}
                  onChange={e => setKwInput(e.target.value)}
                  onKeyDown={e => { if(e.key==="Enter" && e.ctrlKey) handleSearch(); }}
                />
                <div style={{fontSize:10,color:"#94a3b8"}}>Ctrl+Enter to search</div>
                <div className="jv-sbtns">
                  <button className="jv-sbtn go" onClick={handleSearch} disabled={!kwInput.trim()||isLoading}>
                    <svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">
                      <path fillRule="evenodd" d="M9 3.5a5.5 5.5 0 100 11 5.5 5.5 0 000-11zM2 9a7 7 0 1112.452 4.391l3.328 3.329a.75.75 0 11-1.06 1.06l-3.329-3.328A7 7 0 012 9z" clipRule="evenodd"/>
                    </svg>
                    Search All Keywords
                  </button>
                  {searchResults && <button className="jv-sbtn clr" onClick={clearSearch}>Clear</button>}
                </div>
              </div>

              {/* Results grouped by keyword */}
              {searchResults && Object.entries(searchResults.byKeyword).map(([kw, hits]) => (
                <div key={kw} style={{display:"flex",flexDirection:"column",gap:6}}>
                  <div style={{display:"flex",alignItems:"center",gap:6}}>
                    <span style={{background:"#fef3c7",color:"#92400e",fontSize:10.5,fontWeight:700,padding:"2px 8px",borderRadius:6,fontFamily:"monospace"}}>
                      "{kw}"
                    </span>
                    <span style={{fontSize:10,color:"#94a3b8"}}>{hits.length} paragraph{hits.length!==1?"s":""}</span>
                  </div>

                  {hits.length === 0 ? (
                    <div style={{fontSize:11,color:"#94a3b8",padding:"6px 10px",background:"#f8fafc",borderRadius:7}}>
                      No matches in index
                    </div>
                  ) : hits.map(hit => {
                    const fl   = hit.fullText.toLowerCase();
                    const ql   = kw.toLowerCase();
                    const idx  = fl.indexOf(ql);
                    const hi   = idx===-1 ? hit.fullText
                      : `${hit.fullText.slice(0,idx)}<mark>${hit.fullText.slice(idx,idx+kw.length)}</mark>${hit.fullText.slice(idx+kw.length)}`;
                    const isQ  = !!queue.find(x=>x.qid===hit.key);

                    return (
                      <div key={hit.key} className="jv-hit">
                        <div className="jv-hit-top">
                          <span className="jv-hit-nid">ID {hit.nodeId}</span>
                          <span className="jv-hit-file">{hit.nodeFile}</span>
                          <CopyBtn text={buildChunk(hit)} label="Chunk" style={{padding:"2px 7px",fontSize:10,flexShrink:0}}/>
                        </div>
                        <div className="jv-hit-text"
                          dangerouslySetInnerHTML={{__html: hit.fullText.length>100 ? hi.slice(0,100)+"…" : hi}}/>
                        <div className="jv-hit-foot">
                          <button className={`jv-add-btn${isQ?" queued":""}`} onClick={()=>!isQ&&addToQueue(hit)}>
                            {isQ
                              ? <><svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M13.78 4.22a.75.75 0 010 1.06l-7.25 7.25a.75.75 0 01-1.06 0L2.22 9.28a.75.75 0 011.06-1.06L6 10.94l6.72-6.72a.75.75 0 011.06 0z"/></svg>In queue</>
                              : <><svg viewBox="0 0 16 16" fill="currentColor" width="11" height="11"><path d="M7.75 2a.75.75 0 01.75.75V7h4.25a.75.75 0 010 1.5H8.5v4.25a.75.75 0 01-1.5 0V8.5H2.75a.75.75 0 010-1.5H7V2.75A.75.75 0 017.75 2z"/></svg>Queue for Edit</>
                            }
                          </button>
                          <span className="jv-path-tag">para[{hit.paraPath.slice(-2).join("›")}]</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              ))}

              {/* ── STEP 2: Edit Queue ── */}
              {queue.length > 0 && <>
                <div className="jv-sec">
                  <span className="dot pu"/>
                  STEP 2 — Edit Queue
                  <span className="jv-badge pu">{queue.length} item{queue.length!==1?"s":""}</span>
                </div>

                <div style={{fontSize:11,color:"#64748b",lineHeight:1.5}}>
                  Copy each chunk → send to LLM → paste modified JSON back.
                  All applied to <code style={{background:"#f1f5f9",padding:"1px 4px",borderRadius:3}}>master.json</code> in one shot.
                </div>

                {queue.map(item => (
                  <div key={item.qid} className="jv-q-item">
                    <div className="jv-q-hd">
                      <span className="jv-q-nid">ID {item.entry.nodeId}</span>
                      <div className="jv-q-info">
                        <div className="jv-q-file">{item.entry.nodeFile}</div>
                        <div className="jv-q-prev">{item.entry.fullText.slice(0,50)}{item.entry.fullText.length>50?"…":""}</div>
                      </div>
                      <button className="jv-q-rm" onClick={()=>removeFromQueue(item.qid)}>×</button>
                    </div>
                    <div className="jv-q-bd">
                      <CopyBtn text={buildChunk(item.entry)} label="Copy chunk" style={{fontSize:10.5}}/>
                      <textarea
                        className={`jv-edit-ta${item.editErr?" err":""}`}
                        placeholder={'Paste modified paragraph JSON…\n{ "_tag": "w:p", "_children": [...] }'}
                        value={item.editJson}
                        onChange={e => updateQueueItem(item.qid, e.target.value)}
                      />
                      {item.editErr && <div className="jv-q-err">⚠ {item.editErr}</div>}
                    </div>
                  </div>
                ))}

                <button className="jv-apply-btn" onClick={handleApplyAll}
                  disabled={applying || queue.every(i=>!i.editJson.trim())}>
                  {applying
                    ? <><svg className="spinning" viewBox="0 0 24 24" fill="none" width="14" height="14"><circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity=".25"/><path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity=".75"/></svg>Applying…</>
                    : <><svg viewBox="0 0 16 16" fill="currentColor" width="13" height="13"><path d="M11.013 1.427a1.75 1.75 0 012.474 0l1.086 1.086a1.75 1.75 0 010 2.474l-8.61 8.61c-.21.21-.47.364-.756.445l-3.251.93a.75.75 0 01-.927-.928l.929-3.25c.081-.286.235-.547.445-.758l8.61-8.61z"/></svg>
                      Apply {queue.length} Patch{queue.length>1?"es":""} → master.json</>
                  }
                </button>
                {applyErr && <div className="jv-apply-err">⚠ {applyErr}</div>}
              </>}

              {/* ── STEP 3: Patch Log ── */}
              {patchLog.length > 0 && <>
                <div className="jv-sec">
                  <span className="dot gn"/>
                  Applied
                  <span className="jv-badge gn">{patchLog.length} patch{patchLog.length>1?"es":""}</span>
                </div>
                <div className="jv-log">
                  <div className="jv-log-hd"><div className="jv-log-dot"/> master.json updated · index rebuilt</div>
                  <div className="jv-log-list">
                    {patchLog.map(p => (
                      <div key={p.id} className="jv-log-item">
                        <div className="jv-log-file">Node {p.nodeId} · {p.file}</div>
                        <div className="jv-log-prev">{p.prevText||"—"}</div>
                        <div className="jv-log-new">→ {p.newText||"(modified)"}</div>
                        <div className="jv-log-time">{p.time}</div>
                      </div>
                    ))}
                  </div>
                </div>
                <div className="jv-export-card">
                  <div className="jv-export-note">
                    <strong>{patchLog.length} patch{patchLog.length>1?"es":""}</strong> applied.<br/>
                    JSON → XML → ZIP → .docx
                  </div>
                  <button className="jv-export-btn" onClick={handleExport} disabled={exporting}>
                    {exporting ? "Exporting…" : exportDone ? "✓ Exported!"
                      : <><svg viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z"/><path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z"/></svg>Export .docx</>}
                  </button>
                  {exportErr && <div className="jv-apply-err">{exportErr}</div>}
                </div>
              </>}

              {/* Empty state */}
              {!searchResults && queue.length===0 && patchLog.length===0 && (
                <div className="jv-empty">
                  <svg viewBox="0 0 24 24" fill="none" stroke="#e2e8f0" strokeWidth="1.5" width="40" height="40">
                    <path d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                  <div className="jv-empty-t">{isLoading ? "Building index…" : "Index ready"}</div>
                  <div className="jv-empty-s">
                    {isLoading
                      ? "Parsing paragraphs across all XML nodes"
                      : `${indexRef.current.length} paragraphs indexed across ${xmlCount} XML files.\nEnter keywords above to begin.`
                    }
                  </div>
                </div>
              )}

            </div>
          </div>

        </div>
      </div>
    </>
  );
}