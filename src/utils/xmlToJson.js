/**
 * xmlToJson.js
 * ─────────────────────────────────────────────────────────────────────────
 * DOCX  →  Node Map  →  JSON tree  →  (LLM modifies)  →  XML  →  .docx
 * ─────────────────────────────────────────────────────────────────────────
 *
 * Node structure:
 *   nodeId   : 1, 2, 3 … (sequential, stable)
 *   path     : "word/document.xml"
 *   fileName : "document.xml"
 *   type     : "xml" | "binary"
 *   content  : JSON object (xml) | { _binary, _type, _size, _raw } (binary)
 */

/* ── HELPERS ──────────────────────────────────────────────────────────── */

const BINARY_EXTS = new Set([
  "png","jpg","jpeg","gif","bmp","svg","webp","emf","wmf","tiff",
  "ttf","otf","woff","woff2","bin","vml",
]);

const isBinaryPath = (p) => BINARY_EXTS.has(p.split(".").pop().toLowerCase());

/* ── XML → JSON ──────────────────────────────────────────────────────── */

/**
 * Converts a DOM Node recursively into a plain JSON object.
 * Shape: { _tag, _attrs?, _text?, _children? }
 */
export function domNodeToJson(node) {
  if (node.nodeType === 3) {
    const t = node.textContent?.trim();
    return t ? { _text: t } : null;
  }
  if (node.nodeType !== 1) return null;

  const obj = { _tag: node.nodeName };

  if (node.attributes?.length) {
    obj._attrs = {};
    for (const a of node.attributes) obj._attrs[a.name] = a.value;
  }

  const kids = [];
  for (const c of node.childNodes) {
    const p = domNodeToJson(c);
    if (p !== null) kids.push(p);
  }

  if (kids.length === 1 && kids[0]._text !== undefined) {
    obj._text = kids[0]._text;
  } else if (kids.length > 0) {
    obj._children = kids;
  }

  return obj;
}

export function xmlStringToJson(xmlString) {
  try {
    const doc = new DOMParser().parseFromString(xmlString, "application/xml");
    const err = doc.getElementsByTagName("parsererror")[0];
    if (err) return { _error: err.textContent?.slice(0, 200) };
    return domNodeToJson(doc.documentElement);
  } catch (e) {
    return { _error: e.message };
  }
}

/* ── JSON → XML ──────────────────────────────────────────────────────── */

const ESC = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" };
const escXml = (s) => String(s).replace(/[&<>"]/g, c => ESC[c]);

/**
 * Converts a JSON node (produced by domNodeToJson) back into an XML string.
 */
export function jsonToXmlString(node, indent = 0) {
  if (!node || typeof node !== "object") return "";
  if (node._error) return `<!-- parse error: ${node._error} -->`;

  const tag   = node._tag || "node";
  const attrs = node._attrs
    ? Object.entries(node._attrs).map(([k, v]) => ` ${k}="${escXml(v)}"`).join("")
    : "";

  // Self-closing if empty
  if (node._text === undefined && !node._children) {
    return `<${tag}${attrs}/>`;
  }

  // Text-only
  if (node._text !== undefined && !node._children) {
    return `<${tag}${attrs}>${escXml(node._text)}</${tag}>`;
  }

  // Children
  const inner = (node._children || []).map(c => jsonToXmlString(c)).join("");
  return `<${tag}${attrs}>${inner}</${tag}>`;
}

/* ── ZIP → Node Map ──────────────────────────────────────────────────── */

/**
 * Reads every file from a JSZip instance.
 * Returns an ordered array of Node objects:
 * [
 *   { nodeId: 1, path: "[Content_Types].xml", fileName: "[Content_Types].xml",
 *     type: "xml", content: { _tag, … } },
 *   { nodeId: 2, path: "word/document.xml",   fileName: "document.xml",
 *     type: "xml", content: { _tag, … } },
 *   { nodeId: 3, path: "word/media/img1.png", fileName: "img1.png",
 *     type: "binary", content: { _binary:true, _type:"image/png", _size:1234,
 *                                 _raw: Uint8Array } },
 *   …
 * ]
 */
export async function zipToNodeMap(zip) {
  const paths = Object.keys(zip.files).filter(k => !zip.files[k].dir);

  // Deterministic order: root files first, then by path
  paths.sort((a, b) => {
    const da = a.includes("/") ? 1 : 0;
    const db = b.includes("/") ? 1 : 0;
    if (da !== db) return da - db;
    return a.localeCompare(b);
  });

  const nodes = await Promise.all(
    paths.map(async (path, idx) => {
      const entry    = zip.files[path];
      const fileName = path.split("/").pop();
      const ext      = fileName.split(".").pop().toLowerCase();

      if (isBinaryPath(path)) {
        const raw = await entry.async("uint8array");
        return {
          nodeId   : idx + 1,
          path,
          fileName,
          type     : "binary",
          content  : {
            _binary : true,
            _type   : `${ext}`,
            _size   : raw.byteLength,
            _raw    : raw,          // kept for export, not serialised to UI
          },
        };
      } else {
        const xml  = await entry.async("string");
        const json = xmlStringToJson(xml);
        return {
          nodeId   : idx + 1,
          path,
          fileName,
          type     : "xml",
          rawXml   : xml,           // kept for fallback export
          content  : json,
        };
      }
    })
  );

  return nodes;
}

/* ── Node Map → .docx (ArrayBuffer) ─────────────────────────────────── */

/**
 * Takes the (possibly LLM-modified) nodeMap and the original JSZip instance,
 * rebuilds every XML file from its JSON content, repacks the ZIP, and
 * returns a new ArrayBuffer ready for download as .docx.
 */
export async function nodeMapToDocxBuffer(nodeMap, originalZip) {
  const JSZipModule = await import("jszip");
  const JSZip       = JSZipModule.default || JSZipModule;
  const newZip      = new JSZip();

  // Copy every folder structure implicitly via file paths
  for (const node of nodeMap) {
    if (node.type === "binary") {
      // Re-use original binary
      newZip.file(node.path, node.content._raw || originalZip.files[node.path].async("uint8array"));
    } else {
      // Convert (possibly modified) JSON back to XML
      const xmlDecl   = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
      const xmlString = jsonToXmlString(node.content);
      newZip.file(node.path, xmlDecl + xmlString);
    }
  }

  const buffer = await newZip.generateAsync({
    type              : "arraybuffer",
    mimeType          : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression       : "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return buffer;
}