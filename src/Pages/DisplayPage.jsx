import { useEffect, useRef, useState } from "react";
import { renderAsync } from "docx-preview";

export default function DisplayPage({ arrayBuffer, fileName, onBack }) {
  const containerRef  = useRef(null);
  const [isRendering, setIsRendering] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [exportDone,  setExportDone]  = useState(false);
  const [error,       setError]       = useState("");

  useEffect(() => {
    if (!arrayBuffer || !containerRef.current) return;
    setIsRendering(true);
    setError("");

    renderAsync(arrayBuffer, containerRef.current, null, {
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      ignoreLastRenderedPageBreak: true,
      experimental: true,
      trimXmlDeclaration: true,
      useBase64URL: true,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
    })
      .then(() => setIsRendering(false))
      .catch(err => { setError(err.message || "Rendering failed."); setIsRendering(false); });
  }, [arrayBuffer]);

  /* ── Export: download current arrayBuffer as .docx ── */
  const handleExport = async () => {
    if (!arrayBuffer) return;
    setIsExporting(true);
    try {
      const blob = new Blob([arrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      const a   = document.createElement("a");
      a.href    = url;
      a.download = (fileName || "document").replace(/\.docx$/i, "") + "_exported.docx";
      a.click();
      URL.revokeObjectURL(url);
      setExportDone(true);
      setTimeout(() => setExportDone(false), 3000);
    } catch(e) {
      setError("Export failed: " + e.message);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        .dp { font-family:'DM Sans',sans-serif; min-height:100vh; background:#eef2f7; display:flex; flex-direction:column; }

        /* ── Top Bar ── */
        .dp-bar { height:54px; background:#1e3a8a; display:flex; align-items:center; padding:0 24px; gap:12px; position:sticky; top:0; z-index:100; box-shadow:0 2px 16px rgba(15,23,42,0.2); }

        .dp-back { display:flex;align-items:center;gap:6px;background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.18);border-radius:8px;padding:6px 14px;color:#fff;font-size:13px;font-weight:500;cursor:pointer;transition:background 0.15s;white-space:nowrap;flex-shrink:0; }
        .dp-back:hover { background:rgba(255,255,255,0.2); }

        .dp-bar-center { flex:1;display:flex;align-items:center;gap:8px;overflow:hidden;justify-content:center; }
        .dp-file-icon { width:26px;height:26px;background:rgba(255,255,255,0.12);border-radius:6px;display:flex;align-items:center;justify-content:center;flex-shrink:0; }
        .dp-filename { font-size:13px;font-weight:500;color:rgba(255,255,255,0.9);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:280px; }
        @media(max-width:500px){ .dp-filename{max-width:130px;} }

        .dp-bar-right { display:flex;align-items:center;gap:8px;flex-shrink:0; }
        .dp-status-dot { width:8px;height:8px;border-radius:50%;transition:background 0.4s; }
        @keyframes glow{0%,100%{opacity:1}50%{opacity:0.4}} .dp-status-dot.live{background:#4ade80;animation:glow 2s infinite;} .dp-status-dot.loading{background:#fbbf24;}
        .dp-badge { background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.14);border-radius:6px;padding:3px 10px;font-size:11px;font-weight:600;color:rgba(255,255,255,0.55);letter-spacing:0.06em;text-transform:uppercase; }
        @media(max-width:480px){ .dp-badge{display:none;} }

        /* Export button */
        .dp-export-btn { display:flex;align-items:center;gap:6px;background:#fff;border:none;border-radius:8px;padding:6px 14px;font-family:'DM Sans',sans-serif;font-size:13px;font-weight:600;color:#1e3a8a;cursor:pointer;transition:all 0.15s;white-space:nowrap;flex-shrink:0; }
        .dp-export-btn:hover { background:#eff6ff;box-shadow:0 2px 10px rgba(30,58,138,0.2); }
        .dp-export-btn.done { background:#d1fae5;color:#065f46; }
        .dp-export-btn:disabled { background:rgba(255,255,255,0.3);color:rgba(30,58,138,0.5);cursor:not-allowed; }

        /* ── Canvas ── */
        .dp-shell { flex:1;padding:36px 20px 60px;display:flex;flex-direction:column;align-items:center; }
        .dp-label { display:flex;align-items:center;gap:10px;width:100%;max-width:860px;margin-bottom:20px; }
        .dp-label-line { flex:1;height:1px;background:#d1d9e6; }
        .dp-label-text { font-size:10.5px;font-weight:700;color:#94a3b8;letter-spacing:0.14em;text-transform:uppercase;white-space:nowrap; }

        .dp-card { width:100%;max-width:860px;border-radius:16px;overflow:hidden;box-shadow:0 0 0 1px rgba(30,58,138,0.07),0 4px 8px rgba(15,23,42,0.05),0 20px 60px rgba(15,23,42,0.11);background:#fff; }
        .dp-card-stripe { height:4px;background:linear-gradient(90deg,#1e3a8a 0%,#2563eb 55%,#93c5fd 100%); }

        .dp-render-wrap { background:#f0f4f8;min-height:400px; }
        .dp-render-wrap .docx-wrapper { background:#f0f4f8!important;padding:32px 24px!important;display:flex!important;flex-direction:column!important;align-items:center!important;gap:24px!important; }
        .dp-render-wrap .docx-wrapper > section.docx { box-shadow:0 1px 3px rgba(15,23,42,0.08),0 8px 32px rgba(15,23,42,0.10)!important;border-radius:4px!important;margin:0 auto!important; }

        .dp-card-foot { padding:12px 24px;border-top:1px solid #f1f5f9;background:#fff;display:flex;justify-content:space-between;align-items:center; }
        .dp-card-foot span { font-size:11px;color:#cbd5e1; }

        /* Loading */
        .dp-loading { background:#fff;min-height:480px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:14px;padding:60px 20px; }
        .dp-spinner { width:36px;height:36px;border:3px solid #e2e8f0;border-top-color:#1e3a8a;border-radius:50%; }
        @keyframes spin{to{transform:rotate(360deg)}} .dp-spinner{animation:spin 0.7s linear infinite;}
        .dp-loading-title { font-size:14px;color:#475569;font-weight:600; }
        .dp-loading-sub   { font-size:12px;color:#94a3b8; }

        .dp-error { background:#fff;padding:40px 24px;display:flex;align-items:flex-start;gap:10px; }
        .dp-error-text { color:#dc2626;font-size:14px;line-height:1.6; }

        @keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
        .dp-card { animation:fadeIn 0.35s ease both; }

        @keyframes spin2{to{transform:rotate(360deg)}}
        .spin { animation:spin2 0.75s linear infinite;display:inline-block; }
      `}</style>

      <div className="dp">

        {/* ── Top Bar ── */}
        <header className="dp-bar">
          <button className="dp-back" onClick={onBack}>
            <svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">
              <path fillRule="evenodd" d="M17 10a.75.75 0 01-.75.75H5.612l4.158 3.96a.75.75 0 11-1.04 1.08l-5.5-5.25a.75.75 0 010-1.08l5.5-5.25a.75.75 0 111.04 1.08L5.612 9.25H16.25A.75.75 0 0117 10z" clipRule="evenodd"/>
            </svg>
            Back
          </button>

          <div className="dp-bar-center">
            <div className="dp-file-icon">
              <svg viewBox="0 0 16 16" fill="white" width="11" height="11">
                <path d="M4 1h5l4 4v9a1 1 0 01-1 1H4a1 1 0 01-1-1V2a1 1 0 011-1z"/>
              </svg>
            </div>
            <span className="dp-filename">{fileName || "document.docx"}</span>
          </div>

          <div className="dp-bar-right">
            <span className="dp-badge">docx-preview</span>
            <div className={`dp-status-dot ${isRendering ? "loading" : "live"}`} />
          </div>

          {/* Export button */}
          <button
            className={`dp-export-btn${exportDone?" done":""}`}
            onClick={handleExport}
            disabled={isExporting || isRendering}
          >
            {isExporting ? (
              <>
                <svg className="spin" viewBox="0 0 24 24" fill="none" width="14" height="14">
                  <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity="0.25"/>
                  <path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity="0.75"/>
                </svg>
                Exporting…
              </>
            ) : exportDone ? (
              <>
                <svg viewBox="0 0 16 16" fill="currentColor" width="13" height="13">
                  <path fillRule="evenodd" d="M13.78 4.22a.75.75 0 010 1.06l-7.25 7.25a.75.75 0 01-1.06 0L2.22 9.28a.75.75 0 011.06-1.06L6 10.94l6.72-6.72a.75.75 0 011.06 0z" clipRule="evenodd"/>
                </svg>
                Downloaded!
              </>
            ) : (
              <>
                <svg viewBox="0 0 16 16" fill="currentColor" width="13" height="13">
                  <path d="M7.47 10.78a.75.75 0 001.06 0l3.75-3.75a.75.75 0 00-1.06-1.06L8.75 8.44V1.75a.75.75 0 00-1.5 0v6.69L4.78 5.97a.75.75 0 00-1.06 1.06l3.75 3.75z"/>
                  <path d="M3.75 13a.25.25 0 01-.25-.25v-1.5a.75.75 0 00-1.5 0v1.5C2 13.966 2.784 14.75 3.75 14.75h8.5A1.75 1.75 0 0014 13v-1.75a.75.75 0 00-1.5 0V13a.25.25 0 01-.25.25z"/>
                </svg>
                Export .docx
              </>
            )}
          </button>
        </header>

        {/* ── Canvas ── */}
        <main className="dp-shell">
          <div className="dp-label">
            <div className="dp-label-line"/>
            <span className="dp-label-text">Document Canvas</span>
            <div className="dp-label-line"/>
          </div>

          <div className="dp-card">
            <div className="dp-card-stripe"/>

            {isRendering && (
              <div className="dp-loading">
                <div className="dp-spinner"/>
                <div className="dp-loading-title">Rendering document…</div>
                <div className="dp-loading-sub">docx-preview is processing your file</div>
              </div>
            )}

            {error && !isRendering && (
              <div className="dp-error">
                <svg viewBox="0 0 20 20" fill="#ef4444" width="18" height="18" style={{flexShrink:0,marginTop:2}}>
                  <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.28 7.22a.75.75 0 00-1.06 1.06L8.94 10l-1.72 1.72a.75.75 0 101.06 1.06L10 11.06l1.72 1.72a.75.75 0 101.06-1.06L11.06 10l1.72-1.72a.75.75 0 00-1.06-1.06L10 8.94 8.28 7.22z" clipRule="evenodd"/>
                </svg>
                <span className="dp-error-text">{error}</span>
              </div>
            )}

            <div
              ref={containerRef}
              className="dp-render-wrap"
              style={{display: isRendering ? "none" : "block"}}
            />

            {!isRendering && !error && (
              <div className="dp-card-foot">
                <span>docxview · docx-preview</span>
                <span>{new Date().toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"})}</span>
              </div>
            )}
          </div>
        </main>
      </div>
    </>
  );
}