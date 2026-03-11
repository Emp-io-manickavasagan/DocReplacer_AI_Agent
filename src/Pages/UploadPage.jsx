import { useState, useRef, useCallback } from "react";
import EditorPage from "./EditorPage";

export default function UploadPage() {
  const [file,         setFile]         = useState(null);
  const [isDragging,   setIsDragging]   = useState(false);
  const [isLoading,    setIsLoading]    = useState(false);
  const [error,        setError]        = useState("");
  const [progress,     setProgress]     = useState(0);
  const [arrayBuffer,  setArrayBuffer]  = useState(null);
  const [page,         setPage]         = useState("upload"); // "upload" | "editor"
  const inputRef = useRef(null);

  const loadFile = (f) => {
    setError(""); setProgress(0); setPage("upload"); setArrayBuffer(null);
    if (!f) return;
    if (!f.name.toLowerCase().endsWith(".docx")) { setError("Only .docx files are accepted."); return; }
    setFile(f);
  };

  const onDragOver  = useCallback((e) => { e.preventDefault(); setIsDragging(true); }, []);
  const onDragLeave = useCallback(() => setIsDragging(false), []);
  const onDrop      = useCallback((e) => { e.preventDefault(); setIsDragging(false); loadFile(e.dataTransfer.files[0]); }, []);

  const readFile = async () => {
    if (arrayBuffer) return arrayBuffer;
    const buf = await file.arrayBuffer();
    setArrayBuffer(buf);
    return buf;
  };

  const handleOpen = async () => {
    if (!file) return;
    setIsLoading(true); setError(""); setProgress(20);
    try {
      const buf = await readFile(); setProgress(100);
      await new Promise(r => setTimeout(r, 180));
      setArrayBuffer(buf); setPage("editor");
    } catch(err) { setError(err.message); }
    finally { setIsLoading(false); }
  };

  const handleClear = () => {
    setFile(null); setError(""); setProgress(0);
    setArrayBuffer(null); setPage("upload");
    if (inputRef.current) inputRef.current.value = "";
  };

  const fmtBytes = (b) =>
    b < 1024 ? b + " B" : b < 1048576 ? (b/1024).toFixed(1)+" KB" : (b/1048576).toFixed(2)+" MB";

  if (page === "editor") return (
    <EditorPage arrayBuffer={arrayBuffer} fileName={file?.name} onBack={() => setPage("upload")}/>
  );

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Sans:wght@300;400;500;600&display=swap');
        *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

        .up { font-family: 'DM Sans', sans-serif; min-height: 100vh; background: #fff; display: flex; flex-direction: column; }

        .up-nav { height: 62px; border-bottom: 1px solid #f1f5f9; display: flex; align-items: center; justify-content: space-between; padding: 0 40px; }
        @media(max-width:600px){ .up-nav { padding: 0 18px; } }
        .up-nav-logo { display: flex; align-items: center; gap: 10px; }
        .up-nav-icon { width: 32px; height: 32px; background: #1e3a8a; border-radius: 8px; display: flex; align-items: center; justify-content: center; }
        .up-nav-name { font-family: 'DM Serif Display', serif; font-size: 19px; color: #0f172a; letter-spacing: -0.3px; }
        .up-nav-name span { color: #1e3a8a; }
        .up-nav-tag { font-size: 11px; color: #94a3b8; letter-spacing: 0.1em; text-transform: uppercase; font-weight: 500; }
        @media(max-width:480px){ .up-nav-tag { display: none; } }

        .up-hero { flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 56px 16px 72px; }

        .up-title { font-family: 'DM Serif Display', serif; font-size: clamp(2rem,5vw,3rem); color: #0f172a; line-height: 1.1; text-align: center; letter-spacing: -0.5px; margin-bottom: 14px; }
        .up-title span { color: #1e3a8a; }
        .up-desc { font-size: 15px; color: #64748b; line-height: 1.75; font-weight: 300; max-width: 420px; text-align: center; margin-bottom: 16px; }
        .up-desc code { background: #eff6ff; color: #1d4ed8; padding: 2px 7px; border-radius: 5px; font-size: 13px; font-weight: 500; }

        /* mode pills */
        .up-modes { display: flex; gap: 8px; justify-content: center; margin-bottom: 36px; }
        .up-mode { display: flex; align-items: center; gap: 5px; background: #f8faff; border: 1px solid #e0e7ff; border-radius: 20px; padding: 5px 14px; font-size: 12px; color: #475569; font-weight: 500; }
        .up-mode-dot { width: 7px; height: 7px; border-radius: 50%; }

        .up-card { width: 100%; max-width: 480px; border-radius: 20px; border: 1px solid #e2e8f0; overflow: hidden; box-shadow: 0 4px 32px rgba(15,23,42,0.08); }
        .up-card-bar { background: #1e3a8a; padding: 14px 22px; display: flex; align-items: center; gap: 8px; }
        .up-card-bar-label { color: #fff; font-size: 11px; font-weight: 600; letter-spacing: 0.12em; text-transform: uppercase; }
        .up-card-body { background: #fafafa; padding: 22px; display: flex; flex-direction: column; gap: 16px; }

        .up-drop { border: 2px dashed #cbd5e1; border-radius: 14px; padding: 44px 20px; cursor: pointer; text-align: center; background: #fff; transition: all 0.18s ease; user-select: none; }
        .up-drop:hover  { border-color: #93c5fd; background: #f8faff; transform: translateY(-1px); }
        .up-drop.drag   { border-color: #1d4ed8; background: #eff6ff; transform: scale(1.01); }
        .up-drop.loaded { border-color: #93c5fd; background: #f8faff; }
        .up-drop-icon  { width: 50px; height: 50px; border-radius: 12px; display: flex; align-items: center; justify-content: center; margin: 0 auto 12px; }
        .up-drop-title { font-size: 14px; font-weight: 500; color: #1e293b; margin-bottom: 4px; }
        .up-drop-sub   { font-size: 13px; color: #94a3b8; }
        .up-drop-sub u { color: #1d4ed8; text-underline-offset: 3px; }
        .up-drop-hint  { font-size: 11px; color: #cbd5e1; margin-top: 8px; }
        .up-file-name  { font-size: 14px; font-weight: 600; color: #0f172a; margin-bottom: 3px; }
        .up-file-size  { font-size: 12px; color: #94a3b8; }
        .up-file-pill  { display: inline-block; background: #dbeafe; color: #1d4ed8; font-size: 11px; font-weight: 600; padding: 4px 12px; border-radius: 20px; letter-spacing: 0.06em; margin-top: 8px; }

        .up-prog-row   { display: flex; justify-content: space-between; margin-bottom: 6px; }
        .up-prog-label { font-size: 12px; color: #64748b; }
        .up-prog-pct   { font-size: 12px; color: #1d4ed8; font-weight: 600; }
        .up-prog-track { background: #e2e8f0; border-radius: 99px; height: 4px; overflow: hidden; }
        .up-prog-fill  { background: #1e3a8a; height: 100%; border-radius: 99px; transition: width 0.3s ease; }

        .up-err      { display: flex; align-items: flex-start; gap: 8px; background: #fef2f2; border: 1px solid #fecaca; border-radius: 10px; padding: 10px 14px; }
        .up-err-text { color: #dc2626; font-size: 13px; line-height: 1.5; }

        /* Two action buttons stacked */
        .up-btns { display: flex; flex-direction: column; gap: 9px; }
        .up-btn-row { display: flex; gap: 9px; }

        .up-btn { flex: 1; display: flex; align-items: center; justify-content: center; gap: 7px; padding: 12px 0; border-radius: 11px; border: none; font-family: 'DM Sans', sans-serif; font-size: 13.5px; font-weight: 600; letter-spacing: 0.02em; cursor: pointer; transition: all 0.15s; }

        /* Primary — render */
        .up-btn.primary.on  { background: #1e3a8a; color: #fff; }
        .up-btn.primary.on:hover { background: #1e40af; box-shadow: 0 6px 20px rgba(30,58,138,0.25); transform: translateY(-1px); }
        .up-btn.primary.on:active { transform: translateY(0); }
        .up-btn.primary.off { background: #e2e8f0; color: #94a3b8; cursor: not-allowed; }

        /* Secondary — json */
        .up-btn.secondary.on  { background: #fff; color: #1e3a8a; border: 1.5px solid #1e3a8a; }
        .up-btn.secondary.on:hover { background: #eff6ff; box-shadow: 0 4px 14px rgba(30,58,138,0.12); transform: translateY(-1px); }
        .up-btn.secondary.on:active { transform: translateY(0); }
        .up-btn.secondary.off { background: #f8fafc; color: #94a3b8; border: 1.5px solid #e2e8f0; cursor: not-allowed; }

        /* Clear */
        .up-btn.clear { background: #fff; color: #94a3b8; border: 1px solid #e2e8f0; font-size: 13px; flex:0; padding: 12px 16px; }
        .up-btn.clear:hover { border-color: #cbd5e1; color: #64748b; }

        .up-note { margin-top: 22px; display: flex; align-items: center; gap: 5px; color: #cbd5e1; font-size: 12px; }

        @keyframes fadeUp { from { opacity:0; transform:translateY(14px); } to { opacity:1; transform:translateY(0); } }
        .fu  { animation: fadeUp 0.4s ease both; }
        .fu1 { animation-delay: 0.04s; }
        .fu2 { animation-delay: 0.12s; }
        .fu3 { animation-delay: 0.2s;  }
        @keyframes spin { to { transform: rotate(360deg); } }
        .spin { animation: spin 0.75s linear infinite; display: inline-block; }
      `}</style>

      <div className="up">

        <nav className="up-nav fu">
          <div className="up-nav-logo">
            <div className="up-nav-icon">
              <svg viewBox="0 0 20 20" fill="white" width="15" height="15">
                <path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V7.414A2 2 0 0017.414 6L14 2.586A2 2 0 0012.586 2H6a2 2 0 00-2 2z"/>
              </svg>
            </div>
            <div className="up-nav-name">docx<span>view</span></div>
          </div>
          <span className="up-nav-tag">docx-preview · JSZip</span>
        </nav>

        <main className="up-hero">

          <h1 className="up-title fu fu1">Upload a Word Doc<br /><span>Choose Your View</span></h1>
          <p className="up-desc fu fu1">
            Upload a <code>.docx</code> — render it visually with <code>docx-preview</code>, or inspect its raw XML nodes as a structured <code>JSON</code> tree.
          </p>

          {/* Mode pills */}
          <div className="up-modes fu fu1">
            <div className="up-mode">
              <div className="up-mode-dot" style={{background:"#1e3a8a"}}/>
              Render View
            </div>
            <div className="up-mode">
              <div className="up-mode-dot" style={{background:"#7c3aed"}}/>
              JSON Tree
            </div>
          </div>

          <div className="up-card fu fu2">
            <div className="up-card-bar">
              <svg viewBox="0 0 20 20" fill="white" width="14" height="14" style={{opacity:0.8}}>
                <path fillRule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zM6.293 6.707a1 1 0 010-1.414l3-3a1 1 0 011.414 0l3 3a1 1 0 01-1.414 1.414L11 5.414V13a1 1 0 11-2 0V5.414L7.707 6.707a1 1 0 01-1.414 0z" clipRule="evenodd"/>
              </svg>
              <span className="up-card-bar-label">Upload Document</span>
            </div>

            <div className="up-card-body">

              {/* Drop Zone */}
              <div
                className={`up-drop${isDragging?" drag":file?" loaded":""}`}
                onDrop={onDrop} onDragOver={onDragOver} onDragLeave={onDragLeave}
                onClick={() => inputRef.current?.click()}
              >
                <input ref={inputRef} type="file" accept=".docx" hidden onChange={e=>loadFile(e.target.files[0])}/>
                {file ? (
                  <>
                    <div className="up-drop-icon" style={{background:"#dbeafe"}}>
                      <svg viewBox="0 0 24 24" fill="none" width="26" height="26">
                        <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="#1d4ed8" strokeWidth="1.8" strokeLinejoin="round"/>
                        <path d="M14 2v6h6" stroke="#1d4ed8" strokeWidth="1.8" strokeLinejoin="round"/>
                        <path d="M8 13h8M8 17h5" stroke="#1d4ed8" strokeWidth="1.5" strokeLinecap="round"/>
                      </svg>
                    </div>
                    <div className="up-file-name">{file.name}</div>
                    <div className="up-file-size">{fmtBytes(file.size)}</div>
                    <div className="up-file-pill">READY</div>
                  </>
                ) : (
                  <>
                    <div className="up-drop-icon" style={{background:"#f1f5f9"}}>
                      <svg viewBox="0 0 24 24" fill="none" stroke="#94a3b8" strokeWidth="1.5" width="26" height="26">
                        <path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M16 10l-4-4m0 0L8 10m4-4v12" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                    </div>
                    <div className="up-drop-title">Drop your .docx file here</div>
                    <div className="up-drop-sub">or <u>click to browse</u></div>
                    <div className="up-drop-hint">.docx files only</div>
                  </>
                )}
              </div>

              {/* Progress */}
              {isLoading && (
                <div>
                  <div className="up-prog-row">
                    <span className="up-prog-label">Reading file…</span>
                    <span className="up-prog-pct">{progress}%</span>
                  </div>
                  <div className="up-prog-track">
                    <div className="up-prog-fill" style={{width:`${progress}%`}}/>
                  </div>
                </div>
              )}

              {/* Error */}
              {error && (
                <div className="up-err">
                  <svg viewBox="0 0 20 20" fill="#ef4444" width="16" height="16" style={{flexShrink:0,marginTop:1}}>
                    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.28 7.22a.75.75 0 00-1.06 1.06L8.94 10l-1.72 1.72a.75.75 0 101.06 1.06L10 11.06l1.72 1.72a.75.75 0 101.06-1.06L11.06 10l1.72-1.72a.75.75 0 00-1.06-1.06L10 8.94 8.28 7.22z" clipRule="evenodd"/>
                  </svg>
                  <span className="up-err-text">{error}</span>
                </div>
              )}

              {/* Buttons */}
              <div className="up-btns">
                <div className="up-btn-row">
                  {/* Open Editor */}
                  <button
                    className={`up-btn primary ${!file||isLoading?"off":"on"}`}
                    onClick={handleOpen}
                    disabled={!file||isLoading}
                  >
                    {isLoading ? (
                      <>
                        <svg className="spin" viewBox="0 0 24 24" fill="none" width="15" height="15">
                          <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="3" opacity="0.25"/>
                          <path fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z" opacity="0.75"/>
                        </svg>
                        Loading…
                      </>
                    ):(
                      <>
                        <svg viewBox="0 0 20 20" fill="currentColor" width="15" height="15">
                          <path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z"/>
                        </svg>
                        Open in Editor
                      </>
                    )}
                  </button>

                  {/* Clear */}
                  {file && !isLoading && (
                    <button className="up-btn clear" onClick={handleClear}>Clear</button>
                  )}
                </div>
              </div>

            </div>
          </div>

          <div className="up-note fu fu3">
            <svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">
              <path fillRule="evenodd" d="M10 1a4.5 4.5 0 00-4.5 4.5V9H5a2 2 0 00-2 2v6a2 2 0 002 2h10a2 2 0 002-2v-6a2 2 0 00-2-2h-.5V5.5A4.5 4.5 0 0010 1zm3 8V5.5a3 3 0 10-6 0V9h6z" clipRule="evenodd"/>
            </svg>
            Processed entirely in your browser — nothing is sent to any server.
          </div>

        </main>
      </div>
    </>
  );
}