import React, { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { supabase, supabaseConfigured } from './services/supabase';
import { processGroup } from './services/generateReport';
import Header from './components/Header';
import FileUpload from './components/FileUpload';
import History from './components/History';
import './App.css';

function App() {
  const [files, setFiles] = useState([]);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(null);
  const [history, setHistory] = useState([]);
  const [historyLoading, setHistoryLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('generate');
  const [darkMode, setDarkMode] = useState(() => localStorage.getItem('theme') === 'dark');
  const [generatingProgress, setGeneratingProgress] = useState(null);
  const [toast, setToast] = useState(null);
  const [supabaseOnline, setSupabaseOnline] = useState(supabaseConfigured);

  const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', darkMode ? 'dark' : 'light');
    localStorage.setItem('theme', darkMode ? 'dark' : 'light');
  }, [darkMode]);

  const loadHistory = useCallback(async () => {
    setHistoryLoading(true);
    const { data, error: err } = await supabase
      .from('certificate_jobs')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(50);

    if (!err && data) setHistory(data);
    setHistoryLoading(false);
  }, []);

  useEffect(() => {
    loadHistory();
  }, [loadHistory]);

  const extractGroupKey = (filename) => {
    const lower = filename.toLowerCase();
    let base = lower.includes('.') ? lower.substring(0, lower.lastIndexOf('.')) : lower;
    base = base.replace(/\b(acc|vel)\b/g, '').trim().replace(/\s+/g, ' ');
    return base || 'default';
  };

  const classifyFile = (filename) => {
    const lower = filename.toLowerCase();
    const ext = lower.split('.').pop();
    if (!['xlsx', 'pdf'].includes(ext)) return { kind: null, ext, groupKey: null };
    const kind = lower.includes('acc') ? 'acc' : lower.includes('vel') ? 'vel' : null;
    const groupKey = kind ? extractGroupKey(filename) : null;
    return { kind, ext, groupKey };
  };

  const handleFiles = (newFiles) => {
    const oversized = [];
    const classified = Array.from(newFiles).map(f => {
      if (f.size > MAX_FILE_SIZE) { oversized.push(f.name); return null; }
      return { file: f, ...classifyFile(f.name) };
    }).filter(f => f && f.kind && f.groupKey && ['xlsx', 'pdf'].includes(f.ext));

    if (oversized.length > 0) {
      setToast({ message: `${oversized.length} file(s) skipped — exceeds 50 MB limit`, type: 'warning' });
    }

    const replacedCount = classified.filter(newF =>
      files.some(e => e.groupKey === newF.groupKey && e.kind === newF.kind && e.ext === newF.ext)
    ).length;

    setFiles(prev => {
      const combined = [...prev];
      classified.forEach(newF => {
        const existingIdx = combined.findIndex(
          e => e.groupKey === newF.groupKey && e.kind === newF.kind && e.ext === newF.ext
        );
        if (existingIdx >= 0) combined[existingIdx] = newF;
        else combined.push(newF);
      });
      return combined;
    });

    if (replacedCount > 0 && oversized.length === 0) {
      setToast({ message: `${replacedCount} file(s) replaced with newer version`, type: 'info' });
    }

    setError(null);
    setSuccess(null);
  };

  const removeFile = (index) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
  };

  const removeGroup = (groupKey) => {
    setFiles(prev => prev.filter(f => f.groupKey !== groupKey));
  };

  const groups = useMemo(() => {
    const map = {};
    files.forEach((f, idx) => {
      if (!map[f.groupKey]) map[f.groupKey] = [];
      map[f.groupKey].push({ ...f, idx });
    });
    return map;
  }, [files]);

  const groupKeys = Object.keys(groups);

  const validation = useMemo(() => {
    if (groupKeys.length === 0) {
      return { ready: false, level: 'info', message: 'Upload ACC and/or VEL files (Excel + PDF pair) for each instrument.' };
    }

    const missing = [];
    groupKeys.forEach(key => {
      const gFiles = groups[key];
      const hasAccXlsx = gFiles.some(f => f.kind === 'acc' && f.ext === 'xlsx');
      const hasVelXlsx = gFiles.some(f => f.kind === 'vel' && f.ext === 'xlsx');
      const hasAccPdf = gFiles.some(f => f.kind === 'acc' && f.ext === 'pdf');
      const hasVelPdf = gFiles.some(f => f.kind === 'vel' && f.ext === 'pdf');
      const hasAnyAcc = hasAccXlsx || hasAccPdf;
      const hasAnyVel = hasVelXlsx || hasVelPdf;
      if (!hasAnyAcc && !hasAnyVel) { missing.push(`"${key}" has no valid files`); return; }
      if (hasAnyAcc && !hasAccXlsx) missing.push(`"${key}" has ACC PDF but missing ACC Excel`);
      if (hasAnyAcc && !hasAccPdf) missing.push(`"${key}" has ACC Excel but missing ACC PDF`);
      if (hasAnyVel && !hasVelXlsx) missing.push(`"${key}" has VEL PDF but missing VEL Excel`);
      if (hasAnyVel && !hasVelPdf) missing.push(`"${key}" has VEL Excel but missing VEL PDF`);
    });

    if (missing.length > 0) {
      return { ready: false, level: 'error', message: missing.join('. ') + '.' };
    }

    const count = groupKeys.length;
    const label = count === 1 ? '1 instrument' : `${count} instruments`;
    return { ready: true, level: 'success', message: `${label} ready. Click Generate to create certificates.` };
  }, [groups, groupKeys]);

  const handleGenerate = async () => {
    if (!validation.ready || generating) return;
    setGenerating(true);
    setError(null);
    setSuccess(null);

    const allResults = [];
    const total = groupKeys.length;

    try {
      for (let i = 0; i < groupKeys.length; i++) {
        const key = groupKeys[i];
        setGeneratingProgress({ current: i + 1, total, currentGroup: key });
        const groupFiles = groups[key].map(f => f.file);
        try {
          const result = await processGroup(groupFiles);
          allResults.push({ ...result, group: key });
        } catch (e) {
          allResults.push({ status: 'failed', group: key, error: e.message });
        }
      }

      const completed = allResults.filter(r => r.status === 'completed');
      const failed = allResults.filter(r => r.status === 'failed');

      if (completed.length > 0) {
        setSuccess({
          message: `${completed.length} of ${allResults.length} certificate(s) generated successfully!`,
          results: completed,
        });
      }

      if (failed.length > 0 && completed.length === 0) {
        setError(`All ${failed.length} certificate(s) failed to generate.`);
      } else if (failed.length > 0) {
        setError(`${failed.length} failed: ${failed.map(f => f.group).join(', ')}`);
      }

      setFiles([]);
      loadHistory();
    } catch (err) {
      setError(err.message);
    } finally {
      setGenerating(false);
      setGeneratingProgress(null);
    }
  };

  const handleGenerateRef = useRef();
  handleGenerateRef.current = handleGenerate;

  // Enter key shortcut
  useEffect(() => {
    const onKeyDown = (e) => {
      if (e.key === 'Enter' && !e.repeat) handleGenerateRef.current();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  // Auto-dismiss error after 10 seconds
  useEffect(() => {
    if (!error) return;
    const timer = setTimeout(() => setError(null), 10000);
    return () => clearTimeout(timer);
  }, [error]);

  // Auto-dismiss success after 15 seconds
  useEffect(() => {
    if (!success) return;
    const timer = setTimeout(() => setSuccess(null), 15000);
    return () => clearTimeout(timer);
  }, [success]);

  // Auto-dismiss toast after 4 seconds
  useEffect(() => {
    if (!toast) return;
    const timer = setTimeout(() => setToast(null), 4000);
    return () => clearTimeout(timer);
  }, [toast]);

  const downloadAllAsZip = async () => {
    if (!success || success.results.length < 2) return;
    try {
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();
      for (const r of success.results) {
        if (r._pdfBytes) {
          zip.file(r.output_file_name, r._pdfBytes);
        } else {
          const resp = await fetch(r.download_url);
          const blob = await resp.blob();
          zip.file(r.output_file_name, blob);
        }
      }
      const content = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(content);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'calibration_certificates.zip';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch {
      setToast({ message: 'Failed to create ZIP file', type: 'warning' });
    }
  };

  return (
    <div className="app">
      <Header darkMode={darkMode} onToggleDark={() => setDarkMode(d => !d)} />

      {!supabaseOnline && (
        <div className="offline-banner">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <line x1="1" y1="1" x2="23" y2="23"/>
            <path d="M16.72 11.06A10.94 10.94 0 0 1 19 12.55"/>
            <path d="M5 12.55a10.94 10.94 0 0 1 5.17-2.39"/>
            <path d="M10.71 5.05A16 16 0 0 1 22.56 9"/>
            <path d="M1.42 9a15.91 15.91 0 0 1 4.7-2.88"/>
            <path d="M8.53 16.11a6 6 0 0 1 6.95 0"/>
            <line x1="12" y1="20" x2="12.01" y2="20"/>
          </svg>
          <span>Running in offline mode — history and cloud storage unavailable</span>
        </div>
      )}

      <nav className="tab-nav">
        <button
          className={`tab-btn ${activeTab === 'generate' ? 'active' : ''}`}
          onClick={() => setActiveTab('generate')}
        >
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
            <polyline points="14 2 14 8 20 8"/>
            <line x1="12" y1="18" x2="12" y2="12"/>
            <line x1="9" y1="15" x2="15" y2="15"/>
          </svg>
          Generate Certificate
        </button>
        <button
          className={`tab-btn ${activeTab === 'history' ? 'active' : ''}`}
          onClick={() => { setActiveTab('history'); loadHistory(); }}
        >
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="10"/>
            <polyline points="12 6 12 12 16 14"/>
          </svg>
          History
          {history.length > 0 && <span className="badge">{history.length}</span>}
        </button>
      </nav>

      <main className="main-content">
        {activeTab === 'generate' && (
          <div className="generate-tab">
            <FileUpload onFiles={handleFiles} />

            {groupKeys.length > 0 && (
              <div className="file-list">
                <h3>
                  Selected Files &mdash; {groupKeys.length} instrument{groupKeys.length !== 1 && 's'}
                </h3>
                {groupKeys.map(key => (
                  <div key={key} className="instrument-group">
                    <div className="group-header">
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="2" y="3" width="20" height="14" rx="2" ry="2"/>
                        <line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>
                      </svg>
                      <span className="group-name">{key}</span>
                      <span className="group-count">{groups[key].length} file{groups[key].length !== 1 && 's'}</span>
                      <button className="group-remove-btn" onClick={() => removeGroup(key)} title="Remove group">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                          <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                        </svg>
                      </button>
                    </div>
                    <div className="file-chips">
                      {groups[key].map(f => (
                        <div key={f.idx} className={`file-chip ${f.ext}`}>
                          <span className={`file-tag ${f.kind}`}>{f.kind.toUpperCase()}</span>
                          <span className="file-tag ext">{f.ext.toUpperCase()}</span>
                          <span className="file-name">{f.file.name}</span>
                          <button className="remove-btn" onClick={() => removeFile(f.idx)} title="Remove">
                            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                            </svg>
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            )}

            <div className={`validation-bar ${validation.level}`}>
              {validation.level === 'error' && (
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
              )}
              {validation.level === 'warning' && (
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/>
                  <line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
                </svg>
              )}
              {validation.level === 'success' && (
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
              )}
              {validation.level === 'info' && (
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
              )}
              <span>{validation.message}</span>
            </div>

            {error && (
              <div className="alert alert-error">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
                </svg>
                <span>{error}</span>
                <button className="alert-close" onClick={() => setError(null)}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                  </svg>
                </button>
              </div>
            )}

            {success && (
              <div className="alert alert-success">
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>
                </svg>
                <div>
                  <p>{success.message}</p>
                  {success.results.map((r, i) => (
                    <a key={i} href={r.download_url} target="_blank" rel="noopener noreferrer" className="download-link">
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      Download {r.output_file_name}
                    </a>
                  ))}
                  {success.results.length > 1 && (
                    <button className="zip-btn" onClick={downloadAllAsZip}>
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                        <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                      </svg>
                      Download All as ZIP
                    </button>
                  )}
                </div>
                <button className="alert-close" onClick={() => setSuccess(null)}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                  </svg>
                </button>
              </div>
            )}

            {generating && generatingProgress && (
              <div className="progress-section">
                <div className="progress-text">
                  Processing {generatingProgress.current} of {generatingProgress.total}: <strong>{generatingProgress.currentGroup.toUpperCase()}</strong>
                </div>
                <div className="progress-track">
                  <div className="progress-fill" style={{ width: `${(generatingProgress.current / generatingProgress.total) * 100}%` }}></div>
                </div>
              </div>
            )}

            <button
              className="generate-btn"
              disabled={!validation.ready || generating}
              onClick={handleGenerate}
            >
              {generating ? (
                <>
                  <span className="spinner"></span>
                  Generating{generatingProgress ? ` (${generatingProgress.current}/${generatingProgress.total})` : '...'}
                </>
              ) : (
                <>
                  <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                    <line x1="16" y1="13" x2="8" y2="13"/>
                    <line x1="16" y1="17" x2="8" y2="17"/>
                    <polyline points="10 9 9 9 8 9"/>
                  </svg>
                  Generate Certificates
                </>
              )}
            </button>
          </div>
        )}

        {activeTab === 'history' && (
          <History jobs={history} loading={historyLoading} />
        )}
      </main>

      <footer className="app-footer">
        <p>WearCheck ARC &mdash; Condition Monitoring Division</p>
      </footer>

      {toast && (
        <div className={`toast toast-${toast.type}`}>
          <span>{toast.message}</span>
          <button className="toast-close" onClick={() => setToast(null)}>
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
            </svg>
          </button>
        </div>
      )}
    </div>
  );
}

export default App;
