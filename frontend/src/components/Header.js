import React, { useState } from 'react';

function Header({ darkMode, onToggleDark }) {
  const [showInfo, setShowInfo] = useState(false);

  return (
    <>
      <header className="app-header">
        <div className="header-logo">
          <img src="/WearCheck Logo.png" alt="WearCheck" className="header-logo-img" />
          <div>
            <div className="header-title">WearCheck ARC</div>
            <div className="header-subtitle">Calibration Certificate Generator</div>
          </div>
        </div>
        <div className="header-actions">
          <button className="theme-toggle" onClick={onToggleDark} title={darkMode ? 'Switch to light mode' : 'Switch to dark mode'}>
            {darkMode ? (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="5"/>
                <line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/>
                <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>
                <line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/>
                <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>
              </svg>
            ) : (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>
              </svg>
            )}
          </button>
          <button className="info-btn" onClick={() => setShowInfo(true)} title="How it works">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
              <circle cx="12" cy="12" r="11" fill="none" stroke="currentColor" strokeWidth="2"/>
              <text x="12" y="17" textAnchor="middle" fontSize="15" fontWeight="700" fontFamily="Georgia, serif" fill="currentColor">i</text>
            </svg>
            <span>Need help?</span>
          </button>
        </div>
      </header>

      {showInfo && (
        <div className="info-overlay" onClick={() => setShowInfo(false)}>
          <div className="info-modal" onClick={e => e.stopPropagation()}>
            <div className="info-modal-header">
              <h2>How It Works</h2>
              <button className="info-close-btn" onClick={() => setShowInfo(false)}>
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
              </button>
            </div>
            <div className="info-modal-body">
              <div className="info-section">
                <h3>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                    <polyline points="14 2 14 8 20 8"/>
                  </svg>
                  Required Files
                </h3>
                <p>For each test type (ACC or VEL) you need <strong>both</strong> files:</p>
                <ul>
                  <li><strong>Excel file (.xlsx)</strong> — contains calibration data read by the system</li>
                  <li><strong>PDF report (.pdf)</strong> — the original graph report from the analyzer</li>
                </ul>
                <p>You can upload ACC only, VEL only, or both — as long as each type has its matching Excel + PDF pair.</p>
              </div>

              <div className="info-section">
                <h3>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <rect x="2" y="3" width="20" height="14" rx="2" ry="2"/>
                    <line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>
                  </svg>
                  Batch Processing
                </h3>
                <p>Drop files for <strong>multiple instruments</strong> at once. The system groups them automatically by the instrument identifier in the filename.</p>
                <p>For example, dropping these 4 files:</p>
                <div className="info-example">
                  B2140 1237749 Acc.xlsx<br/>
                  B2140 1237749 Vel.xlsx<br/>
                  B2140 1237749 Acc.pdf<br/>
                  B2140 1237749 Vel.pdf
                </div>
                <p>...creates one group: <strong>B2140 1237749</strong></p>
              </div>

              <div className="info-section">
                <h3>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
                  </svg>
                  Naming Convention
                </h3>
                <p>File names <strong>must</strong> contain:</p>
                <ul>
                  <li><strong>"Acc"</strong> or <strong>"Vel"</strong> — to identify the test type</li>
                  <li>A <strong>common identifier</strong> — so files can be grouped together (e.g., model + serial number)</li>
                </ul>
              </div>

              <div className="info-section">
                <h3>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>
                  </svg>
                  Output
                </h3>
                <p>For each instrument group, the system generates a full calibration report PDF containing:</p>
                <ul>
                  <li>A branded <strong>Certificate of Calibration</strong> cover page</li>
                  <li>The original <strong>ACC/VEL PDF graph reports</strong> appended</li>
                </ul>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

export default Header;
