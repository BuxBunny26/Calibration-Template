import React, { useRef, useState } from 'react';

function FileUpload({ onFiles }) {
  const inputRef = useRef(null);
  const [dragover, setDragover] = useState(false);

  const handleDrop = (e) => {
    e.preventDefault();
    setDragover(false);
    if (e.dataTransfer.files.length > 0) {
      onFiles(e.dataTransfer.files);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    setDragover(true);
  };

  const handleDragLeave = () => setDragover(false);

  const handleClick = () => inputRef.current?.click();

  const handleChange = (e) => {
    if (e.target.files.length > 0) {
      onFiles(e.target.files);
      e.target.value = '';
    }
  };

  return (
    <div
      className={`upload-zone ${dragover ? 'dragover' : ''}`}
      onClick={handleClick}
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
    >
      <div className="upload-icon">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="17 8 12 3 7 8"/>
          <line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
      </div>
      <div className="upload-title">Drop files here or click to browse</div>
      <div className="upload-subtitle">Upload ACC and/or VEL files (.xlsx + .pdf) for each instrument</div>
      <div className="upload-hint">Each test type needs a matching pair: Excel (.xlsx) + PDF report</div>
      <input
        ref={inputRef}
        className="upload-input"
        type="file"
        multiple
        accept=".xlsx,.pdf"
        onChange={handleChange}
      />
    </div>
  );
}

export default FileUpload;
