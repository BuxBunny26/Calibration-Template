import React, { useState, useMemo } from 'react';
import { supabase } from '../services/supabase';

function formatDate(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  return d.toLocaleDateString('en-ZA', {
    year: 'numeric', month: 'short', day: 'numeric',
    hour: '2-digit', minute: '2-digit',
  });
}

function History({ jobs, loading }) {
  const [searchSerial, setSearchSerial] = useState('');
  const [searchModel, setSearchModel] = useState('');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');

  const models = useMemo(() => {
    if (!jobs) return [];
    const set = new Set();
    jobs.forEach(j => { if (j.equipment_model) set.add(j.equipment_model); });
    return [...set].sort();
  }, [jobs]);

  const filtered = useMemo(() => {
    if (!jobs) return [];
    return jobs.filter(job => {
      if (searchSerial) {
        const serial = (job.serial_number || '').toLowerCase();
        if (!serial.includes(searchSerial.toLowerCase())) return false;
      }
      if (searchModel && job.equipment_model !== searchModel) return false;
      if (dateFrom) {
        const calDate = job.calibration_date || '';
        if (calDate < dateFrom) return false;
      }
      if (dateTo) {
        const calDate = job.calibration_date || '';
        if (calDate > dateTo) return false;
      }
      return true;
    });
  }, [jobs, searchSerial, searchModel, dateFrom, dateTo]);

  const hasFilters = searchSerial || searchModel || dateFrom || dateTo;

  const clearFilters = () => {
    setSearchSerial('');
    setSearchModel('');
    setDateFrom('');
    setDateTo('');
  };

  if (loading) {
    return (
      <div className="loading-skeleton">
        {[1, 2, 3].map(i => (
          <div key={i} className="skeleton-card">
            <div className="skeleton-line" style={{ width: '70%' }}></div>
            <div className="skeleton-line" style={{ width: '40%' }}></div>
          </div>
        ))}
      </div>
    );
  }

  if (!jobs || jobs.length === 0) {
    return (
      <div className="history-empty">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
          <line x1="3" y1="9" x2="21" y2="9"/>
          <line x1="9" y1="21" x2="9" y2="9"/>
        </svg>
        <p>No certificates generated yet.</p>
        <p>Upload files and generate your first certificate!</p>
      </div>
    );
  }

  const getDownloadUrl = (job) => {
    if (!job.output_file_path) return null;
    const { data } = supabase.storage
      .from('calibration-files')
      .getPublicUrl(job.output_file_path);
    return data?.publicUrl;
  };

  return (
    <div className="history-section">
      <div className="history-filters">
        <div className="filter-row">
          <div className="filter-field">
            <label>Serial Number</label>
            <div className="filter-input-wrap">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
              </svg>
              <input
                type="text"
                placeholder="Search serial..."
                value={searchSerial}
                onChange={e => setSearchSerial(e.target.value)}
              />
            </div>
          </div>
          <div className="filter-field">
            <label>Equipment Model</label>
            <select value={searchModel} onChange={e => setSearchModel(e.target.value)}>
              <option value="">All models</option>
              {models.map(m => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Cal. Date From</label>
            <input type="date" value={dateFrom} onChange={e => setDateFrom(e.target.value)} />
          </div>
          <div className="filter-field">
            <label>Cal. Date To</label>
            <input type="date" value={dateTo} onChange={e => setDateTo(e.target.value)} />
          </div>
        </div>
        {hasFilters && (
          <div className="filter-status">
            <span>{filtered.length} of {jobs.length} result{jobs.length !== 1 && 's'}</span>
            <button className="filter-clear-btn" onClick={clearFilters}>Clear filters</button>
          </div>
        )}
      </div>

      {filtered.length === 0 ? (
        <div className="history-empty">
          <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
            <circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>
          </svg>
          <p>No certificates match your filters.</p>
        </div>
      ) : (
        filtered.map(job => {
          const downloadUrl = getDownloadUrl(job);
          return (
            <div key={job.id} className="history-card">
              <div className={`job-icon ${job.status}`}>
                {job.status === 'completed' && (
                  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>
                  </svg>
                )}
                {job.status === 'failed' && (
                  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>
                  </svg>
                )}
                {(job.status === 'pending' || job.status === 'processing') && (
                  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>
                  </svg>
                )}
              </div>
              <div className="job-info">
                <div className="job-title">
                  {job.output_file_name || `${job.equipment_model || 'Unknown'} — ${job.serial_number || ''}`}
                </div>
                <div className="job-meta">
                  <span>{formatDate(job.created_at)}</span>
                  {job.equipment_model && <span>{job.equipment_model}</span>}
                  {job.serial_number && <span>S/N {job.serial_number}</span>}
                  {job.calibration_date && <span>Cal: {job.calibration_date}</span>}
                  <span className={`status-badge ${job.status}`}>{job.status}</span>
                </div>
                {job.status === 'failed' && job.error_message && (
                  <div style={{ fontSize: 12, color: '#DC2626', marginTop: 4 }}>{job.error_message}</div>
                )}
              </div>
              <div className="job-actions">
                {job.status === 'completed' && downloadUrl ? (
                  <a href={downloadUrl} target="_blank" rel="noopener noreferrer" className="download-btn">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                      <polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>
                    </svg>
                    Download
                  </a>
                ) : job.status === 'completed' && (
                  <span className="download-unavailable">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>
                    </svg>
                    File unavailable
                  </span>
                )}
              </div>
            </div>
          );
        })
      )}
    </div>
  );
}

export default History;
