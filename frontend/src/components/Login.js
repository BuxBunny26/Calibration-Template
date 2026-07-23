import React, { useState, useEffect } from 'react';
import { supabase } from '../services/supabase';

const ALLOWED_DOMAIN = 'wearcheckrs.com';

const EyeIcon = ({ open }) => open ? (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94"/>
    <path d="M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19"/>
    <line x1="1" y1="1" x2="23" y2="23"/>
  </svg>
) : (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/>
    <circle cx="12" cy="12" r="3"/>
  </svg>
);

const LockIcon = () => (
  <svg className="login-input-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
    <path d="M7 11V7a5 5 0 0110 0v4"/>
  </svg>
);

export default function Login() {
  const [step, setStep] = useState('login'); // 'login' | 'forgot' | 'forgot-sent' | 'set-password'
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [newPassword, setNewPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [showPassword, setShowPassword] = useState(false);
  const [showNewPassword, setShowNewPassword] = useState(false);

  useEffect(() => {
    if (window.location.hash.includes('type=recovery')) {
      setStep('set-password');
    }
  }, []);

  const isValidDomain = (addr) =>
    addr.toLowerCase().trim().endsWith(`@${ALLOWED_DOMAIN}`);

  const handleSignIn = async (e) => {
    e.preventDefault();
    setError(null);
    const trimmed = email.trim().toLowerCase();
    if (!isValidDomain(trimmed)) {
      setError(`Only @${ALLOWED_DOMAIN} email addresses are allowed.`);
      return;
    }
    setLoading(true);
    const { error: authError } = await supabase.auth.signInWithPassword({ email: trimmed, password });
    setLoading(false);
    if (authError) setError(authError.message);
  };

  const handleForgot = async (e) => {
    e.preventDefault();
    setError(null);
    const trimmed = email.trim().toLowerCase();
    if (!isValidDomain(trimmed)) {
      setError(`Only @${ALLOWED_DOMAIN} email addresses are allowed.`);
      return;
    }
    setLoading(true);
    const { error: authError } = await supabase.auth.resetPasswordForEmail(trimmed, {
      redirectTo: window.location.origin,
    });
    setLoading(false);
    if (authError) {
      setError(authError.message);
    } else {
      setStep('forgot-sent');
    }
  };

  const handleSetPassword = async (e) => {
    e.preventDefault();
    setError(null);
    if (newPassword.length < 8) {
      setError('Password must be at least 8 characters.');
      return;
    }
    if (newPassword !== confirmPassword) {
      setError('Passwords do not match.');
      return;
    }
    setLoading(true);
    const { error: authError } = await supabase.auth.updateUser({ password: newPassword });
    setLoading(false);
    if (authError) setError(authError.message);
    // On success, onAuthStateChange in App.js handles the session
  };

  const ErrorMsg = () => error ? (
    <div className="login-error">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
        <circle cx="12" cy="12" r="10"/>
        <line x1="12" y1="8" x2="12" y2="12"/>
        <line x1="12" y1="16" x2="12.01" y2="16"/>
      </svg>
      {error}
    </div>
  ) : null;

  const CardHeader = () => (
    <>
      <div className="login-logo">
        <img src="/WearCheck Logo.png" alt="WearCheck" className="login-logo-img" />
      </div>
      <div className="login-header">
        <h1 className="login-title">WearCheck ARC</h1>
        <p className="login-subtitle">Calibration Certificate Generator</p>
      </div>
    </>
  );

  // ── Set new password (after clicking reset link in email) ──
  if (step === 'set-password') {
    return (
      <div className="login-page">
        <div className="login-card">
          <CardHeader />
          <form className="login-form" onSubmit={handleSetPassword} noValidate>
            <p className="login-instructions" style={{ textAlign: 'left' }}>
              Choose a new password for your account.
            </p>
            <div className="login-field">
              <label className="login-label" htmlFor="new-password">New password</label>
              <div className="login-input-wrapper">
                <LockIcon />
                <input
                  id="new-password"
                  type={showNewPassword ? 'text' : 'password'}
                  className="login-input login-input--pw"
                  placeholder="Minimum 8 characters"
                  value={newPassword}
                  onChange={(e) => { setNewPassword(e.target.value); setError(null); }}
                  autoFocus
                  required
                />
                <button type="button" className="login-pw-toggle" onClick={() => setShowNewPassword(v => !v)} aria-label="Toggle password visibility">
                  <EyeIcon open={showNewPassword} />
                </button>
              </div>
            </div>
            <div className="login-field">
              <label className="login-label" htmlFor="confirm-password">Confirm password</label>
              <div className="login-input-wrapper">
                <LockIcon />
                <input
                  id="confirm-password"
                  type={showNewPassword ? 'text' : 'password'}
                  className="login-input login-input--pw"
                  placeholder="Re-enter password"
                  value={confirmPassword}
                  onChange={(e) => { setConfirmPassword(e.target.value); setError(null); }}
                  required
                />
              </div>
            </div>
            <ErrorMsg />
            <button type="submit" className="login-btn" disabled={loading || !newPassword || !confirmPassword}>
              {loading ? <span className="login-spinner" /> : 'Set password & sign in'}
            </button>
          </form>
        </div>
      </div>
    );
  }

  // ── Forgot password — email input ──
  if (step === 'forgot') {
    return (
      <div className="login-page">
        <div className="login-card">
          <CardHeader />
          <form className="login-form" onSubmit={handleForgot} noValidate>
            <p className="login-instructions" style={{ textAlign: 'left' }}>
              Enter your work email and we'll send a password reset link.
            </p>
            <div className="login-field">
              <label className="login-label" htmlFor="forgot-email">Work email address</label>
              <div className="login-input-wrapper">
                <svg className="login-input-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                  <polyline points="22,6 12,13 2,6"/>
                </svg>
                <input
                  id="forgot-email"
                  type="email"
                  className="login-input"
                  placeholder={`you@${ALLOWED_DOMAIN}`}
                  value={email}
                  onChange={(e) => { setEmail(e.target.value); setError(null); }}
                  autoFocus
                  required
                />
              </div>
            </div>
            <ErrorMsg />
            <button type="submit" className="login-btn" disabled={loading || !email.trim()}>
              {loading ? <span className="login-spinner" /> : 'Send reset link'}
            </button>
            <button type="button" className="login-back-btn" onClick={() => { setStep('login'); setError(null); }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="15 18 9 12 15 6"/>
              </svg>
              Back to sign in
            </button>
          </form>
        </div>
      </div>
    );
  }

  // ── Forgot password — confirmation ──
  if (step === 'forgot-sent') {
    return (
      <div className="login-page">
        <div className="login-card">
          <CardHeader />
          <div className="login-form">
            <div className="login-sent-notice login-sent-notice--large">
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                <polyline points="22,6 12,13 2,6"/>
              </svg>
              <div>
                <strong>Check your inbox</strong>
                <span>A password reset link was sent to</span>
                <span className="login-sent-email">{email.trim().toLowerCase()}</span>
              </div>
            </div>
            <p className="login-instructions">
              Click the link in the email to set a new password. The link expires in 1 hour.
            </p>
            <button type="button" className="login-back-btn" style={{ alignSelf: 'center' }} onClick={() => { setStep('login'); setError(null); }}>
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="15 18 9 12 15 6"/>
              </svg>
              Back to sign in
            </button>
          </div>
        </div>
      </div>
    );
  }

  // ── Default: email + password sign in ──
  return (
    <div className="login-page">
      <div className="login-card">
        <CardHeader />
        <form className="login-form" onSubmit={handleSignIn} noValidate>
          <div className="login-field">
            <label className="login-label" htmlFor="login-email">Work email address</label>
            <div className="login-input-wrapper">
              <svg className="login-input-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                <polyline points="22,6 12,13 2,6"/>
              </svg>
              <input
                id="login-email"
                type="email"
                className="login-input"
                placeholder={`you@${ALLOWED_DOMAIN}`}
                value={email}
                onChange={(e) => { setEmail(e.target.value); setError(null); }}
                autoComplete="email"
                autoFocus
                required
              />
            </div>
          </div>
          <div className="login-field">
            <div className="login-field-row">
              <label className="login-label" htmlFor="login-password">Password</label>
              <button type="button" className="login-link" onClick={() => { setStep('forgot'); setError(null); }}>
                Forgot password?
              </button>
            </div>
            <div className="login-input-wrapper">
              <LockIcon />
              <input
                id="login-password"
                type={showPassword ? 'text' : 'password'}
                className="login-input login-input--pw"
                placeholder="Your password"
                value={password}
                onChange={(e) => { setPassword(e.target.value); setError(null); }}
                autoComplete="current-password"
                required
              />
              <button type="button" className="login-pw-toggle" onClick={() => setShowPassword(v => !v)} aria-label="Toggle password visibility">
                <EyeIcon open={showPassword} />
              </button>
            </div>
          </div>
          <ErrorMsg />
          <button type="submit" className="login-btn" disabled={loading || !email.trim() || !password}>
            {loading ? (
              <span className="login-spinner" />
            ) : (
              <>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M15 3h4a2 2 0 012 2v14a2 2 0 01-2 2h-4"/>
                  <polyline points="10 17 15 12 10 7"/>
                  <line x1="15" y1="12" x2="3" y2="12"/>
                </svg>
                Sign In
              </>
            )}
          </button>
          <p className="login-hint">
            Access is restricted to <strong>@{ALLOWED_DOMAIN}</strong> accounts.
          </p>
        </form>
      </div>
    </div>
  );
}
