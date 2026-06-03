import React, { useState } from 'react';
import { supabase } from '../services/supabase';

const ALLOWED_DOMAIN = 'wearcheckrs.com';

export default function Login() {
  const [step, setStep] = useState('email'); // 'email' | 'sent'
  const [email, setEmail] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [resendCooldown, setResendCooldown] = useState(0);

  const isValidDomain = (addr) =>
    addr.toLowerCase().trim().endsWith(`@${ALLOWED_DOMAIN}`);

  const sendLink = async (addr) => {
    const { error: authError } = await supabase.auth.signInWithOtp({
      email: addr,
      options: {
        shouldCreateUser: true,
        emailRedirectTo: window.location.origin,
      },
    });
    return authError;
  };

  const handleSend = async (e) => {
    e.preventDefault();
    setError(null);

    const trimmed = email.trim().toLowerCase();
    if (!isValidDomain(trimmed)) {
      setError(`Only @${ALLOWED_DOMAIN} email addresses are allowed.`);
      return;
    }

    setLoading(true);
    const authError = await sendLink(trimmed);
    setLoading(false);

    if (authError) {
      setError(authError.message);
      return;
    }

    setStep('sent');
    startResendCooldown();
  };

  const handleResend = async () => {
    if (resendCooldown > 0) return;
    setError(null);
    setLoading(true);
    const authError = await sendLink(email.trim().toLowerCase());
    setLoading(false);
    if (authError) {
      setError(authError.message);
    } else {
      startResendCooldown();
    }
  };

  const startResendCooldown = () => {
    setResendCooldown(60);
    const interval = setInterval(() => {
      setResendCooldown((c) => {
        if (c <= 1) { clearInterval(interval); return 0; }
        return c - 1;
      });
    }, 1000);
  };

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-logo">
          <img src="/WearCheck Logo.png" alt="WearCheck" className="login-logo-img" />
        </div>

        <div className="login-header">
          <h1 className="login-title">WearCheck ARC</h1>
          <p className="login-subtitle">Calibration Certificate Generator</p>
        </div>

        {step === 'email' ? (
          <form className="login-form" onSubmit={handleSend} noValidate>
            <div className="login-field">
              <label className="login-label" htmlFor="login-email">
                Work email address
              </label>
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

            {error && (
              <div className="login-error">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="8" x2="12" y2="12"/>
                  <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                {error}
              </div>
            )}

            <button type="submit" className="login-btn" disabled={loading || !email.trim()}>
              {loading ? (
                <span className="login-spinner" />
              ) : (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <line x1="22" y1="2" x2="11" y2="13"/>
                    <polygon points="22 2 15 22 11 13 2 9 22 2"/>
                  </svg>
                  Send Sign-In Link
                </>
              )}
            </button>

            <p className="login-hint">
              Access is restricted to <strong>@{ALLOWED_DOMAIN}</strong> accounts.
            </p>
          </form>
        ) : (
          <div className="login-form">
            <div className="login-sent-notice login-sent-notice--large">
              <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
                <polyline points="22,6 12,13 2,6"/>
              </svg>
              <div>
                <strong>Check your inbox</strong>
                <span>A sign-in link has been sent to</span>
                <span className="login-sent-email">{email.trim().toLowerCase()}</span>
              </div>
            </div>

            <p className="login-instructions">
              Open the email from <strong>WearCheck ARC</strong> and click the <strong>Log In</strong> link. You'll be signed in automatically — no password needed.
            </p>

            {error && (
              <div className="login-error">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="8" x2="12" y2="12"/>
                  <line x1="12" y1="16" x2="12.01" y2="16"/>
                </svg>
                {error}
              </div>
            )}

            <div className="login-resend-row">
              <button
                type="button"
                className="login-back-btn"
                onClick={() => { setStep('email'); setError(null); }}
              >
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <polyline points="15 18 9 12 15 6"/>
                </svg>
                Change email
              </button>
              <button
                type="button"
                className="login-resend-btn"
                onClick={handleResend}
                disabled={resendCooldown > 0 || loading}
              >
                {loading ? <span className="login-spinner login-spinner--sm" /> : resendCooldown > 0 ? `Resend in ${resendCooldown}s` : 'Resend link'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
