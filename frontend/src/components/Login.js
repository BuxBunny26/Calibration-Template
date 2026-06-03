import React, { useState } from 'react';
import { supabase } from '../services/supabase';

const ALLOWED_DOMAIN = 'wearcheckrs.com';

export default function Login() {
  const [step, setStep] = useState('email'); // 'email' | 'otp'
  const [email, setEmail] = useState('');
  const [otp, setOtp] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [resendCooldown, setResendCooldown] = useState(0);

  const isValidDomain = (addr) =>
    addr.toLowerCase().trim().endsWith(`@${ALLOWED_DOMAIN}`);

  const handleSendOtp = async (e) => {
    e.preventDefault();
    setError(null);

    const trimmed = email.trim().toLowerCase();
    if (!isValidDomain(trimmed)) {
      setError(`Only @${ALLOWED_DOMAIN} email addresses are allowed.`);
      return;
    }

    setLoading(true);
    const { error: authError } = await supabase.auth.signInWithOtp({
      email: trimmed,
      options: { shouldCreateUser: true },
    });
    setLoading(false);

    if (authError) {
      setError(authError.message);
      return;
    }

    setStep('otp');
    startResendCooldown();
  };

  const handleVerifyOtp = async (e) => {
    e.preventDefault();
    setError(null);

    const token = otp.trim();
    if (token.length !== 6 || !/^\d+$/.test(token)) {
      setError('Enter the 6-digit code from your email.');
      return;
    }

    setLoading(true);
    const { error: verifyError } = await supabase.auth.verifyOtp({
      email: email.trim().toLowerCase(),
      token,
      type: 'email',
    });
    setLoading(false);

    if (verifyError) {
      setError('Invalid or expired code. Please try again.');
    }
    // On success, supabase.auth.onAuthStateChange fires in App.js automatically
  };

  const handleResend = async () => {
    if (resendCooldown > 0) return;
    setError(null);
    setLoading(true);
    const { error: authError } = await supabase.auth.signInWithOtp({
      email: email.trim().toLowerCase(),
      options: { shouldCreateUser: true },
    });
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
          <form className="login-form" onSubmit={handleSendOtp} noValidate>
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
                  Send Login Code
                </>
              )}
            </button>

            <p className="login-hint">
              Access is restricted to <strong>@{ALLOWED_DOMAIN}</strong> accounts.
            </p>
          </form>
        ) : (
          <form className="login-form" onSubmit={handleVerifyOtp} noValidate>
            <div className="login-sent-notice">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polyline points="20 6 9 17 4 12"/>
              </svg>
              <div>
                <strong>Code sent</strong>
                <span>Check your inbox at {email.trim().toLowerCase()}</span>
              </div>
            </div>

            <div className="login-field">
              <label className="login-label" htmlFor="login-otp">
                6-digit verification code
              </label>
              <div className="login-input-wrapper">
                <svg className="login-input-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="3" y="11" width="18" height="11" rx="2" ry="2"/>
                  <path d="M7 11V7a5 5 0 0 1 10 0v4"/>
                </svg>
                <input
                  id="login-otp"
                  type="text"
                  inputMode="numeric"
                  className="login-input login-otp-input"
                  placeholder="000000"
                  value={otp}
                  onChange={(e) => { setOtp(e.target.value.replace(/\D/g, '').slice(0, 6)); setError(null); }}
                  autoComplete="one-time-code"
                  autoFocus
                  maxLength={6}
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

            <button type="submit" className="login-btn" disabled={loading || otp.length !== 6}>
              {loading ? (
                <span className="login-spinner" />
              ) : (
                <>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
                  </svg>
                  Verify & Sign In
                </>
              )}
            </button>

            <div className="login-resend-row">
              <button
                type="button"
                className="login-back-btn"
                onClick={() => { setStep('email'); setOtp(''); setError(null); }}
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
                {resendCooldown > 0 ? `Resend in ${resendCooldown}s` : 'Resend code'}
              </button>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}
