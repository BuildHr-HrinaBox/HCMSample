import React, { useState } from 'react';

const SimpleLogin = () => {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');

  const handleSubmit = (e) => {
    e.preventDefault();
    // For development: allow any login and set auth in localStorage so ProtectedRoute passes
    if (typeof window !== 'undefined' && window.catalyst && window.catalyst.auth) {
      window.catalyst.auth.signIn({ email, password }).catch(() => {});
    }
    // If no Catalyst, just show message
    alert('Use Catalyst auth or configure login.');
  };

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#FFEFD5', padding: 20 }}>
      <div style={{ background: '#fff', padding: 40, borderRadius: 12, boxShadow: '0 4px 20px rgba(0,0,0,0.08)', maxWidth: 400, width: '100%' }}>
        <h2 style={{ margin: '0 0 24px', color: '#333' }}>HCM Login</h2>
        <form onSubmit={handleSubmit}>
          <div style={{ marginBottom: 16 }}>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600, color: '#555' }}>Email</label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              style={{ width: '100%', padding: '10px 12px', border: '1px solid #ddd', borderRadius: 8, boxSizing: 'border-box' }}
              placeholder="your@email.com"
            />
          </div>
          <div style={{ marginBottom: 24 }}>
            <label style={{ display: 'block', marginBottom: 8, fontWeight: 600, color: '#555' }}>Password</label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              style={{ width: '100%', padding: '10px 12px', border: '1px solid #ddd', borderRadius: 8, boxSizing: 'border-box' }}
              placeholder="••••••••"
            />
          </div>
          <button
            type="submit"
            style={{
              width: '100%',
              padding: '12px',
              background: '#8B4513',
              color: '#fff',
              border: 'none',
              borderRadius: 8,
              fontWeight: 600,
              cursor: 'pointer',
            }}
          >
            Sign In
          </button>
        </form>
        <p style={{ marginTop: 16, fontSize: 12, color: '#888' }}>
          HR Compliance Management – use Catalyst auth in production.
        </p>
      </div>
    </div>
  );
};

export default SimpleLogin;
