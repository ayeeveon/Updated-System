//creotec-certficate\frontend\src\components\ForgotPasswordForm.jsx
import React, { useState, useRef } from 'react';

export default function ForgotPasswordForm({ onBackToLogin }) {
  const [email, setEmail] = useState('');
  const [codeSent, setCodeSent] = useState(false);
  const [code, setCode] = useState(['', '', '', '', '', '']);
  const inputsRef = useRef([]);

  const handleSendCode = async (e) => {
    e.preventDefault();
    if (!email) {
      alert('Please enter your email.');
      return;
    }
    // TODO: Add API call to send verification code here
    // For demo, codeSent is set to true
    setCodeSent(true);
  };

  const handleCodeChange = (e, idx) => {
    const val = e.target.value;
    if (!/^\d?$/.test(val)) return;

    const newCode = [...code];
    newCode[idx] = val;
    setCode(newCode);

    if (val && idx < 5) {
      inputsRef.current[idx + 1].focus();
    }
  };

  const handleKeyDown = (e, idx) => {
    if (e.key === 'Backspace' && !code[idx] && idx > 0) {
      inputsRef.current[idx - 1].focus();
    }
  };

  const handleVerifyCode = () => {
    if (code.some(d => d === '')) {
      alert('Please enter the full 6-digit code.');
      return;
    }
    const enteredCode = code.join('');
    // TODO: Verify code via API here
    alert(`Code entered: ${enteredCode}`);
  };

  return (
    <div className="max-w-md w-full p-4 text-white">
      {!codeSent ? (
        <form onSubmit={handleSendCode} className="space-y-6">
          <h2 className="text-xl font-bold mb-4 text-white opacity-40">Forgot Password</h2>
          <input
            type="email"
            placeholder="Enter your email"
            value={email}
            onChange={e => setEmail(e.target.value)}
            className="w-full p-4 rounded-md bg-zinc-700 placeholder:text-zinc-400 focus:ring-2 focus:ring-violet-500"
            required
          />
          <button
            type="submit"
            className="w-full bg-violet-600 hover:bg-violet-700 transition p-4 rounded-md font-semibold"
          >
            Send Code
          </button>
          <button
            type="button"
            onClick={onBackToLogin}
            className="w-full text-center text-violet-400 hover:underline mt-4"
          >
            Back to Login
          </button>
        </form>
      ) : (
        <div className="space-y-6">
          <h2 className="text-2xl font-bold mb-4">Enter 6-Digit Code</h2>
          <div className="flex justify-center space-x-2">
            {code.map((digit, idx) => (
              <input
                key={idx}
                type="text"
                inputMode="numeric"
                maxLength={1}
                value={digit}
                ref={el => (inputsRef.current[idx] = el)}
                onChange={e => handleCodeChange(e, idx)}
                onKeyDown={e => handleKeyDown(e, idx)}
                className="w-12 h-12 text-center rounded-md bg-zinc-700 text-white text-xl focus:ring-2 focus:ring-violet-500"
              />
            ))}
          </div>
          <button
            type="button"
            onClick={handleVerifyCode}
            disabled={code.some(d => d === '')}
            className={`w-full p-4 rounded-md font-semibold transition ${
              code.some(d => d === '') ? 'bg-violet-400/60 cursor-not-allowed' : 'bg-violet-600 hover:bg-violet-700'
            }`}
          >
            Verify Code
          </button>
          <button
            type="button"
            onClick={onBackToLogin}
            className="w-full text-center text-violet-400 hover:underline"
          >
            Back to Login
          </button>
        </div>
      )}
    </div>
  );
}
