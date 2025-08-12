import { useState } from 'react';
import { togglePasswordVisibility, useAuth } from '../utils/auth';
import { useNavigate } from 'react-router-dom';
import Toast from './Toast';
import LoadingSpinner from './LoadingSpinner';

export default function LoginForm({ isLoading, setIsLoading, onForgotPassword }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showToast, setShowToast] = useState(false);
  const [toastMessage, setToastMessage] = useState('');
  const [toastType, setToastType] = useState('info');
  const { login } = useAuth();
  const navigate = useNavigate();

  const handleLogin = async (e) => {
    e.preventDefault();
    if (!username || !password) {
      triggerToast('Please fill in both fields.', 'error');
      return;
    }

    setIsLoading(true);

    try {
      const result = await login(username, password);
      if (result.success) {
        triggerToast('Login successful!', 'success');
        setTimeout(() => {
          navigate('/home');
        }, 1200);
      } else {
        triggerToast(result.message || 'Invalid username or password.', 'error');
      }
    } catch (err) {
      triggerToast(`Network error: ${err.message}`, 'error');
    } finally {
      setIsLoading(false);
    }
  };

  const triggerToast = (message, type = 'info') => {
    setToastMessage(message);
    setToastType(type);
    setShowToast(true);
    setTimeout(() => setShowToast(false), 3000);
  };

  return (
    <form onSubmit={handleLogin} className="space-y-6 w-full max-w-md">
      <div>
        <input
          type="text"
          name="username"
          placeholder="Username"
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          disabled={isLoading}
          className="w-full p-4 rounded-md bg-zinc-800 text-white placeholder:text-zinc-400 focus:ring-2 focus:ring-violet-500"
        />
      </div>

      <div className="relative">
        <input
          id="password"
          type="password"
          name="password"
          placeholder="Enter your password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          disabled={isLoading}
          className="w-full p-4 pr-20 rounded-md bg-zinc-800 text-white placeholder:text-zinc-400 focus:ring-2 focus:ring-violet-500"
        />
        <button
          type="button"
          onClick={() => togglePasswordVisibility('password')}
          className="absolute right-4 top-1/2 transform -translate-y-1/2 text-violet-300 hover:text-violet-400"
        >
          Show
        </button>
      </div>

      <div className="text-right">
        <a 
          href="#" 
          className="text-violet-400 hover:underline text-sm"
           onClick={(e) => {
            e.preventDefault();
            if (typeof onForgotPassword === 'function') onForgotPassword();
          }}>
            Forgot Password?
        </a>
      </div>

      <button
        type="submit"
        disabled={isLoading}
        className="w-full bg-violet-600 hover:bg-violet-700 transition p-4 rounded-md text-white font-semibold"
      >
        {isLoading ? 'Logging in...' : 'Login'}
      </button>

      {showToast && <Toast message={toastMessage} type={toastType} />}
    </form>
  );
}
