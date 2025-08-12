// src/context/AuthContext.jsx
import React, { createContext, useContext, useState, useEffect } from 'react';
import { sendLoginRequest } from '../utils/api';

const AuthContext = createContext();

export function AuthProvider({ children }) {
  const [user, setUser] = useState(null);

  // Load user from storage on mount
  useEffect(() => {
    const stored = localStorage.getItem('authUser');
    if (stored) setUser(JSON.parse(stored));
  }, []);

  const login = async (username, password) => {
    try {
      const result = await sendLoginRequest(username, password);
      if (result.success) {
        setUser(result.data.user); // Set user with { id, username } from backend
        localStorage.setItem('authUser', JSON.stringify(result.data.user));
      }
      return result;
    } catch (error) {
      console.error('Login error:', error.message);
      return { success: false, message: error.message };
    }
  };

  const logout = () => {
    setUser(null);
    localStorage.removeItem('authUser');
  };

  const isAuthenticated = !!user;

  return (
    <AuthContext.Provider value={{ user, login, logout, isAuthenticated }}>
      {children}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  return useContext(AuthContext);
}

export const togglePasswordVisibility = (inputId) => {
  const input = document.getElementById(inputId);
  if (input) {
    input.type = input.type === 'password' ? 'text' : 'password';
    const button = input.nextElementSibling;
    if (button) {
      button.textContent = input.type === 'password' ? 'Show' : 'Hide';
    }
  }
};
