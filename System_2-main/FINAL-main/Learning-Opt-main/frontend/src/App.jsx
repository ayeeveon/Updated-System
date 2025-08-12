import React from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom';
import LoginPage from './pages/LoginPage';
import CertificateGenerator from './pages/CertificateGenerator';
import { AuthProvider, useAuth } from './utils/auth';
import TESDAPage from './pages/TESDAPage';
import Home from './pages/Home';
import History from './pages/History';  
import Immersion from './pages/Immersion';
import ImmersionRecords from './pages/ImmersionRecords';
import ImmersionTechnical from './pages/ImmersionTechnical'; // import technical page
import ImmersionProduction from './pages/ImmersionProduction'; // import production page

function PrivateRoute({ children }) {
  const { isAuthenticated } = useAuth();
  return isAuthenticated ? children : <Navigate to="/login" replace />;
}

const App = () => (
  <AuthProvider>
    <Router>
      <Routes>
        <Route path="/" element={<Navigate to="/login" replace />} />
        <Route path="/login" element={<LoginPage />} />
        <Route path="/home" element={<PrivateRoute><Home /></PrivateRoute>} />
        <Route path="/history" element={<PrivateRoute><History /></PrivateRoute>} />
        <Route path="/generate" element={<PrivateRoute><CertificateGenerator /></PrivateRoute>} />
        <Route path="/tesda" element={<PrivateRoute><TESDAPage /></PrivateRoute>} />
        <Route path="/immersion" element={<PrivateRoute><Immersion /></PrivateRoute>} />
        <Route path="/immersion/records/:filename" element={<PrivateRoute><ImmersionRecords /></PrivateRoute>} />
        <Route path="/immersion/production/:filename" element={<PrivateRoute><ImmersionProduction /></PrivateRoute>} />
        <Route path="/immersion/technical/:filename" element={<PrivateRoute><ImmersionTechnical /></PrivateRoute>} />

      </Routes>
    </Router>
  </AuthProvider>
);

export default App;
