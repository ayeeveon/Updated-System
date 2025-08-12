// File: src/pages/LoginPage.jsx
import React, { useState } from 'react';
import LoginForm from '../components/LoginForm';
import ForgotPasswordForm from '../components/ForgotPasswordForm';
import LoadingSpinner from '../components/LoadingSpinner';
import '../index.css';

export default function LoginPage() {
  const [isLoading, setIsLoading] = useState(false);
  const [forgotPasswordActive, setForgotPasswordActive] = useState(false);
  const [showSignUp, setShowSignUp] = useState(false);

  return (
    <div className="relative min-h-screen flex justify-center items-center bg-[#2d2b38] p-5 overflow-hidden">
      {isLoading && <LoadingSpinner overlay size={50} color="white" />}
      <div className="fixed inset-0 z-0 pointer-events-none" id="backgroundAnimation"></div>

      <div className="login-container flex flex-col md:flex-row w-full max-w-[1100px] h-auto md:h-[450px] bg-[#2d2b38] rounded-[20px] overflow-hidden shadow-xl relative z-10">

        {/* Left violet panel with logo */}
        <div className="hidden md:flex md:w-[60%] flex-col justify-center  items-center-safe bg-gradient-to-br from-[#7a61df] to-[#b3a1fc] p-10">
          <img
            src="/CreotecLogo.png"
            alt="Creotec Logo"
            className="w-[500px] h-[300px] object-contain mb-6"
          />
         
        </div>

            
        <div className="w-full md:w-[52%] flex flex-col justify-center items-center bg-[#1a1a1a] p-8 md:p-10">
          <div className="flex flex-col justify-center items-center mb-6 text-center">
            <h1 className="text-[#8770e6] font-bold text-3xl leading-tight">
              Learning Operations <br /> and Linkages
            </h1>
          </div>

          {forgotPasswordActive ? (
            <ForgotPasswordForm onBackToLogin={() => setForgotPasswordActive(false)} />
          ) : (
            <LoginForm
              isLoading={isLoading}
              setIsLoading={setIsLoading}
              onForgotPassword={() => setForgotPasswordActive(true)}
              onSignUp={() => setShowSignUp(true)}
            />
          )}
        </div>


      </div>
    </div>
  );
}