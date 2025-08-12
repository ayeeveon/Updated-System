import React from "react";
import Dashboard from "../components/Dashboard";
import { useNavigate } from "react-router-dom";

function HomePage() {
  const navigate = useNavigate();

  const logout = () => {
    localStorage.removeItem("token");
    navigate("/login");
  };

  return (
    <div className="relative flex min-h-screen bg-[#1f1f1f] text-white overflow-hidden">
      {/* Background */}

      <Dashboard />

      {/* Main Content */}
      <div className="flex-1 relative px-10 py-8 overflow-y-auto">

        {/* Centered Content */}
        <div className="min-h-screen flex items-center justify-center">
          <div className="bg-[#a0a0a0b7] rounded-xl p-10 shadow-xl text-center flex flex-col items-center w-full max-w-6xl">
            {/* Logo */}
            <img
              src="/CreotecLogo.png"
              alt="Creotec Logo"
              className="w-[350px] h-[120px] object-contain mb-3 mt-2"
            />

            {/* Welcome Text */}
            <p className="text-2xl font-semibold mt-2">
              Welcome to our internal certificate and training records system.
            </p>
            <p className="text-lg text-gray-300 max-w-4xl mt-4">
              This platform allows you to automatically generate certificates for
              OJT and Immersion programs and conveniently create TESDA records by
              uploading an Excel file. It’s a fast and efficient way to manage
              your documents.
            </p>

            {/* Feature Cards */}
            <div className="mt-10 flex flex-wrap justify-center gap-6">
              <div className="bg-[#3e3e3e] hover:scale-[1.03] transition-all p-6 rounded-xl shadow-md w-72">
                <h2 className="text-xl font-semibold text-purple-300 mb-2">
                  Fast Certificate Generator
                </h2>
                <p className="text-sm text-gray-400">
                  Easily generate and download certificates with template-based
                  automation.
                </p>
              </div>

              <div className="bg-[#3e3e3e] hover:scale-[1.03] transition-all p-6 rounded-xl shadow-md w-72">
                <h2 className="text-xl font-semibold text-purple-300 mb-2">
                  TESDA Integration
                </h2>
                <p className="text-sm text-gray-400">
                  Manage TESDA-related records, training documents, and performance
                  tracking.
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default HomePage;
