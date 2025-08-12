import React from 'react';
import { NavLink } from 'react-router-dom';

const Sidebar = ({ onLogoutClick }) => {
  return (
    <aside className="fixed top-0 left-0 w-64 h-screen bg-gradient-to-br from-[#6b55c7] to-[#4c3a91] text-white p-10 z-10">
      <h2 className="text-2xl font-bold tracking-wide mb-10">DASHBOARD</h2>
      <nav className="flex flex-col space-y-4">

        <NavLink 
          to="/home" 
          className="flex items-center self-center gap-2 px-4 py-1 rounded border-2 w-58 border-purple-950 hover:bg-purple-950 transition-colors"
        >
          <img src="/Home.svg" alt="Home Icon" className="w-10 h-10" />
          <span>HOME</span>
        </NavLink>

        <NavLink 
          to="/history" 
          className="flex items-center self-center gap-2 px-4 py-1 rounded border-2 w-58 border-purple-950 hover:bg-purple-950 transition-colors"
        >
          <img src="/History.svg" alt="History Icon" className="w-10 h-10" />
          <span>History</span>
        </NavLink>

        <NavLink 
          to="/generate" 
          className="flex items-center self-center gap-2 px-4 py-1 rounded border-2 w-58 border-purple-950 hover:bg-purple-950 transition-colors"
        >
          <img src="/Certification.svg" alt="Certificate Generator Icon" className="w-10 h-10" />
          <span>Certificate Generator</span>
        </NavLink>

        <NavLink 
          to="/tesda" 
          className="flex items-center self-center gap-2 px-4 py-1 rounded border-2 w-58 border-purple-950 hover:bg-purple-950 transition-colors"
        >
          <img src="/TESDA.svg" alt="TESDA Logo" className="w-10 h-10" />
          <span>TESDA</span>
        </NavLink>
      </nav>

      {/* Logout Button */}
      <div className="absolute bottom-6 left-0 w-full px-10">
        <button
          onClick={onLogoutClick}
          className="w-full bg-red-600 hover:bg-red-700 text-white py-2 rounded text-center"
        >
          Logout
        </button>
      </div>
    </aside>
  );
};

export default Sidebar;