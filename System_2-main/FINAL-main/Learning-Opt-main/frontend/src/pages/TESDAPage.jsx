// src/pages/TESDAPage.jsx
import React, { useState } from "react";
import { useAuth } from "../utils/auth";
import { useNavigate } from "react-router-dom";

import TemplateSelection from "../components/TESDASelection";
import CertificateDetails from "../components/DetailsForm";
import Signatories from "../components/SignatoriesForm";
import ExcelUploader from "../components/ExcelUploader";
import DownloadLinks from "../components/DownloadLinks";
import Dashboard from "../components/Dashboard";

function TESDAPage() {
  const { logout } = useAuth();
  const navigate = useNavigate();

  // --- UI state (unchanged) ---
  const [selectedTemplate, setSelectedTemplate] = useState("tesda");
  const [excelData, setExcelData] = useState([]);
  const [previewLoading, setPreviewLoading] = useState(false);
  const [generatedFiles, setGeneratedFiles] = useState([]);

  // --- NEW: capture the raw uploaded File for FormData ---
  const [selectedFile, setSelectedFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const [errorMsg, setErrorMsg] = useState("");

  // Keep your "Detected Columns" preview
  const handleExcelParsed = (rows) => {
    setExcelData(rows);
  };

  const detectedColumns = excelData.length > 0 ? Object.keys(excelData[0]) : [];

  // === Reference logic: send FormData(file [+ optional mapping]) to Flask ===
  const handleGenerate = async () => {
    setErrorMsg("");

    if (!selectedFile) {
      alert("No data file selected. Please upload an Excel file first.");
      return;
    }

    try {
      setBusy(true);

      const formData = new FormData();
      formData.append("file", selectedFile);

      // If you add a mapping UI later, append it like this:
      // if (mappingJson.trim()) formData.append("mapping", mappingJson);

      // Call Flask directly; avoid proxy confusion
      const res = await fetch("http://127.0.0.1:5000/api/generate", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const msg = await res.json().catch(() => ({}));
        throw new Error(msg.error || `Generate failed (${res.status})`);
      }

      // Download resulting workbook (same pattern as your reference)
      const blob = await res.blob();
      const cd = res.headers.get("Content-Disposition") || "";
      const match = cd.match(/filename="?([^"]+)"?/i);
      const filename = match?.[1] || `filled_${Date.now()}.xlsx`;

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      // Keep your DownloadLinks section; no visual change
      setGeneratedFiles((prev) => prev);
    } catch (err) {
      console.error("Error generating workbook:", err);
      setErrorMsg(err.message || "Generation failed");
      alert(err.message || "Generation failed");
    } finally {
      setBusy(false);
      setPreviewLoading(false);
    }
  };

  // (Optional) Your existing preview still works with parsed rows
  const handlePreview = async () => {
    if (excelData.length === 0) {
      alert("No data loaded. Please upload an Excel file first.");
      return;
    }
    try {
      const response = await fetch("http://localhost:5000/generate/preview", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ template: selectedTemplate, rows: excelData }),
      });
      const html = await response.text();
      const previewWindow = window.open("", "_blank");
      previewWindow.document.write(html);
    } catch (err) {
      console.error("Error previewing certificate:", err);
    }
  };

  return (
    <div className="relative flex bg-zinc-800 min-h-screen text-white overflow-hidden">
      {/* Background */}
      <div
        className="fixed inset-0 bg-cover bg-center blur-lg -z-10"
      ></div>

      {/* Optional dark overlay */}
      <div className="absolute inset-0 bg-black/10 -z-10"></div>
      <Dashboard />

      {/* Main Content on the right */}
      <div className="flex-1 p-3 mt-13 mr-9 overflow-y-auto">

        {/* Excel Upload */}
          <div className="w-full max-w-[1800px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
          <h3 className="section-header">Upload Data File</h3>

          {/* IMPORTANT: pass onFileSelected so we get the raw File */}
          <ExcelUploader
            onDataParsed={handleExcelParsed}
            onFileSelected={setSelectedFile}
          />

          {detectedColumns.length > 0 && (
            <div className="mt-4 bg-violet-950 p-4 rounded-md text-white">
              <h4 className="font-semibold mb-2">Detected Columns:</h4>
              <div className="flex flex-wrap gap-2">
                {detectedColumns.map((col, idx) => (
                  <div
                    key={idx}
                    className="bg-violet-600 text-white text-sm px-3 py-1 rounded-full shadow-sm max-w-max"
                  >
                    {col}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Template Selection */}
          <div className="w-full max-w-[1600px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
          <h3 className="section-header">Select a Template</h3>
          <TemplateSelection
            selected={selectedTemplate}
            onChange={setSelectedTemplate}
          />
        </div>

        {/* Custom Template Details */}
        {selectedTemplate === "custom" && (
          <div className="w-full max-w-[1600px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
            <h3 className="section-header">Certificate Customization</h3>
            <CertificateDetails />
            <Signatories />
          </div>
        )}

        {/* Generate Buttons */}
          <div className="w-full max-w-[1600px] bg-zinc-500 p-2 rounded-md mx-auto mb-8">
          <button
            className="rounded-md bg-[#3737ff7e] m-3 p-2"
            onClick={handlePreview}
            disabled={previewLoading}
          >
            {previewLoading ? "Loading..." : "Preview"}
          </button>

          <button
            className="rounded-md bg-[#a361ef] p-2"
            id="generate-btn"
            onClick={handleGenerate}
            disabled={busy}
          >
            {busy ? "Generating..." : "Generate Certificates"}
          </button>
        </div>

        {/* Download Links */}
        {generatedFiles.length > 0 && <DownloadLinks files={generatedFiles} />}

        {/* Optional inline error (keeps design minimal) */}
        {errorMsg && (
          <p className="mt-2 text-red-400 text-sm">{errorMsg}</p>
        )}
      </div>
    </div>
  );
}

export default TESDAPage;
