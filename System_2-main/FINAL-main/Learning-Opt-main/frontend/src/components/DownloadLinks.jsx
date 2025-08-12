import React, { useState, useEffect } from 'react';
import axios from 'axios';

export default function DownloadLinks({ files, setFiles }) {
  const [fileNames, setFileNames] = useState(
    files.reduce((acc, file, idx) => ({ ...acc, [idx]: file }), {})
  );

  const handleFileNameChange = (idx, newName) => {
    setFileNames(prev => ({ ...prev, [idx]: newName }));
  };

  const handleDownloadAll = async () => {
    for (let idx = 0; idx < files.length; idx++) {
      await downloadFile(files[idx], idx);
      await new Promise(resolve => setTimeout(resolve, 500)); // Add 500ms delay between downloads
    }
  };

  const fetchCertificates = async () => {
    try {
      const res = await axios.get("http://localhost:5000/api/certificates");
      if (setFiles) {
        setFiles(res.data);
      }
    } catch (err) {
      console.error("Failed to fetch certificates", err);
    }
  };

  const downloadFile = async (file, idx) => {
    try {
      const fileUrl = `http://localhost:5000/static/generated/${file}?t=${Date.now()}`;
      const a = document.createElement('a');
      a.href = fileUrl;

      // Ensure filename ends with .pptx
      const baseName = fileNames[idx] || file;
      a.download = baseName.endsWith('.pptx') ? baseName : `${baseName}.pptx`;

      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      await fetchCertificates(); // Optional refresh
    } catch (e) {
      console.error('Download failed', e);
    }
  };

  // Optional helper (not used but kept for dev convenience)
  const getZipUrl = () => {
    const qs = files.map(f => `files=${encodeURIComponent(f)}`).join('&');
    return `http://localhost:5000/generate/zip?${qs}`;
  };

  return (
      <div className="w-full max-w-[1800px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
      <h3 className="text-white text-3xl font-bold mb-3 ml-0">Download Certificates</h3>

      <div className="mb-2">
        <button
          type="button"
          onClick={handleDownloadAll}
          className="bg-green-600 text-white px-4 py-2 rounded-md hover:bg-green-700"
        >
          Download All
        </button>
      </div>

      <div className="flex flex-col gap-1">
        {files.map((file, idx) => (
          <div
            key={idx}
            className="certDetails flex gap-2 bg-[#ffffff1a] p-2 mt-2 items-center rounded-md"
          >
            <input
              type="text"
              value={fileNames[idx] || file}
              onChange={(e) => handleFileNameChange(idx, e.target.value)}
              className="flex-1 mainInput h-1rem rounded-md p-1 bg-[#ffffff1a] text-white border border-[#ffffff3a] focus:border-[#a361ef]"
              placeholder="Enter file name"
            />
            <button
              type="button"
              onClick={() => downloadFile(file, idx)}
              className="text-xs font-semibold text-[#613d89] bg-[#a361ef90] px-3 py-1 rounded-md hover:bg-[#a361ef]"
            >
              Download
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}
