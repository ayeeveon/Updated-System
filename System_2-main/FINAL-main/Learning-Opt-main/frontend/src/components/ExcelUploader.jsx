// src/components/ExcelUploader.jsx
import React, { useCallback, useState } from 'react';
import * as XLSX from 'xlsx';

export default function ExcelUploader({ onDataParsed, onFileSelected }) {
  const [fileName, setFileName] = useState('');

  const handleFile = useCallback((file) => {
    if (!file) return;
    setFileName(file.name);

    // Send the raw File up to the parent page
    try {
      if (typeof onFileSelected === 'function') onFileSelected(file);
    } catch {}

    // Keep Detected Columns working
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (typeof onDataParsed === 'function') onDataParsed(json);
    };

    reader.readAsArrayBuffer(file);
  }, [onDataParsed, onFileSelected]);

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();

    const file = e.dataTransfer.files[0];
    if (file && file.name.match(/\.(xlsx|xls|csv)$/)) {
      handleFile(file);
    } else {
      alert('Invalid file format. Please upload .xlsx, .xls, or .csv.');
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleChange = (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
  };

  return (
    <label
      htmlFor="excel-upload"
      onDrop={handleDrop}
      onDragOver={handleDragOver}
      className="upload-excel cursor-pointer border-dashed border-2 rounded-md p-6 text-center text-white bg-zinc-700 transition block"
    >
      <p className="mb-2">Drop your Excel file here or click to upload</p>
      <p className="text-sm text-zinc-400">Supported formats: .xlsx, .xls, .csv</p>
      <span className="text-violet-400 underline ml-1">Browse Files</span>

      <input
        id="excel-upload"
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={handleChange}
        className="hidden"
      />

      {fileName && (
        <div className="mt-2 text-sm text-white rounded px-3 py-2 ">
          Uploaded File: <span className="font-semibold">{fileName}</span>
        </div>
      )}
    </label>
  );
}
