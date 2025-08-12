import React, { useState, useEffect } from "react";
import Dashboard from "../components/Dashboard";

function History() {
  const [certificates, setCertificates] = useState([]);
  const [tesdaRecords, setTesdaRecords] = useState([]);
  const [filterType, setFilterType] = useState("all");
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState({ file: "", type: "" });


  const downloadFile = async (filename) => {
    try {
      const res = await fetch(`http://localhost:5000/static/generated/${filename}`);
      if (!res.ok) throw new Error("Download failed");
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);

      const filetype = filename.endsWith(".pptx") ? "Certificate" : "TESDA";
      await fetch("http://localhost:5000/api/download-history", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename, filetype }),
      });
    } catch (e) {
      console.error("Download failed:", e);
    }
  };

  const deleteFile = async (filename, filetype) => {
    const isCertificate = filetype !== "TESDA";
    const endpoint = isCertificate
      ? `http://localhost:5000/generate/delete_certificate?filename=${encodeURIComponent(filename)}`
      : `http://localhost:5000/api/delete_excel?filename=${encodeURIComponent(filename)}`;

    try {
      const res = await fetch(endpoint, { method: "DELETE" });
      if (!res.ok) throw new Error("Delete failed");

      if (isCertificate) {
        setCertificates((prev) => prev.filter((f) => f !== filename));
      } else {
        setTesdaRecords((prev) => prev.filter((f) => f !== filename));
      }

      const storedDates = JSON.parse(localStorage.getItem("fileDates") || "{}");
      delete storedDates[filename];
      localStorage.setItem("fileDates", JSON.stringify(storedDates));
    } catch (e) {
      console.error("Delete failed:", e);
    }
  };

  useEffect(() => {
    const fetchData = async () => {
      try {
        const certRes = await fetch("http://localhost:5000/api/certificates");
        const certData = await certRes.json();
        setCertificates(Array.isArray(certData) ? certData : []);

        const tesdaRes = await fetch("http://localhost:5000/api/tesda");
        const tesdaData = await tesdaRes.json();
        setTesdaRecords(Array.isArray(tesdaData) ? tesdaData : []);
      } catch (err) {
        console.error("Error fetching history:", err);
      }
    };
    fetchData();
  }, []);

  useEffect(() => {
    const allFiles = [...certificates, ...tesdaRecords];
    const storedDates = JSON.parse(localStorage.getItem("fileDates") || "{}");
    let updated = false;

    allFiles.forEach((file) => {
      if (!storedDates[file]) {
        storedDates[file] = new Date().toISOString();
        updated = true;
      }
    });

    if (updated) {
      localStorage.setItem("fileDates", JSON.stringify(storedDates));
    }
  }, [certificates, tesdaRecords]);

  const containerStyle = {
    backgroundColor: "#696b6c",
    color: "white",
    margin: "auto",
    padding: "1rem",
    marginLeft: "2px",
    boxShadow: "0 5px 7px #0000004d",
    borderRadius: "0.5rem", 
  };

  const sectionTitleStyle = {
    fontSize: "1.8rem",
    fontWeight: "700",
    marginBottom: "1rem",
    color: "#f3eaff",
    borderBottom: "2px solid #a361ef",
    paddingBottom: "0.5rem",
  };

  const storedDates = JSON.parse(localStorage.getItem("fileDates") || "{}");

  const combined = [
    ...certificates.map((file) => {
      const lower = file.toLowerCase();
      let type = "Certificate";
      if (lower.includes("immersion")) type = "Work Immersion";
      else if (lower.includes("ojt")) type = "OJT Certificate";
      return { file, type, date: storedDates[file] || new Date().toISOString() };
    }),
    ...tesdaRecords.map((file) => ({
      file,
      type: "TESDA",
      date: storedDates[file] || new Date().toISOString(),
    })),
  ];

  const filtered = combined.filter((record) => {
    if (filterType === "all") return true;
    return record.type.toLowerCase().includes(filterType);
  });

  return (
    <Dashboard>
      {/* Background */}
      <div className="fixed inset-0 bg-cover bg-center bg-zinc-800 -z-10"></div>
      <div className="absolute inset-0 bg-black/10 -z-10"></div>

      <div className="flex-1 p-3 mt-10 mr-5 ml-6 overflow-y-auto">
        <div style={containerStyle}>
          <h2 style={sectionTitleStyle}>Generated Files</h2>

          <div className="mb-4 flex gap-4">
            {["all", "tesda", "ojt", "immersion"].map((type) => (
              <button
                key={type}
                onClick={() => setFilterType(type)}
                className={`px-4 py-2 rounded ${
                  filterType === type ? "bg-[#a361ef] text-white" : "bg-gray-200 text-black"
                }`}
              >
                {type === "all"
                  ? "Show All"
                  : type === "tesda"
                  ? "TESDA"
                  : type === "ojt"
                  ? "OJT Certificate"
                  : "Work Immersion"}
              </button>
            ))}
          </div>

          <table className="w-full table-auto border border-collapse border-[#2c2c2c] text-center text-white">
            <thead>
              <tr className="bg-[#4c3a91]">
                <th className="border border-[#4c3a91] px-4 py-2">#</th>
                <th className="border border-[#4c3a91] px-4 py-2">Filename</th>
                <th className="border border-[#4c3a91] px-4 py-2">Type</th>
                <th className="border border-[#4c3a91] px-4 py-2">Date Added</th>
                <th className="border border-[#4c3a91]">Action</th>
              </tr>
            </thead>
            <tbody>
              {filtered.length === 0 ? (
                <tr>
                  <td colSpan={5} className="border-[#2c2c2c] px-4 py-2 text-gray-300">
                    No records found.
                  </td>
                </tr>
              ) : (
                filtered.map((record, i) => (
                  <tr key={i} className={i % 2 === 0 ? "bg-[#2c2c2c]" : "bg-[#3a3a3a]"}>
                    <td className="px-4 py-2">{i + 1}</td>
                    <td className=" px-4 py-2">{record.file}</td>
                    <td className="px-4 py-2">{record.type}</td>
                    <td className=" px-4 py-2">
                      {new Date(record.date).toLocaleString()}
                    </td>
                    <td className="px-4 py-2 space-x-2">
                      <button
                        className="bg-[#a361ef] hover:bg-purple-700 text-white font-semibold py-1 px-3 rounded"
                        onClick={() => downloadFile(record.file)}
                      >
                        Download
                      </button>
                      
                      <button
                        className="bg-red-600 hover:bg-red-700 text-white font-semibold py-1 px-3 rounded"
                        onClick={() => {
                          setDeleteTarget({ file: record.file, type: record.type });
                          setShowDeleteConfirm(true);
                        }}
                      >
                        Delete
                      </button>

                      {showDeleteConfirm && (
                        <div className="fixed inset-0 bg-black/5 flex justify-center items-center z-50">
                          <div className="bg-white text-black p-6 rounded-md max-w-sm w-full">
                            <h2 className="text-lg font-semibold mb-4 text-center">Confirm Delete</h2>
                            <p className="mb-6">Are you sure you want to delete this file?</p>
                            <div className="flex justify-end gap-4">
                              <button
                                className="px-4 py-2 rounded-md border border-gray-400"
                                onClick={() => setShowDeleteConfirm(false)}
                              >
                                Cancel
                              </button>
                              <button
                                className="px-4 py-2 rounded-md bg-red-600 text-white hover:bg-red-700"
                                onClick={() => {
                                  deleteFile(deleteTarget.file, deleteTarget.type);
                                  setShowDeleteConfirm(false);
                                }}
                              >
                                Delete
                              </button>
                            </div>
                          </div>
                        </div>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </Dashboard>
  );
}

export default History;
