import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";
import Dashboard from "../components/Dashboard";
import { useNavigate } from "react-router-dom";

function Immersion() {
  const [data, setData] = useState(() => {
    const stored = localStorage.getItem("immersionData");
    return stored ? JSON.parse(stored) : [];
  });

  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);
  const [deleteTargetIndex, setDeleteTargetIndex] = useState(null);

  const navigate = useNavigate();

  useEffect(() => {
    const dataToStore = data.map(({ downloadUrl, ...rest }) => rest);
    localStorage.setItem("immersionData", JSON.stringify(dataToStore));
  }, [data]);

 const handleFileUpload = (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const workbook = XLSX.read(e.target.result, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    if (jsonData.length === 0) return;

    const cleanedData = jsonData.map((row) => ({
      "Last Name": row["LAST NAME"] || "",
      "First Name": row["FIRST NAME"] || "",
      "Middle Name": row["MIDDLE NAME"] || "",
      Strand: row["STRAND"] || "",
      Department: row["DEPARTMENT"] || "",
    }));

    const newEntry = {
      fileName: file.name,
      date: new Date().toLocaleString(),
      content: cleanedData,
    };

    // Prepend newEntry to show newest first
    setData((prevData) => [newEntry, ...prevData]);
  };
  reader.readAsBinaryString(file);
};



  const handleDownload = (rowData, index) => {
    const formData = new FormData();

    const worksheet = XLSX.utils.json_to_sheet(rowData.content);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    const excelBlob = new Blob(
      [XLSX.write(workbook, { bookType: "xlsx", type: "array" })],
      {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }
    );

    formData.append("file", excelBlob, rowData.fileName);

    fetch("http://localhost:5000/fill-template", {
      method: "POST",
      body: formData,
    })
      .then((res) => {
        if (!res.ok) throw new Error("Failed to generate filled template");
        return res.blob();
      })
      .then((blob) => {
        const url = window.URL.createObjectURL(blob);

        setData((prevData) =>
          prevData.map((item, i) =>
            i === index ? { ...item, downloadUrl: url } : item
          )
        );

        const a = document.createElement("a");
        a.href = url;
        a.download = "TESDA_Filled_Template.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
      })
      .catch((err) => {
        console.error("Download error:", err);
        alert("Something went wrong while generating the filled template.");
      });
  };

  const confirmDelete = (index) => {
    setDeleteTargetIndex(index);
    setShowDeleteConfirm(true);
  };

  const deleteRow = () => {
    if (deleteTargetIndex !== null) {
      const newData = data.filter((_, idx) => idx !== deleteTargetIndex);
      setData(newData);
      setDeleteTargetIndex(null);
      setShowDeleteConfirm(false);
    }
  };

  const handleViewDetails = (index) => {
    const fileData = data[index];
    // Navigate to immersion records with filename in URL and state
    navigate(`/immersion/records/${encodeURIComponent(fileData.fileName)}`, {
      state: { row: fileData },
    });
  };

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

  return (
    <Dashboard>
      <div className="fixed inset-0 bg-cover bg-center bg-zinc-800 -z-10"></div>
      <div className="absolute inset-0 bg-black/10 -z-10"></div>

      <div className="flex-1 p-3 mt-10 mr-5 ml-6 overflow-y-auto">
        <div style={containerStyle}>
          <h2 style={sectionTitleStyle}>Work Immersion Records</h2>

          <div className="mb-4">
            <label className="bg-[#a361ef] hover:bg-purple-700 text-white font-semibold py-2 px-4 rounded cursor-pointer inline-block">
              Upload Excel File
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleFileUpload}
                className="hidden"
              />
            </label>
          </div>

          <table className="w-full table-auto border border-collapse border-[#2c2c2c] text-center text-white">
            <thead>
              <tr className="bg-[#4c3a91]">
                <th className="border border-[#4c3a91] px-4 py-2 w-1/15">#</th>
                <th className="border border-[#4c3a91] px-4 py-2 w-4/15">Filename</th>
                <th className="border border-[#4c3a91] px-4 py-2 w-2/15">Date</th>
                <th className="border border-[#4c3a91] px-4 py-2 w-2/15">Action</th>
              </tr>
            </thead>
            <tbody>
              {data.length === 0 ? (
                <tr className="bg-[#2c2c2c]">
                  <td colSpan={4} className="px-4 py-4 text-gray-400">
                    No Excel file uploaded yet.
                  </td>
                </tr>
              ) : (
                data.map((row, index) => (
                  <tr
                    key={index}
                    className={index % 2 === 0 ? "bg-[#2c2c2c]" : "bg-[#3a3a3a]"}
                  >
                    <td className="px-4 py-2">{index + 1}</td>
                    <td
                      className="px-4 py-2 text-blue-400 underline cursor-pointer hover:text-blue-300"
                      onClick={() => handleViewDetails(index)}
                    >
                      {row.fileName}
                    </td>
                    <td className="px-4 py-2">{row.date}</td>
                    <td className="px-4 py-2 space-x-2">
                      <button
                        onClick={() => handleDownload(row, index)}
                        className="bg-[#a361ef] hover:bg-purple-700 text-white font-semibold py-1 px-3 rounded"
                      >
                        Download
                      </button>
                      {row.downloadUrl && (
                        <button
                          onClick={() => {
                            const a = document.createElement("a");
                            a.href = row.downloadUrl;
                            a.download = "TESDA_Filled_Template.xlsx";
                            document.body.appendChild(a);
                            a.click();
                            a.remove();
                          }}
                          className="bg-green-600 hover:bg-green-700 text-white font-semibold py-1 px-3 rounded"
                        >
                          Re-download
                        </button>
                      )}
                      <button
                        onClick={() => confirmDelete(index)}
                        className="bg-red-600 hover:bg-red-700 text-white font-semibold py-1 px-3 rounded"
                      >
                        Delete
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {showDeleteConfirm && (
          <div className="fixed inset-0 bg-black/5 flex justify-center items-center z-50">
            <div className="bg-white text-black p-6 rounded-md max-w-sm w-full">
              <h2 className="text-lg font-semibold mb-4 text-center">Confirm Delete</h2>
              <p className="mb-6">Are you sure you want to delete this row?</p>
              <div className="flex justify-end gap-4">
                <button
                  className="px-4 py-2 rounded-md border border-gray-400"
                  onClick={() => setShowDeleteConfirm(false)}
                >
                  Cancel
                </button>
                <button
                  className="px-4 py-2 rounded-md bg-red-600 text-white"
                  onClick={deleteRow}
                >
                  Delete
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </Dashboard>
  );
}

export default Immersion;
