import React, { useState } from "react";
import axios from "axios";

function GradePreview() {
  const [students, setStudents] = useState([]);
  const [file, setFile] = useState(null);
  const [topper, setTopper] = useState(null);
  const [activeDept, setActiveDept] = useState("TECHNICAL");

  const handleUpload = async () => {
    if (!file) return;
    const form = new FormData();
    form.append("file", file);

    try {
      const res = await axios.post(
        "http://localhost:5000/api/upload/trainee",
        form
      );

      const dataWithGrades = res.data.students.map((student) => {
        const dept = student.department?.trim().toUpperCase();
        const gradeCount =
          dept === "PRODUCTION" || dept === "Production" ? 18 : 15;

        return {
          ...student,
          ...Object.fromEntries(
            Array.from({ length: gradeCount }, (_, i) => [`${i + 1}G`, ""])
          ),
        };
      });

      setStudents(dataWithGrades);
    } catch (err) {
      console.error(err);
      alert("Upload failed");
    }
  };

  const handleGradeChange = (index, field, value) => {
    const updated = [...students];
    updated[index][field] = value;
    setStudents(updated);

    const valid = updated.filter((s) => s["1G"]);
    if (valid.length > 0) {
      const top = valid.reduce((a, b) =>
        Number(a["1G"]) > Number(b["1G"]) ? a : b
      );
      setTopper(top);
    }
  };

  // Departments array stays the same
  const departments = ["TECHNICAL", "PRODUCTION", "SUPPORT"];

  const filteredStudents = students.filter((s) => {
    const dept = s.department?.trim().toUpperCase() || "";
    if (activeDept === "TECHNICAL") {
      return dept === "TECHNICAL" || dept === "IT";
    }
    if (activeDept === "PRODUCTION") {
      return dept === "PRODUCTION";
    }
    if (activeDept === "SUPPORT") {
      return dept !== "TECHNICAL" && dept !== "IT" && dept !== "PRODUCTION";
    }
    return false;
  });

  const gradeCount = activeDept === "PRODUCTION" ? 18 : 15;

  return (
    <div className="p-4">
      <h2 className="text-lg font-bold mb-2">Upload Trainee Excel File</h2>
      <input type="file" onChange={(e) => setFile(e.target.files[0])} />
      <button
        className="ml-2 bg-blue-500 text-white px-4 py-1 rounded"
        onClick={handleUpload}
      >
        Upload
      </button>

      {/* Tabs */}
      {students.length > 0 && (
        <div className="mt-4 flex gap-2">
          {departments.map((dept) => (
            <button
              key={dept}
              onClick={() => setActiveDept(dept)}
              className={`px-4 py-2 rounded ${
                activeDept === dept
                  ? "bg-green-600 text-white"
                  : "bg-gray-300 text-black"
              }`}
            >
              {dept}
            </button>
          ))}
        </div>
      )}

      {/* Top Performer */}
      {topper && (
        <div className="mt-4 p-2 bg-yellow-200 rounded shadow">
          🏆 Top Performer:{" "}
          <strong>
            {topper.last_name}, {topper.first_name}
          </strong>{" "}
          – 1G: {topper["1G"]}
          <strong>{topper.over_all}</strong>
        </div>
      )}

      {/* Table */}
      {filteredStudents.length > 0 && (
        <>
          <table className="mt-4 w-full border text-sm">
            <thead>
              <tr className="bg-gray-200">
                <th>Name</th>
                <th>Strand</th>
                <th>Department</th>
                <th>School</th>
                <th>Batch</th>
                <th>Date</th>
                <th>Performance Appraisal</th> {/* over_all */}
                {[...Array(gradeCount)].map((_, idx) => (
                  <th key={idx}>{`${idx + 1}G`}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredStudents.map((s, i) => (
                <tr key={i} className="text-center">
                  <td>
                    {s.last_name}, {s.first_name} {s.middle_name}
                  </td>
                  <td>{s.strand}</td>
                  <td>{s.department}</td>
                  <td>{s.school}</td>
                  <td>{s.batch}</td>
                  <td>{s.date_of_immersion}</td>

                  {/* Performance Appraisal input */}
                  <td>
                    <input
                      type="number"
                      value={s.over_all || ""}
                      onChange={(e) =>
                        handleGradeChange(
                          students.indexOf(s),
                          "over_all",
                          e.target.value
                        )
                      }
                      className="w-14 text-center bg-yellow-100"
                    />
                  </td>

                  {[...Array(gradeCount)].map((_, idx) => (
                    <td key={idx}>
                      <input
                        type="number"
                        id={`${idx + 1}G`}
                        value={s[`${idx + 1}G`] || ""}
                        onChange={(e) =>
                          handleGradeChange(
                            students.indexOf(s),
                            `${idx + 1}G`,
                            e.target.value
                          )
                        }
                        className="w-14 text-center bg-blue-100"
                      />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>

          {/* Generate Excel Button */}
          <button
            className="mt-4 bg-green-600 text-white px-4 py-2 rounded"
            onClick={async () => {
              try {
                const res = await axios.post(
                  "http://localhost:5000/api/generate/excel",
                  { students },
                  { responseType: "blob" }
                );

                const blob = new Blob([res.data], {
                  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = "IMMERSION-GENERATED.xlsx";
                a.click();
                window.URL.revokeObjectURL(url);
              } catch (err) {
                console.error(err);
                alert("Upload failed");
              }
            }}
          >
            Generate Excel Report
          </button>
        </>
      )}
    </div>
  );
}

export default GradePreview;