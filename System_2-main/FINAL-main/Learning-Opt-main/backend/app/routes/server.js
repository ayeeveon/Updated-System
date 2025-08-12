const express = require("express");
const cors = require("cors");
const mysql = require("mysql2");

const app = express();
const port = 3000;

app.use(cors());
app.use(express.json());

// MySQL connection
const db = mysql.createConnection({
  host: "localhost",
  user: "root",     // ⬅️ replace with your actual MySQL username
  password: "GiharuAlfonso1031",   // ⬅️ replace with your actual MySQL password
  database: "creo_certificate",   // ⬅️ replace with your actual database name
});

// Check DB connection
db.connect((err) => {
  if (err) {
    console.error("MySQL connection error:", err.message);
    return;
  }
  console.log("✅ Connected to MySQL database");
});

// API: Get recent certificates
app.get("/api/certificates", (req, res) => {
  db.query(
    "SELECT title FROM certificates ORDER BY created_at DESC LIMIT 20",
    (err, results) => {
      if (err) return res.status(500).json({ error: err.message });
      const titles = results.map((row) => row.title);
      res.json(titles);
    }
  );
});

// API: Get recent TESDA records
app.get("/api/tesda", (req, res) => {
  db.query(
    "SELECT title FROM tesda_records ORDER BY created_at DESC LIMIT 20",
    (err, results) => {
      if (err) return res.status(500).json({ error: err.message });
      const titles = results.map((row) => row.title);
      res.json(titles);
    }
  );
});

// Start server
app.listen(port, () => {
  console.log(`🚀 Server running at http://localhost:${port}`);
});
