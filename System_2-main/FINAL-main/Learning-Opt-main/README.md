<p align="center"> 
  <img src="https://i.imgur.com/OTTnnBI.png" width="200" alt="Creotec Logo" />
</p>

<h1 align="center">
  CREOTEC: CERTIFICATE GENERATOR SYSTEM
</h1>

<h3 align="center">CS 331 - CS Internship</h3>
<h5 align="center">Batangas State University - Alangilan, Mid Semester 2025</h5>

<h5 align="center">Members</h5>
<p align="center">
  Abella, Vince Jericho V. <br>
  Domingo, Joy Susette V. <br>
  Montenegro, Ciavel Anby P. <br>
  Tolentino, John Benedict A.
</p>

---

## 📌 About

**Creotec Certificate Generator** is a web-based application designed to streamline the creation of official certificates for OJT, work immersion, internship, and other academic or organizational events.

Built with **React**, **TailwindCSS**, and **Flask**, this system allows admins to upload recipient data (CSV/Excel), choose certificate templates, customize contents, preview, and print certificates — all in one page.

---

## ✨ Features

### 🖥️ Web Application Features
- **Login & Authentication**
  - Secure login with future Firebase integration
- **Data Upload**
  - Supports CSV or Excel uploads for bulk certificate generation
- **Template Selection**
  - Preloaded templates (OJT, Internship, Work Immersion)
- **Live Certificate Preview**
  - Real-time WYSIWYG preview panel
- **Print-Ready Output**
  - Generates printable, downloadable certificates with accurate layout
- **Custom Signatories**
  - Upload names and designations of authorized signers
- **Responsive UI**
  - Modern UI with mobile/tablet support using TailwindCSS

---

# Setup & Run Instructions

## 1. Install dependencies (frontend & backend)

Run this command in the **frontend** directory:

```bash
npm run install-all
```

This will:

- Install all frontend dependencies

- Create and activate a Python virtual environment in the backend folder

- Install backend Python dependencies from requirements.txt

**Note:** The activation command in the script is for Windows (`venv\Scripts\activate`).  
For macOS/Linux, you might need to modify it to `source venv/bin/activate`.

## 2. Start frontend and backend servers concurrently

Run this command in the **frontend** directory:

```bash
npm run start-all
```

This will start both frontend (Vite) and backend (Flask) servers concurrently.

## Notes

- The backend requires Python 3.8+ and the dependencies listed in `backend/requirements.txt`.

- Make sure you run the `npm run install-all` **inside the frontend directory** to properly set up both frontend and backend.

- The Flask backend runs on port 5000 by default; the frontend (Vite) runs on port 5173.

- For production deployment, consider using a production-ready WSGI server (e.g., Gunicorn) instead of Flask’s development server.

- If you encounter permission issues with the virtual environment activation, try running your terminal as Administrator or adjust execution policies (especially on Windows PowerShell).

- To stop both servers started by `npm run start-all`, press `Ctrl + C` in the terminal.

---

## 🧠 OOP Principles Applied

<ul>
  <li><strong>Encapsulation</strong> – Backend services encapsulate logic for certificate generation and PDF rendering via Flask.</li>
  <li><strong>Abstraction</strong> – Common utility functions and hooks abstract repetitive UI behaviors and API calls.</li>
  <li><strong>Inheritance</strong> – Certificate templates inherit a base layout style and override properties for customization.</li>
  <li><strong>Polymorphism</strong> – Template rendering functions support dynamic content substitution based on user input.</li>
</ul>

---

## 🌐 Tech Stack

| Tech            | Usage                         |
|-----------------|-------------------------------|
| React + Vite    | Frontend Framework            |
| TailwindCSS     | UI Styling                    |
| Flask (Python)  | Backend API and PDF generation|
| MySQL           | User & Certificate data       |
| SheetJS         | Excel/CSV parsing             |

---

## 🎯 Purpose

This project enhances administrative productivity by reducing manual certificate creation and ensures uniform, error-free document generation. It was developed with real deployment in mind for use in schools, events, or organizations needing mass certificate production.

---


