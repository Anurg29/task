# 🏢 TaxLedger QC Pro 

**An AI-Powered Property Tax Reconciliation & Quality Control System**

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![React](https://img.shields.io/badge/Frontend-React.js-61DAFB?logo=react&logoColor=black)
![FastAPI](https://img.shields.io/badge/Backend-FastAPI-009688?logo=fastapi&logoColor=white)
![Python](https://img.shields.io/badge/Processing-Python_3.10-3776AB?logo=python&logoColor=white)

---

## 📖 About the Project
**TaxLedger QC Pro** is an intelligent, high-performance web application designed to automatically cross-verify and reconcile master property tax records (Excel) against raw municipal tax ledger reports (PDFs). It instantly audits up to 23 distinct tax heads and property details, detecting missing properties, numerical discrepancies, and typographical errors in regional languages (Marathi) in real-time.

## ✨ Key Features
- **Intelligent Fuzzy Matching:** Employs advanced Levenshtein distance algorithms to handle minor spelling mistakes in Marathi text (e.g., `शासकीय` vs `शमसककय`).
- **AI Self-Learning (Typo Dictionary):** Users can 'Approve' minor typos directly from the UI. The system remembers these corrections globally for all future audits.
- **Regex Normalization:** Automatically standardizes complex Property IDs by stripping dashes, slashes, and spaces (e.g., `75-1052` perfectly matches `75/1052`).
- **High-Performance Processing:** Utilizes Python's `ProcessPoolExecutor` (Multiprocessing) to divide large PDF reports across multiple CPU cores, increasing processing speed by up to 5x.
- **Live Progress Tracking:** Integrated WebSockets provide users with a real-time progress bar while analyzing massive 500+ page PDFs.
- **Dynamic Excel Exports:** Exports the final audited results into a beautifully formatted, color-coded Excel file using `ExcelJS`, highlighting missing fields in red and matches in green.

---

## 🏗️ Architecture

The application is built using a modern decoupled Microservice-style architecture for maximum scalability and performance.

### 1. Frontend (User Interface)
- **Framework:** React.js powered by Vite for blazing-fast Hot Module Replacement (HMR) and optimized production builds.
- **Design System:** Custom CSS implementing a "Glassmorphism" aesthetic with smooth transitions and micro-animations. No heavy CSS frameworks were used, ensuring lightning-fast load times.
- **Deployment:** Hosted seamlessly on **Netlify**.

### 2. Backend (API & Core Logic)
- **Framework:** **FastAPI** (Python), chosen for its exceptional concurrency handling and ASGI support.
- **Data Extraction:** `pdfplumber` is utilized for robust spatial text extraction from unstructured PDF documents.
- **Data Manipulation:** `pandas` is used to load, query, and join the Master Excel database rapidly.
- **Deployment:** Hosted on **Render**, acting as the computational powerhouse.

### 3. Communication & Storage
- **REST APIs:** For robust, stateless data transfers (e.g., file uploads).
- **WebSockets:** For bidirectional, real-time communication during heavy data processing.
- **Local JSON Data-Store:** Persists the AI Typo Dictionary for ongoing system intelligence.

---

## 🚀 How to Run Locally

### Backend Setup
```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
```
The API will start at `http://localhost:8000`.

### Frontend Setup
```bash
cd frontend
npm install
npm run dev
```
The UI will start at `http://localhost:5173`.

---

## 👨‍💻 Developed By
**Anurag Dinesh Rokade**  
*Full Stack Developer & AI Enthusiast*  
[GitHub Profile](https://github.com/Anurg29)

---
*Built with ❤️ for perfectly accurate property tax systems.*
