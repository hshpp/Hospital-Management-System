<div align="center">

# 🏥 Hospital Management System (HMS)
### Modern, All-in-One Clinic & Hospital Management Desktop Application

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?style=for-the-badge&logo=python)](https://python.org)
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green?style=for-the-badge)](https://docs.python.org/3/library/tkinter.html)
[![SQLite](https://img.shields.io/badge/Database-SQLite-lightgrey?style=for-the-badge&logo=sqlite)](https://sqlite.org)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

**A comprehensive desktop application designed to automate and streamline all operations of a clinic or hospital.**

[Key Features](#-key-features) • [Tech Stack](#-tech-stack) • [Installation](#-installation--usage) • [Screenshots](#-screenshots) • [Developer](#-developer)

</div>

---

## 📖 About The Project

This application was built to solve the fragmentation of clinic management tools. It provides a centralized, secure, and intuitive platform for managing patients, doctors, appointments, medical records, laboratory tests, pharmacy inventory, and financial reports. 

Whether you are running a small private practice or a multi-specialty clinic, this software covers the entire patient lifecycle from registration to final billing.

---

## ✨ Key Features

### 🔐 Secure Authentication & Access Control
*   Role-based access control (Admin, Doctor, Receptionist, Pharmacist).
*   SHA-256 password hashing.
*   Account recovery via secure recovery codes.

### 👥 Patient & Doctor Management
*   Complete CRUD operations for Patients and Doctors.
*   Patient Timeline feature (visualizes full medical history).
*   Advanced search and filtering.

### 📅 Appointment System
*   Schedule, update, and track appointments.
*   Real-time alerts for today's and tomorrow's appointments.
*   Conflict prevention (prevents double-booking a doctor).

### 📋 Electronic Health Records (EHR)
*   Record vital signs, diagnoses, and symptoms.
*   **Smart Prescription Writer:** Auto-fills dosage and duration based on the selected medicine.
*   Automated PDF generation for printed prescriptions.

### 🧪 Laboratory Module
*   Manage both **Direct** (walk-in) and **Doctor-Requested** lab tests.
*   Enter lab results directly into the patient's EHR.
*   Dedicated financial tracking for lab services.

### 💊 Pharmacy & Inventory Control
*   Real-time stock management.
*   **Direct Sales** (POS-style) and **Prescription Dispensing**.
*   Automated low-stock alerts.
*   Detailed revenue tracking for both sales channels.

### 💰 Advanced Financial & Billing System
*   Automated invoice generation for Visits, Pharmacy, and Lab.
*   **Dynamic Reports:** View revenue for Daily, Weekly, Monthly, and Yearly periods.
*   **Excel Export with Native Charts:** Generates highly formatted `.xlsx` files with embedded bar and pie charts.
*   PDF Receipt printing with full Persian language support.

### 💾 System Administration
*   **Automated Backups:** Daily database snapshots with an option for secondary (external) backup locations.
*   **Audit Log:** Tracks every action (Create, Update, Delete, Login) performed by users.
*   **Modern UI:** Fully responsive design with customizable printer settings and a Dark Mode option.

---

## 🛠️ Tech Stack

*   **Core:** Python 3
*   **GUI:** Tkinter (Customized with Canvas for a modern look)
*   **Database:** SQLite3
*   **Data Visualization:** Matplotlib
*   **Excel Engine:** OpenPyXL
*   **PDF Generation:** FPDF2, Arabic Reshaper (for Persian text), Python-Bidi

---

## 🚀 Installation & Usage

> ⚠️ **Note:** For non-developers, a standalone `.exe` version is available in the *Releases* section. No Python installation is required!

### Run from Source (For Developers)
1. **Clone the repository:**
   ```bash
   git clone https://github.com/hshpp/Hospital-Management-System.git
   cd Hospital-Management-System
   
Set up a virtual environment (Recommended):

python -m venv venv
source venv/bin/activate  # On Windows use: venv\Scripts\activate

Install dependencies:

pip install -r requirements.txt

Initialize the Database:

python setup_database.py

Run the application:

python hospital_gui_full.py

🔑 Default Credentials
Username: admin

Password: admin123

📸 Screenshots
Dashboard	Pharmacy
https://via.placeholder.com/600x400.png?text=Dashboard+Screenshot	https://via.placeholder.com/600x400.png?text=Pharmacy+Screenshot
Medical Record	Reports
https://via.placeholder.com/600x400.png?text=Medical+Record+Screenshot	https://via.placeholder.com/600x400.png?text=Reports+Screenshot

👨‍💻 Developer
Hossein Shahabi

LinkedIn: [https://www.linkedin.com/in/%D8%AD%D8%B3%DB%8C%D9%86-%D8%B4%D9%87%D8%A7%D8%A8%DB%8C-90a49844b/?lipi=urn%3Ali%3Apage%3Ad_flagship3_profile_view_base_contact_details%3BTNQkoSWnTdqExHUuHqAMLg%3D%3D]

Email : spacex.nasa.starlink@gmail.com

GitHub: @hshpp

