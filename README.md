# repair_manager
Repair Manager is a local desktop application built with **SQLite** for managing clients, jobs, and invoices for repair businesses.

This was created using AI, it works fine and is a simple program. Unless I find something not working correctly it will not be updated. If you improve it.... great. Do what you will with it if you find it useful.

## Features
- Manage **clients**, **jobs**, and **invoices**
- Mark invoices as **Paid** or **Outstanding**
- Search and filter clients, jobs, and invoices
- Export and import CSV data
- Generate printable **PDF invoices** (with company logo)
- Preferences panel for theme, font size, and company details
- SQLite database stored locally for easy backup
- Cross-platform: **Linux, Windows, macOS**

---

## Requirements
Install the dependencies using:
```bash
pip install -r requirements.txt

---

python3 repair_manager.py(Unless your environment uses python repair_manager.py

---

Database Schema
Tables:
clients

CREATE TABLE clients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    email TEXT,
    phone TEXT,
    address TEXT
);

jobs

CREATE TABLE jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL,
    created_on TEXT NOT NULL,
    type_id INTEGER NOT NULL,
    status_id INTEGER NOT NULL,
    description TEXT,
    FOREIGN KEY(client_id) REFERENCES clients(id)
);

invoices

CREATE TABLE invoices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    job_id INTEGER NOT NULL,
    issue_date TEXT NOT NULL,
    due_date TEXT,
    status TEXT CHECK(status IN ('Outstanding', 'Paid')) NOT NULL DEFAULT 'Outstanding',
    total REAL NOT NULL,
    FOREIGN KEY(job_id) REFERENCES jobs(id)
);

---

Example SQL Queries

List all outstanding invoices:

SELECT * FROM invoices WHERE status = 'Outstanding';

Mark an invoice as paid:

UPDATE invoices SET status = 'Paid' WHERE id = 1;

Get jobs for a specific client:

SELECT * FROM jobs WHERE client_id = 2;

---

