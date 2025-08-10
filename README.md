# CRM Application

A desktop **Customer Relationship Management (CRM)** tool built with **PyQt6**.
It features a dashboard with quick statistics and charts, full record management for invoices and work orders, CSV/JSON import/export, PDF generation, and configurable business details.

---

## **Features Overview**
- **Dashboard**
  - Displays total counts for customers, invoices, and work orders.
  - Shows financial totals for invoices.
  - Displays a summary chart using matplotlib.
  - Auto-refreshes when the app starts.
- **Invoices Tab**
  - Create, view, edit, and delete invoices.
  - Quick "New" and "Delete" buttons.
  - Color-coded statuses.
  - PDF export using stored business info.
- **Work Orders Tab**
  - Full CRUD (Create, Read, Update, Delete).
  - Quick "New" and "Delete" buttons.
  - Status highlighting.
- **Business Info Tab**
  - Store name, address, phone, email, and website.
  - Information is auto-used in PDF exports.
- **Import/Export**
  - Export invoices or work orders to CSV or JSON.
  - Exports are timestamped automatically.
  - Import CSV/JSON with preview before overwrite.
  - Can overwrite existing records silently.
- **Column Management**
  - Auto-resizing columns for better readability.
- **Data Storage**
  - SQLite backend with automatic migrations.
  - New fields and schema updates are handled at startup.

---

## **Installation**

You can install dependencies **either** via your system's package manager (**recommended for Kubuntu 24.04**) **or** via `pip`.

### **Option 1: Install from Kubuntu 24.04 Repositories**
```bash
sudo apt update
sudo apt install python3-pyqt6 python3-pyqt6.qtcharts python3-reportlab python3-pandas python3-matplotlib
```

### **Option 2: Install via pip**
```bash
pip install -r requirements.txt
```

---

## **requirements.txt** (for pip)
```
PyQt6>=6.5
pandas>=2.0
reportlab>=4.0
matplotlib>=3.7
```

---

## **Running the Application**
```bash
python3 crm_app.py
```

---

## **Tab-by-Tab Functionality**

### **Dashboard**
- Displays live counts for:
  - Customers
  - Invoices
  - Work Orders
- Shows total value of all invoices.
- Visual chart for quick trend overview.

### **Invoices**
- Create, edit, delete invoices.
- Quick action buttons for faster workflow.
- Status highlighting for overdue or paid invoices.
- PDF export with business branding.

### **Work Orders**
- Full CRUD capabilities.
- Quick action buttons.
- Color-coded statuses for progress tracking.

### **Business Info**
- Store company name, address, phone, email, and website.
- Data is embedded in generated PDFs.

### **Import/Export**
- Export selected table to CSV or JSON.
- Files automatically timestamped.
- Import from CSV/JSON with preview.
- Supports silent overwriting of existing records.
- Auto-resizes columns after import.

---

## **Data Storage**
- SQLite database with automatic schema migrations.
- Handles adding new columns without manual intervention.

---

## **License**
MIT License – free to use, modify, and distribute.
