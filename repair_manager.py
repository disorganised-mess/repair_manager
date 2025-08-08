#!/usr/bin/env python3
"""
repair_manager.py — Single-file PyQt6 application with SQLite and PDF export.

Requirements:
    pip install PyQt6 reportlab

Run:
    python repair_manager.py
"""
import sys
import os
import sqlite3
import shutil
import csv
import subprocess
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, date

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QPlainTextEdit, QTableWidget,
    QTableWidgetItem, QTabWidget, QMessageBox, QFileDialog, QComboBox, QDialog,
    QDialogButtonBox, QSlider, QDoubleSpinBox, QDateEdit, QSplitter, QHeaderView
)
from PyQt6.QtGui import QAction, QFont, QIcon
from PyQt6.QtCore import Qt, QDate

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# -------- Config --------
DB_FILE = "repair_manager.db"
DEFAULT_BRAND_COLOR = "#2E86C1"

# -------- Utilities --------
def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def backup_db_to(folder):
    if not os.path.exists(DB_FILE):
        return False, "No DB file to backup"
    try:
        if not os.path.isdir(folder):
            return False, "Target is not a directory"
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = os.path.join(folder, f"repair_manager_{ts}.db")
        shutil.copy2(DB_FILE, dest)
        return True, dest
    except Exception as e:
        return False, str(e)

def money(v):
    d = Decimal(str(v or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"${d:,.2f}"

def safe_float(v):
    try:
        return float(v)
    except:
        return 0.0

# -------- Initialize DB & sample data --------
def init_db(sample_data=True):
    existed = os.path.exists(DB_FILE)
    conn = get_conn()
    c = conn.cursor()
    c.executescript("""
    PRAGMA foreign_keys = ON;
    CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT,
        email TEXT,
        address TEXT
    );
    CREATE TABLE IF NOT EXISTS jobs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        description TEXT,
        status TEXT,
        created_on TEXT,
        FOREIGN KEY(client_id) REFERENCES clients(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job_id INTEGER NOT NULL,
        total REAL NOT NULL DEFAULT 0.0,
        status TEXT NOT NULL DEFAULT 'Outstanding',
        issued_on TEXT,
        FOREIGN KEY(job_id) REFERENCES jobs(id) ON DELETE CASCADE
    );
    CREATE TABLE IF NOT EXISTS invoice_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        invoice_id INTEGER NOT NULL,
        description TEXT,
        qty REAL NOT NULL DEFAULT 1,
        unit_price REAL NOT NULL DEFAULT 0.0,
        tax_rate REAL NOT NULL DEFAULT 0.0,
        line_total REAL NOT NULL DEFAULT 0.0,
        FOREIGN KEY(invoice_id) REFERENCES invoices(id) ON DELETE CASCADE
    );
    """)
    conn.commit()

    def sset(k, v):
        c.execute("INSERT OR IGNORE INTO settings (key,value) VALUES (?,?)", (k, v))

    # default settings
    sset("company_name", "Your Company")
    sset("company_address", "123 Business Rd\nCity, State ZIP")
    sset("company_phone", "")
    sset("company_email", "")
    sset("company_logo", "")  # path to logo image
    sset("ui_theme", "light")
    sset("ui_font_size", "10")
    sset("brand_color", DEFAULT_BRAND_COLOR)
    sset("invoice_header_template", "{company_name}\n{company_address}")
    sset("invoice_footer_template", "Thank you for your business.")
    conn.commit()

    if (not existed) and sample_data:
        # sample clients
        c.execute("INSERT INTO clients (name,phone,email,address) VALUES (?,?,?,?)",
                  ("Alice Johnson", "555-0100", "alice@example.com", "100 Maple St\nTown"))
        c.execute("INSERT INTO clients (name,phone,email,address) VALUES (?,?,?,?)",
                  ("TechCorp LLC", "555-0200", "info@techcorp.example", "200 Corporate Way\nCity"))
        conn.commit()
        c.execute("SELECT id FROM clients WHERE name=?", ("Alice Johnson",))
        alice = c.fetchone()[0]
        c.execute("SELECT id FROM clients WHERE name=?", ("TechCorp LLC",))
        tc = c.fetchone()[0]
        today = date.today().isoformat()
        c.execute("INSERT INTO jobs (client_id,description,status,created_on) VALUES (?,?,?,?)",
                  (alice, "Laptop LCD replacement", "Open", today))
        c.execute("INSERT INTO jobs (client_id,description,status,created_on) VALUES (?,?,?,?)",
                  (tc, "Network rack maintenance", "In Progress", today))
        conn.commit()
        c.execute("SELECT id FROM jobs WHERE client_id=? LIMIT 1", (alice,))
        job1 = c.fetchone()[0]
        c.execute("INSERT INTO invoices (job_id,total,status,issued_on) VALUES (?,?,?,?)",
                  (job1, 160.00, "Outstanding", today))
        conn.commit()
        c.execute("SELECT id FROM invoices WHERE job_id=? LIMIT 1", (job1,))
        inv1 = c.fetchone()[0]
        c.execute("INSERT INTO invoice_items (invoice_id,description,qty,unit_price,tax_rate,line_total) VALUES (?,?,?,?,?,?)",
                  (inv1, "LCD Screen - Model X", 1, 120.00, 0.0, 120.00))
        c.execute("INSERT INTO invoice_items (invoice_id,description,qty,unit_price,tax_rate,line_total) VALUES (?,?,?,?,?,?)",
                  (inv1, "Labor", 1, 40.00, 0.0, 40.00))
        conn.commit()

    conn.close()

# -------- Settings helpers --------
def get_setting(key, default=""):
    conn = get_conn(); c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key=?", (key,))
    r = c.fetchone(); conn.close()
    return r[0] if r else default

def set_setting(key, value):
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)", (key, str(value)))
    conn.commit(); conn.close()

# -------- CRUD helpers --------
def list_clients():
    conn = get_conn(); c = conn.cursor()
    c.execute("SELECT id,name,phone,email,address FROM clients ORDER BY name")
    rows = c.fetchall(); conn.close(); return rows

def add_client(name, phone="", email="", address=""):
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT INTO clients (name,phone,email,address) VALUES (?,?,?,?)", (name, phone, email, address))
    conn.commit(); conn.close()

def update_client(client_id, name, phone, email, address):
    conn = get_conn(); c = conn.cursor()
    c.execute("UPDATE clients SET name=?,phone=?,email=?,address=? WHERE id=?", (name, phone, email, address, client_id))
    conn.commit(); conn.close()

def delete_client(client_id):
    conn = get_conn(); c = conn.cursor()
    c.execute("DELETE FROM clients WHERE id=?", (client_id,))
    conn.commit(); conn.close()

def list_jobs(client_id=None):
    conn = get_conn(); c = conn.cursor()
    if client_id and client_id != 0:
        c.execute("""SELECT j.id,j.client_id,j.description,j.status,j.created_on,c.name as client_name
                     FROM jobs j JOIN clients c ON j.client_id=c.id WHERE j.client_id=? ORDER BY j.created_on DESC""", (client_id,))
    else:
        c.execute("""SELECT j.id,j.client_id,j.description,j.status,j.created_on,c.name as client_name
                     FROM jobs j JOIN clients c ON j.client_id=c.id ORDER BY j.created_on DESC""")
    rows = c.fetchall(); conn.close(); return rows

def add_job(client_id, description, status="Open", created_on=None):
    if not created_on: created_on = date.today().isoformat()
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT INTO jobs (client_id,description,status,created_on) VALUES (?,?,?,?)", (client_id, description, status, created_on))
    conn.commit(); conn.close()

def update_job(job_id, description, status):
    conn = get_conn(); c = conn.cursor()
    c.execute("UPDATE jobs SET description=?, status=? WHERE id=?", (description, status, job_id))
    conn.commit(); conn.close()

def delete_job(job_id):
    conn = get_conn(); c = conn.cursor()
    c.execute("DELETE FROM jobs WHERE id=?", (job_id,))
    conn.commit(); conn.close()

def list_invoices(status_filter=None):
    conn = get_conn(); c = conn.cursor()
    if status_filter and status_filter != "All":
        c.execute("""SELECT i.id,i.job_id,i.total,i.status,i.issued_on, j.description, c.name
                     FROM invoices i
                     JOIN jobs j ON i.job_id=j.id
                     JOIN clients c ON j.client_id=c.id
                     WHERE i.status=? ORDER BY i.issued_on DESC""", (status_filter,))
    else:
        c.execute("""SELECT i.id,i.job_id,i.total,i.status,i.issued_on, j.description, c.name
                     FROM invoices i
                     JOIN jobs j ON i.job_id=j.id
                     JOIN clients c ON j.client_id=c.id
                     ORDER BY i.issued_on DESC""")
    rows = c.fetchall(); conn.close(); return rows

def add_invoice(job_id, issued_on=None):
    if not issued_on: issued_on = date.today().isoformat()
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT INTO invoices (job_id,total,status,issued_on) VALUES (?,?,?,?)", (job_id, 0.0, "Outstanding", issued_on))
    invoice_id = c.lastrowid
    conn.commit(); conn.close()
    return invoice_id

def update_invoice_total(invoice_id):
    conn = get_conn(); c = conn.cursor()
    c.execute("SELECT SUM(line_total) as s FROM invoice_items WHERE invoice_id=?", (invoice_id,))
    row = c.fetchone()
    total = row["s"] if row and row["s"] is not None else 0.0
    c.execute("UPDATE invoices SET total=? WHERE id=?", (float(total), invoice_id))
    conn.commit(); conn.close()
    return float(total)

def update_invoice_status(invoice_id, status):
    conn = get_conn(); c = conn.cursor()
    if status not in ("Outstanding", "Paid"):
        status = "Outstanding"
    c.execute("UPDATE invoices SET status=? WHERE id=?", (status, invoice_id))
    conn.commit(); conn.close()

def delete_invoice(invoice_id):
    conn = get_conn(); c = conn.cursor()
    c.execute("DELETE FROM invoices WHERE id=?", (invoice_id,))
    conn.commit(); conn.close()

def get_invoice_items(invoice_id):
    conn = get_conn(); c = conn.cursor()
    c.execute("SELECT id,description,qty,unit_price,tax_rate,line_total FROM invoice_items WHERE invoice_id=? ORDER BY id", (invoice_id,))
    rows = c.fetchall(); conn.close(); return rows

def add_invoice_item(invoice_id, description, qty, unit_price, tax_rate=0.0):
    line_total = float(Decimal(str(qty)) * Decimal(str(unit_price)))
    conn = get_conn(); c = conn.cursor()
    c.execute("INSERT INTO invoice_items (invoice_id,description,qty,unit_price,tax_rate,line_total) VALUES (?,?,?,?,?,?)",
              (invoice_id, description, qty, unit_price, tax_rate, line_total))
    conn.commit(); conn.close()
    update_invoice_total(invoice_id)

def update_invoice_item(item_id, description, qty, unit_price, tax_rate):
    line_total = float(Decimal(str(qty)) * Decimal(str(unit_price)))
    conn = get_conn(); c = conn.cursor()
    c.execute("UPDATE invoice_items SET description=?,qty=?,unit_price=?,tax_rate=?,line_total=? WHERE id=?",
              (description, qty, unit_price, tax_rate, line_total, item_id))
    conn.commit()
    c.execute("SELECT invoice_id FROM invoice_items WHERE id=?", (item_id,))
    row = c.fetchone()
    conn.close()
    if row:
        update_invoice_total(row["invoice_id"])

def delete_invoice_item(item_id):
    conn = get_conn(); c = conn.cursor()
    c.execute("SELECT invoice_id FROM invoice_items WHERE id=?", (item_id,))
    row = c.fetchone()
    if row:
        invoice_id = row["invoice_id"]
        c.execute("DELETE FROM invoice_items WHERE id=?", (item_id,))
        conn.commit(); conn.close()
        update_invoice_total(invoice_id)
    else:
        conn.close()

# -------- Dialogs --------
class ClientDialog(QDialog):
    def __init__(self, parent=None, client=None):
        super().__init__(parent)
        self.setWindowTitle("Client")
        self.resize(420, 320)
        v = QVBoxLayout(self)
        form = QFormLayout()
        self.name = QLineEdit()
        self.phone = QLineEdit()
        self.email = QLineEdit()
        self.address = QPlainTextEdit()
        form.addRow("Name:", self.name)
        form.addRow("Phone:", self.phone)
        form.addRow("Email:", self.email)
        form.addRow("Address:", self.address)
        v.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject)
        v.addWidget(buttons)
        if client:
            self.name.setText(client["name"])
            self.phone.setText(client["phone"] or "")
            self.email.setText(client["email"] or "")
            self.address.setPlainText(client["address"] or "")

    def get(self):
        return self.name.text().strip(), self.phone.text().strip(), self.email.text().strip(), self.address.toPlainText().strip()

class JobDialog(QDialog):
    def __init__(self, parent=None, clients=None, job=None):
        super().__init__(parent)
        self.setWindowTitle("Job")
        v = QVBoxLayout(self)
        form = QFormLayout()
        self.client_combo = QComboBox()
        for c in clients:
            self.client_combo.addItem(c["name"], c["id"])
        self.description = QTextEdit()
        self.status = QComboBox(); self.status.addItems(["Open", "In Progress", "Completed"])
        self.date_edit = QDateEdit(); self.date_edit.setCalendarPopup(True); self.date_edit.setDate(QDate.currentDate())
        form.addRow("Client:", self.client_combo)
        form.addRow("Description:", self.description)
        form.addRow("Status:", self.status)
        form.addRow("Created On:", self.date_edit)
        v.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject)
        v.addWidget(buttons)
        if job:
            idx = self.client_combo.findData(job["client_id"]); self.client_combo.setCurrentIndex(idx)
            self.description.setPlainText(job["description"])
            self.status.setCurrentText(job["status"])
            try:
                d = QDate.fromString(job["created_on"], "yyyy-MM-dd")
                if d.isValid(): self.date_edit.setDate(d)
            except: pass

    def get(self):
        return self.client_combo.currentData(), self.description.toPlainText().strip(), self.status.currentText(), self.date_edit.date().toString("yyyy-MM-dd")

class InvoiceDialog(QDialog):
    def __init__(self, parent=None, jobs=None, invoice=None):
        super().__init__(parent)
        self.setWindowTitle("Invoice")
        v = QVBoxLayout(self)
        form = QFormLayout()
        self.job_combo = QComboBox()
        for j in jobs:
            self.job_combo.addItem(f"{j['id']} - {j['description'][:60]}", j['id'])
        self.issued_on = QDateEdit(); self.issued_on.setCalendarPopup(True); self.issued_on.setDate(QDate.currentDate())
        self.status = QComboBox(); self.status.addItems(["Outstanding","Paid"])
        form.addRow("Job:", self.job_combo)
        form.addRow("Issued On:", self.issued_on)
        form.addRow("Status:", self.status)
        v.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject)
        v.addWidget(buttons)
        if invoice:
            idx = self.job_combo.findData(invoice["job_id"]); self.job_combo.setCurrentIndex(idx)
            try:
                d = QDate.fromString(invoice["issued_on"], "yyyy-MM-dd")
                if d.isValid(): self.issued_on.setDate(d)
            except: pass
            self.status.setCurrentText(invoice["status"])

    def get(self):
        return self.job_combo.currentData(), self.issued_on.date().toString("yyyy-MM-dd"), self.status.currentText()

class InvoiceItemDialog(QDialog):
    def __init__(self, parent=None, item=None):
        super().__init__(parent)
        self.setWindowTitle("Invoice Item")
        v = QVBoxLayout(self)
        form = QFormLayout()
        self.description = QLineEdit()
        self.qty = QDoubleSpinBox(); self.qty.setMinimum(0); self.qty.setDecimals(2); self.qty.setValue(1)
        self.unit_price = QDoubleSpinBox(); self.unit_price.setMinimum(0); self.unit_price.setDecimals(2)
        self.tax_rate = QDoubleSpinBox(); self.tax_rate.setMinimum(0); self.tax_rate.setMaximum(100); self.tax_rate.setDecimals(2)
        form.addRow("Description:", self.description)
        form.addRow("Quantity:", self.qty)
        form.addRow("Unit Price (USD):", self.unit_price)
        form.addRow("Tax Rate (%):", self.tax_rate)
        v.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept); buttons.rejected.connect(self.reject)
        v.addWidget(buttons)
        if item:
            self.description.setText(item["description"])
            self.qty.setValue(item["qty"])
            self.unit_price.setValue(item["unit_price"])
            self.tax_rate.setValue(item["tax_rate"])

    def get(self):
        return self.description.text().strip(), float(self.qty.value()), float(self.unit_price.value()), float(self.tax_rate.value())

# -------- Tabs --------
class ClientsTab(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("Search:"))
        self.search = QLineEdit(); self.search.setPlaceholderText("Search by name, phone, email"); self.search.textChanged.connect(self.refresh)
        top.addWidget(self.search)
        import_btn = QPushButton("Import CSV"); import_btn.clicked.connect(self.import_csv)
        export_btn = QPushButton("Export CSV"); export_btn.clicked.connect(self.export_csv)
        top.addWidget(import_btn); top.addWidget(export_btn)
        v.addLayout(top)
        self.table = QTableWidget(); self.table.setColumnCount(5); self.table.setHorizontalHeaderLabels(["ID","Name","Phone","Email","Address"]); self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        v.addWidget(self.table)
        h = QHBoxLayout()
        add = QPushButton("Add"); add.clicked.connect(self.add)
        edit = QPushButton("Edit"); edit.clicked.connect(self.edit)
        delete = QPushButton("Delete"); delete.clicked.connect(self.delete)
        h.addWidget(add); h.addWidget(edit); h.addWidget(delete)
        v.addLayout(h)
        self.refresh()

    def refresh(self):
        term = self.search.text().strip().lower()
        rows = list_clients()
        if term:
            rows = [r for r in rows if term in r["name"].lower() or term in (r["phone"] or "").lower() or term in (r["email"] or "").lower()]
        self.rows = rows
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(str(r["id"])))
            self.table.setItem(i, 1, QTableWidgetItem(r["name"]))
            self.table.setItem(i, 2, QTableWidgetItem(r["phone"] or ""))
            self.table.setItem(i, 3, QTableWidgetItem(r["email"] or ""))
            self.table.setItem(i, 4, QTableWidgetItem(r["address"] or ""))
        self.table.resizeColumnsToContents()

    def selected_id(self):
        r = self.table.currentRow()
        if r < 0: return None
        return int(self.table.item(r, 0).text())

    def add(self):
        dlg = ClientDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            name, phone, email, address = dlg.get()
            if name:
                add_client(name, phone, email, address)
                self.refresh()

    def edit(self):
        cid = self.selected_id()
        if not cid:
            QMessageBox.information(self, "Select", "Select a client to edit.")
            return
        client = next((r for r in self.rows if r["id"] == cid), None)
        dlg = ClientDialog(self, client)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            name, phone, email, address = dlg.get()
            update_client(cid, name, phone, email, address)
            self.refresh()

    def delete(self):
        cid = self.selected_id()
        if not cid:
            QMessageBox.information(self, "Select", "Select a client to delete.")
            return
        ans = QMessageBox.question(self, "Confirm", "Delete client and all related jobs/invoices? This cannot be undone.")
        if ans == QMessageBox.StandardButton.Yes:
            delete_client(cid)
            self.refresh()

    def export_csv(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Export Clients CSV", "clients.csv", "CSV Files (*.csv)")
        if not fname: return
        if not fname.lower().endswith(".csv"): fname += ".csv"
        with open(fname, "w", newline='', encoding='utf-8') as f:
            writer = csv.writer(f); writer.writerow(["id","name","phone","email","address"])
            for r in self.rows: writer.writerow([r["id"], r["name"], r["phone"] or "", r["email"] or "", r["address"] or ""])
        QMessageBox.information(self, "Export", f"Exported {len(self.rows)} clients to {fname}")

    def import_csv(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Import Clients CSV", "", "CSV Files (*.csv)")
        if not fname: return
        added = 0
        with open(fname, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = row.get("name") or row.get("Name") or ""
                if not name: continue
                phone = row.get("phone") or ""
                email = row.get("email") or ""
                address = row.get("address") or ""
                add_client(name, phone, email, address); added += 1
        self.refresh()
        QMessageBox.information(self, "Import", f"Imported {added} clients from CSV")

class JobsTab(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self)
        f = QHBoxLayout()
        f.addWidget(QLabel("Client:"))
        self.client_filter = QComboBox(); self.client_filter.addItem("All", 0); self.client_filter.currentIndexChanged.connect(self.refresh)
        f.addWidget(self.client_filter)
        f.addWidget(QLabel("Search:"))
        self.search = QLineEdit(); self.search.setPlaceholderText("Search description"); self.search.textChanged.connect(self.refresh)
        f.addWidget(self.search)
        import_btn = QPushButton("Import CSV"); import_btn.clicked.connect(self.import_csv)
        export_btn = QPushButton("Export CSV"); export_btn.clicked.connect(self.export_csv)
        f.addWidget(import_btn); f.addWidget(export_btn)
        v.addLayout(f)
        self.table = QTableWidget(); self.table.setColumnCount(5); self.table.setHorizontalHeaderLabels(["ID","Client","Description","Status","Created On"]); self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        v.addWidget(self.table)
        h = QHBoxLayout()
        add = QPushButton("Add"); add.clicked.connect(self.add)
        edit = QPushButton("Edit"); edit.clicked.connect(self.edit)
        delete = QPushButton("Delete"); delete.clicked.connect(self.delete)
        h.addWidget(add); h.addWidget(edit); h.addWidget(delete)
        v.addLayout(h)
        self.load_clients(); self.refresh()

    def load_clients(self):
        self.client_filter.blockSignals(True)
        self.client_filter.clear(); self.client_filter.addItem("All", 0)
        rows = list_clients()
        for r in rows: self.client_filter.addItem(r["name"], r["id"])
        self.client_filter.blockSignals(False)

    def refresh(self):
        cid = self.client_filter.currentData() or 0
        term = self.search.text().strip().lower()
        rows = list_jobs(cid if cid else None)
        if term:
            rows = [r for r in rows if term in (r["description"] or "").lower() or term in (r["client_name"] or "").lower()]
        self.rows = rows
        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(str(r["id"])))
            self.table.setItem(i, 1, QTableWidgetItem(r["client_name"]))
            self.table.setItem(i, 2, QTableWidgetItem(r["description"] or ""))
            self.table.setItem(i, 3, QTableWidgetItem(r["status"] or ""))
            self.table.setItem(i, 4, QTableWidgetItem(r["created_on"] or ""))
        self.table.resizeColumnsToContents()

    def selected_id(self):
        r = self.table.currentRow(); return None if r < 0 else int(self.table.item(r,0).text())

    def add(self):
        clients = list_clients()
        if not clients:
            QMessageBox.information(self, "No Clients", "Add a client first.")
            return
        dlg = JobDialog(self, clients=[dict(x) for x in clients])
        if dlg.exec() == QDialog.DialogCode.Accepted:
            client_id, description, status, created_on = dlg.get()
            add_job(client_id, description, status, created_on)
            self.load_clients(); self.refresh()

    def edit(self):
        jid = self.selected_id()
        if not jid:
            QMessageBox.information(self, "Select", "Select a job to edit.")
            return
        job = next((r for r in self.rows if r["id"] == jid), None)
        clients = list_clients()
        dlg = JobDialog(self, clients=[dict(x) for x in clients], job=job)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            client_id, description, status, created_on = dlg.get()
            update_job(jid, description, status)
            self.load_clients(); self.refresh()

    def delete(self):
        jid = self.selected_id()
        if not jid:
            QMessageBox.information(self, "Select", "Select a job to delete.")
            return
        ans = QMessageBox.question(self, "Confirm", "Delete job and related invoices? This cannot be undone.")
        if ans == QMessageBox.StandardButton.Yes:
            delete_job(jid)
            self.load_clients(); self.refresh()

    def export_csv(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Export Jobs CSV", "jobs.csv", "CSV Files (*.csv)")
        if not fname: return
        if not fname.lower().endswith(".csv"): fname += ".csv"
        rows = list_jobs()
        with open(fname, "w", newline='', encoding='utf-8') as f:
            w = csv.writer(f); w.writerow(["id","client_id","description","status","created_on"])
            for r in rows: w.writerow([r["id"], r["client_id"], r["description"] or "", r["status"] or "", r["created_on"] or ""])
        QMessageBox.information(self, "Exported", f"Exported {len(rows)} jobs to {fname}")

    def import_csv(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Import Jobs CSV", "", "CSV Files (*.csv)")
        if not fname: return
        added = 0
        with open(fname, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                client_id = row.get("client_id")
                if not client_id: continue
                description = row.get("description") or ""
                status = row.get("status") or "Open"
                created_on = row.get("created_on") or date.today().isoformat()
                try:
                    add_job(int(client_id), description, status, created_on)
                    added += 1
                except:
                    continue
        self.load_clients(); self.refresh()
        QMessageBox.information(self, "Imported", f"Imported {added} jobs from CSV")

class InvoicesTab(QWidget):
    def __init__(self):
        super().__init__()
        v = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("Filter:"))
        self.filter_box = QComboBox(); self.filter_box.addItems(["All","Outstanding","Paid"]); self.filter_box.currentIndexChanged.connect(self.refresh)
        top.addWidget(self.filter_box)
        top.addWidget(QLabel("Search:"))
        self.search = QLineEdit(); self.search.setPlaceholderText("Search by client or job"); self.search.textChanged.connect(self.refresh)
        top.addWidget(self.search)
        v.addLayout(top)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        left = QWidget(); lv = QVBoxLayout(left)
        self.table = QTableWidget(); self.table.setColumnCount(6); self.table.setHorizontalHeaderLabels(["ID","Client","Job","Total","Status","Issued On"]); self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.itemSelectionChanged.connect(self.refresh_preview)
        lv.addWidget(self.table)
        btn_h = QHBoxLayout()
        add = QPushButton("New Invoice"); add.clicked.connect(self.add_invoice)
        edit = QPushButton("Edit Invoice Lines"); edit.clicked.connect(self.edit_invoice)
        mark_paid = QPushButton("Mark Paid"); mark_paid.clicked.connect(self.mark_paid)
        mark_out = QPushButton("Mark Outstanding"); mark_out.clicked.connect(self.mark_outstanding)
        export = QPushButton("Export CSV"); export.clicked.connect(self.export_csv)
        import_btn = QPushButton("Import CSV"); import_btn.clicked.connect(self.import_csv)
        pdf_btn = QPushButton("Print / PDF"); pdf_btn.clicked.connect(self.generate_pdf)
        batch_pdf = QPushButton("Batch Export PDFs"); batch_pdf.clicked.connect(self.batch_export_pdfs)
        btn_h.addWidget(add); btn_h.addWidget(edit); btn_h.addWidget(mark_paid); btn_h.addWidget(mark_out)
        btn_h.addWidget(export); btn_h.addWidget(import_btn); btn_h.addWidget(pdf_btn); btn_h.addWidget(batch_pdf)
        lv.addLayout(btn_h)
        left.setLayout(lv)

        right = QWidget(); rv = QVBoxLayout(right)
        rv.addWidget(QLabel("Invoice Preview"))
        self.preview = QTextEdit(); self.preview.setReadOnly(True)
        rv.addWidget(self.preview)
        right.setLayout(rv)

        splitter.addWidget(left); splitter.addWidget(right)
        splitter.setStretchFactor(0,3); splitter.setStretchFactor(1,2)
        v.addWidget(splitter)
        self.refresh()

    def refresh(self):
        status = self.filter_box.currentText(); term = self.search.text().strip().lower()
        rows = list_invoices(None if status == "All" else status)
        if term:
            rows = [r for r in rows if term in (r["name"] or "").lower() or term in (r["description"] or "").lower()]
        self.rows = rows
        self.table.setRowCount(len(rows))
        for i,r in enumerate(rows):
            self.table.setItem(i,0,QTableWidgetItem(str(r["id"])))
            self.table.setItem(i,1,QTableWidgetItem(r["name"]))
            self.table.setItem(i,2,QTableWidgetItem(r["description"] or ""))
            self.table.setItem(i,3,QTableWidgetItem(money(r["total"])))
            self.table.setItem(i,4,QTableWidgetItem(r["status"]))
            self.table.setItem(i,5,QTableWidgetItem(r["issued_on"] or ""))
        self.table.resizeColumnsToContents()
        self.refresh_preview()

    def selected_id(self):
        r = self.table.currentRow(); return None if r < 0 else int(self.table.item(r,0).text())

    def add_invoice(self):
        jobs = list_jobs()
        if not jobs:
            QMessageBox.information(self, "No Jobs", "Create a job first.")
            return
        dlg = InvoiceDialog(self, jobs=[dict(x) for x in jobs])
        if dlg.exec() == QDialog.DialogCode.Accepted:
            job_id, issued_on, status = dlg.get()
            invoice_id = add_invoice(job_id, issued_on)
            update_invoice_status(invoice_id, status)
            # open lines editor
            self.edit_invoice(invoice_id)
            self.refresh()

    def edit_invoice(self, invoice_id_override=None):
        iid = invoice_id_override or self.selected_id()
        if not iid:
            QMessageBox.information(self, "Select", "Select an invoice to edit.")
            return
        dlg = InvoiceLinesEditor(self, invoice_id=iid)
        dlg.exec()
        update_invoice_total(iid)
        self.refresh()

    def mark_paid(self):
        iid = self.selected_id();
        if not iid: QMessageBox.information(self, "Select", "Select an invoice"); return
        update_invoice_status(iid, "Paid"); self.refresh()

    def mark_outstanding(self):
        iid = self.selected_id()
        if not iid: QMessageBox.information(self, "Select", "Select an invoice"); return
        update_invoice_status(iid, "Outstanding"); self.refresh()

    def export_csv(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Export Invoices CSV", "invoices.csv", "CSV Files (*.csv)")
        if not fname: return
        if not fname.lower().endswith(".csv"): fname += ".csv"
        rows = self.rows
        with open(fname, "w", newline='', encoding='utf-8') as f:
            w = csv.writer(f); w.writerow(["id","job_id","total","status","issued_on"])
            for r in rows:
                w.writerow([r["id"], r["job_id"], r["total"], r["status"], r["issued_on"] or ""])
        QMessageBox.information(self, "Exported", f"Exported {len(rows)} invoices to {fname}")

    def import_csv(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Import Invoices CSV", "", "CSV Files (*.csv)")
        if not fname: return
        added = 0
        with open(fname, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                job_id = row.get("job_id")
                if not job_id: continue
                issued_on = row.get("issued_on") or date.today().isoformat()
                status = row.get("status") or "Outstanding"
                try:
                    inv_id = add_invoice(int(job_id), issued_on)
                    update_invoice_status(inv_id, status)
                    added += 1
                except:
                    continue
        self.refresh()
        QMessageBox.information(self, "Imported", f"Imported {added} invoices from CSV")

    def refresh_preview(self):
        sel = self.selected_id()
        if not sel:
            self.preview.clear(); return
        conn = get_conn(); c = conn.cursor()
        c.execute("""SELECT i.id,i.total,i.status,i.issued_on,j.id as jobid,j.description as job_desc,c.name as client_name,c.address as client_address,c.phone as client_phone,c.email as client_email
                     FROM invoices i JOIN jobs j ON i.job_id=j.id JOIN clients c ON j.client_id=c.id WHERE i.id=?""", (sel,))
        r = c.fetchone(); conn.close()
        if not r:
            self.preview.clear(); return
        header_tpl = get_setting("invoice_header_template", "")
        footer_tpl = get_setting("invoice_footer_template", "")
        header = header_tpl.replace("{company_name}", get_setting("company_name")).replace("{company_address}", get_setting("company_address"))
        footer = footer_tpl.replace("{company_name}", get_setting("company_name"))
        logo = get_setting("company_logo","")
        html = f"<h2>{get_setting('company_name')}</h2>"
        if logo and os.path.exists(logo):
            html += f"<img src='file://{logo}' style='max-height:80px;'><br/>"
        html += f"<div>{header}</div><hr/>"
        html += f"<div><b>Invoice #{r['id']}</b> &nbsp; Issued: {r['issued_on']} &nbsp; Status: <b>{r['status']}</b></div>"
        html += f"<div><b>Client:</b> {r['client_name']}</div>"
        html += f"<div><b>Job:</b> {r['jobid']} - {r['job_desc']}</div>"
        items = get_invoice_items(sel)
        html += "<table border='1' cellpadding='4' cellspacing='0' style='border-collapse:collapse;width:100%;'><tr><th>Description</th><th>Qty</th><th>Unit</th><th>Tax%</th><th>Line Total</th></tr>"
        for it in items:
            html += f"<tr><td>{it['description']}</td><td>{it['qty']}</td><td>{money(it['unit_price'])}</td><td>{it['tax_rate']}</td><td>{money(it['line_total'])}</td></tr>"
        html += f"<tr><td colspan='4' align='right'><b>Total</b></td><td>{money(r['total'])}</td></tr></table>"
        html += f"<div>{footer}</div>"
        self.preview.setHtml(html)

    def generate_pdf(self):
        sel = self.selected_id()
        if not sel:
            QMessageBox.information(self, "Select", "Select an invoice to print.")
            return
        fname, _ = QFileDialog.getSaveFileName(self, "Save Invoice PDF", f"invoice_{sel}.pdf", "PDF Files (*.pdf)")
        if not fname: return
        if not fname.lower().endswith(".pdf"): fname += ".pdf"
        self._generate_invoice_pdf_file(sel, fname)
        QMessageBox.information(self, "PDF Saved", f"Invoice saved to {fname}")

    def _generate_invoice_pdf_file(self, invoice_id, fname):
        conn = get_conn(); c = conn.cursor()
        c.execute("""SELECT i.id,i.total,i.status,i.issued_on,j.id as jobid,j.description as job_desc,c.name as client_name,c.address as client_address,c.phone as client_phone,c.email as client_email
                     FROM invoices i JOIN jobs j ON i.job_id=j.id JOIN clients c ON j.client_id=c.id WHERE i.id=?""", (invoice_id,))
        r = c.fetchone()
        items = get_invoice_items(invoice_id)
        conn.close()
        if not r: return
        doc = SimpleDocTemplate(fname, pagesize=letter, leftMargin=inch, rightMargin=inch, topMargin=inch, bottomMargin=inch)
        styles = getSampleStyleSheet()
        brand_color = get_setting("brand_color", DEFAULT_BRAND_COLOR)
        title_style = ParagraphStyle("title", parent=styles['Title'], textColor=brand_color)
        normal = styles['Normal']
        elems = []
        logo = get_setting("company_logo","")
        if logo and os.path.exists(logo):
            try:
                im = Image(logo)
                max_h = 0.8 * inch
                im.drawHeight = max_h
                im.drawWidth = im.drawHeight * (im.imageWidth / im.imageHeight)
                elems.append(im)
            except Exception as e:
                print("Logo load error:", e)
        if get_setting("company_name"):
            elems.append(Paragraph(get_setting("company_name"), title_style))
        if get_setting("company_address"):
            elems.append(Paragraph(get_setting("company_address").replace("\n", "<br/>"), normal))
        elems.append(Spacer(1, 12))
        elems.append(Paragraph(f"<b>Invoice:</b> {r['id']}", styles['Heading3']))
        elems.append(Paragraph(f"Issued On: {r['issued_on'] or ''}", normal))
        elems.append(Paragraph(f"Status: {r['status']}", normal))
        elems.append(Spacer(1, 8))
        elems.append(Paragraph(f"<b>Client:</b> {r['client_name']}", normal))
        if r['client_address']:
            elems.append(Paragraph(r['client_address'].replace("\n","<br/>"), normal))
        elems.append(Spacer(1, 12))
        elems.append(Paragraph(f"<b>Job:</b> {r['jobid']} - {r['job_desc']}", normal))
        elems.append(Spacer(1, 12))
        data = [["Description","Qty","Unit Price","Tax %","Line Total"]]
        for it in items:
            data.append([it["description"] or "", f"{it['qty']}", money(it["unit_price"]), f"{it['tax_rate']}", money(it["line_total"])])
        data.append(["", "", "", "Total", money(r["total"])])
        table = Table(data, colWidths=[3.5*inch, 0.7*inch, 1.0*inch, 0.7*inch, 1.0*inch])
        table.setStyle(TableStyle([
            ('GRID',(0,0),(-1,-1),0.5,colors.grey),
            ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#f2f2f2")),
            ('ALIGN',(1,1),(-1,-1),'RIGHT'),
            ('SPAN', (0, len(data)-1), (3, len(data)-1)),
            ('BACKGROUND', (3,len(data)-1),(3,len(data)-1), colors.HexColor(brand_color)),
            ('TEXTCOLOR', (3,len(data)-1),(3,len(data)-1), colors.white),
            ('ALIGN', (3,len(data)-1),(3,len(data)-1), 'RIGHT'),
        ]))
        elems.append(table)
        elems.append(Spacer(1,18))
        elems.append(Paragraph(get_setting("invoice_footer_template","Thank you for your business."), normal))
        elems.append(Spacer(1,8))
        elems.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ParagraphStyle("meta", parent=normal, fontSize=8, textColor=colors.grey)))
        doc.build(elems)
        try:
            if sys.platform.startswith("linux"):
                subprocess.Popen(["xdg-open", fname])
            elif sys.platform.startswith("darwin"):
                subprocess.Popen(["open", fname])
            elif sys.platform.startswith("win"):
                os.startfile(fname)
        except Exception as e:
            print("Open PDF failed:", e)

    def batch_export_pdfs(self):
        folder = QFileDialog.getExistingDirectory(self, "Select output folder for PDFs", os.getcwd())
        if not folder: return
        opt, ok2 = QMessageBox.getItem(self, "Batch Export", "Which invoices to export?", ["All", "Outstanding", "Selected"], 0, False)
        if not ok2: return
        ids = []
        if opt == "All":
            rows = list_invoices(None)
            ids = [r["id"] for r in rows]
        elif opt == "Outstanding":
            rows = list_invoices("Outstanding")
            ids = [r["id"] for r in rows]
        else:
            sel_rows = self.table.selectionModel().selectedRows()
            ids = [int(self.table.item(r.row(),0).text()) for r in sel_rows]
        if not ids:
            QMessageBox.information(self, "No invoices", "No invoices found for selection.")
            return
        for iid in ids:
            fname = os.path.join(folder, f"invoice_{iid}.pdf")
            self._generate_invoice_pdf_file(iid, fname)
        QMessageBox.information(self, "Batch Export", f"Exported {len(ids)} invoices to {folder}")

class InvoiceLinesEditor(QDialog):
    def __init__(self, parent, invoice_id):
        super().__init__(parent)
        self.setWindowTitle(f"Invoice #{invoice_id} — Lines Editor")
        self.invoice_id = invoice_id
        self.resize(900, 520)
        v = QVBoxLayout(self)
        info_h = QHBoxLayout()
        info_h.addWidget(QLabel(f"Invoice ID: {invoice_id}"))
        self.total_label = QLabel("Total: $0.00"); info_h.addWidget(self.total_label)
        info_h.addStretch()
        v.addLayout(info_h)
        self.table = QTableWidget(); self.table.setColumnCount(6); self.table.setHorizontalHeaderLabels(["ID","Description","Qty","Unit Price","Tax %","Line Total"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        v.addWidget(self.table)
        h = QHBoxLayout()
        add = QPushButton("Add Line"); add.clicked.connect(self.add_line)
        edit = QPushButton("Edit Line"); edit.clicked.connect(self.edit_line)
        delete = QPushButton("Delete Line"); delete.clicked.connect(self.delete_line)
        save = QPushButton("Save & Close"); save.clicked.connect(self.accept)
        h.addWidget(add); h.addWidget(edit); h.addWidget(delete); h.addStretch(); h.addWidget(save)
        v.addLayout(h)
        self.refresh()

    def refresh(self):
        rows = get_invoice_items(self.invoice_id)
        self.rows = rows
        self.table.setRowCount(len(rows))
        total = 0.0
        for i, r in enumerate(rows):
            self.table.setItem(i, 0, QTableWidgetItem(str(r["id"])))
            self.table.setItem(i, 1, QTableWidgetItem(r["description"] or ""))
            self.table.setItem(i, 2, QTableWidgetItem(str(r["qty"])))
            self.table.setItem(i, 3, QTableWidgetItem(money(r["unit_price"])))
            self.table.setItem(i, 4, QTableWidgetItem(str(r["tax_rate"])))
            self.table.setItem(i, 5, QTableWidgetItem(money(r["line_total"])))
            total += r["line_total"]
        update_invoice_total(self.invoice_id)
        self.total_label.setText(f"Total: {money(total)}")
        self.table.resizeColumnsToContents()

    def selected_item_id(self):
        r = self.table.currentRow()
        if r < 0: return None
        return int(self.table.item(r,0).text())

    def add_line(self):
        dlg = InvoiceItemDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            description, qty, unit_price, tax_rate = dlg.get()
            add_invoice_item(self.invoice_id, description, qty, unit_price, tax_rate)
            self.refresh()

    def edit_line(self):
        item_id = self.selected_item_id()
        if not item_id:
            QMessageBox.information(self, "Select", "Select a line to edit.")
            return
        item = next((r for r in self.rows if r["id"] == item_id), None)
        dlg = InvoiceItemDialog(self, item)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            description, qty, unit_price, tax_rate = dlg.get()
            update_invoice_item(item_id, description, qty, unit_price, tax_rate)
            self.refresh()

    def delete_line(self):
        item_id = self.selected_item_id()
        if not item_id:
            QMessageBox.information(self, "Select", "Select a line to delete.")
            return
        ans = QMessageBox.question(self, "Confirm", "Delete this line item?")
        if ans == QMessageBox.StandardButton.Yes:
            delete_invoice_item(item_id)
            self.refresh()

class SettingsTab(QWidget):
    def __init__(self, refresh_callback=None, apply_theme_callback=None):
        super().__init__()
        self.refresh_callback = refresh_callback
        self.apply_theme_callback = apply_theme_callback
        v = QVBoxLayout(self)
        form = QFormLayout()
        self.company_name = QLineEdit(get_setting("company_name",""))
        self.company_address = QPlainTextEdit(get_setting("company_address",""))
        self.company_phone = QLineEdit(get_setting("company_phone",""))
        self.company_email = QLineEdit(get_setting("company_email",""))
        self.logo_path = QLineEdit(get_setting("company_logo",""))
        browse = QPushButton("Browse...")
        browse.clicked.connect(self.browse_logo)
        logo_h = QHBoxLayout(); logo_h.addWidget(self.logo_path); logo_h.addWidget(browse)
        form.addRow("Company Name:", self.company_name)
        form.addRow("Company Address:", self.company_address)
        form.addRow("Phone:", self.company_phone)
        form.addRow("Email:", self.company_email)
        form.addRow(QLabel("Logo Path:"), logo_h)
        v.addLayout(form)
        v.addWidget(QLabel("Invoice Header Template (placeholders {company_name},{company_address},{invoice_number},{invoice_date},{invoice_status}):"))
        self.header_tpl = QPlainTextEdit(get_setting("invoice_header_template","{company_name}\n{company_address}"))
        v.addWidget(self.header_tpl)
        v.addWidget(QLabel("Invoice Footer Template:"))
        self.footer_tpl = QPlainTextEdit(get_setting("invoice_footer_template","Thank you for your business."))
        v.addWidget(self.footer_tpl)
        pref_h = QHBoxLayout()
        pref_h.addWidget(QLabel("Theme:"))
        self.theme = QComboBox(); self.theme.addItems(["light","dark"]); self.theme.setCurrentText(get_setting("ui_theme","light"))
        pref_h.addWidget(self.theme)
        pref_h.addWidget(QLabel("Font size:"))
        self.font_slider = QSlider(Qt.Orientation.Horizontal); self.font_slider.setMinimum(8); self.font_slider.setMaximum(20)
        try:
            self.font_slider.setValue(int(get_setting("ui_font_size","10")))
        except: self.font_slider.setValue(10)
        pref_h.addWidget(self.font_slider)
        pref_h.addWidget(QLabel("Brand Color (hex):"))
        self.brand_color = QLineEdit(get_setting("brand_color", DEFAULT_BRAND_COLOR))
        pref_h.addWidget(self.brand_color)
        v.addLayout(pref_h)
        btn_h = QHBoxLayout()
        save = QPushButton("Save All"); save.clicked.connect(self.save_all)
        backup = QPushButton("Backup DB"); backup.clicked.connect(self.backup_db)
        btn_h.addWidget(save); btn_h.addWidget(backup)
        v.addLayout(btn_h)

    def browse_logo(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select Logo", "", "Images (*.png *.jpg *.jpeg *.svg)")
        if fname:
            self.logo_path.setText(fname); set_setting("company_logo", fname)
            if self.apply_theme_callback: self.apply_theme_callback()

    def save_all(self):
        set_setting("company_name", self.company_name.text())
        set_setting("company_address", self.company_address.toPlainText())
        set_setting("company_phone", self.company_phone.text())
        set_setting("company_email", self.company_email.text())
        set_setting("company_logo", self.logo_path.text())
        set_setting("invoice_header_template", self.header_tpl.toPlainText())
        set_setting("invoice_footer_template", self.footer_tpl.toPlainText())
        set_setting("ui_theme", self.theme.currentText())
        set_setting("ui_font_size", str(self.font_slider.value()))
        set_setting("brand_color", self.brand_color.text() or DEFAULT_BRAND_COLOR)
        if self.refresh_callback: self.refresh_callback()
        if self.apply_theme_callback: self.apply_theme_callback()
        QMessageBox.information(self, "Saved", "Settings saved.")

    def backup_db(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Backup Folder", os.getcwd())
        if not folder: return
        ok, result = backup_db_to(folder)
        if ok:
            QMessageBox.information(self, "Backup Created", f"Backup saved to:\n{result}")
        else:
            QMessageBox.warning(self, "Backup Failed", f"Backup failed: {result}")

# -------- Main Window --------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Repair Manager")
        self.resize(1200, 800)
        toolbar = self.addToolBar("Main")
        refresh_action = QAction(QIcon.fromTheme("view-refresh"), "Refresh", self)
        refresh_action.triggered.connect(self.refresh_all); toolbar.addAction(refresh_action)
        self.clients_tab = ClientsTab()
        self.jobs_tab = JobsTab()
        self.invoices_tab = InvoicesTab()
        self.settings_tab = SettingsTab(refresh_callback=self.refresh_all, apply_theme_callback=self.apply_theme)
        tabs = QTabWidget()
        tabs.addTab(self.clients_tab, "Clients")
        tabs.addTab(self.jobs_tab, "Jobs")
        tabs.addTab(self.invoices_tab, "Invoices")
        tabs.addTab(self.settings_tab, "Settings")
        self.setCentralWidget(tabs)
        self.apply_theme()

    def apply_theme(self):
        theme = get_setting("ui_theme", "light")
        try:
            fs = int(get_setting("ui_font_size", "10"))
        except:
            fs = 10
        if theme == "dark":
            base_style = "QWidget { background: #2b2b2b; color: #f0f0f0; } QLineEdit, QPlainTextEdit, QTextEdit { background: #3c3c3c; color: #f0f0f0; }"
        else:
            base_style = "QWidget { background: #ffffff; color: #000000; } QLineEdit, QPlainTextEdit, QTextEdit { background: #ffffff; color: #000000; }"
        self.setStyleSheet(base_style)
        font = QFont(); font.setPointSize(fs); QApplication.instance().setFont(font)
        try:
            self.invoices_tab.refresh_preview()
        except: pass

    def refresh_all(self):
        try: self.clients_tab.refresh()
        except: pass
        try: self.jobs_tab.load_clients(); self.jobs_tab.refresh()
        except: pass
        try: self.invoices_tab.refresh()
        except: pass

# -------- Entry point --------
def main():
    init_db(sample_data=True)
    app = QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
