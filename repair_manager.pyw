\
#!/usr/bin/env python3
"""
repair_manager.py - Extended Computer Repair Manager (SQLite + Tkinter)
Features:
- Add/Edit Customers
- Add Equipment linked to Customers
- Technicians management and assignment
- Parts inventory and parts used per work order
- Create / Close Work Orders
- Log Work Details (multiple per work order)
- Reports: Work History by Customer, Open Work Orders
- Export reports to CSV
- Printable work order slip PDF generation
"""

import os, sqlite3, csv
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.styles import getSampleStyleSheet

BASE_DIR = os.path.dirname(__file__)
DB_FILE = os.path.join(BASE_DIR, "repair_shop.db")

# --- Database Helpers ---
def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    with open(os.path.join(BASE_DIR, "database_schema.sql"), "r") as f:
        cur.executescript(f.read())
    conn.commit()
    conn.close()

def run_query(query, params=()):
    conn = get_conn(); cur = conn.cursor()
    cur.execute(query, params); conn.commit(); last = cur.lastrowid; conn.close(); return last

def fetch_all(query, params=()):
    conn = get_conn(); cur = conn.cursor(); cur.execute(query, params); rows = cur.fetchall(); conn.close(); return rows

# --- CRUD ---
def add_customer(first, last, phone, address, email):
    return run_query("INSERT INTO Customers (first_name, last_name, phone, address, email) VALUES (?, ?, ?, ?, ?)", (first, last, phone, address, email))

def list_customers():
    return fetch_all("SELECT * FROM Customers ORDER BY last_name, first_name")

def add_equipment(customer_id, make, model, cpu, ram, storage, os_, serial):
    return run_query("INSERT INTO Equipment (customer_id, make, model, cpu, ram, storage, os, serial_number) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (customer_id, make, model, cpu, ram, storage, os_, serial))

def list_equipment_by_customer(customer_id):
    return fetch_all("SELECT * FROM Equipment WHERE customer_id=?", (customer_id,))

def add_technician(name):
    return run_query("INSERT INTO Technicians (name) VALUES (?)", (name,))

def list_technicians():
    return fetch_all("SELECT * FROM Technicians ORDER BY name")

def add_part(sku, desc, qty, cost):
    return run_query("INSERT INTO Parts (sku, description, quantity, unit_cost) VALUES (?, ?, ?, ?)", (sku, desc, qty, cost))

def list_parts():
    return fetch_all("SELECT * FROM Parts ORDER BY description")

def create_workorder(equipment_id, technician_id, work_description):
    today = date.today().isoformat()
    return run_query("INSERT INTO WorkOrders (equipment_id, technician_id, date_opened, status, work_description) VALUES (?, ?, ?, 'Open', ?)", (equipment_id, technician_id, today, work_description))

def close_workorder(workorder_id):
    today = date.today().isoformat()
    run_query("UPDATE WorkOrders SET date_closed=?, status='Closed' WHERE workorder_id=?", (today, workorder_id))

def add_work_detail(workorder_id, description):
    today = date.today().isoformat()
    return run_query("INSERT INTO WorkDetails (workorder_id, date, description) VALUES (?, ?, ?)", (workorder_id, today, description))

def use_part(workorder_id, part_id, qty):
    # decrement inventory and add PartsUsed record
    run_query("INSERT INTO PartsUsed (workorder_id, part_id, quantity) VALUES (?, ?, ?)", (workorder_id, part_id, qty))
    run_query("UPDATE Parts SET quantity = quantity - ? WHERE part_id = ?", (qty, part_id))

# --- Reports ---
def get_open_workorders():
    return fetch_all("SELECT wo.workorder_id, wo.date_opened, e.equipment_id, e.make, e.model, c.first_name, c.last_name, t.name as technician, wo.work_description FROM WorkOrders wo JOIN Equipment e ON wo.equipment_id=e.equipment_id JOIN Customers c ON e.customer_id=c.customer_id LEFT JOIN Technicians t ON wo.technician_id=t.technician_id WHERE wo.status='Open' ORDER BY wo.date_opened")

def get_workorders_by_customer(customer_id):
    return fetch_all("SELECT wo.workorder_id, wo.date_opened, wo.date_closed, wo.status, wo.work_description, e.make, e.model, e.serial_number FROM WorkOrders wo JOIN Equipment e ON wo.equipment_id=e.equipment_id WHERE e.customer_id=? ORDER BY wo.date_opened DESC", (customer_id,))

def get_workdetails_for_workorder(workorder_id):
    return fetch_all("SELECT * FROM WorkDetails WHERE workorder_id=? ORDER BY date", (workorder_id,))

def search_customers(term):
    like = f"%{term}%"
    return fetch_all("SELECT * FROM Customers WHERE first_name LIKE ? OR last_name LIKE ? OR phone LIKE ? OR email LIKE ? ORDER BY last_name, first_name", (like, like, like, like))

def search_workorders(term):
    like = f"%{term}%"
    return fetch_all("SELECT wo.*, e.make, e.model FROM WorkOrders wo JOIN Equipment e ON wo.equipment_id=e.equipment_id WHERE wo.work_description LIKE ? OR e.serial_number LIKE ? OR wo.workorder_id LIKE ? ORDER BY wo.date_opened DESC", (like, like, like))

# --- PDF print ---
def generate_workorder_pdf(workorder_id, filepath):
    rows = fetch_all("SELECT wo.*, e.*, c.first_name, c.last_name, c.phone, c.email, t.name as technician FROM WorkOrders wo JOIN Equipment e ON wo.equipment_id=e.equipment_id JOIN Customers c ON e.customer_id=c.customer_id LEFT JOIN Technicians t ON wo.technician_id=t.technician_id WHERE wo.workorder_id=?", (workorder_id,))
    if not rows:
        raise ValueError("Work order not found")
    wo = rows[0]
    details = get_workdetails_for_workorder(workorder_id)
    parts = fetch_all("SELECT p.sku, p.description, pu.quantity FROM PartsUsed pu JOIN Parts p ON pu.part_id=p.part_id WHERE pu.workorder_id=?", (workorder_id,))
    doc = SimpleDocTemplate(filepath, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph(f"Work Order #{workorder_id}", styles['Title']))
    story.append(Spacer(1,6))
    story.append(Paragraph(f"Customer: {wo['first_name']} {wo['last_name']}", styles['Normal']))
    story.append(Paragraph(f"Phone: {wo['phone'] or ''}   Email: {wo['email'] or ''}", styles['Normal']))
    story.append(Paragraph(f"Equipment: {wo['make']} {wo['model']}   Serial: {wo['serial_number']}", styles['Normal']))
    story.append(Paragraph(f"Technician: {wo['technician'] or ''}", styles['Normal']))
    story.append(Paragraph(f"Status: {wo['status']}   Opened: {wo['date_opened']}   Closed: {wo['date_closed'] or ''}", styles['Normal']))
    story.append(Spacer(1,8))
    if details:
        story.append(Paragraph("Work Details:", styles['Heading3']))
        for d in details:
            story.append(Paragraph(f"{d['date']}: {d['description']}", styles['Normal']))
    if parts:
        story.append(Spacer(1,8))
        story.append(Paragraph("Parts Used:", styles['Heading3']))
        tdata = [["SKU","Description","Qty"]]+[[p['sku'], p['description'], str(p['quantity'])] for p in parts]
        tbl = Table(tdata, hAlign='LEFT')
        story.append(tbl)
    doc.build(story)

# --- GUI ---
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Computer Repair Manager - Full")
        self.geometry("1000x600")
        self.create_widgets()

    def create_widgets(self):
        frm = ttk.Frame(self); frm.pack(fill="both", expand=True, padx=10, pady=10)
        btn_frame = ttk.Frame(frm); btn_frame.pack(side="top", fill="x")
        actions = [
            ("Add Customer", self.ui_add_customer),
            ("Add Equipment", self.ui_add_equipment),
            ("Add Technician", self.ui_add_technician),
            ("Add Part", self.ui_add_part),
            ("Create Work Order", self.ui_create_workorder),
            ("Log Work Detail", self.ui_log_workdetail),
            ("Use Part on WO", self.ui_use_part),
            ("Close Work Order", self.ui_close_workorder),
            ("Open Work Orders", self.ui_open_workorders),
            ("Work History by Customer", self.ui_work_history_by_customer),
            ("Search", self.ui_search)
        ]
        for (t, cmd) in actions:
            ttk.Button(btn_frame, text=t, command=cmd).pack(side="left", padx=4, pady=4)
        self.tree = ttk.Treeview(frm, columns=("A","B","C","D","E","F","G","H"), show="headings")
        self.tree.pack(fill="both", expand=True, pady=8)
        for i in range(1,9):
            self.tree.heading(f"#{i}", text=f"Col{i}")
            self.tree.column(f"#{i}", width=120)
        self.status = ttk.Label(self, text="Ready", anchor="w"); self.status.pack(side="bottom", fill="x")

    def clear_tree(self):
        for i in self.tree.get_children(): self.tree.delete(i)

    def set_status(self, msg): self.status.config(text=msg)

    # UI: Add customer
    def ui_add_customer(self):
        win = tk.Toplevel(self); win.title("Add Customer")
        labels = ("First Name","Last Name","Phone","Address","Email"); entries=[]
        for r,lab in enumerate(labels):
            ttk.Label(win, text=lab).grid(row=r, column=0, sticky="w", padx=5, pady=3)
            e=ttk.Entry(win); e.grid(row=r, column=1, padx=5, pady=3); entries.append(e)
        def save():
            if not entries[0].get() or not entries[1].get():
                messagebox.showerror("Error","First and Last name required"); return
            add_customer(entries[0].get(), entries[1].get(), entries[2].get(), entries[3].get(), entries[4].get())
            messagebox.showinfo("Saved","Customer added"); win.destroy()
        ttk.Button(win, text="Save", command=save).grid(row=len(labels), columnspan=2, pady=6)

    def ui_add_equipment(self):
        customers = list_customers = fetch_all("SELECT customer_id, first_name || ' ' || last_name AS name FROM Customers")
        if not customers: messagebox.showinfo("Info","Please add customers first"); return
        win = tk.Toplevel(self); win.title("Add Equipment")
        ttk.Label(win, text="Customer").grid(row=0, column=0, padx=5, pady=3)
        cust_var = tk.StringVar(); cust_combo = ttk.Combobox(win, textvariable=cust_var, values=[f\"{r['customer_id']} - {r['name']}\" for r in customers]); cust_combo.grid(row=0, column=1, padx=5, pady=3)
        labels = ("Make","Model","CPU","RAM","Storage","OS","Serial Number"); entries=[]
        for r,lab in enumerate(labels, start=1):
            ttk.Label(win, text=lab).grid(row=r, column=0, sticky="w", padx=5, pady=3)
            e=ttk.Entry(win); e.grid(row=r, column=1, padx=5, pady=3); entries.append(e)
        def save():
            try: cust_id = int(cust_var.get().split(" - ")[0])
            except Exception: messagebox.showerror("Error","Select a valid customer"); return
            add_equipment(cust_id, *(e.get() for e in entries)); messagebox.showinfo("Saved","Equipment added"); win.destroy()
        ttk.Button(win, text="Save", command=save).grid(row=len(labels)+1, columnspan=2, pady=6)

    def ui_add_technician(self):
        name = simpledialog.askstring("Technician","Enter technician name:")
        if not name: return
        add_technician(name); messagebox.showinfo("Saved","Technician added")

    def ui_add_part(self):
        win = tk.Toplevel(self); win.title("Add Part")
        labels = ("SKU","Description","Quantity","Unit Cost"); entries=[]
        for r,lab in enumerate(labels):
            ttk.Label(win, text=lab).grid(row=r, column=0, sticky="w", padx=5, pady=3)
            e=ttk.Entry(win); e.grid(row=r, column=1, padx=5, pady=3); entries.append(e)
        def save():
            try:
                qty = int(entries[2].get())
            except:
                qty = 0
            try:
                cost = float(entries[3].get())
            except:
                cost = 0.0
            add_part(entries[0].get(), entries[1].get(), qty, cost); messagebox.showinfo("Saved","Part added"); win.destroy()
        ttk.Button(win, text="Save", command=save).grid(row=len(labels), columnspan=2, pady=6)

    def ui_create_workorder(self):
        equipment = fetch_all("SELECT equipment_id, make, model FROM Equipment")
        technicians = list_technicians()
        if not equipment: messagebox.showinfo("Info","Add equipment first"); return
        win = tk.Toplevel(self); win.title("Create Work Order")
        ttk.Label(win, text="Equipment").grid(row=0, column=0, padx=5, pady=3)
        eq_var=tk.StringVar(); ttk.Combobox(win, textvariable=eq_var, values=[f\"{r['equipment_id']} - {r['make']} {r['model']}\" for r in equipment]).grid(row=0,column=1,padx=5,pady=3)
        ttk.Label(win, text="Technician (optional)").grid(row=1,column=0,padx=5,pady=3)
        tech_var=tk.StringVar(); ttk.Combobox(win, textvariable=tech_var, values=[f\"{t['technician_id']} - {t['name']}\" for t in technicians]).grid(row=1,column=1,padx=5,pady=3)
        ttk.Label(win, text="Work Description").grid(row=2,column=0,padx=5,pady=3)
        desc=ttk.Entry(win,width=50); desc.grid(row=2,column=1,padx=5,pady=3)
        def save():
            try: eq_id=int(eq_var.get().split(" - ")[0])
            except: messagebox.showerror("Error","Select valid equipment"); return
            tech_id=None
            if tech_var.get(): tech_id=int(tech_var.get().split(" - ")[0])
            create_workorder(eq_id, tech_id, desc.get()); messagebox.showinfo("Saved","Work order created"); win.destroy()
        ttk.Button(win, text="Create", command=save).grid(row=3,columnspan=2,pady=6)

    def ui_log_workdetail(self):
        wos = fetch_all("SELECT workorder_id FROM WorkOrders ORDER BY workorder_id DESC")
        if not wos: messagebox.showinfo("Info","No work orders found"); return
        win = tk.Toplevel(self); win.title("Log Work Detail")
        ttk.Label(win, text="Work Order ID").grid(row=0,column=0,padx=5,pady=3)
        wo_var=tk.StringVar(); ttk.Combobox(win, textvariable=wo_var, values=[str(r['workorder_id']) for r in wos]).grid(row=0,column=1,padx=5,pady=3)
        ttk.Label(win, text="Description").grid(row=1,column=0,padx=5,pady=3)
        txt=tk.Text(win,width=60,height=8); txt.grid(row=1,column=1,padx=5,pady=3)
        def save():
            try: wid=int(wo_var.get())
            except: messagebox.showerror("Error","Select valid work order"); return
            add_work_detail(wid, txt.get("1.0","end").strip()); messagebox.showinfo("Saved","Work detail logged"); win.destroy()
        ttk.Button(win, text="Save", command=save).grid(row=2,columnspan=2,pady=6)

    def ui_use_part(self):
        parts = list_parts()
        wos = fetch_all("SELECT workorder_id FROM WorkOrders ORDER BY workorder_id DESC")
        if not parts: messagebox.showinfo("Info","No parts in inventory"); return
        if not wos: messagebox.showinfo("Info","No work orders"); return
        win = tk.Toplevel(self); win.title("Use Part on Work Order")
        ttk.Label(win, text="Work Order ID").grid(row=0,column=0,padx=5,pady=3)
        wo_var=tk.StringVar(); ttk.Combobox(win, textvariable=wo_var, values=[str(r['workorder_id']) for r in wos]).grid(row=0,column=1,padx=5,pady=3)
        ttk.Label(win, text="Part").grid(row=1,column=0,padx=5,pady=3)
        part_var=tk.StringVar(); ttk.Combobox(win, textvariable=part_var, values=[f\"{p['part_id']} - {p['description']} (Qty {p['quantity']})\" for p in parts]).grid(row=1,column=1,padx=5,pady=3)
        ttk.Label(win, text="Quantity").grid(row=2,column=0,padx=5,pady=3); qty_ent=ttk.Entry(win); qty_ent.grid(row=2,column=1,padx=5,pady=3)
        def save():
            try: wid=int(wo_var.get()); part_id=int(part_var.get().split(" - ")[0]); qty=int(qty_ent.get())
            except: messagebox.showerror("Error","Invalid selection/quantity"); return
            use_part(wid, part_id, qty); messagebox.showinfo("Saved","Part used"); win.destroy()
        ttk.Button(win, text="Save", command=save).grid(row=3,columnspan=2,pady=6)

    def ui_close_workorder(self):
        rows = fetch_all("SELECT workorder_id FROM WorkOrders WHERE status='Open' ORDER BY workorder_id")
        if not rows: messagebox.showinfo("Info","No open work orders"); return
        win = tk.Toplevel(self); win.title("Close Work Order")
        ttk.Label(win, text="Open WorkOrder ID").grid(row=0,column=0,padx=5,pady=3)
        wo_var=tk.StringVar(); ttk.Combobox(win, textvariable=wo_var, values=[str(r['workorder_id']) for r in rows]).grid(row=0,column=1,padx=5,pady=3)
        def closeit():
            try: wid=int(wo_var.get())
            except: messagebox.showerror("Error","Select valid work order"); return
            close_workorder(wid); messagebox.showinfo("Closed",f"Work order {wid} closed"); win.destroy()
        ttk.Button(win, text="Close", command=closeit).grid(row=1,columnspan=2,pady=6)

    def ui_open_workorders(self):
        data = get_open_workorders(); self.clear_tree()
        cols=("WO ID","Opened","Equip ID","Make","Model","Customer","Technician","Description")
        self.tree["columns"]=cols
        for i,col in enumerate(cols): self.tree.heading(f"#{i+1}", text=col); self.tree.column(f"#{i+1}", width=120)
        for r in data:
            cust = f\"{r['first_name']} {r['last_name']}\" if r['first_name'] else ""
            self.tree.insert("", "end", values=(r['workorder_id'], r['date_opened'], r['equipment_id'], r['make'], r['model'], cust, r['technician'] or "", r['work_description']))
        self.set_status(f"{len(data)} open work orders")

    def ui_work_history_by_customer(self):
        customers = fetch_all("SELECT customer_id, first_name || ' ' || last_name AS name FROM Customers")
        if not customers: messagebox.showinfo("Info","No customers found"); return
        win = tk.Toplevel(self); win.title("Work History by Customer")
        ttk.Label(win, text="Customer").grid(row=0,column=0,padx=5,pady=3)
        cust_var=tk.StringVar(); cust_combo=ttk.Combobox(win, textvariable=cust_var, values=[f\"{r['customer_id']} - {r['name']}\" for r in customers]); cust_combo.grid(row=0,column=1,padx=5,pady=3)
        tree = ttk.Treeview(win, columns=("WO ID","Opened","Closed","Status","Equipment","Serial","Description"), show="headings")
        for col in ("WO ID","Opened","Closed","Status","Equipment","Serial","Description"): tree.heading(col, text=col); tree.column(col, width=120)
        tree.grid(row=1, column=0, columnspan=3, pady=6)
        def show():
            try: cid=int(cust_var.get().split(" - ")[0])
            except: messagebox.showerror("Error","Select a valid customer"); return
            rows=get_workorders_by_customer(cid)
            for r in rows:
                equip = f\"{r['make']} {r['model']} (ID {r['equipment_id']})\"
                tree.insert("", "end", values=(r['workorder_id'], r['date_opened'], r['date_closed'] or "", r['status'], equip, r['serial_number'], r['work_description']))
        ttk.Button(win, text="Show", command=show).grid(row=0,column=2,padx=6)
        def export_csv():
            try: cid=int(cust_var.get().split(" - ")[0])
            except: messagebox.showerror("Error","Select a valid customer"); return
            rows=get_workorders_by_customer(cid)
            if not rows: messagebox.showinfo("Info","No records to export"); return
            fn=filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files","*.csv")])
            if not fn: return
            with open(fn, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f); writer.writerow(["WorkOrderID","DateOpened","DateClosed","Status","Make","Model","Serial","Description"])
                for r in rows: writer.writerow([r['workorder_id'], r['date_opened'], r['date_closed'] or "", r['status'], r['make'], r['model'], r['serial_number'], r['work_description']])
            messagebox.showinfo("Saved", f"Exported {len(rows)} rows to {fn}")
        ttk.Button(win, text="Export CSV", command=export_csv).grid(row=2, column=0, pady=6)
        def print_pdf():
            try: cid=int(cust_var.get().split(" - ")[0])
            except: messagebox.showerror("Error","Select a valid customer"); return
            rows=get_workorders_by_customer(cid)
            if not rows: messagebox.showinfo("Info","No records"); return
            fn=filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files","*.pdf")])
            if not fn: return
            # simple PDF: list workorders
            doc = SimpleDocTemplate(fn, pagesize=letter); styles = getSampleStyleSheet(); story=[]
            story.append(Paragraph(f"Work History for Customer {cid}", styles['Title'])); story.append(Spacer(1,6))
            for r in rows: story.append(Paragraph(f"WO {r['workorder_id']}: {r['date_opened']} - {r['status']} - {r['work_description']}", styles['Normal']))
            doc.build(story); messagebox.showinfo("Saved", f"PDF saved to {fn}")
        ttk.Button(win, text="Print PDF", command=print_pdf).grid(row=2, column=1, pady=6)

    def ui_search(self):
        win = tk.Toplevel(self); win.title("Search")
        ttk.Label(win, text="Search term").grid(row=0,column=0,padx=5,pady=3); term_ent=ttk.Entry(win); term_ent.grid(row=0,column=1,padx=5,pady=3)
        tree = ttk.Treeview(win, columns=("Type","ID","Info"), show="headings"); tree.heading("Type", text="Type"); tree.heading("ID", text="ID"); tree.heading("Info", text="Info"); tree.grid(row=1,column=0,columnspan=3)
        def dosearch():
            term = term_ent.get().strip()
            tree.delete(*tree.get_children())
            if not term: return
            custs = search_customers(term)
            for c in custs: tree.insert("", "end", values=("Customer", c['customer_id'], f\"{c['first_name']} {c['last_name']} - {c['phone']}\"))
            wos = search_workorders(term)
            for w in wos: tree.insert("", "end", values=("WorkOrder", w['workorder_id'], f\"{w['date_opened']} - {w['make']} {w['model']} - {w['work_description']}\"))
        ttk.Button(win, text="Search", command=dosearch).grid(row=0,column=2,padx=6)

# helper for fetch_all used in methods above
def fetch_all(query, params=()):
    conn = sqlite3.connect(DB_FILE); conn.row_factory = sqlite3.Row; cur = conn.cursor(); cur.execute(query, params); rows = cur.fetchall(); conn.close(); return rows

if __name__ == "__main__":
    init_db()
    # insert a default technician and part if none exist
    if not fetch_all("SELECT * FROM Technicians LIMIT 1"):
        add_technician("Default Tech")
    if not fetch_all("SELECT * FROM Parts LIMIT 1"):
        add_part("BAT-001", "Replacement Battery", 5, 25.0)
    app = App(); app.mainloop()
