import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import sqlite3
import pandas as pd
from datetime import datetime
import os
import glob
import time

class SurbhiCollectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Surbhi Collection - Premium Retail Engine")
        self.root.geometry("1350x800")  
        self.root.state('zoomed')       
        self.root.configure(bg="#f8f9fa")
        
        self.db_name = "surbhi_shop.db"
        self.init_database()
        
        # Hardware & Global Config
        self.paper_size = tk.StringVar(value="3_inch")
        self.is_return_mode = tk.BooleanVar(value=False)
        
        # --- Multi-Bill Management Architecture ---
        self.draft_slots = {}
        self.active_slot_id = 1
        self.slot_counter = 1
        
        # --- UI STYLING ARCHITECTURE ---
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("TNotebook", background="#f8f9fa", borderwidth=0)
        self.style.configure("TNotebook.Tab", font=("Segoe UI", 11, "bold"), padding=[20, 8], background="#e9ecef", foreground="#495057")
        self.style.map("TNotebook.Tab", background=[("selected", "#ffffff")], foreground=[("selected", "#2c3e50")])
        self.style.configure("TLabelframe", background="#ffffff", bordercolor="#dee2e6", borderwidth=1)
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 10, "bold"), foreground="#2c3e50", background="#ffffff")
        self.style.configure("TLabel", background="#ffffff", font=("Segoe UI", 10), foreground="#495057")
        self.style.configure("TRadiobutton", background="#ffffff", font=("Segoe UI", 10))
        self.style.configure("TCheckbutton", background="#f1f2f6", font=("Segoe UI", 10, "bold"), foreground="#c0392b")
        
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=28, background="#ffffff", fieldbackground="#ffffff", borderwidth=0)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#34495e", foreground="white", relief="flat", anchor="center")
        self.style.map("Treeview.Heading", background=[("active", "#2c3e50")])
        
        self.main_notebook = ttk.Notebook(self.root)
        self.main_notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.billing_tab = ttk.Frame(self.main_notebook, style="TNotebook")
        self.admin_tab = ttk.Frame(self.main_notebook, style="TNotebook")
        
        self.main_notebook.add(self.billing_tab, text="  🛒  Billing Counter  ")
        self.main_notebook.add(self.admin_tab, text="  ⚙️  Admin Dashboard & Central Registry  ")
        
        self.setup_billing_ui()
        self.setup_admin_ui()
        
        # Initialize First Bill Slot
        self.create_new_bill_slot()

    def get_db_connection(self):
        conn = sqlite3.connect(self.db_name, timeout=20.0)
        conn.execute("PRAGMA journal_mode=WAL;")
        return conn

    def init_database(self):
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS inventory (
                item_no TEXT PRIMARY KEY,
                item_name TEXT NOT NULL,
                commission_pct REAL NOT NULL
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS salesmen (
                salesman_no TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                phone TEXT,
                address TEXT,
                base_salary REAL DEFAULT 0
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sales_ledger (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bill_no TEXT,
                sale_date TEXT,
                sale_year INTEGER,
                item_no TEXT,
                qty REAL,
                unit_price REAL,
                raw_total REAL,
                allocated_discount REAL,
                final_item_price REAL,
                calculated_commission REAL,
                salesman_no TEXT,
                payment_mode TEXT,
                customer_phone TEXT,
                customer_name TEXT
            )
        ''')
        conn.commit()
        conn.close()

    def get_archive_paths(self):
        now = datetime.now()
        current_year = now.strftime("Year_%Y")
        current_month_idx = now.strftime("%m")
        month_name = now.strftime("%b")
        
        base_year_dir = os.path.join(os.getcwd(), current_year)
        base_bills_dir = os.path.join(base_year_dir, "1_All_Bills_Logs")
        
        monthly_bills_dir = os.path.join(base_bills_dir, f"{current_month_idx}_{month_name}_Bills")
        
        excel_dir = os.path.join(base_year_dir, "12_Months_Excel_Summaries")
        cust_dir = os.path.join(base_year_dir, "Customer_Profiles_Ledger")
        
        for path in [monthly_bills_dir, excel_dir, cust_dir]:
            if not os.path.exists(path):
                os.makedirs(path)
                
        sales_analysis_path = os.path.join(excel_dir, f"sales analysis_{current_month_idx}_{month_name}.xlsx")
        salesman_data_path = os.path.join(excel_dir, f"salesman data_{current_month_idx}_{month_name}.xlsx")
        cust_excel_path = os.path.join(cust_dir, "Master_Customer_Database.xlsx")
        
        return monthly_bills_dir, sales_analysis_path, salesman_data_path, cust_excel_path

    def format_bill_seq(self, seq_num):
        return f"0{seq_num}" if seq_num < 10 else f"{seq_num}"

    def get_next_bill_number(self, is_return=False):
        now = datetime.now()
        current_month = now.strftime("%m")
        current_year = now.strftime("%y")
        month_year_suffix = f"{current_month}{current_year}"
        prefix = "R_" if is_return else ""
        
        existing_seqs = []

        conn = self.get_db_connection()
        cursor = conn.cursor()
        
        if is_return:
            query_pattern = f"R_%{month_year_suffix}"
        else:
            query_pattern = f"%{month_year_suffix}"

        cursor.execute("SELECT bill_no FROM sales_ledger WHERE bill_no LIKE ?", (query_pattern,))
        for r in cursor.fetchall():
            if r[0]:
                b_str = r[0]
                if is_return and not b_str.startswith("R_"):
                    continue
                if not is_return and b_str.startswith("R_"):
                    continue
                try:
                    clean_str = b_str.replace("R_", "").replace(month_year_suffix, "")
                    existing_seqs.append(int(clean_str))
                except Exception:
                    pass
        conn.close()

        for slot_data in self.draft_slots.values():
            b_no = slot_data.get("bill_no", "")
            if is_return and b_no.startswith("R_") and b_no.endswith(month_year_suffix):
                try:
                    existing_seqs.append(int(b_no.replace("R_", "").replace(month_year_suffix, "")))
                except Exception:
                    pass
            elif not is_return and not b_no.startswith("R_") and b_no.endswith(month_year_suffix):
                try:
                    existing_seqs.append(int(b_no.replace(month_year_suffix, "")))
                except Exception:
                    pass

        next_seq = max(existing_seqs) + 1 if existing_seqs else 1
        return f"{prefix}{self.format_bill_seq(next_seq)}{month_year_suffix}"

    def setup_billing_ui(self):
        header = tk.Frame(self.billing_tab, bg="#2c3e50")
        header.pack(fill="x", ipady=12)
        tk.Label(header, text="SURBHI COLLECTION", font=("Segoe UI", 24, "bold"), fg="#ffffff", bg="#2c3e50").pack()
        
        draft_bar = tk.Frame(self.billing_tab, bg="#34495e")
        draft_bar.pack(fill="x", padx=10, pady=2)
        
        tk.Label(draft_bar, text="Active Billing Counter Slots:", font=("Segoe UI", 10, "bold"), fg="white", bg="#34495e").pack(side="left", padx=10, pady=5)
        
        self.slots_frame = tk.Frame(draft_bar, bg="#34495e")
        self.slots_frame.pack(side="left", fill="x", expand=True)
        
        tk.Button(draft_bar, text="➕ New Bill Slot", font=("Segoe UI", 9, "bold"), bg="#27ae60", fg="white", relief="flat", padx=10, command=self.create_new_bill_slot).pack(side="right", padx=10, pady=5)
        
        main_body = tk.Frame(self.billing_tab, bg="#f8f9fa")
        main_body.pack(fill="both", expand=True, padx=10, pady=5)
        
        left_panel = tk.Frame(main_body, width=340, bg="#f8f9fa")
        left_panel.pack(side="left", fill="y", padx=5)
        
        settings_frame = ttk.LabelFrame(left_panel, text=" Hardware & Settings ", padding=10)
        settings_frame.pack(fill="x", pady=5)
        ttk.Label(settings_frame, text="Thermal Width:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Radiobutton(settings_frame, text="2 Inch", variable=self.paper_size, value="2_inch").grid(row=0, column=1, padx=10, pady=5)
        ttk.Radiobutton(settings_frame, text="3 Inch", variable=self.paper_size, value="3_inch").grid(row=0, column=2, padx=10, pady=5)
        
        cust_frame = ttk.LabelFrame(left_panel, text=" Customer Profile (Optional) ", padding=10)
        cust_frame.pack(fill="x", pady=5)
        
        ttk.Label(cust_frame, text="Customer Name:").pack(anchor="w", pady=2)
        self.cust_name_entry = tk.Entry(cust_frame, font=("Segoe UI", 11), bd=1, relief="solid", bg="#ffffff")
        self.cust_name_entry.pack(fill="x", pady=3, ipady=3)
        self.cust_name_entry.bind("<KeyRelease>", self.save_current_slot_state)
        
        ttk.Label(cust_frame, text="Mobile Number (Enter '-' if empty):").pack(anchor="w", pady=2)
        self.cust_phone_entry = tk.Entry(cust_frame, font=("Segoe UI", 11), bd=1, relief="solid", bg="#ffffff")
        self.cust_phone_entry.pack(fill="x", pady=3, ipady=3)
        self.cust_phone_entry.bind("<KeyRelease>", self.save_current_slot_state)

        right_panel = tk.Frame(main_body, bg="#ffffff", bd=1, relief="solid")
        right_panel.pack(side="right", fill="both", expand=True, padx=5)
        
        status_bar = tk.Frame(right_panel, bg="#ffffff")
        status_bar.pack(fill="x", padx=10, pady=8)
        self.bill_lbl = tk.Label(status_bar, text="INVOICE LOG NUMBER: # ", font=("Segoe UI", 13, "bold"), fg="#e74c3c", bg="#ffffff")
        self.bill_lbl.pack(side="left")
        
        entry_frame = tk.Frame(right_panel, bg="#f1f2f6")
        entry_frame.pack(fill="x", padx=10, pady=5, ipady=8)
        
        def make_entry_box(parent, label, width):
            box = tk.Frame(parent, bg="#f1f2f6")
            tk.Label(box, text=label, font=("Segoe UI", 9, "bold"), bg="#f1f2f6", fg="#2c3e50").pack(anchor="w", padx=5, pady=2)
            ent = tk.Entry(box, font=("Segoe UI", 11), width=width, bd=1, relief="solid", bg="#ffffff")
            ent.pack(pady=3, ipady=3, padx=5)
            return box, ent

        b0, self.salesman_entry = make_entry_box(entry_frame, "Salesman No:", 10)
        b0.grid(row=0, column=0, padx=5, pady=5)

        b1, self.item_search_entry = make_entry_box(entry_frame, "Item No / Search Name:", 20)
        b1.grid(row=0, column=1, padx=5, pady=5)
        
        self.resolved_name_lbl = tk.Label(entry_frame, text="", font=("Segoe UI", 10, "bold"), fg="#27ae60", bg="#f1f2f6", width=16, anchor="w")
        self.resolved_name_lbl.grid(row=0, column=2, padx=5, pady=15)
        
        b3, self.qty_entry = make_entry_box(entry_frame, "Qty:", 6)
        b3.grid(row=0, column=3, padx=5, pady=5)

        b2, self.price_entry = make_entry_box(entry_frame, "Original Price (₹):", 11)
        b2.grid(row=0, column=4, padx=5, pady=5)
        
        # Checkbox for Returned Mode
        self.return_chk = ttk.Checkbutton(entry_frame, text="↩ RETURN", variable=self.is_return_mode, command=self.handle_return_toggle, style="TCheckbutton")
        self.return_chk.grid(row=0, column=5, padx=8, pady=5)
        
        tk.Button(entry_frame, text="[+] Add Item", font=("Segoe UI", 10, "bold"), bg="#27ae60", fg="white", activebackground="#2196f3", activeforeground="white", relief="flat", padx=12, command=self.add_item_to_grid).grid(row=0, column=6, padx=10, pady=10)

        # Keyboard Navigation Keybindings
        self.cust_name_entry.bind("<Return>", lambda e: self.cust_phone_entry.focus())
        self.cust_phone_entry.bind("<Return>", lambda e: self.salesman_entry.focus())
        self.cust_phone_entry.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.cust_name_entry))

        self.salesman_entry.bind("<Return>", lambda e: self.item_search_entry.focus())
        self.salesman_entry.bind("<Right>", lambda e: self.focus_if_at_end(e, self.item_search_entry))
        self.salesman_entry.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.cust_phone_entry))

        self.item_search_entry.bind("<KeyRelease>", self.handle_realtime_item_search)
        self.item_search_entry.bind("<Return>", lambda e: self.qty_entry.focus())
        self.item_search_entry.bind("<Right>", lambda e: self.focus_if_at_end(e, self.qty_entry))
        self.item_search_entry.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.salesman_entry))

        self.qty_entry.bind("<Return>", lambda e: self.price_entry.focus())
        self.qty_entry.bind("<Right>", lambda e: self.focus_if_at_end(e, self.price_entry))
        self.qty_entry.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.item_search_entry))

        self.price_entry.bind("<Return>", lambda e: self.add_item_to_grid())
        self.price_entry.bind("<Right>", lambda e: self.focus_if_at_end(e, action=self.add_item_to_grid))
        self.price_entry.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.qty_entry))
        
        grid_container = tk.Frame(right_panel, bg="#ffffff")
        grid_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        columns = ("sm_no", "item_no", "item_name", "qty", "price", "disc_rate", "total")
        self.tree = ttk.Treeview(grid_container, columns=columns, show="headings")
        self.tree.heading("sm_no", text="SM NO")
        self.tree.heading("item_no", text="ITEM NO")
        self.tree.heading("item_name", text="ITEM NAME")
        self.tree.heading("qty", text="QTY")
        self.tree.heading("price", text="ORIGINAL PRICE (₹)")
        self.tree.heading("disc_rate", text="DISCOUNTED RATE (₹)")
        self.tree.heading("total", text="FINAL TOTAL")
        
        self.tree.column("sm_no", width=80, anchor="center")
        self.tree.column("item_no", width=100, anchor="center")
        self.tree.column("item_name", width=220, anchor="center")
        self.tree.column("qty", width=80, anchor="center")
        self.tree.column("price", width=120, anchor="center")
        self.tree.column("disc_rate", width=140, anchor="center")
        self.tree.column("total", width=140, anchor="center")
        
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.handle_double_click_edit)
        self.tree.bind("<Delete>", self.delete_selected_cart_item)
        
        sb = ttk.Scrollbar(grid_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        
        bottom_row = tk.Frame(right_panel, bg="#f8f9fa")
        bottom_row.pack(fill="x", side="bottom", ipady=8)
        
        self.summary_lbl = tk.Label(bottom_row, text="Unique Items: 0  |  Total Qty: 0\nGrand Total: ₹0.00  |  Total Discount (Saved): ₹0.00", font=("Segoe UI", 12, "bold"), fg="#2c3e50", bg="#f8f9fa", justify="left")
        self.summary_lbl.pack(side="left", padx=15, pady=5)
        
        disc_frame = tk.Frame(bottom_row, bg="#f8f9fa")
        disc_frame.pack(side="right", padx=10)
        tk.Label(disc_frame, text="Flat Bill Disc %:", font=("Segoe UI", 10, "bold"), bg="#f8f9fa", fg="#2c3e50").pack(anchor="w")
        self.global_discount_entry = tk.Entry(disc_frame, font=("Segoe UI", 12, "bold"), width=8, bd=1, relief="solid", justify="center")
        self.global_discount_entry.pack(ipady=4)
        self.global_discount_entry.insert(0, "10") 
        self.global_discount_entry.bind("<KeyRelease>", lambda e: self.calculate_bill_totals())
        
        tk.Button(bottom_row, text="PRINT & SAVE INVOICE", font=("Segoe UI", 11, "bold"), bg="#2e7d32", fg="white", activebackground="#1b5e20", relief="flat", padx=20, pady=8, command=self.commit_and_print_invoice).pack(side="right", padx=10, pady=5)
        tk.Button(bottom_row, text="Reset Slot", font=("Segoe UI", 10), bg="#7f8c8d", fg="white", activebackground="#616161", relief="flat", padx=5, pady=8, command=self.clear_complete_billing_grid).pack(side="right", padx=5, pady=5)

        hist_bar = tk.Frame(self.billing_tab, bg="#e2e8f0")
        hist_bar.pack(fill="x", padx=10, pady=(0, 5))
        tk.Label(hist_bar, text="📜 Recent Bills Adjustment:", font=("Segoe UI", 9, "bold"), bg="#e2e8f0", fg="#2c3e50").pack(side="left", padx=8, pady=4)
        
        tk.Button(hist_bar, text="🔍 Look Up & Edit Completed Bill", font=("Segoe UI", 9, "bold"), bg="#34495e", fg="white", relief="flat", padx=10, command=self.open_recent_bills_lookup).pack(side="left", padx=5, pady=2)

    def handle_return_toggle(self):
        is_ret = self.is_return_mode.get()
        new_b_no = self.get_next_bill_number(is_return=is_ret)
        self.draft_slots[self.active_slot_id]["bill_no"] = new_b_no
        self.draft_slots[self.active_slot_id]["is_return"] = is_ret
        self.bill_lbl.config(text=f"INVOICE LOG NUMBER: # {new_b_no}")

    def focus_if_at_beginning(self, event, target_widget):
        if event.widget.index(tk.INSERT) == 0:
            target_widget.focus()
            target_widget.icursor(tk.END)
            return "break"

    def focus_if_at_end(self, event, target_widget=None, action=None):
        if event.widget.index(tk.INSERT) == len(event.widget.get()):
            if action:
                action()
            elif target_widget:
                target_widget.focus()
                target_widget.icursor(0)
            return "break"

    def setup_admin_ui(self):
        admin_main = tk.Frame(self.admin_tab, bg="#f8f9fa")
        admin_main.pack(fill="both", expand=True, padx=10, pady=10)
        
        entry_side = tk.Frame(admin_main, bg="#f8f9fa")
        entry_side.pack(side="left", fill="y", padx=5)
        
        prod_lf = ttk.LabelFrame(entry_side, text=" Stock Inventory Item Manager ", padding=15)
        prod_lf.pack(fill="x", pady=5)
        
        tk.Label(prod_lf, text="Item Code (Target):").pack(anchor="w", pady=2)
        self.adm_item_code = tk.Entry(prod_lf, font=("Segoe UI", 11), bd=1, relief="solid", width=30)
        self.adm_item_code.pack(fill="x", pady=4, ipady=3)
        
        tk.Label(prod_lf, text="Item Name / Type:").pack(anchor="w", pady=2)
        self.adm_item_name = tk.Entry(prod_lf, font=("Segoe UI", 11), bd=1, relief="solid", width=30)
        self.adm_item_name.pack(fill="x", pady=4, ipady=3)
        
        tk.Label(prod_lf, text="Salesman Comm %:").pack(anchor="w", pady=2)
        self.adm_item_comm = tk.Entry(prod_lf, font=("Segoe UI", 11), bd=1, relief="solid", width=30)
        self.adm_item_comm.pack(fill="x", pady=4, ipady=3)
        
        self.adm_item_code.bind("<Return>", lambda e: self.adm_item_name.focus())
        self.adm_item_name.bind("<Return>", lambda e: self.adm_item_comm.focus())
        self.adm_item_comm.bind("<Return>", lambda e: self.admin_save_item())

        self.adm_item_name.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.adm_item_code))
        self.adm_item_comm.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.adm_item_name))

        prod_btn_frame = tk.Frame(prod_lf, bg="#ffffff")
        prod_btn_frame.pack(pady=12)
        tk.Button(prod_btn_frame, text="✨ Add Item", bg="#2c3e50", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.admin_save_item, padx=10, pady=4).pack(side="left", padx=5)
        tk.Button(prod_btn_frame, text="✏️ Update", bg="#e67e22", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.admin_update_item, padx=10, pady=4).pack(side="left", padx=5)
        tk.Button(prod_btn_frame, text="🗑️ Delete", bg="#c0392b", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.admin_delete_item, padx=10, pady=4).pack(side="left", padx=5)
        
        sm_lf = ttk.LabelFrame(entry_side, text=" Salesmen Registry Manager ", padding=15)
        sm_lf.pack(fill="x", pady=15)
        
        tk.Label(sm_lf, text="Salesman No/ID:").pack(anchor="w", pady=2)
        self.adm_sm_id = tk.Entry(sm_lf, font=("Segoe UI", 11), bd=1, relief="solid", width=30)
        self.adm_sm_id.pack(fill="x", pady=4, ipady=3)
        
        tk.Label(sm_lf, text="Full Name:").pack(anchor="w", pady=2)
        self.adm_sm_name = tk.Entry(sm_lf, font=("Segoe UI", 11), bd=1, relief="solid", width=30)
        self.adm_sm_name.pack(fill="x", pady=4, ipady=3)
        
        self.adm_sm_id.bind("<Return>", lambda e: self.adm_sm_name.focus())
        self.adm_sm_name.bind("<Return>", lambda e: self.admin_save_salesman())
        self.adm_sm_name.bind("<Left>", lambda e: self.focus_if_at_beginning(e, self.adm_sm_id))

        sm_btn_frame = tk.Frame(sm_lf, bg="#ffffff")
        sm_btn_frame.pack(pady=12)
        tk.Button(sm_btn_frame, text="👤 Register", bg="#2980b9", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.admin_save_salesman, padx=10, pady=4).pack(side="left", padx=5)
        tk.Button(sm_btn_frame, text="✏️ Update Profile", bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.admin_update_salesman, padx=10, pady=4).pack(side="left", padx=5)
        tk.Button(sm_btn_frame, text="🗑️ Remove Staff", bg="#c0392b", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.admin_delete_salesman, padx=10, pady=4).pack(side="left", padx=5)

        view_side = tk.Frame(admin_main, bg="#f8f9fa")
        view_side.pack(side="right", fill="both", expand=True, padx=5)
        
        lbl_ctrls = tk.Frame(view_side, bg="#f8f9fa")
        lbl_ctrls.pack(fill="x", pady=5)
        tk.Label(lbl_ctrls, text="📊 Central Core Data Registries", font=("Segoe UI", 12, "bold"), fg="#2c3e50", bg="#f8f9fa").pack(side="left")
        tk.Button(lbl_ctrls, text="🔄 Refresh Systems", bg="#2c3e50", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", command=self.refresh_admin_views, padx=12, pady=4).pack(side="right", padx=2)
        
        self.admin_views_nb = ttk.Notebook(view_side)
        self.admin_views_nb.pack(fill="both", expand=True, pady=5)
        
        self.p_inventory_frame = ttk.Frame(self.admin_views_nb)
        self.p_staff_frame = ttk.Frame(self.admin_views_nb)
        
        self.admin_views_nb.add(self.p_inventory_frame, text=" Products Inventory ")
        self.admin_views_nb.add(self.p_staff_frame, text=" Registered Staff ")
        
        self.build_tabular_views()
        self.refresh_admin_views()

    def build_tabular_views(self):
        self.tree_inv = ttk.Treeview(self.p_inventory_frame, columns=("code", "name", "comm"), show="headings")
        self.tree_inv.heading("code", text="ITEM CODE", anchor="center")
        self.tree_inv.heading("name", text="PRODUCT TYPE / NAME", anchor="center")
        self.tree_inv.heading("comm", text="SALESMAN COMMISSION %", anchor="center")
        self.tree_inv.column("code", anchor="center")
        self.tree_inv.column("name", anchor="center")
        self.tree_inv.column("comm", anchor="center")
        self.tree_inv.pack(fill="both", expand=True)
        
        self.tree_staff = ttk.Treeview(self.p_staff_frame, columns=("id", "name"), show="headings")
        self.tree_staff.heading("id", text="SALESMAN ID", anchor="center")
        self.tree_staff.heading("name", text="STAFF MEMBER FULL NAME", anchor="center")
        self.tree_staff.column("id", anchor="center")
        self.tree_staff.column("name", anchor="center")
        self.tree_staff.pack(fill="both", expand=True)

    def create_new_bill_slot(self):
        slot_id = self.slot_counter
        self.slot_counter += 1
        
        is_ret = self.is_return_mode.get()
        new_bill_no = self.get_next_bill_number(is_return=is_ret)
        
        self.draft_slots[slot_id] = {
            "cart_items": [],
            "cust_name": "",
            "cust_phone": "",
            "disc_pct": "10",
            "bill_no": new_bill_no,
            "is_return": is_ret
        }
        self.switch_to_slot(slot_id)

    def switch_to_slot(self, slot_id):
        if self.active_slot_id in self.draft_slots:
            self.save_current_slot_state()
            
        self.active_slot_id = slot_id
        slot_data = self.draft_slots[slot_id]
        
        self.cust_name_entry.delete(0, tk.END)
        self.cust_name_entry.insert(0, slot_data["cust_name"])
        
        self.cust_phone_entry.delete(0, tk.END)
        self.cust_phone_entry.insert(0, slot_data["cust_phone"])
        
        self.global_discount_entry.delete(0, tk.END)
        self.global_discount_entry.insert(0, slot_data["disc_pct"])
        
        self.is_return_mode.set(slot_data.get("is_return", False))
        self.bill_lbl.config(text=f"INVOICE LOG NUMBER: # {slot_data['bill_no']}")
        
        self.calculate_bill_totals()
        self.render_slot_buttons()
        self.cust_name_entry.focus()

    def save_current_slot_state(self, event=None):
        if self.active_slot_id in self.draft_slots:
            self.draft_slots[self.active_slot_id]["cust_name"] = self.cust_name_entry.get().strip()
            self.draft_slots[self.active_slot_id]["cust_phone"] = self.cust_phone_entry.get().strip()
            self.draft_slots[self.active_slot_id]["disc_pct"] = self.global_discount_entry.get().strip()
            self.draft_slots[self.active_slot_id]["is_return"] = self.is_return_mode.get()

    def render_slot_buttons(self):
        for widget in self.slots_frame.winfo_children():
            widget.destroy()
            
        for s_id in sorted(self.draft_slots.keys()):
            b_text = f" Bill #{s_id} " if s_id != self.active_slot_id else f" ▶ Bill #{s_id} ◀ "
            b_bg = "#2980b9" if s_id == self.active_slot_id else "#7f8c8d"
            
            slot_btn = tk.Button(self.slots_frame, text=b_text, bg=b_bg, fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, command=lambda sid=s_id: self.switch_to_slot(sid))
            slot_btn.pack(side="left", padx=3, pady=5)
            
            if len(self.draft_slots) > 1:
                close_btn = tk.Button(self.slots_frame, text="✕", bg="#c0392b", fg="white", font=("Segoe UI", 8, "bold"), relief="flat", command=lambda sid=s_id: self.close_slot(sid))
                close_btn.pack(side="left", padx=(0, 6), pady=5)

    def close_slot(self, slot_id):
        if len(self.draft_slots) <= 1:
            return
        was_active = (self.active_slot_id == slot_id)
        del self.draft_slots[slot_id]
        if was_active:
            self.switch_to_slot(list(self.draft_slots.keys())[0])
        else:
            self.switch_to_slot(self.active_slot_id)

    @property
    def cart_items(self):
        return self.draft_slots[self.active_slot_id]["cart_items"]

    @cart_items.setter
    def cart_items(self, value):
        self.draft_slots[self.active_slot_id]["cart_items"] = value

    @property
    def current_bill_no(self):
        return self.draft_slots[self.active_slot_id]["bill_no"]

    def admin_save_item(self):
        i_code, name, comm = self.adm_item_code.get().strip(), self.adm_item_name.get().strip(), self.adm_item_comm.get().strip()
        if not i_code or not name or not comm:
            messagebox.showerror("Error", "Please fill all item entry fields.")
            return
        try:
            conn = self.get_db_connection(); cursor = conn.cursor()
            cursor.execute("INSERT INTO inventory VALUES (?, ?, ?)", (i_code, name, float(comm)))
            conn.commit(); conn.close()
            messagebox.showinfo("Success", f"'{name}' successfully configured!")
            self.adm_item_code.delete(0, tk.END)
            self.adm_item_name.delete(0, tk.END)
            self.adm_item_comm.delete(0, tk.END)
            self.adm_item_code.focus()
            self.refresh_admin_views()
        except Exception as e: messagebox.showerror("Database Error", f"Failed: {e}")

    def admin_update_item(self):
        i_code, name, comm = self.adm_item_code.get().strip(), self.adm_item_name.get().strip(), self.adm_item_comm.get().strip()
        if not i_code:
            messagebox.showerror("Error", "Specify structural Target Item Code.")
            return
        conn = self.get_db_connection(); cursor = conn.cursor()
        if name: cursor.execute("UPDATE inventory SET item_name = ? WHERE item_no = ?", (name, i_code))
        if comm: cursor.execute("UPDATE inventory SET commission_pct = ? WHERE item_no = ?", (float(comm), i_code))
        conn.commit(); conn.close()
        messagebox.showinfo("Success", "Product registry entry altered.")
        self.refresh_admin_views()

    def admin_delete_item(self):
        i_code = self.adm_item_code.get().strip()
        if not i_code: return
        if messagebox.askyesno("Confirm", "Permanently clear item from catalog?"):
            conn = self.get_db_connection(); cursor = conn.cursor()
            cursor.execute("DELETE FROM inventory WHERE item_no = ?", (i_code,))
            conn.commit(); conn.close()
            self.refresh_admin_views()

    def admin_save_salesman(self):
        s_id, name = self.adm_sm_id.get().strip(), self.adm_sm_name.get().strip()
        if not s_id or not name:
            messagebox.showerror("Error", "Staff ID and Full Name are strictly mandatory.")
            return
        try:
            conn = self.get_db_connection(); cursor = conn.cursor()
            cursor.execute("INSERT INTO salesmen VALUES (?, ?, '', '', 0.0)", (s_id, name))
            conn.commit(); conn.close()
            messagebox.showinfo("Success", f"Salesman '{name}' fully registered.")
            self.adm_sm_id.delete(0, tk.END)
            self.adm_sm_name.delete(0, tk.END)
            self.adm_sm_id.focus()
            self.refresh_admin_views()
        except Exception as e: messagebox.showerror("Database Error", f"Failed: {e}")

    def admin_update_salesman(self):
        s_id, name = self.adm_sm_id.get().strip(), self.adm_sm_name.get().strip()
        if not s_id or not name: return
        conn = self.get_db_connection(); cursor = conn.cursor()
        cursor.execute("UPDATE salesmen SET name = ? WHERE salesman_no = ?", (name, s_id))
        conn.commit(); conn.close()
        messagebox.showinfo("Success", "Profile record modified successfully.")
        self.refresh_admin_views()

    def admin_delete_salesman(self):
        s_id = self.adm_sm_id.get().strip()
        if not s_id: return
        if messagebox.askyesno("Confirm Removal", "Delete staff profile completely?"):
            conn = self.get_db_connection(); cursor = conn.cursor()
            cursor.execute("DELETE FROM salesmen WHERE salesman_no = ?", (s_id,))
            conn.commit(); conn.close()
            self.refresh_admin_views()

    def refresh_admin_views(self):
        conn = self.get_db_connection()
        for item in self.tree_inv.get_children(): self.tree_inv.delete(item)
        for r in conn.execute("SELECT item_no, item_name, commission_pct FROM inventory").fetchall():
            self.tree_inv.insert("", "end", values=r)
            
        for item in self.tree_staff.get_children(): self.tree_staff.delete(item)
        for r in conn.execute("SELECT salesman_no, name FROM salesmen").fetchall():
            self.tree_staff.insert("", "end", values=r)
        conn.close()

    def handle_realtime_item_search(self, event):
        query = self.item_search_entry.get().strip()
        if not query:
            self.resolved_name_lbl.config(text="")
            return
        conn = self.get_db_connection(); cursor = conn.cursor()
        cursor.execute("SELECT item_no, item_name FROM inventory WHERE item_no = ? LIMIT 1", (query,))
        row = cursor.fetchone()
        if row:
            self.resolved_name_lbl.config(text=row[1], fg="#2ed573")
        else:
            cursor.execute("SELECT item_no, item_name FROM inventory WHERE item_name LIKE ? LIMIT 1", (f"%{query}%",))
            row_by_name = cursor.fetchone()
            if row_by_name: self.resolved_name_lbl.config(text=f"↳ {row_by_name[0]}", fg="#2980b9")
            else: self.resolved_name_lbl.config(text="Not Registered", fg="#ff4757")
        conn.close()

    def add_item_to_grid(self):
        query = self.item_search_entry.get().strip()
        sm_no = self.salesman_entry.get().strip()
        
        conn = self.get_db_connection(); cursor = conn.cursor()
        cursor.execute("SELECT item_no, item_name, commission_pct FROM inventory WHERE item_no = ? OR item_name LIKE ? LIMIT 1", (query, f"%{query}%"))
        matched_row = cursor.fetchone()
        cursor.execute("SELECT name FROM salesmen WHERE salesman_no = ?", (sm_no,))
        sm_exists = cursor.fetchone(); conn.close()
        
        if not matched_row:
            messagebox.showerror("Error", "Product configuration missing in registry!")
            return
        if not sm_exists:
            messagebox.showerror("Error", f"Salesman ID '{sm_no}' is not registered!")
            return
            
        item_code, item_name, comm_pct = matched_row
        try:
            price = float(self.price_entry.get())
            raw_qty = float(self.qty_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid pricing or item volume amounts entered.")
            return

        qty = -abs(raw_qty) if self.is_return_mode.get() else abs(raw_qty)
            
        try:
            disc_pct = float(self.global_discount_entry.get() if self.global_discount_entry.get() else 0)
        except ValueError: 
            disc_pct = 0.0

        discounted_rate_per_unit = price * (1 - (disc_pct / 100.0))
        final_gross_total = discounted_rate_per_unit * qty
        
        record = {
            "item_no": item_code, "item_name": item_name, "qty": qty, "price": price, 
            "disc_rate": discounted_rate_per_unit, "total": final_gross_total, 
            "comm_pct": comm_pct, "salesman_no": sm_no
        }
        self.cart_items.append(record)
        
        self.calculate_bill_totals()
        self.clear_item_input_fields(keep_customer_and_salesman=True)

    def delete_selected_cart_item(self, event=None):
        selected_uid = self.tree.selection()
        if not selected_uid:
            return
            
        row_vals = self.tree.item(selected_uid[0], "values")
        sm_no = row_vals[0]
        item_no = row_vals[1]
        qty = float(row_vals[3])
        price = float(row_vals[4])
        
        for idx, item in enumerate(self.cart_items):
            if (item["salesman_no"] == sm_no and 
                item["item_no"] == item_no and 
                item["qty"] == qty and 
                item["price"] == price):
                self.cart_items.pop(idx)
                break
                
        self.tree.delete(selected_uid[0])
        self.calculate_bill_totals()

    def handle_double_click_edit(self, event):
        selected_uid = self.tree.selection()
        if not selected_uid: return
            
        row_vals = self.tree.item(selected_uid[0], "values")
        sm_to_edit = row_vals[0]
        item_no_to_edit = row_vals[1]
        qty_to_edit = float(row_vals[3])
        price_to_edit = float(row_vals[4])
        
        self.clear_item_input_fields(keep_customer_and_salesman=True)
        
        self.salesman_entry.delete(0, tk.END)
        self.salesman_entry.insert(0, sm_to_edit)
        self.item_search_entry.insert(0, item_no_to_edit)
        self.qty_entry.insert(0, f"{abs(qty_to_edit):g}")
        self.price_entry.insert(0, f"{price_to_edit:.2f}")
        
        self.handle_realtime_item_search(None)
        
        for idx, item in enumerate(self.cart_items):
            if item["item_no"] == item_no_to_edit and item["qty"] == qty_to_edit and item["price"] == price_to_edit:
                self.cart_items.pop(idx)
                break
                
        self.tree.delete(selected_uid[0])
        self.calculate_bill_totals()
        self.price_entry.focus()

    def calculate_bill_totals(self):
        unique_count = len(self.cart_items)
        total_units = sum(float(i["qty"]) for i in self.cart_items)
        
        try:
            disc_pct = float(self.global_discount_entry.get() if self.global_discount_entry.get() else 0)
        except ValueError: 
            disc_pct = 0.0

        gross_undiscounted_amt = 0.0
        total_discount_sum = 0.0
        grand_total = 0.0
        
        for item in self.cart_items:
            raw_item_total = item["price"] * item["qty"]
            item_discount = raw_item_total * (disc_pct / 100.0)
            
            item["disc_rate"] = item["price"] * (1 - (disc_pct / 100.0))
            item["total"] = raw_item_total - item_discount
            
            gross_undiscounted_amt += raw_item_total
            total_discount_sum += item_discount
            grand_total += item["total"]
            
        for row in self.tree.get_children(): self.tree.delete(row)
        for i in self.cart_items:
            self.tree.insert("", "end", values=(
                i["salesman_no"], i["item_no"], i["item_name"], f"{i['qty']:g}", f"{i['price']:.2f}", f"{i['disc_rate']:.2f}", f"{i['total']:.2f}"
            ))
        
        self.summary_lbl.config(
            text=f"Unique Items: {unique_count}  |  Total Qty: {total_units:g}\nGrand Total: ₹{gross_undiscounted_amt:.2f}  |  Total Discount (Saved): ₹{total_discount_sum:.2f}"
        )

    def clear_item_input_fields(self, keep_customer_and_salesman=False):
        # Clear item entry fields
        self.item_search_entry.delete(0, tk.END)
        self.price_entry.delete(0, tk.END)
        self.qty_entry.delete(0, tk.END)
        self.resolved_name_lbl.config(text="")
        
        if keep_customer_and_salesman:
            # Clear Salesman No field and return focus to Salesman No
            self.salesman_entry.delete(0, tk.END)
            self.salesman_entry.focus()
        else:
            self.salesman_entry.delete(0, tk.END)
            self.cust_name_entry.delete(0, tk.END)
            self.cust_phone_entry.delete(0, tk.END)
            
            if self.active_slot_id in self.draft_slots:
                self.draft_slots[self.active_slot_id]["cust_name"] = ""
                self.draft_slots[self.active_slot_id]["cust_phone"] = ""
                
            self.cust_name_entry.focus()

    def commit_and_print_invoice(self):
        if not self.cart_items: return
        
        committed_bill_no = self.current_bill_no
        c_phone = self.cust_phone_entry.get().strip()
        c_name = " ".join(self.cust_name_entry.get().split())
        
        if not c_name: c_name = "Guest Customer"
        if not c_phone: c_phone = "-"
        
        now_dt = datetime.now()
        date_str = now_dt.strftime("%Y-%m-%d")
        year_val = now_dt.year
        
        try:
            disc_pct = float(self.global_discount_entry.get() if self.global_discount_entry.get() else 0)
        except ValueError: 
            disc_pct = 0.0
            
        gross_total = sum((float(i["price"]) * float(i["qty"])) for i in self.cart_items)
        saved = gross_total * (disc_pct / 100.0)
        original_grand_total = round(gross_total - saved)
        
        conn = self.get_db_connection(); cursor = conn.cursor()
        for item in self.cart_items:
            raw_total = item["price"] * item["qty"]
            allocated_item_disc = raw_total * (disc_pct / 100.0)
            final_item_price = raw_total - allocated_item_disc
            calculated_item_comm = final_item_price * (item["comm_pct"] / 100.0)
            
            cursor.execute('''
                INSERT INTO sales_ledger (
                    bill_no, sale_date, sale_year, item_no, qty, unit_price, raw_total, 
                    allocated_discount, final_item_price, calculated_commission, salesman_no, payment_mode, customer_phone, customer_name
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                committed_bill_no, date_str, year_val, item["item_no"], item["qty"], item["price"], raw_total,
                allocated_item_disc, final_item_price, calculated_item_comm, item["salesman_no"], "N/A", c_phone, c_name
            ))
        conn.commit(); conn.close()
        
        monthly_bills_dir, sales_analysis_path, salesman_data_path, customer_ledger_path = self.get_archive_paths()
        
        receipt_text = self.build_receipt_string(c_name, c_phone, disc_pct, bill_no=committed_bill_no)
        invoice_path = os.path.join(monthly_bills_dir, f"Bill_{committed_bill_no}.txt")
        with open(invoice_path, "w", encoding="utf-8") as f: f.write(receipt_text)
            
        self.generate_excel_reports(sales_analysis_path, salesman_data_path)
        self.update_customer_ledger_excel(customer_ledger_path)
            
        cleaned_receipt = receipt_text.replace("₹", "Rs.").replace("<bold>", "").replace("</bold>", "")
        
        first_copy = cleaned_receipt + ("\n" * 9)
        second_copy = cleaned_receipt + ("\n" * 9)

        try:
            import win32print
            printer_name = win32print.GetDefaultPrinter()
            
            def send_raw_job(job_title, payload_text):
                hPrinter = win32print.OpenPrinter(printer_name)
                try:
                    hJob = win32print.StartDocPrinter(hPrinter, 1, (job_title, None, "RAW"))
                    try:
                        win32print.StartPagePrinter(hPrinter)
                        win32print.WritePrinter(hPrinter, payload_text.encode("ascii", "replace"))
                        win32print.EndPagePrinter(hPrinter)
                    finally:
                        win32print.EndDocPrinter(hPrinter)
                finally:
                    win32print.ClosePrinter(hPrinter)

            send_raw_job(f"Bill_{committed_bill_no}_Copy1", first_copy)
            time.sleep(5)
            send_raw_job(f"Bill_{committed_bill_no}_Copy2", second_copy)

        except Exception as print_err:
            messagebox.showwarning(
                "Printer Error", 
                f"Failed to issue command to default printer:\n{print_err}\n\nPlease check if 'pypiwin32' is installed and your thermal printer is online."
            )

        # Reset return mode state back to False after saving the bill
        self.is_return_mode.set(False)
            
        if len(self.draft_slots) > 1:
            del self.draft_slots[self.active_slot_id]
            next_slot_id = list(self.draft_slots.keys())[0]
            self.switch_to_slot(next_slot_id)
        else:
            is_ret = False
            self.draft_slots[self.active_slot_id] = {
                "cart_items": [],
                "cust_name": "",
                "cust_phone": "",
                "disc_pct": "10",
                "bill_no": self.get_next_bill_number(is_return=is_ret),
                "is_return": is_ret
            }
            self.switch_to_slot(self.active_slot_id)
            
        self.clear_item_input_fields(keep_customer_and_salesman=False)
        self.refresh_admin_views()
        
        self.prompt_post_print_settlement(committed_bill_no, original_grand_total, sales_analysis_path, salesman_data_path)
        
    def prompt_post_print_settlement(self, bill_no, original_amount, sales_analysis_path, salesman_data_path):
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Post-Print Settlement - Bill #{bill_no}")
        dlg.geometry("440x240")
        dlg.configure(bg="#ffffff")
        dlg.resizable(False, False)
        
        tk.Label(dlg, text=f"Bill #{bill_no} Saved Successfully!", font=("Segoe UI", 12, "bold"), fg="#27ae60", bg="#ffffff").pack(pady=(15, 5))
        tk.Label(dlg, text=f"Original Printed Total: ₹{original_amount:.2f}", font=("Segoe UI", 10, "bold"), fg="#2c3e50", bg="#ffffff").pack(pady=2)
        tk.Label(dlg, text="Did customer pay less or alternate amount?", font=("Segoe UI", 9), fg="#7f8c8d", bg="#ffffff").pack(pady=2)
        
        amt_frame = tk.Frame(dlg, bg="#ffffff")
        amt_frame.pack(pady=10)
        tk.Label(amt_frame, text="Actual Cash Received: ₹", font=("Segoe UI", 10, "bold"), bg="#ffffff").pack(side="left")
        
        ent = tk.Entry(amt_frame, font=("Segoe UI", 11, "bold"), width=10, bd=1, relief="solid")
        ent.pack(side="left", padx=5)
        ent.insert(0, f"{original_amount}")
        
        def save_adjustment():
            try:
                actual_given = float(ent.get().strip())
                if actual_given != original_amount:
                    self.apply_bill_adjustment(bill_no, original_amount, actual_given, sales_analysis_path, salesman_data_path)
                dlg.destroy()
            except ValueError:
                messagebox.showerror("Error", "Enter valid amount.", parent=dlg)
                
        btn_frame = tk.Frame(dlg, bg="#ffffff")
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text="Update Received Amount", font=("Segoe UI", 9, "bold"), bg="#2980b9", fg="white", relief="flat", padx=10, pady=5, command=save_adjustment).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Skip (Keep Full Amount)", font=("Segoe UI", 9), bg="#95a5a6", fg="white", relief="flat", padx=10, pady=5, command=dlg.destroy).pack(side="left", padx=5)

    def apply_bill_adjustment(self, bill_no, original_amount, actual_given, sales_analysis_path, salesman_data_path):
        ratio = actual_given / original_amount if original_amount != 0 else 1.0
        conn = self.get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, final_item_price, calculated_commission FROM sales_ledger WHERE bill_no = ?", (bill_no,))
        rows = cursor.fetchall()
        
        for row_id, old_final_price, old_comm in rows:
            new_final_price = old_final_price * ratio
            new_comm = old_comm * ratio
            cursor.execute("""
                UPDATE sales_ledger 
                SET final_item_price = ?, calculated_commission = ? 
                WHERE id = ?
            """, (new_final_price, new_comm, row_id))
            
        conn.commit()
        conn.close()
        
        self.generate_excel_reports(sales_analysis_path, salesman_data_path)
        messagebox.showinfo("Bill Updated", f"Bill #{bill_no} recalculated & adjusted to ₹{actual_given:.2f}.")

    def open_recent_bills_lookup(self):
        lookup_win = tk.Toplevel(self.root)
        lookup_win.title("Recent Bills Registry & Adjustments")
        lookup_win.geometry("800x480")
        lookup_win.configure(bg="#f8f9fa")
        
        tk.Label(lookup_win, text="Select Recent Bill to Modify, Cancel or Remove Items", font=("Segoe UI", 11, "bold"), fg="#2c3e50", bg="#f8f9fa").pack(pady=10)
        
        tree_frame = tk.Frame(lookup_win, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        cols = ("bill_no", "date", "cust", "phone", "total")
        b_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        b_tree.heading("bill_no", text="BILL NO")
        b_tree.heading("date", text="DATE")
        b_tree.heading("cust", text="CUSTOMER NAME")
        b_tree.heading("phone", text="PHONE")
        b_tree.heading("total", text="FINAL BILL AMT (₹)")
        
        b_tree.column("bill_no", width=120, anchor="center")
        b_tree.column("date", width=100, anchor="center")
        b_tree.column("cust", width=180, anchor="w")
        b_tree.column("phone", width=110, anchor="center")
        b_tree.column("total", width=120, anchor="center")
        b_tree.pack(fill="both", expand=True)
        
        def reload_recent_bills():
            for row in b_tree.get_children():
                b_tree.delete(row)
            conn = self.get_db_connection(); cursor = conn.cursor()
            cursor.execute("""
                SELECT bill_no, sale_date, customer_name, customer_phone, SUM(final_item_price) 
                FROM sales_ledger 
                GROUP BY bill_no 
                ORDER BY id DESC LIMIT 25
            """)
            for row in cursor.fetchall():
                b_tree.insert("", "end", values=(row[0], row[1], row[2], row[3], f"{row[4]:.2f}"))
            conn.close()

        reload_recent_bills()
        
        act_frame = tk.Frame(lookup_win, bg="#f8f9fa")
        act_frame.pack(fill="x", padx=15, pady=10)
        
        def edit_selected_bill():
            sel = b_tree.selection()
            if not sel: return
            val = b_tree.item(sel[0], "values")
            b_no = val[0]
            curr_amt = float(val[4])
            
            new_val = simpledialog.askfloat("Adjust Bill Total", f"Updating Bill #{b_no}\nCurrent recorded amount: ₹{curr_amt:.2f}\nEnter actual total collected:", parent=lookup_win)
            if new_val is not None:
                _, sales_analysis_path, salesman_data_path, _ = self.get_archive_paths()
                self.apply_bill_adjustment(b_no, curr_amt, new_val, sales_analysis_path, salesman_data_path)
                reload_recent_bills()

        def cancel_selected_bill():
            sel = b_tree.selection()
            if not sel: return
            val = b_tree.item(sel[0], "values")
            b_no = val[0]
            if messagebox.askyesno("Cross Out / Void Bill", f"Are you sure you want to cancel Bill #{b_no}? This will clear amounts in ledger.", parent=lookup_win):
                _, sales_analysis_path, salesman_data_path, _ = self.get_archive_paths()
                self.apply_bill_adjustment(b_no, float(val[4]), 0.0, sales_analysis_path, salesman_data_path)
                reload_recent_bills()

        def manage_items_in_bill():
            sel = b_tree.selection()
            if not sel: return
            b_no = b_tree.item(sel[0], "values")[0]
            self.open_bill_item_return_window(b_no, reload_recent_bills)

        tk.Button(act_frame, text="📦 Cancel Specific Item", font=("Segoe UI", 9, "bold"), bg="#e67e22", fg="white", relief="flat", padx=12, pady=5, command=manage_items_in_bill).pack(side="left", padx=5)
        tk.Button(act_frame, text="✏️ Adjust Received Amount", font=("Segoe UI", 9, "bold"), bg="#2980b9", fg="white", relief="flat", padx=12, pady=5, command=edit_selected_bill).pack(side="left", padx=5)
        tk.Button(act_frame, text="❌ Strike Out / Cancel Bill", font=("Segoe UI", 9, "bold"), bg="#c0392b", fg="white", relief="flat", padx=12, pady=5, command=cancel_selected_bill).pack(side="left", padx=5)

    def process_item_quantity_reduction(self, entry_id, bill_no, current_qty, unit_price, comm_pct):
        qty_to_return = simpledialog.askfloat(
            "Partial Return", 
            f"Current Quantity: {current_qty:g}\nHow many pieces is the customer returning?",
            minvalue=0.1, 
            maxvalue=abs(float(current_qty))
        )

        if qty_to_return is None:
            return

        new_qty = current_qty - qty_to_return

        conn = self.get_db_connection()
        cursor = conn.cursor()

        if new_qty == 0:
            cursor.execute("DELETE FROM sales_ledger WHERE id = ?", (entry_id,))
            msg = f"Item completely removed from Bill #{bill_no}."
        else:
            cursor.execute("SELECT raw_total, allocated_discount FROM sales_ledger WHERE id = ?", (entry_id,))
            row = cursor.fetchone()
            
            raw_total = row[0] if row else (current_qty * unit_price)
            allocated_discount = row[1] if row else 0.0
            
            discount_per_unit = allocated_discount / current_qty if current_qty != 0 else 0.0
            
            new_raw_total = new_qty * unit_price
            new_allocated_discount = new_qty * discount_per_unit
            new_final_price = new_raw_total - new_allocated_discount
            new_commission = new_final_price * (comm_pct / 100.0)

            cursor.execute("""
                UPDATE sales_ledger 
                SET qty = ?, 
                    raw_total = ?, 
                    allocated_discount = ?, 
                    final_item_price = ?, 
                    calculated_commission = ?
                WHERE id = ?
            """, (new_qty, new_raw_total, new_allocated_discount, new_final_price, new_commission, entry_id))
            
            msg = f"Bill #{bill_no} updated: Returned {qty_to_return:g} pcs. Remaining Qty: {new_qty:g}."

        conn.commit()
        conn.close()

        _, sales_analysis_path, salesman_data_path, customer_ledger_path = self.get_archive_paths()
        self.generate_excel_reports(sales_analysis_path, salesman_data_path)
        self.update_customer_ledger_excel(customer_ledger_path)

        messagebox.showinfo("Return Processed", f"{msg}\nDatabase, Excel summaries & commissions updated!")

    def open_bill_item_return_window(self, bill_no, refresh_parent_callback=None):
        item_win = tk.Toplevel(self.root)
        item_win.title(f"Item Return & Refunds - Bill #{bill_no}")
        item_win.geometry("750x380")
        item_win.configure(bg="#ffffff")

        tk.Label(item_win, text=f"Purchased Items for Bill #{bill_no}", font=("Segoe UI", 11, "bold"), bg="#ffffff", fg="#2c3e50").pack(pady=10)

        t_frame = tk.Frame(item_win, bg="#ffffff")
        t_frame.pack(fill="both", expand=True, padx=15, pady=5)

        cols = ("id", "item_no", "item_name", "qty", "unit_price", "final_price", "comm_pct")
        i_tree = ttk.Treeview(t_frame, columns=cols, show="headings")
        i_tree.heading("id", text="ENTRY ID")
        i_tree.heading("item_no", text="ITEM CODE")
        i_tree.heading("item_name", text="ITEM NAME")
        i_tree.heading("qty", text="QTY")
        i_tree.heading("unit_price", text="UNIT PRICE")
        i_tree.heading("final_price", text="TOTAL PAID")
        i_tree.heading("comm_pct", text="COMM %")

        i_tree.column("id", width=60, anchor="center")
        i_tree.column("item_no", width=90, anchor="center")
        i_tree.column("item_name", width=180, anchor="center")
        i_tree.column("qty", width=60, anchor="center")
        i_tree.column("unit_price", width=90, anchor="center")
        i_tree.column("final_price", width=100, anchor="center")
        i_tree.column("comm_pct", width=70, anchor="center")
        i_tree.pack(fill="both", expand=True)

        def load_bill_items():
            for row in i_tree.get_children():
                i_tree.delete(row)
            conn = self.get_db_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT s.id, s.item_no, COALESCE(i.item_name, s.item_no), s.qty, s.unit_price, s.final_item_price, COALESCE(i.commission_pct, 0)
                FROM sales_ledger s
                LEFT JOIN inventory i ON s.item_no = i.item_no
                WHERE s.bill_no = ?
            """, (bill_no,))
            for r in cursor.fetchall():
                i_tree.insert("", "end", values=(r[0], r[1], r[2], f"{r[3]:g}", f"{r[4]:.2f}", f"{r[5]:.2f}", f"{r[6]:.2f}"))
            conn.close()

        load_bill_items()

        def process_item_cancellation():
            sel = i_tree.selection()
            if not sel:
                messagebox.showerror("Selection Error", "Select an item to adjust or return.", parent=item_win)
                return
            
            vals = i_tree.item(sel[0], "values")
            entry_id = int(vals[0])
            current_qty = float(vals[3])
            unit_price = float(vals[4])
            comm_pct = float(vals[6])

            self.process_item_quantity_reduction(entry_id, bill_no, current_qty, unit_price, comm_pct)
            load_bill_items()
            if refresh_parent_callback:
                refresh_parent_callback()

        act_frame = tk.Frame(item_win, bg="#ffffff")
        act_frame.pack(fill="x", padx=15, pady=10)

        tk.Button(act_frame, text="↩️ Return / Reduce Item Qty", font=("Segoe UI", 9, "bold"), bg="#e67e22", fg="white", relief="flat", padx=12, pady=5, command=process_item_cancellation).pack(side="left")
        tk.Button(act_frame, text="Close", font=("Segoe UI", 9), bg="#7f8c8d", fg="white", relief="flat", padx=12, pady=5, command=item_win.destroy).pack(side="right")

    def generate_excel_reports(self, sales_analysis_path, salesman_data_path):
        current_month_pattern = datetime.now().strftime("%Y-%m") + "%"
        conn = self.get_db_connection()
        
        query_total = """
            SELECT SUM(final_item_price) as total_amt 
            FROM sales_ledger
            WHERE sale_date LIKE ?
        """
        df_total = pd.read_sql_query(query_total, conn, params=(current_month_pattern,))
        total_monthly_sales = float(df_total['total_amt'].iloc[0]) if not df_total.empty and df_total['total_amt'].iloc[0] else 0.0
        
        summary_data = [
            {"Metric": "Total Monthly Sales", "Amount (₹)": total_monthly_sales}
        ]
        df_sales_analysis = pd.DataFrame(summary_data)
        
        with pd.ExcelWriter(sales_analysis_path, engine="openpyxl") as writer:
            df_sales_analysis.to_excel(writer, sheet_name="Sales Summary", index=False)
            
        query_sm = """
            SELECT 
                s.salesman_no as [Salesman No], 
                COALESCE(sm.name, 'Unknown') as [Salesman Name],
                COUNT(DISTINCT s.bill_no) as [Total Orders], 
                SUM(s.final_item_price) as [Total Sales], 
                SUM(s.calculated_commission) as [Commission Earned]
            FROM sales_ledger s
            LEFT JOIN salesmen sm ON s.salesman_no = sm.salesman_no
            WHERE s.sale_date LIKE ?
            GROUP BY s.salesman_no
        """
        df_sm = pd.read_sql_query(query_sm, conn, params=(current_month_pattern,))
        conn.close()
        
        if df_sm.empty:
            df_sm = pd.DataFrame(columns=["Salesman No", "Salesman Name", "Total Orders", "Total Sales", "Commission Earned"])
            
        with pd.ExcelWriter(salesman_data_path, engine="openpyxl") as writer:
            df_sm.to_excel(writer, sheet_name="Salesman Records", index=False)

    def update_customer_ledger_excel(self, excel_path):
        current_year = datetime.now().year
        conn = self.get_db_connection()
        
        query_yearly_customers = """
            SELECT 
                customer_name as [customer name],
                MAX(CASE WHEN customer_phone != '-' THEN customer_phone ELSE '-' END) as [mobile no],
                COUNT(DISTINCT bill_no) as [no of times entry],
                SUM(final_item_price) as [total purchase]
            FROM sales_ledger
            WHERE customer_name IS NOT NULL 
              AND customer_name NOT IN ('', 'Guest Customer') 
              AND sale_year = ?
            GROUP BY LOWER(TRIM(customer_name))
            ORDER BY [total purchase] DESC
        """
        df_customers = pd.read_sql_query(query_yearly_customers, conn, params=(current_year,))
        conn.close()
        
        if df_customers.empty:
            df_customers = pd.DataFrame(columns=["customer name", "mobile no", "no of times entry", "total purchase"])
        
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df_customers.to_excel(writer, sheet_name="Yearly Customer Summary", index=False)

    def build_receipt_string(self, c_name, c_phone, disc_pct, bill_no):
        import math

        def round_half_up(n):
            return math.floor(n + 0.5)

        now_dt = datetime.now()
        date_str = now_dt.strftime("%d/%m/%y")
        time_str = now_dt.strftime("%H:%M")
        
        unique_items = len(self.cart_items)
        total_qty = sum(float(i["qty"]) for i in self.cart_items)
        
        raw_gross = sum((float(i["price"]) * float(i["qty"])) for i in self.cart_items)
        gross_total = round_half_up(raw_gross)
        
        saved = round_half_up(gross_total * (disc_pct / 100.0))
        net_total = gross_total - saved
        final_total = round_half_up(net_total)
        
        is_return_bill = bill_no.startswith("R_") or any(i["qty"] < 0 for i in self.cart_items)
        
        W = 38
        SEP_MARKER = "__LINE_SEP__"
        
        lines = []
        
        lines.append("SURBHI COLLECTION".center(W))
        lines.append("Shukleshwar Road, Rahuri,".center(W))
        lines.append("Ahilyanagar(MH)".center(W))
        lines.append("Contact No: 7588297096".center(W))
        lines.append("")
        
        lines.append(f"BILL NO: {bill_no}".ljust(19) + f"DATE: {date_str}".rjust(19))
        lines.append(f"TIME: {time_str}".rjust(W))
        lines.append(SEP_MARKER)
        
        lines.append(f"Customer: {c_name}"[:W])
        lines.append(f"Mobile: {c_phone}"[:W])
        lines.append(SEP_MARKER)
        
        lines.append(f"{'SM':<4}{'ITEM':<14}{'QTY':>6}{'RATE':>6}{'AMT':>8}")
        lines.append(SEP_MARKER)
        
        for item in self.cart_items:
            sm = str(item["salesman_no"])[:3].ljust(4)
            name = str(item["item_name"])[:13].ljust(14)
            
            qty_num = item['qty']
            qty_str = f"{int(qty_num)}" if qty_num.is_integer() else f"{qty_num:g}"
            qty_val = f"{qty_str}".rjust(6)
            
            item_price_rounded = round_half_up(item['price'])
            item_amt_rounded = round_half_up(item['price'] * item['qty'])
            
            price_val = f"{item_price_rounded}".rjust(6)
            amt_val = f"{item_amt_rounded}".rjust(8)
            
            lines.append(f"{sm}{name}{qty_val}{price_val}{amt_val}")
            
        lines.append(SEP_MARKER)
        
        tot_qty_str = f"{int(total_qty)}" if total_qty.is_integer() else f"{total_qty:g}"
        summary_str = f"ITEMS: {unique_items} | QTY: {tot_qty_str}"
        gross_val_str = f"Rs.{gross_total}"
        lines.append(f"{summary_str:<30}{gross_val_str:>8}")
        
        disc_label = f"TOTAL SAVED ({disc_pct:g}%):"
        saved_val_str = f"Rs.{saved}"
        lines.append(f"{disc_label:<30}{saved_val_str:>8}")
        lines.append(SEP_MARKER)
        
        if is_return_bill:
            net_label = "NET AMOUNT (RETURNED):"
        else:
            net_label = "NET AMOUNT:"
            
        net_val_str = f"Rs.{final_total}"
        lines.append(f"{net_label:<30}{net_val_str:>8}")
        lines.append(SEP_MARKER)
        
        lines.append("No Guarantee / No Return".center(W))
        lines.append("Thank You for Shopping! Visit Again")
        
        side_pad = "  "
        full_width = W + (len(side_pad) * 2)
        full_line_sep = "-" * full_width
        
        formatted_lines = []
        for line in lines:
            if line == SEP_MARKER:
                formatted_lines.append(full_line_sep)
            else:
                formatted_lines.append(f"{side_pad}{line}{side_pad}")
        
        return "\n".join(formatted_lines)

    def clear_complete_billing_grid(self):
        self.cart_items = []
        for row in self.tree.get_children(): 
            self.tree.delete(row)
        self.calculate_bill_totals()
        self.clear_item_input_fields(keep_customer_and_salesman=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = SurbhiCollectionApp(root)
    root.mainloop()