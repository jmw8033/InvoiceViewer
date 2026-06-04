from tkinter import ttk, messagebox, scrolledtext, font, simpledialog
from tkcalendar import DateEntry
from collections import defaultdict
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
import os, pymssql, time, threading, re, queue, sys, json, gc
import babel.numbers


INVOICE_DIR = r"S:\Titan_DM\Titan_Filing\AP_Invoices"
invoice_re = re.compile(r'(\d+)')

class InvoiceViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Titan Invoice Viewer")
        self.iconbitmap(default="icon.ico")
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.tk.call("tk", "scaling", 1.33)
        self.gui_queue = queue.Queue()
        self.geometry("1400x700")

        self.invoices = {} # List of dictionaries, {"VendorID", "InvoiceNum", "InvoiceDate", "ExtAmount", "Filepath"}
        self.company_ids = set() # Set of company IDs for quick lookup
        self.sort_col = "Date"
        self.sort_desc = True # descending
        self.broken_companies = []
        self.broken_invoices = []
        self.by_vendor_invoice = []
        self.checks_by_vendor_invoice = defaultdict(list)
        self.accounts_by_vendor_invoice = defaultdict(list)
        self.account_description_by_account = {}
        self.missing_invoices = []

        # Ignore list for private vendors
        self.ignoring = True
        self.ignore_list = set()
        with open("ignore.json", "r") as f:
            self.ignore_list = set(json.load(f))
        self.protocol("WM_DELETE_WINDOW", self.on_exit)  # runs exit protocol on window closed

        # Loading info
        self.create_loading_screen()

        # Get data
        self.after(0, lambda: threading.Thread(target=self.load_data, daemon=True).start())
        self.loading_loop_id = self.after(50, self.loading_loop)


    def create_loading_screen(self):
        self.loading_bg = tk.PhotoImage(file="logo.png")
        self.loading_canvas = tk.Canvas(self, bg="white", width=1220, height=700)
        self.loading_canvas.pack(expand=True, fill="both", side="top", anchor="w")
        self.loading_canvas.background = self.loading_bg
        self.loading_canvas.create_image(1220/2, 0, anchor="n", image=self.loading_bg)
        
        tk.Label(self.loading_canvas, text="Welcome to Titan Invoice Viewer", font=("TKDefaultFont", 24, "bold"), bg="white").pack(side="top", anchor="w")
        tk.Label(self.loading_canvas, text="Please press the Help button for more information about this program", font=("TKDefaultFont", 20), bg="white").pack(side="top", anchor="w")
        tk.Label(self.loading_canvas, text="Loading Titan invoices, Please wait...", font=("TKDefaultFont", 16), bg="white").pack(side="top", anchor="w")


    def loading_update(self, msg, color="#000000"):
        #print(msg)
        self.gui_queue.put((msg, color))


    def loading_loop(self):
        try:
            while True:
                msg, color = self.gui_queue.get_nowait()
                tk.Label(self.loading_canvas, text=msg, font=("TKDefaultFont", 16), fg=color, bg="white").pack(side="top", anchor="w")
        except queue.Empty:
            pass
        self.loading_loop_id = self.after(50, self.loading_loop)


    def load_data(self):
        self.t0 = time.perf_counter()

        with ThreadPoolExecutor(max_workers=2) as pool:
            _ = pool.submit(self.load_database)
            file_index = pool.submit(self.load_files)

            file_index = file_index.result()
            _ = _.result()

        # Match files with invoices
        t0 = time.perf_counter()
        for row in self.invoices:
            try:
                row["Filepath"] = file_index.pop((row["VendorID"], row["InvoiceNum"]))
            except:
                self.missing_invoices.append((row["VendorID"], row["InvoiceNum"], row["InvoiceDate"].strftime("%m-%d-%Y")))
        t1 = time.perf_counter()
        self.loading_update(f"Invoice files loaded in {t1 - t0:.2f} seconds.")
        self.loading_update((f"{len(self.broken_companies)} broken titan entries found."), color="#FF0000")
        self.loading_update(f"{len(self.missing_invoices)} missing invoice files.", color="#FF0000")

        # Check for errors
        if len(file_index) > 0:
            for file in file_index:
                self.broken_invoices.append(file_index[file])
            self.loading_update(f"{len(file_index)} invoice files without matches.", color="#FF0000")

        self.after(100, self.load_gui)

    
    def load_gui(self):
        # Destroy loading screen
        self.loading_canvas.destroy()

        # Create Frames
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        self.create_treeview()
        self.create_filter_frame()
        self.create_summary_bar()
        self.error_popup = ErrorPopup(self, self.broken_companies, self.broken_invoices, self.missing_invoices)
        self.help_popup = HelpPopup(self)

        # Ignore list image
        self.ignore_photo = tk.PhotoImage(file="leaf.png")
        self.ignore_label = tk.Label(self.filter_frame, image=self.ignore_photo)
        self.bind("<Control-F9>", self.add_ignore)
        self.bind("<Control-F10>", self.toggle_ignore_list)


    def load_database(self):
        def load_header():
            # Connect to the database and fetch invoice data
            t0 = time.perf_counter()
            conn = pymssql.connect(
                server="ACAPP1",
                user="titan",
                password="titan",
                database="titan",
            )

            with conn.cursor(as_dict=True) as cur:
                cur.execute("""
                    SELECT APH.VendorID, APH.InvoiceNum, APH.InvoiceDate, APH.Subtotal, APH.Payments, APH.PlantID, V.CompanyName
                    FROM AP_Header APH
                    JOIN Vendors V ON APH.VendorID = V.VendorId
                """)
                data = cur.fetchall()

            self.invoices = [row for row in data if row["VendorID"] and row["InvoiceNum"] and row["InvoiceDate"] 
                            and row["Subtotal"] is not None and row["Payments"] is not None]
            self.broken_companies = [row for row in data if not (row["VendorID"] and row["InvoiceNum"] and row["InvoiceDate"] 
                                     and row["Subtotal"] is not None and row["Payments"] is not None)]
            self.company_ids = {(row["VendorID"], row["CompanyName"], row["VendorID"] in self.ignore_list) for row in self.invoices if row["VendorID"] and row["CompanyName"]}
            self.by_vendor_invoice = {(row["VendorID"], row["InvoiceNum"]): row for row in self.invoices}

            t1 = time.perf_counter()
            self.loading_update(f"Invoice data loaded in {t1 - t0:.2f} seconds.")
            conn.close()

        def load_checks():
            t0 = time.perf_counter()
            # Connect to the database and fetch check data
            conn = pymssql.connect(
                server="ACAPP1",
                user="titan",
                password="titan",
                database="titan",
            )

            with conn.cursor(as_dict=True) as cur:
                cur.execute("""
                    SELECT CH.CheckNum, CH.CheckDate, CD.InvoiceNum, CD.Amount, CH.VendorID
                    FROM Check_Header CH
                    JOIN Check_Detail CD ON CH.CheckID = CD.CheckID
                """)
                data=cur.fetchall()
            for row in data:
                self.checks_by_vendor_invoice[(row["VendorID"], row["InvoiceNum"])].append((row["CheckNum"], row["CheckDate"], row["Amount"]))
            conn.close()
            t1 = time.perf_counter()
            self.loading_update(f"Check data loaded in {t1 - t0:.2f} seconds.")
            conn.close()

        def load_accounts():
            t0 = time.perf_counter()
            # Connect to the database and fetch check data
            conn = pymssql.connect(
                server="ACAPP1",
                user="titan",
                password="titan",
                database="titan",
            )

            with conn.cursor(as_dict=True) as cur:
                cur.execute("""
                    SELECT APH.VendorID, APH.InvoiceNum, APD.Account, APD.ExtAmount, GL.AccountDescription
                    FROM AP_Header APH
                    JOIN AP_Detail APD ON APH.RecordNum = APD.RecordNum
                    JOIN GL_Accounts GL ON APD.Account = GL.AccountNumber
                """)
                data=cur.fetchall()
            for row in data:
                self.accounts_by_vendor_invoice[(row["VendorID"], row["InvoiceNum"])].append((row["Account"], row["ExtAmount"]))
                self.account_description_by_account[row["Account"]] = row["AccountDescription"]

            conn.close()
            t1 = time.perf_counter()
            self.loading_update(f"Check data loaded in {t1 - t0:.2f} seconds.")
            conn.close()
        
        
        with ThreadPoolExecutor(max_workers=3) as pool:
            _ = pool.submit(load_header).result()
            _ = pool.submit(load_checks).result()
            _ = pool.submit(load_accounts).result()
        t1 = time.perf_counter()
        self.loading_update(f"Database loaded in {t1 - self.t0:.2f} seconds.")


    def load_files(self):
        # Scan the invoice directory and create file lookup table
        t0 = time.perf_counter()
        file_index = {}
        for file in os.scandir(INVOICE_DIR):
            file = file.name
            fname = file.split("_") # COMPANY_INVOICE_MM-DD-YYYY_RANDOMINT
            if fname and len(fname) >= 3:
                if fname[0] == "PUCA": # Special case, they have underscore in company ID
                    fname[0] = "PUCA_150"
                    fname[1] = fname[2] # Invoice number

                fname[1] = fname[1].replace("[slash]", "/").replace("[quote]", '"')
                file_index[(fname[0], fname[1])] = os.path.join(INVOICE_DIR,  file)

        t1 = time.perf_counter()
        self.loading_update(f"Invoice files scanned in {t1 - t0:.2f} seconds.")
        return file_index


    def create_filter_frame(self):
        self.filter_frame = ttk.Frame(self, height=60)
        self.filter_frame.grid(row=0, column=0, sticky="ew")
        self.filter_frame.grid_propagate(False)

        # Row 0

        # Company Search bar
        ttk.Label(self.filter_frame, text="Company ID:").grid(row=0, column=0, padx=5)
        self.company_entry = AutoCompleteEntry(self)
        self.company_entry.grid(row=0, column=1, padx=5, pady=5)

        # Account Search bar
        ttk.Label(self.filter_frame, text="Account:").grid(row=0, column=2, padx=5)
        self.account_entry = tk.Entry(self.filter_frame)
        self.account_entry.grid(row=0, column=3, padx=5)
        self.account_text = tk.StringVar()
        self.account_entry["textvariable"] = self.account_text
        self.prev_account_text = ""
        # When account text changes, rebuild using the new account filter
        self.account_text.trace_add("write", lambda *_: self.company_entry.on_select(source="account"))

        # Date ranges
        ttk.Label(self.filter_frame, text="Start date:").grid(row=0, column=4, padx=5)
        self.start_entry = DateEntry(self.filter_frame, width=10, date_pattern="mm/dd/yyyy")
        self.start_entry.set_date("01/01/2014")
        self.start_entry.grid(row=0, column=5, padx=5)
        self.start_entry.bind("<<DateEntrySelected>>", self.company_entry.on_select)
        self.start_entry.bind("<Return>", self.company_entry.on_select)

        ttk.Label(self.filter_frame, text="End date:").grid(row=0, column=6, padx=5)
        self.end_entry = DateEntry(self.filter_frame, width=10, date_pattern="mm/dd/yyyy")
        self.end_entry.grid(row=0, column=7, padx=5)
        self.end_entry.bind("<Return>", self.company_entry.on_select)
        self.end_entry.bind("<<DateEntrySelected>>", self.company_entry.on_select)

        # All companies checkbox
        self.all_companies = tk.BooleanVar()
        self.all_companies_cb = ttk.Checkbutton(self.filter_frame, text="Search All Companies", variable=self.all_companies, command=self.company_entry.toggle_all_companies, takefocus=False)
        self.all_companies_cb.grid(row=0, column=8, padx=5)

        # PDF Only Checkbox
        self.pdf_only = tk.BooleanVar()
        self.pdf_cb = ttk.Checkbutton(self.filter_frame, text="File Available Only", variable=self.pdf_only, command=self.company_entry.on_select, takefocus=False)
        self.pdf_cb.grid(row=0, column=9, padx=5)

        # Far right frame for buttons
        ttk.Label(self.filter_frame, text="                                                                                ").grid(row=0, column=10) # Spacer
        self.right_button_frame = ttk.Frame(self.filter_frame)
        self.right_button_frame.grid(row=0, column=11, padx=5, sticky="e")
        # Refresh button
        self.refresh_button = tk.Button(self.right_button_frame, text="⭮", command=self.restart)
        self.refresh_button.pack(side="left", padx=5)

        # Help button
        self.help_button = tk.Button(self.right_button_frame, text="Help", command=lambda *_: self.help_popup.toggle())
        self.help_button.pack(side="left", padx=5)

        # Errors button
        self.errors_button = tk.Button(self.right_button_frame, text="Errors", command=lambda *_: self.error_popup.toggle())
        self.errors_button.pack(side="left", padx=5)

        # Row 1

        # Invoice search
        ttk.Label(self.filter_frame, text="Invoice:").grid(row=1, column=0, padx=5)
        self.invoice_entry = tk.Entry(self.filter_frame)
        self.invoice_entry.grid(row=1, column=1, padx=5)
        self.invoice_text = tk.StringVar()
        self.prev_invoice_text = ""
        self.invoice_entry["textvariable"] = self.invoice_text
        self.invoice_text.trace_add("write", lambda *_: self.company_entry.on_select(source="invoice"))

        # Plant Filter Dropdown
        ttk.Label(self.filter_frame, text="Plant:").grid(row=1, column=2, padx=5)
        self.plant_var = tk.StringVar(value="Both")
        self.plant_cb = ttk.Combobox(self.filter_frame, textvariable=self.plant_var, values=["Both", "ACP", "APC"], width=6, state="readonly")
        self.plant_cb.grid(row=1, column=3, padx=5, sticky="w")
        self.plant_cb.bind("<<ComboboxSelected>>", self.company_entry.on_select)

        # Date Filter Dropdown
        ttk.Label(self.filter_frame, text="Date Filter:").grid(row=1, column=4, padx=5)
        self.date_filter_var = tk.StringVar(value="Invoice Date")
        self.date_filter_cb = ttk.Combobox(self.filter_frame, textvariable=self.date_filter_var, values=["Invoice Date", "Check Date"], width=12, state="readonly")
        self.date_filter_cb.grid(row=1, column=5, padx=5)
        self.date_filter_cb.bind("<<ComboboxSelected>>", self.company_entry.on_select)


    def create_summary_bar(self):
        # Bottom level summary bar for totals and selected sum
        self.summary_frame = ttk.Frame(self, relief="groove", borderwidth=1, padding=(6, 4))
        self.summary_frame.grid(row=2, column=0, sticky="ew")
        
        for i in range(5):
            self.summary_frame.columnconfigure(i, weight=1)

        self.amount_label = ttk.Label(self.summary_frame, text="0 invoices found.", font=("TKDefaultFont", 10, "bold"))
        self.amount_label.grid(row=0, column=0, sticky="w", padx=10)

        self.account_sum = tk.StringVar(value="Account Total: $0")
        self.account_sum_label = ttk.Label(self.summary_frame, textvariable=self.account_sum, font=("TKDefaultFont", 10))
        self.account_sum_label.grid(row=0, column=1, sticky="w")

        self.selected_sum = tk.StringVar(value="Selected Total: $0.00")
        self.selected_sum_label = ttk.Label(self.summary_frame, textvariable=self.selected_sum, font=("TKDefaultFont", 10))
        self.selected_sum_label.grid(row=0, column=2, sticky="w")

        self.invoice_total = tk.StringVar(value="Invoice Total: $0")
        self.invoice_total_label = ttk.Label(self.summary_frame, textvariable=self.invoice_total, font=("TKDefaultFont", 10, "bold"))
        self.invoice_total_label.grid(row=0, column=3, sticky="w")

        self.balance_total = tk.StringVar(value="Balance Total: $0")
        self.balance_total_label = ttk.Label(self.summary_frame, textvariable=self.balance_total, font=("TKDefaultFont", 10, "bold"))
        self.balance_total_label.grid(row=0, column=4, sticky="w", padx=10)


    def create_treeview(self):
        self.tree_frame = ttk.Frame(self, height=600)
        self.tree_frame.grid(row=1, column=0, sticky="nsew")

        # Column setup
        self.tree = ttk.Treeview(self.tree_frame, columns=("Vendor", "Company Name", "GL Account", "Invoice", "Date", "Invoice Amount", "Balance", 
                                                           "Check Number", "Check Date", "File Available", "Filepath"), show='tree headings')
        self.tree.column("#0", width=0, stretch=False)
        self.tree.column("Vendor", width=50, anchor="center")
        self.tree.column("Company Name", width=110, anchor="center")
        self.tree.column("GL Account", width=140, anchor="center")
        self.tree.column("Invoice", width=110, anchor="center")
        self.tree.column("Date", width=40, anchor="center")
        self.tree.column("Invoice Amount", width=50, anchor="center")
        self.tree.column("Balance", width=50, anchor="center")
        self.tree.column("Check Number", width=70, anchor="center")
        self.tree.column("Check Date", width=40, anchor="center")
        self.tree.column("File Available", width=30, anchor="center")
        self.tree.column("Filepath", width=0, stretch=False)

        self.tree.heading("Vendor", text="Vendor", command=lambda: self.sort_by("Vendor"))
        self.tree.heading("Company Name", text="Company Name", command=lambda: self.sort_by("Company Name"))
        self.tree.heading("GL Account", text="GL Account", command=lambda: self.sort_by("GL Account"))
        self.tree.heading("Invoice", text="Invoice", command=lambda: self.sort_by("Invoice"))
        self.tree.heading("Date", text="Date  ▼", command=lambda: self.sort_by("Date"))
        self.tree.heading("Invoice Amount", text="Invoice Amount", command=lambda: self.sort_by("Invoice Amount"))
        self.tree.heading("Balance", text="Balance", command=lambda: self.sort_by("Balance"))
        self.tree.heading("Check Number", text="Check Number", command=lambda: self.sort_by("Check Number"))
        self.tree.heading("Check Date", text="Check Date", command=lambda: self.sort_by("Check Date"))
        self.tree.heading("File Available", text="File Available", command=lambda: self.sort_by("File Available"))
        self.tree.heading("Filepath", text="")
        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(self.tree_frame, command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self.update_selected_sum)

        # Styles
        self.style.configure("Treeview", rowheight=20) 
        self.tree.tag_configure("oddrow",  background="#f7f7f7")
        self.tree.tag_configure("evenrow", background="#ffffff")
        self.tree.tag_configure("checkrow", background="#fdfaf1")
        
    
    def update_selected_sum(self, *_):
        total = 0
        for item in self.tree.selection():
            vals = self.tree.item(item, "values")
            amt_str = vals[5]  # Invoice Amount column — empty on subrows
            if amt_str:
                total += float(amt_str.replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
        total_str = f"${total:,.2f}" if total >= 0 else f"(${abs(total):,.2f})"
        self.selected_sum.set(f"Selected: {total_str}")


    def show_invoices(self, company, invoice_prefix, account_filter): 
        invoice_count = 0
        invoice_total = 0
        balance_total = 0
        values = []
        plant_filter = self.plant_var.get()
        date_filter = self.date_filter_var.get()

        for entry in self.invoices:
            # Check Plant ID, 110 for Concrete, 410 for Precast
            plant_id = entry["PlantID"]
            if plant_filter == "ACP" and plant_id != "110":
                continue
            if plant_filter == "APC" and plant_id != "410":
                continue

            # Check company, invoice prefix, and account filter
            vendor = str(entry["VendorID"])
            if (self.all_companies.get() and not vendor.lower().startswith(company.lower())) or (self.ignoring and vendor in self.ignore_list):
                continue
            if not self.all_companies.get() and vendor.lower() != company.lower():
                continue
            if account_filter:
                gl_accounts = self.accounts_by_vendor_invoice[(vendor, str(entry["InvoiceNum"]))].copy()
                matched = False
                for account, amount in gl_accounts:
                    if account and self.account_match_filter(account_filter, account):
                        matched = True
                        break
                if not matched:
                    continue

            # Check if invoice date is between start and end dates
            invoice =  str(entry["InvoiceNum"])
            date = entry["InvoiceDate"].date()
            if date_filter == "Invoice Date":
                if date < self.start_entry.get_date() or date > self.end_entry.get_date():
                    continue
            elif date_filter == "Check Date":
                checks = self.checks_by_vendor_invoice[(vendor, invoice)]
                if not checks:
                    continue
                has_valid_check_date = False
                for check_num, check_date, check_amount in checks:
                    if check_date.date() >= self.start_entry.get_date() and check_date.date() <= self.end_entry.get_date():
                        has_valid_check_date = True
                        break
                if not has_valid_check_date:
                    continue 

            # Check if Has File Only is checked
            filepath = entry.get("Filepath", "")
            has_filepath = "✔" if filepath else ""
            if self.pdf_only.get() == 1 and not has_filepath:
                continue

            # If we passed all the filters, add the invoice to the list
            if not invoice.lower().startswith(invoice_prefix.lower()):
                continue
            company_name = entry["CompanyName"]
            amount = entry['Subtotal']
            check_number = ""
            check_date= ""

            # Add subrows for checks
            checks = self.checks_by_vendor_invoice[(vendor, invoice)]
            if len(checks) == 1:
                check_number, check_date, _ = checks[0]
                check_date = check_date.strftime("%m-%d-%Y")

            # Get Balance
            payments = entry["Payments"]
            if payments == amount:
                balance = "Paid In Full"
            else:
                balance = amount - payments
                balance_total += balance
                balance = f"${balance:,.2f}" if balance >= 0 else f"(${abs(balance):,.2f})"
            invoice_total += amount
            amount = f"${amount:,.2f}" if amount >= 0 else f"(${abs(amount):,.2f})"

            # get GL Account, just like checks
            gl_account = self.accounts_by_vendor_invoice[(vendor, invoice)].copy()                    
            if len(gl_account) == 1:
                gl_account = gl_account[0]
                gl_account = f"{gl_account[0]} - {self.account_description_by_account.get(gl_account[0], '')}"
            else:
                gl_account = ""

            values.append((vendor, company_name, gl_account, invoice, date, amount, balance, check_number, check_date, has_filepath, filepath))
            invoice_count += 1

        invoice_total =  f"${invoice_total:,.2f}" if invoice_total >= 0 else f"(${abs(invoice_total):,.2f})"
        balance_total = f"${balance_total:,.2f}" if balance_total >= 0 else f"(${abs(balance_total):,.2f})"
        self.invoice_total.set(f"Invoice Total: {invoice_total}")
        self.balance_total.set(f"Balance Total: {balance_total}")

        return invoice_count, values


    def filter_rows(self, company, invoice_prefix, account_filter): # Filter current rows based on company name, invoice, and gl account
        invoice_total = 0
        balance_total = 0
        i = 1

        for row in self.tree.get_children():
            values = self.tree.item(row, "values")
            # Check company and invoice prefix
            if not values[0].lower().startswith(company.lower()) or not values[3].lower().startswith(invoice_prefix.lower()):
                self.tree.delete(row)
            elif values[2] in ("▼", "▲"): # Has subrows, need to check each subrow for account filter
                subrows = self.tree.get_children(row)
                remove_row = True
                for subrow in subrows:
                    subvalues = self.tree.item(subrow, "values")
                    if self.account_match_filter(account_filter, subvalues[2]):
                        remove_row = False
                    else:
                        self.tree.delete(subrow)
                if remove_row:
                    self.tree.delete(row)
            elif not self.account_match_filter(account_filter, values[2]): # No subrows, just check account filter
                self.tree.delete(row)

        # Recalculate totals and set row colors
        for row in self.tree.get_children():
            i += 1
            # Set color tag
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.item(row, tags=tag)
            # Format totals.;
            invoice_total += float(values[5].replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
            if values[6] != "Paid In Full":
                balance_total += float(values[6].replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))

        invoice_total = f"${invoice_total:,.2f}" if invoice_total >= 0 else f"(${abs(invoice_total):,.2f})"
        balance_total = f"${balance_total:,.2f}" if balance_total >= 0 else f"(${abs(balance_total):,.2f})"
        self.invoice_total.set(f"Invoice Total: {invoice_total}")
        self.balance_total.set(f"Total: {balance_total}")

        self.update_account_sum()
        
        return len(self.tree.get_children())
    

    def update_account_sum(self):
        total = 0
        for row in self.tree.get_children():
            subrows = self.tree.get_children(row)
            children = [c for c in subrows if self.tree.set(c, "GL Account")]
            if children:
                for c in children:
                    subvals = self.tree.item(c, "values")
                    amt = subvals[3] # GL Account Amount
                    total += float(amt.replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
            else:
                # Get main row amount from Invoice Amount
                vals = self.tree.item(row, "values")
                amt = vals[5]
                total += float(amt.replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
        total = f"${total:,.2f}" if total >= 0 else f"(${abs(total):,.2f})"
        self.account_sum.set(f"Account Sum: {total}")


    def sort_by(self, col, values=None, header_pressed=True):
        account_filter = self.account_text.get()
        if header_pressed:
            if col == self.sort_col:
                self.sort_desc = not self.sort_desc
            else:
                self.sort_col = col
                self.sort_desc = True   

        if not values:
            values = [self.tree.item(i, "values") for i in self.tree.get_children()]
            self.tree.delete(*self.tree.get_children())

        def invoice_key(inv):
            if inv.isdigit():
                return [(0, int(inv))]
            key = []
            return [(1,)] + [(0, int(p)) if p.isdigit() else (1, p.lower()) for p in invoice_re.split(inv) if p]

        keymap = {
            "Vendor": lambda x: x[0],
            "Company Name": lambda x: x[1],
            "GL Account": lambda x: x[2],
            "Invoice": lambda x: invoice_key(str(x[3])),
            "Date": lambda x: x[4],
            "Invoice Amount": lambda x: float(x[5].replace("$", "").replace("(", "-").replace(",", "").replace(")", "")),
            "Balance": lambda x: float(x[6].replace("$", "").replace("(", "-").replace(",", "").replace(")", "")) if x[6] != "Paid In Full" else 0,
            "Check Number": lambda x: x[7],
            "Check Date": lambda x: datetime.strptime(x[8], "%m-%d-%Y")  if x[8] else datetime(2000, 1, 1),
            "File Available": lambda x: x[9]
        }
        reverse = self.sort_desc
        if col == "Vendor" or col == "Invoice" or col == "Company Name":
            reverse = not reverse

        values.sort(key=keymap[col], reverse=reverse)
        for i, row in enumerate(values):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            iid = self.tree.insert("", "end", values=row, tags=tag)

            # Add subrows GL Accounts
            gl_accounts = self.accounts_by_vendor_invoice[(row[0], row[3])].copy()
            if len(gl_accounts) > 1:
                gl_accounts_copy = gl_accounts.copy()
                for account, amount in gl_accounts_copy:
                    if account is None:
                        gl_accounts.remove((account, amount))
                    if not self.account_match_filter(account_filter, account):
                        gl_accounts.remove((account, amount))
                    
                if len(gl_accounts) == 0:
                    self.tree.delete(iid)
                    continue
                else:
                    # Sort GL accounts by account number
                    gl_accounts.sort(key=lambda x: x[0])
                    self.tree.set(iid, "GL Account", "▼")
                    for i in range(len(gl_accounts)):
                        acct, amt = gl_accounts[i]
                        if amt is None:
                            amt = 0
                        else:
                            amt = f"${amt:,.2f}" if amt >= 0 else f"(${abs(amt):,.2f})"
                        acct = f"{acct} - {self.account_description_by_account.get(acct, '')}"
                        self.tree.insert(iid, "end", values=("", "", acct, amt, "", "", "", "", "", "", ""), tags="checkrow")

            # Add subrows Checks
            checks = self.checks_by_vendor_invoice[(row[0], row[3])]
            if len(checks) > 1:
                # Sort checks by check date, newest first
                checks.sort(key=lambda x: x[1], reverse=True)
                for cnum, cdate, camt in checks:
                    if self.date_filter_var.get() == "Check Date":
                        if cdate.date() < self.start_entry.get_date() or cdate.date() > self.end_entry.get_date():
                            continue
                    self.tree.set(iid, "Check Number", "▼")
                    cdate = cdate.strftime("%m-%d-%Y")
                    camt = f"${camt:,.2f}" if camt >= 0 else f"(${abs(camt):,.2f})"
                    self.tree.insert(iid, "end", values=("", "", "", "", "", "", camt, cnum, cdate, "", ""), tags="checkrow")

        arrow = "  ▼" if self.sort_desc else "  ▲"
        for c in self.tree["columns"]:
            text = c + arrow if c == col else c
            self.tree.heading(c, text=text)
        
        self.update_account_sum()
        return "break"


    def account_match_filter(self, account_filter, account):
        filter = account_filter.lower()
        account = "" if account is None else str(account).strip().lower()
        description = self.account_description_by_account.get(account, "").lower()

        if filter.isdigit():
            return account.startswith(filter)
        else:
            return filter in account or filter in description

    
    def restart(self):
        # Save ignore json
        with open("ignore.json", "w") as f:
            json.dump(list(self.ignore_list), f)
        
        for w in self.winfo_children():
            w.destroy()
        
        self.after_cancel(self.loading_loop_id)
        self.invoices.clear()
        self.company_ids.clear()
        self.sort_col = "Date"
        self.sort_desc = True
        self.broken_companies.clear()
        self.broken_invoices.clear()
        self.by_vendor_invoice.clear()
        self.checks_by_vendor_invoice.clear()
        self.accounts_by_vendor_invoice.clear()
        self.missing_invoices.clear()

        self.columnconfigure(0, weight=0)
        self.rowconfigure(0, weight=0)

        gc.collect()

        self.create_loading_screen()
        self.loading_loop_id = self.after(50, self.loading_loop)
        threading.Thread(target=self.load_data, daemon=True).start()


    def toggle_ignore_list(self, event):
        self.ignoring = not self.ignoring
        if self.ignoring:
            self.ignore_label.grid_forget()
        else:
            self.ignore_label.grid(row=0, column=11, sticky="w", padx=12)
        self.company_entry.on_select()
        return "break"


    def add_ignore(self, event):
        vendors = ", ".join([("\n" * ((i) % 7 == 0)) + s for i, s in enumerate(self.ignore_list)])
        new_item = simpledialog.askstring("Hidden Vendors", "Current Hidden Vendors:\n" + vendors + "\n\nEnter new Vendor ID (case doesn't matter)", parent=self)
        if new_item:
            self.ignore_list.add(new_item.upper())
        return "break"
    

    def on_exit(self):
        with open("ignore.json", "w") as f:
            json.dump(list(self.ignore_list), f)
        self.destroy()


class AutoCompleteEntry(tk.Entry):
    def __init__(self, root: InvoiceViewer, *a, **kw):
        super().__init__(root.filter_frame, *a, **kw)
        self.invoices = root.invoices
        self.company_ids = root.company_ids
        self.tree = root.tree
        self.root = root
        self.listbox = None
        self.company = tk.StringVar()
        self.prev_company = ""
        self["textvariable"] = self.company 

        self.text_trace = self.company.trace_add("write", self.show_suggestions)
        self.bind("<Return>", self.on_select)
        self.bind("<Up>", lambda *_: self.listbox_move("up"))
        self.bind("<Down>", lambda *_: self.listbox_move("down"))
        self.bind("<Escape>", self.close_listbox)
        self.tree.bind("<ButtonPress-1>", self.on_row_click, True)
        #self.tree.bind("<Double-1>", self.open_file)


    def show_suggestions(self, *_):
        # Get current text and find matches
        text = self.company.get()
        if not text:
            self.close_listbox()
            return

        matches = [w for w in self.company_ids if w[0].lower().startswith(text.lower()) and (not self.root.ignoring or not w[2])]
        if not matches:
            self.close_listbox()
            return

        if self.listbox is None:
            self.listbox = ttk.Treeview(self.root, columns=("id", "name"), show="tree", height=8)
            self.listbox.heading("id", text="ID")
            self.listbox.heading("name", text="Name")
            self.listbox.bind("<ButtonRelease-1>", self.on_select)
            self.listbox.bind("<Return>", self.on_select)
            self.listbox.bind("<Up>", lambda e: self.listbox_move("up"))
            self.listbox.bind("<Down>", lambda e: self.listbox_move("down"))

        self.listbox.delete(*self.listbox.get_children())
        matches.sort()  # Sort matches alphabetically
        for w in matches:
            self.listbox.insert("", tk.END, values=(w[0], w[1]))

        # position the listbox just under the entry widget
        x = self.winfo_x()
        y = self.winfo_y() + self.winfo_height() + 7
        self.listbox.place(x=x, y=y)


    def on_select(self, *_, source=None):
        # Get selected company and update treeview
        if not self.root.all_companies.get():
            if self.listbox: 
                selection = self.listbox.selection()
                if selection:
                    self.company.set(self.listbox.item(selection[0], "values")[0])
                else:
                    items = self.listbox.get_children()
                    if len(items) == 1:
                        self.company.set(self.listbox.item(items[0], "values")[0])

            company = self.company.get()
            if not any(company in tup for tup in self.company_ids):   
                self.tree.delete(*self.tree.get_children())
                return
        else:
            company = self.company.get()
            
        if self.root.ignoring and company in self.root.ignore_list:
            self.tree.delete(*self.tree.get_children())
            return
        invoice_prefix = self.root.invoice_text.get()
        account_filter = self.root.account_text.get()

        # If user adds text, just need to filter not re add all rows
        narrow = False
        if ((source == "company" and company.startswith(self.prev_company) and not company == self.prev_company) or
            (source == "invoice" and invoice_prefix.startswith(self.root.prev_invoice_text) and not invoice_prefix == self.root.prev_invoice_text) or
            (source == "account" and account_filter.startswith(self.root.prev_account_text) and not account_filter == self.root.prev_account_text)):
                narrow = True
        self.prev_company = company
        self.root.prev_invoice_text = invoice_prefix
        self.root.prev_account_text = account_filter
        
        if narrow:
            # Filter current rows
            invoice_count = self.root.filter_rows(company, invoice_prefix, account_filter)
        else:
            self.tree.delete(*self.tree.get_children())
            # Update treeview with invoices for selected company
            invoice_count, values = self.root.show_invoices(company, invoice_prefix, account_filter)
            # Resort
            self.root.sort_by(self.root.sort_col, values, False)
        

        if invoice_count == 0:
            self.root.amount_label.config(text="No invoices found.")
        elif invoice_count == 1:
            self.root.amount_label.config(text="1 invoice found.")
        else:
            self.root.amount_label.config(text=f"{invoice_count} invoices found.")
        self.close_listbox()


    def listbox_move(self, dir):
        if not self.listbox:
            return
        
        dir = 1 if dir == "down" else -1
        
        rows = self.listbox.get_children()
        if not rows:
            return

        current = self.listbox.selection()
        i = rows.index(current[0]) if current else -1
        i = (i + dir) % len(rows)
        self.listbox.selection_set(rows[i])
        self.listbox.focus(rows[i])
        self.listbox.see(rows[i])


    def on_row_click(self, event):
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
    
        row = self.tree.identify_row(event.y)
        if not row:
            return
        
        col_num = self.tree.identify_column(event.x)
        col = self.tree.heading(col_num)["text"]

        if col == "GL Account" or col == "GL Account  ▲" or col == "GL Account  ▼":
            self.toggle_gl_accounts(row)
            return "break"
        elif col == "Check Number" or col == "Check Number  ▲" or col == "Check Number  ▼":
            self.toggle_checks(row)
            return "break"
        else:
            self.open_file(event)


    def toggle_checks(self, row):
        subrows = self.tree.get_children(row)
        has_check = False
        for subrow in subrows:
            if self.tree.set(subrow, "Check Number"):
                has_check = True
                break

        if has_check:
            is_open = self.tree.item(row, "open")
            if is_open:
                self.tree.set(row, "Check Number", "▼")
                if self.tree.set(row, "GL Account") in ("▲", "▼"):
                    self.tree.set(row, "GL Account", "▼")
            else:
                self.tree.set(row, "Check Number", "▲")
                if self.tree.set(row, "GL Account") in ("▲", "▼"):
                    self.tree.set(row, "GL Account", "▲")
            self.tree.item(row, open=not is_open)


    def toggle_gl_accounts(self, row):
        subrows = self.tree.get_children(row)
        has_gl_account = False
        for subrow in subrows:
            if self.tree.set(subrow, "GL Account"):
                has_gl_account = True
                break

        if has_gl_account:
            is_open = self.tree.item(row, "open")
            if is_open:
                self.tree.set(row, "GL Account", "▼")
                if self.tree.set(row, "Check Number") in ("▲", "▼"):
                    self.tree.set(row, "Check Number", "▼")
            else:
                self.tree.set(row, "GL Account", "▲")
                if self.tree.set(row, "Check Number") in ("▲", "▼"):
                    self.tree.set(row, "Check Number", "▲")
            self.tree.item(row, open=not is_open)
    

    def toggle_all_companies(self):
        if self.root.all_companies.get():
            self.company.trace_remove("write", self.text_trace)
            self.text_trace = self.company.trace_add("write", lambda *_: self.on_select(source="company"))
        else:
            self.company.trace_remove("write", self.text_trace)
            self.text_trace = self.company.trace_add("write", self.show_suggestions)
        self.on_select()


    def close_listbox(self, *_):
        if self.listbox:
            self.listbox.destroy()
            self.listbox = None

    
    def open_file(self, event):
        try:
            selection = self.tree.selection()
            if not selection:
                return
            
            if self.tree.identify_region(event.x, event.y) != "cell":
                return
        
            row = self.tree.identify_row(event.y)
            if not row:
                return
            
            if selection[0] != row:
                return
            
            item = selection[0]
            filepath = self.tree.set(item, "Filepath")
            if not filepath:
                return
            if not os.path.exists(filepath):
                messagebox.showerror("Error", "File not found.")
                return
            
            os.startfile(filepath)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {e}")


class ErrorPopup(tk.Toplevel):
    def __init__(self, root, terrors, ierrors, missing:list[tuple], **kw):
        super().__init__(root, **kw)
        self.root = root
        self.title("Error Page")
        self.wm_attributes("-toolwindow", True)
        self.withdraw()
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.toggle)
        self.bind("<space>", self.toggle)
        self.geometry("750x400")

        frame = tk.Frame(self)
        frame.pack(expand=True, fill="both")

        self.text = scrolledtext.ScrolledText(frame, font=font.Font(family="Consolas", size=9), wrap="none")
        self.text.pack(expand=True, fill="both")
        hscroll = tk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        hscroll.pack(side="bottom", fill="x")
        self.text.configure(xscrollcommand=hscroll.set)
        self.text.tag_configure("bold", font=font.Font(family="Consolas", size=12, weight="bold"))
        
        # Titan errors
        terrors.sort(key=lambda x: x["VendorID"])
        self.text.insert(tk.END, f" {len(terrors)} Titan Invoice Errors\n", ("bold",))
        for row in terrors:
            self.text.insert(tk.END, f" -{row}\n")

        # File errors
        self.text.insert(tk.END, f"\n {len(ierrors)} Invoice File Errors\n", ("bold",))
        for row in ierrors:
            self.text.insert(tk.END, f" -{row}\n")

        # Missing invoice files
        missing.sort(key=lambda x: (x[0], -datetime.strptime(x[2], "%m-%d-%Y").timestamp()))
        self.text.insert(tk.END, f"\n {len(missing)} Invoices Missing Files\n", ("bold",))
        for row in missing:
            self.text.insert(tk.END, f" -{row}\n")

    
    def toggle(self, *_):
        if self.state() == "withdrawn":
            self.show()
        else:
            self.withdraw()


    def show(self):
        # center on top of anchor window
        ax, ay = self.root.winfo_rootx(), self.root.winfo_rooty()
        aw, ah = self.root.winfo_width(), self.root.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        x = ax + (aw - w) // 2
        y = ay + (ah - h) // 3
        self.geometry(f"+{x}+{y}")
        self.deiconify()


class HelpPopup(tk.Toplevel):
    def __init__(self, root, **kw):
        super().__init__(root, **kw)
        self.root = root
        self.title("Help Page")
        self.wm_attributes("-toolwindow", True)
        self.withdraw()
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.toggle)
        self.bind("<space>", self.toggle)
        self.geometry("700x600")

        frame = tk.Frame(self)
        frame.pack(expand=True, fill="both")

        self.text = scrolledtext.ScrolledText(frame, font=font.Font(family="Consolas", size=9), wrap="none")
        self.text.pack(expand=True, fill="both")

        
        self.text.tag_configure("header", font=font.Font(family="Consolas", size=10, weight="bold"), spacing1=8)
        self.text.tag_configure("title",  font=font.Font(family="Consolas", size=12, weight="bold"), spacing1=4)
        self.text.tag_configure("body",   font=font.Font(family="Consolas", size=9), lmargin1=10, lmargin2=10)
        self.text.tag_configure("indent", font=font.Font(family="Consolas", size=9), lmargin1=26, lmargin2=26)
        self.text.tag_configure("footer", font=font.Font(family="Consolas", size=9), spacing1=10)

        def h(text): self.text.insert(tk.END, text + "\n", "header")
        def b(text): self.text.insert(tk.END, text + "\n", "body")
        def i(text): self.text.insert(tk.END, text + "\n", "indent")
        def f(text): self.text.insert(tk.END, text + "\n", "footer")

        self.text.insert(tk.END, "Titan Invoice Viewer — Help\n", "title")

        h("GETTING STARTED")
        b("When the program opens it loads all invoice data from the Titan database and scans")
        b("the invoice file directory. This takes a few seconds. Once loading is complete the")
        b("main invoice table and search bar will appear.")
        b("• Press the ⭮ Refresh button (top-right) at any time to reload the latest data.")
        b("• Press the Errors button to see any invoices or files that had problems loading.")

        h("SEARCHING FOR INVOICES")
        b("Company ID  — Type a vendor ID into the Company ID box. A suggestion list will")
        b("appear; click a result or press Enter to load that vendor's invoices.")
        b("")
        b("Search All Companies  — Check this box to show invoices across all vendors at once.")
        b("In this mode the Company ID box becomes a prefix filter: typing 'AC' shows every")
        b("vendor whose ID starts with 'AC', rather than requiring an exact match.")
        b("")
        b("Invoice  — Type in the Invoice box to narrow results to invoices whose number")
        b("starts with your entry.")
        b("")
        b("Account  — Filter by GL account number or description. Entering digits matches")
        b("account numbers starting with those digits. Entering text matches any account")
        b("whose number or description contains that text.")
        b("")
        b("Plant  — Filter invoices by which plant they belong to:")
        i("Both  — Show invoices from all plants (default)")
        i("ACP   — Show only Atlantic Concrete invoices (Plant ID 110)")
        i("APC   — Show only Atlantic Precast invoices (Plant ID 410)")
        b("")
        b("Start / End Date  — Only invoices within this date range will be shown. The")
        b("Date Filter dropdown controls what kind of date is being filtered:")
        i("Invoice Date  — Filters by the date printed on the invoice (default)")
        i("Check Date    — Filters by the date the payment check was issued.")
        b("                Note: in Check Date mode, invoices with no associated checks")
        b("                are hidden, and only checks within the date range are shown")
        b("                when expanding a multi-check row.")
        b("")
        b("File Available Only  — Check this box to hide any invoices that do not have a")
        b("PDF file stored on the network drive.")

        h("THE INVOICE TABLE")
        b("Each row is one invoice. The columns show:")
        i("Vendor          — The vendor ID code")
        i("Company Name    — The vendor's full company name")
        i("GL Account      — The expense account this invoice was coded to")
        i("Invoice         — The invoice number")
        i("Date            — The invoice date")
        i("Invoice Amount  — The total dollar amount of the invoice")
        i("Balance         — The remaining unpaid amount, or 'Paid In Full'")
        i("Check Number    — The check used to pay this invoice")
        i("Check Date      — The date that check was issued")
        i("File Available  — A ✔ means a PDF of the invoice is stored on the network")

        h("EXPANDING ROWS — GL ACCOUNTS AND CHECKS")
        b("Some invoices are split across multiple expense accounts or were paid by more")
        b("than one check. These rows show a ▼ arrow in the GL Account or Check Number")
        b("column.")
        b("• Click the ▼ in the GL Account column to expand and see each individual account")
        b("  line with its account number, description, and amount. Account lines are")
        b("  sorted by account number.")
        b("• Click the ▼ in the Check Number column to expand and see each check with its")
        b("  number, date, and payment amount. Checks are sorted newest first.")
        b("• Click the ▲ again to collapse the row.")

        h("OPENING INVOICE FILES")
        b("Double left-click any row that has a ✔ in the File Available column to open the")
        b("invoice PDF.")

        h("SORTING")
        b("Click any column header to sort the table by that column. Click the same header")
        b("again to reverse the sort direction. The active sort column is marked with ▲ or ▼.")

        h("SELECTING ROWS AND TOTALS")
        b("Click a row to select it. The totals bar at the bottom of the window shows:")
        i("Account Sum    — Sum of GL distribution amounts for all visible invoices")
        i("Selected Total — Sum of Invoice Amount for only the rows you have selected")
        i("Invoice Total  — Sum of Invoice Amount for all visible invoices")
        i("Balance Total  — Sum of outstanding balances for all visible invoices")
        b("To select multiple rows hold Ctrl and click individual rows, or hold Shift and")
        b("click to select a continuous range.")

        f("Questions and suggestions: jmwesthoff@atlanticconcrete.com")


    def toggle(self, *_):
        if self.state() == "withdrawn":
            self.show()
        else:
            self.withdraw()


    def show(self):
        # center on top of anchor window
        ax, ay = self.root.winfo_rootx(), self.root.winfo_rooty()
        aw, ah = self.root.winfo_width(), self.root.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        x = ax + (aw - w) // 2
        y = ay + (ah - h) // 3
        self.geometry(f"+{x}+{y}")
        self.deiconify()
    

if __name__ == "__main__":
    invoice_viewer = InvoiceViewer()
    invoice_viewer.mainloop()