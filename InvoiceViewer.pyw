from tkinter import ttk, messagebox, scrolledtext, font
from tkcalendar import DateEntry
from collections import defaultdict
from datetime import datetime
from PIL import Image, ImageTk
from concurrent.futures import ThreadPoolExecutor
import tkinter as tk
import os, pymssql, time, threading, re, queue


INVOICE_DIR = r"S:\Titan_DM\Titan_Filing\AP_Invoices"
invoice_re = re.compile(r'(\d+)')

class InvoiceViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Titan Invoice Viewer")
        self.iconbitmap( "icon.ico")
        self.geometry("1220x700")
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.gui_queue = queue.Queue()

        self.invoices = {} # List of dictionaries, {"VendorID", "InvoiceNum", "InvoiceDate", "ExtAmount", "Filepath"}
        self.company_ids = set() # Set of company IDs for quick lookup
        self.sort_col = "Date"
        self.sort_desc = True # descending
        self.broken_companies = []
        self.broken_invoices = []
        self.by_vendor_invoice = []
        self.checks_by_invoice = []
        self.missing_invoices = []

        # Loading info
        photoimage = ImageTk.PhotoImage(Image.open("logo.png").resize((800, 700)))
        self.loading_canvas = tk.Canvas(self, bg="white", width=1220, height=700)
        self.loading_canvas.pack(expand=True, fill="both", side="top", anchor="w")
        self.loading_canvas.background = photoimage
        self.loading_canvas.create_image(1220/2, 0, anchor="n", image=photoimage)
        
        tk.Label(self.loading_canvas, text="Welcome to Titan Invoice Viewer", font=("TKDefaultFont", 24, "bold"), bg="white").pack(side="top", anchor="w")
        tk.Label(self.loading_canvas, text="Please press the Help button for more information about this program", font=("TKDefaultFont", 20), bg="white").pack(side="top", anchor="w")
        tk.Label(self.loading_canvas, text="Loading Titan invoices, Please wait...", font=("TKDefaultFont", 16), bg="white").pack(side="top", anchor="w")

        self.after(0, lambda: threading.Thread(target=self.load_data, daemon=True).start())
        self.after(50, self.loading_loop)


    def loading_update(self, msg, color="#000000"):
        print(msg)
        self.gui_queue.put((msg, color))


    def loading_loop(self):
        try:
            while True:
                msg, color = self.gui_queue.get_nowait()
                tk.Label(self.loading_canvas, text=msg, font=("TKDefaultFont", 16), fg=color, bg="white").pack(side="top", anchor="w")
        except queue.Empty:
            pass
        self.after(50, self.loading_loop)


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
                self.missing_invoices.insert(0, row["VendorID"] + " - " + row["InvoiceNum"] + " - " + row["InvoiceDate"].strftime("%m-%d-%Y"))
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
        self.error_popup = ErrorPopup(self, self.broken_companies, self.broken_invoices, self.missing_invoices)
        self.help_popup = HelpPopup(self)
        self.bind("<space>", lambda *_: self.error_popup.toggle())


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
                    SELECT APH.VendorID, APH.InvoiceNum, APH.InvoiceDate, APH.Subtotal, APH.Payments, V.CompanyName
                    FROM AP_Header APH
                    JOIN Vendors V ON APH.VendorID = V.VendorId
                """)
                data = cur.fetchall()
            
            self.invoices = [row for row in data if row["VendorID"] and row["InvoiceNum"] and row["InvoiceDate"] 
                            and row["Subtotal"] is not None and row["Payments"] is not None]
            self.broken_companies = [row for row in data if not (row["VendorID"] and row["InvoiceNum"] and row["InvoiceDate"] 
                                     and row["Subtotal"] is not None and row["Payments"] is not None)]
            self.company_ids = {(row["VendorID"], row["CompanyName"]) for row in self.invoices if row["VendorID"] and row["CompanyName"]}
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
            self.checks_by_invoice = defaultdict(list)
            for row in data:
                self.checks_by_invoice[row["InvoiceNum"]].append((row["CheckNum"], row["CheckDate"], row["Amount"], row["VendorID"]))
            conn.close()
            t1 = time.perf_counter()
            self.loading_update(f"Check data loaded in {t1 - t0:.2f} seconds.")
            conn.close()
        
        
        with ThreadPoolExecutor(max_workers=2) as pool:
            _ = pool.submit(load_header)
            _ = pool.submit(load_checks)
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

        # Company Search bar
        ttk.Label(self.filter_frame, text="Company ID:").grid(row=0, column=0, padx=5)
        self.company_entry = AutoCompleteEntry(self)
        self.company_entry.grid(row=0, column=1, padx=5, pady=5)

        # Date ranges
        ttk.Label(self.filter_frame, text="Start date:").grid(row=0, column=2, padx=5)
        ttk.Label(self.filter_frame, text="End date:").grid(row=0, column=4, padx=5)

        self.start_entry = DateEntry(self.filter_frame, width=10, date_pattern="mm/dd/yyyy")
        self.end_entry = DateEntry(self.filter_frame, width=10, date_pattern="mm/dd/yyyy")
        self.start_entry.set_date("01/01/2014")
        self.start_entry.grid(row=0, column=3, padx=5)
        self.end_entry.grid(row=0, column=5, padx=5)
        self.start_entry.bind("<<DateEntrySelected>>", self.company_entry.on_select)
        self.end_entry.bind("<<DateEntrySelected>>", self.company_entry.on_select)

        # PDF Only Checkbox
        self.pdf_only = tk.BooleanVar()
        self.pdf_cb = ttk.Checkbutton(self.filter_frame, text="File Available Only", variable=self.pdf_only, command=self.company_entry.on_select, takefocus=False)
        self.pdf_cb.grid(row=0, column=6, padx=5)

        # All companies checkbox
        self.all_companies = tk.BooleanVar()
        self.all_companies_cb = ttk.Checkbutton(self.filter_frame, text="Search All Companies", variable=self.all_companies, command=self.company_entry.toggle_all_companies, takefocus=False)
        self.all_companies_cb.grid(row=0, column=7, padx=5)

        # Help button
        self.help_button = tk.Button(self.filter_frame, text="Help", command=lambda *_: self.help_popup.toggle())
        self.help_button.place(x=1175, y=5)

        # Invoice search
        ttk.Label(self.filter_frame, text="Invoice:").grid(row=1, column=0, padx=5)
        self.invoice_entry = tk.Entry(self.filter_frame)
        self.invoice_entry.grid(row=1, column=1, padx=5)
        self.invoice_text = tk.StringVar()
        self.prev_invoice_text = ""
        self.invoice_entry["textvariable"] = self.invoice_text
        self.invoice_text.trace_add("write", lambda *_: self.company_entry.on_select(source="invoice"))

        # Amount found label
        self.amount_label = ttk.Label(self.filter_frame, text="")
        self.amount_label.grid(row=1, column=2, padx=5, columnspan=6, sticky="w")

        # Invoice and Balance totals
        self.invoice_total = tk.StringVar()
        self.invoice_total.set("Total: $0")
        self.invoice_total_label = ttk.Label(self.filter_frame, textvariable=self.invoice_total)
        self.invoice_total_label.place(x=600, y=40)

        self.balance_total = tk.StringVar()
        self.balance_total.set("Total: $0")
        self.balance_total_label = ttk.Label(self.filter_frame, textvariable=self.balance_total)
        self.balance_total_label.place(x=725, y=40)


    def create_treeview(self):
        self.tree_frame = ttk.Frame(self, height=600)
        self.tree_frame.grid(row=1, column=0, sticky="nsew")

        # Column setup
        self.tree = ttk.Treeview(self.tree_frame, columns=("Vendor", "Company Name", "Invoice", "Date", "Invoice Amount", "Balance", 
                                                           "Check Number", "Check Date", "File Available", "Filepath"), show='tree headings')
        self.tree.column("#0", width=0, stretch=False)
        self.tree.column("Vendor", width=50, anchor="center")
        self.tree.column("Company Name", width=110, anchor="center")
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

        # Styles
        self.style.configure("Treeview", rowheight=20) 
        self.tree.tag_configure("oddrow",  background="#f7f7f7")
        self.tree.tag_configure("evenrow", background="#ffffff")
        self.tree.tag_configure("checkrow", background="#fdfaf1")
        
    
    def show_invoices(self, company, invoice_prefix):
        invoice_count = 0
        invoice_total = 0
        balance_total = 0
        values = []

        for entry in self.invoices:
            vendor = entry["VendorID"]
            
            if self.all_companies.get() and not vendor.lower().startswith(company.lower()):
                continue
            if not self.all_companies.get() and vendor != company:
                continue

            # Check if invoice date is between start and end dates
            date = entry["InvoiceDate"].date()
            if date < self.start_entry.get_date() or date > self.end_entry.get_date():
                continue

            # Check if Has File Only is checked
            filepath = entry.get("Filepath", "")
            has_filepath = "✔" if filepath else ""
            if self.pdf_only.get() == 1 and not has_filepath:
                continue

            invoice = entry["InvoiceNum"]
            if not invoice.lower().startswith(invoice_prefix.lower()):
                continue
            company_name = entry["CompanyName"]
            amount = entry['Subtotal']
            check_number = ""
            check_date= ""

            # Add subrows for checks
            checks = self.checks_by_invoice[invoice]
            if len(checks) == 1:
                check_number, check_date, _, _ = checks[0]
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

            values.append((vendor, company_name, invoice, date, amount, balance, check_number, check_date, has_filepath, filepath))
            invoice_count += 1

        invoice_total =  f"${invoice_total:,.2f}" if invoice_total >= 0 else f"(${abs(invoice_total):,.2f})"
        balance_total = f"${balance_total:,.2f}" if balance_total >= 0 else f"(${abs(balance_total):,.2f})"
        self.invoice_total.set(f"Total: {invoice_total}")
        self.balance_total.set(f"Total: {balance_total}")

        return invoice_count, values


    def filter_rows(self, company, invoice_prefix): # Filter current rows based on company name and invoice
        invoice_total = 0
        balance_total = 0
        for row in self.tree.get_children():
            values = self.tree.item(row, "values")
            if not values[0].lower().startswith(company.lower()) or not values[2].lower().startswith(invoice_prefix.lower()):
                self.tree.delete(row)
            else:
                invoice_total += float(values[4].replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
                if values[5] != "Paid In Full":
                    balance_total += float(values[5].replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
        invoice_total = f"${invoice_total:,.2f}" if invoice_total >= 0 else f"(${abs(invoice_total):,.2f})"
        balance_total = f"${balance_total:,.2f}" if balance_total >= 0 else f"(${abs(balance_total):,.2f})"
        self.invoice_total.set(f"Total: {invoice_total}")
        self.balance_total.set(f"Total: {balance_total}")
        
        return len(self.tree.get_children())


    def sort_by(self, col, values=None):
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
            "Invoice": lambda x: invoice_key(str(x[2])),
            "Date": lambda x: x[3],
            "Invoice Amount": lambda x: float(x[4].replace("$", "").replace("(", "-").replace(",", "").replace(")", "")),
            "Balance": lambda x: float(x[5].replace("$", "").replace("(", "-").replace(",", "").replace(")", "")) if x[5] != "Paid In Full" else 0,
            "Check Number": lambda x: x[6],
            "Check Date": lambda x: datetime.strptime(x[7], "%m-%d-%Y")  if x[7] else datetime(2000, 1, 1),
            "File Available": lambda x: x[8]
        }
        reverse = self.sort_desc
        if col == "Vendor" or col == "Invoice":
            reverse = not reverse

        values.sort(key=keymap[col], reverse=reverse)

        for i, row in enumerate(values):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            iid = self.tree.insert("", "end", values=row, tags=tag)

            # Add subrows for checks
            checks = self.checks_by_invoice[row[2]]
            if len(checks) > 1:
                for cnum, cdate, camt, ven in checks:
                    if row[0] != ven:
                        continue
                    self.tree.set(iid, "Check Number", "▼")
                    cdate = cdate.strftime("%m-%d-%Y")
                    camt = f"${camt:,.2f}" if camt >= 0 else f"(${abs(camt):,.2f})"
                    self.tree.insert(iid, "end", values=("", "", "", "", "", camt, cnum, cdate, "", ""), tags="checkrow")

        arrow = "  ▼" if self.sort_desc else "  ▲"
        for c in self.tree["columns"]:
            text = c + arrow if c == col else c
            self.tree.heading(c, text=text)
        return "break"

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
        self.tree.bind("<ButtonPress-1>", self.toggle_checks, True)
        self.tree.bind("<Double-1>", self.open_file)


    def show_suggestions(self, *_):
        # Get current text and find matches
        text = self.company.get()
        if not text:
            self.close_listbox()
            return

        matches = [w for w in self.company_ids if w[0].lower().startswith(text.lower())]
        if not matches:
            self.close_listbox()
            return

        if self.listbox is None:
            self.listbox = ttk.Treeview(self.root, columns=("id", "name"), show="tree", height=8)
            self.listbox.column("#0", width=0, stretch=False)
            self.listbox.column("id", width=80)
            self.listbox.column("name", width=300)
            self.listbox.bind("<ButtonRelease-1>", self.on_select)
            self.listbox.bind("<Return>", self.on_select)
            self.listbox.bind("<Up>", lambda: self.listbox_move("up"))
            self.listbox.bind("<Down>", lambda: self.listbox_move("down"))

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
            if company not in dict(self.company_ids):
                return
        else:
            company = self.company.get()
        invoice_prefix = self.root.invoice_text.get()

        # If user adds text, just need to filter not re add all rows
        narrow = False
        if ((source == "company" and company.startswith(self.prev_company)) or
            (source == "invoice" and invoice_prefix.startswith(self.root.prev_invoice_text))):
                narrow = True
        self.prev_company = company
        self.root.prev_invoice_text = invoice_prefix
        
        if narrow:
            # Filter current rows
            invoice_count = self.root.filter_rows(company, invoice_prefix)
        else:
            # Update treeview with invoices for selected company
            invoice_count, values = self.root.show_invoices(company, invoice_prefix)
            # Resort
            self.root.sort_by(self.root.sort_col, values)
        

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


    def toggle_checks(self, event):
        row = self.tree.identify_row(event.y)
        if not row:
            return
        
        col_num = self.tree.identify_column(event.x)
        col = self.tree.heading(col_num)["text"]
        if col != "Check Number" and col != "Check Number  ▲" and col != "Check Number  ▼":
            return
        
        if self.tree.get_children(row):
            is_open = self.tree.item(row, "open")
            self.tree.item(row, open=not is_open)
        return "break"
    

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
        return "break"


class ErrorPopup(tk.Toplevel):
    def __init__(self, root, terrors, ierrors, missing, **kw):
        super().__init__(root, **kw)
        self.root = root
        self.title("Error Page")
        self.wm_attributes("-toolwindow", True)
        self.withdraw()
        self.attributes("-topmost", True)
        self.protocol("WM_DELETE_WINDOW", self.toggle)
        self.bind("<space>", self.toggle)

        frame = tk.Frame(self)
        frame.pack(expand=True, fill="both")

        self.text = scrolledtext.ScrolledText(frame, font=font.Font(family="Consolas", size=9), wrap="none")
        self.text.pack(expand=True, fill="both")
        hscroll = tk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        hscroll.pack(side="bottom", fill="x")
        self.text.configure(xscrollcommand=hscroll.set)
        self.text.tag_configure("bold", font=font.Font(family="Consolas", size=12, weight="bold"))
        
        # Titan errors
        self.text.insert(tk.END, f" {len(terrors)} Titan Invoice Errors\n", ("bold",))
        for row in terrors:
            self.text.insert(tk.END, f" -{row}\n")

        # File errors
        self.text.insert(tk.END, f"\n {len(ierrors)} Invoice File Errors\n", ("bold",))
        for row in ierrors:
            self.text.insert(tk.END, f" -{row}\n")

        # Missing invoice files
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

        frame = tk.Frame(self)
        frame.pack(expand=True, fill="both")

        about_text = """Welcome to Titan Invoice Viewer, here are some useful tips for operating this program
- Enter a company ID into the search bar to view all invoices for that company
- Selecting "Search All Companies" will show every invoice
- It also turns the Company ID bar to a filter only showing companies starting with your entry
- The Invoice bar will only show invoices starting with your entry
- Double left-click a row to open the invoice file if it is available
- Left-click a column header to sort by that column, click again to swap direction
- Left-click the ▼ in the Check Number column to display all associated checks
- Press space to open and close an error tab to see broken Titan entries and invoice files
- All questions and suggestions can be directed to jmwesthoff@atlanticconcrete.com"""
        tk.Label(frame, text=about_text, justify="left", font=("Consolas", 12)).pack(anchor="w")

    
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