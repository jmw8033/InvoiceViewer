import tkinter as tk
from tkinter import ttk, messagebox, font
from tkcalendar import DateEntry
from collections import defaultdict
from datetime import datetime, date
import os
import pymssql
import time

INVOICE_DIR = r"S:\Titan_DM\Titan_Filing\AP_Invoices"

class InvoiceViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Titan Invoice Viewer")
        self.geometry("900x600")
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.invoices = {} # List of dictionaries, {"VendorID", "InvoiceNum", "InvoiceDate", "ExtAmount", "Filepath"}
        self.company_ids = set() # Set of company IDs for quick lookup
        self.error_count = 0
        self.sort_col = "Date"
        self.sort_desc = True # descending

        # Load Data
        print("Loading data...")
        start_time = time.perf_counter()
        self.load_database()
        self.load_files()
        end_time = time.perf_counter()

        # Check for errors
        if self.error_count > 0:
            print(f"{self.error_count} errors.")
        print(f"Data loaded in {end_time - start_time:.2f} seconds.")

        # Create Frames
        self.create_treeview()
        self.create_filter_frame()
        self.create_footer_frame()

        self.mainloop()


    def load_database(self):
        # Connect to the database and fetch invoice data
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
                ORDER BY APH.VendorID, APH.InvoiceDate, APH.InvoiceNum
            """)
            data=cur.fetchall()
        
        self.invoices = [row for row in data if row["VendorID"] and row["InvoiceNum"] and row["InvoiceDate"] 
                          and row["Subtotal"] is not None and row["Payments"] is not None]
        self.broken_companies = [row for row in data if not (row["VendorID"] and row["InvoiceNum"] and row["InvoiceDate"] 
                                 and row["Subtotal"] is not None and row["Payments"] is not None)]
        [print(row) for row in self.broken_companies]
        print(f"{len(self.broken_companies)} broken entries found.")
        self.company_ids = {(row["VendorID"], row["CompanyName"]) for row in self.invoices if row["VendorID"] and row["CompanyName"]}
        self.by_vendor_invoice = {(row["VendorID"], row["InvoiceNum"]): row for row in self.invoices}

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


    def load_files(self):
        # Scan the invoice directory and match files to database entries
        for file in os.listdir(INVOICE_DIR):
            try:
                fname = file.split("_") # COMPANY_INVOICE_MM-DD-YYYY_RANDOMINT
                if fname and len(fname) >= 3:
                    if fname[0] == "PUCA": # Special case, they have underscore in company ID
                        fname[0] = "PUCA_150"
                        fname[1] = fname[2] # Invoice number

                    fname[1] = fname[1].replace("[slash]", "/").replace("[quote]", '"')
                    key = (fname[0], fname[1])
                    row = self.by_vendor_invoice.get(key)
                    if row:
                        row["Filepath"] = os.path.join(INVOICE_DIR,  file)
                    else:
                        print(f"No match found for file: {file}, {fname}")
                        self.error_count += 1
            except Exception as e:
                print(f"Parsing failed for file {fname} - {e}")
                self.error_count += 1
                continue


    def create_filter_frame(self):
        self.filter_frame = ttk.Frame(self, border=2)
        self.filter_frame.pack(fill='x', padx=10, pady=10, side="top")

        # Search bar
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
        self.checkbutton = ttk.Checkbutton(self.filter_frame, text="File Available Only", variable=self.pdf_only, command=self.company_entry.on_select, takefocus=False)
        self.checkbutton.grid(row=0, column=6, padx=5)

        # Amount found label
        self.amount_label = ttk.Label(self.filter_frame, text="")
        self.amount_label.grid(row=1, column=0, padx=5, columnspan=3, sticky="w")



    def create_treeview(self):
        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(expand=True, fill='both', padx=10, pady=5, side="bottom")

        # Column setup
        self.tree = ttk.Treeview(self.tree_frame, columns=("Invoice", "Date", "Invoice Amount", "Balance", 
                                                           "Check Number", "Check Date", "File Available", "Filepath"), show='tree headings')
        self.tree.column("#0", width=0, stretch=False)
        self.tree.column("Invoice", width=110, anchor="center")
        self.tree.column("Date", width=40, anchor="center")
        self.tree.column("Invoice Amount", width=50, anchor="center")
        self.tree.column("Balance", width=50, anchor="center")
        self.tree.column("Check Number", width=70, anchor="center")
        self.tree.column("Check Date", width=40, anchor="center")
        self.tree.column("File Available", width=30, anchor="center")
        self.tree.column("Filepath", width=0, stretch=False)

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


    def create_footer_frame(self):
        pass

    
    def sort_by(self, col):
        if col == self.sort_col:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_col = col
            self.sort_desc = True

        items = [(k, (self.tree.set(k, self.sort_col))) for k in self.tree.get_children()]
        if self.sort_col == "Invoice Amount" or self.sort_col == "Balance":
            items.sort(key=lambda x: float(x[1].replace("$", "").replace("(", "-").replace(",", "").replace(")", ""))
                       if x[1] != "Paid In Full" else 0, reverse=self.sort_desc)
        elif self.sort_col == "Date" or self.sort_col == "Check Date":
            items.sort(key=lambda x: datetime.strptime(x[1], "%m-%d-%Y") 
                       if x[1] else datetime(2000, 1, 1), reverse=self.sort_desc)
        else:
            items.sort(key=lambda x: x[1], reverse=self.sort_desc)

        for i, (iid, val) in enumerate(items):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.move(iid, "", i)
            self.tree.item(iid, tags=tag)

        arrow = "  ▼" if self.sort_desc else "  ▲"
        for c in self.tree["columns"]:
            text = c + arrow if c == col else c
            self.tree.heading(c, text=text)


class AutoCompleteEntry(tk.Entry):
    def __init__(self, root: InvoiceViewer, *a, **kw):
        super().__init__(root.filter_frame, *a, **kw)
        self.invoices = root.invoices
        self.company_ids = root.company_ids
        self.tree = root.tree
        self.root = root
        self.listbox = None
        self.text = tk.StringVar()
        self["textvariable"] = self.text 

        self.text.trace_add("write", self.show_suggestions)
        self.bind("<Return>", self.on_select)
        self.bind("<Up>", lambda x: self.listbox_move("up"))
        self.bind("<Down>", lambda x: self.listbox_move("down"))
        self.bind("<Escape>", self.close_listbox)
        self.tree.bind("<ButtonPress-1>", self.toggle_checks, True)
        self.tree.bind("<Double-1>", self.open_file)


    def show_suggestions(self, *_):
        # Get current text and find matches
        text = self.text.get()
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
            self.listbox.bind("<ButtonPress-1>", self.on_select)
            self.listbox.bind("<Return>", self.on_select)
            self.listbox.bind("<Up>", lambda: self.listbox_move("up"))
            self.listbox.bind("<Down>", lambda: self.listbox_move("down"))

        self.listbox.delete(*self.listbox.get_children())
        matches.sort()  # Sort matches alphabetically
        for w in matches:
            self.listbox.insert("", tk.END, values=(w[0], w[1]))

        # position the listbox just under the entry widget
        x = self.winfo_x() + 10
        y = self.winfo_y() + self.winfo_height() + 11
        self.listbox.place(x=x, y=y)


    def on_select(self, *_):
        # Get selected company and update treeview
        if self.listbox: 
            selection = self.listbox.selection()
            if selection:
                self.text.set(self.listbox.item(selection[0], "values")[0])
            else:
                items = self.listbox.get_children()
                if len(items) == 1:
                    self.text.set(self.listbox.item(items[0], "values")[0])

        company = self.text.get()
        if company not in dict(self.company_ids):
            return
        
        # Update treeview with invoices for selected company
        invoice_count = 0
        for row in self.tree.get_children():
            self.tree.delete(row)
        for i, entry in enumerate(self.invoices):
            if entry["VendorID"] != company:
                continue
            try:
                # Check if invoice date is between start and end dates
                date = entry["InvoiceDate"].date()
                if date < self.root.start_entry.get_date() or date > self.root.end_entry.get_date():
                    continue
                date = date.strftime("%m-%d-%Y")

                # Check if Has File Only is checked
                filepath = entry.get("Filepath", "")
                has_filepath = "✔" if filepath else ""
                if self.root.pdf_only.get() == 1 and not has_filepath:
                    continue

                invoice = entry["InvoiceNum"]
                amount = entry['Subtotal']
                check_number = ""
                check_date= ""
                
                checks = self.root.checks_by_invoice[invoice]
                if len(checks) == 1:
                    check_number, check_date, _, _ = checks[0]
                    check_date = check_date.strftime("%m-%d-%Y")

                # Get Balance
                payments = entry["Payments"]
                if payments == amount:
                    balance = "Paid In Full"
                else:
                    balance = amount - payments
                    balance = f"${balance:,.2f}" if balance >= 0 else f"(${abs(balance):,.2f})"
                amount = f"${amount:,.2f}" if amount >= 0 else f"(${abs(amount):,.2f})"
                    
                # If duplicate invoice, add random int to end of invoice
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                iid = f"{invoice}_{os.urandom(4).hex()}"
                invoice = self.tree.insert("", "end", iid=iid, 
                                 values=(invoice, date, amount, balance, check_number, check_date,
                                        has_filepath, filepath), tags=tag)
                
                # Add subrows for checks
                if len(checks) > 1:
                    self.tree.set(iid, "Check Number", "▼")
                    for number, date, amount, vendor in checks:
                        if vendor != company:
                            continue
                        date = date.strftime("%m-%d-%Y")
                        amount = f"${amount:,.2f}" if amount >= 0 else f"(${abs(amount):,.2f})"
                        self.tree.insert(invoice, "end", values=("", "", "", amount, number, date, ""), tags="checkrow")
                invoice_count += 1
            except Exception as e:
                print(f"Error processing entry {entry}: {e}")

        if invoice_count == 0:
            self.root.amount_label.config(text="No invoices found.")
        elif invoice_count == 1:
            self.root.amount_label.config(text="1 invoice found.")
        else:
            self.root.amount_label.config(text=f"{invoice_count} invoices found.")
        self.close_listbox()

        # Resort
        self.root.sort_desc = not self.root.sort_desc
        self.root.sort_by(self.root.sort_col)


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
        
        if self.tree.get_children(row):
            is_open = self.tree.item(row, "open")
            self.tree.item(row, open=not is_open)
        return "break"


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

if __name__ == "__main__":
    InvoiceViewer = InvoiceViewer()
