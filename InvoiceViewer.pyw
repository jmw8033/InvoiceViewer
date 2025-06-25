import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import os, re, datetime as dt
from dateutil import parser
import pymssql
import time

INVOICE_DIR = r"S:\Titan_DM\Titan_Filing\AP_Invoices"

class InvoiceViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Invoice Viewer")
        self.geometry("800x600")

        # Load Data
        print("Loading data...")
        start_time = time.perf_counter()
        self.error_count = 0

        self.companies = {} # List of dictionaries, {"VendorID", "InvoiceNum", "InvoiceDate", "ExtAmount", "Filepath"}
        self.company_ids = set() # Set of company IDs for quick lookup
        self.load_database()
        self.load_files()

        end_time = time.perf_counter()
        if self.error_count > 0:
            print(f"{self.error_count} errors.")
        print(f"Data loaded in {end_time - start_time:.2f} seconds.")

        # Create Filter Frame
        self.filter_frame = ttk.Frame(self)
        self.filter_frame.pack(fill='x', padx=10, pady=10)

        # Create Treeview Frame
        self.tree_frame = ttk.Frame(self)
        self.tree_frame.pack(expand=True, fill='both', padx=10, pady=5)

        self.create_treeview()
        self.create_filter_frame()
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
                SELECT APH.VendorID, APH.InvoiceNum, APH.InvoiceDate, APH.Subtotal
                FROM AP_Header APH
            """)
            data=cur.fetchall()
        
        self.companies = data
        self.company_ids = {row["VendorID"] for row in self.companies if row["VendorID"]}
        self.by_vendor_invoice = {
            (row["VendorID"], row["InvoiceNum"]): row for row in self.companies
        }


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
                    row =  self.by_vendor_invoice.get(key)
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
        # Search bar
        ttk.Label(self.filter_frame, text="Company:").grid(row=0, column=0, padx=5)
        self.company_entry = AutoCompleteEntry(self)
        self.company_entry.grid(row=0, column=1, padx=5)

        # Date ranges
        ttk.Label(self.filter_frame, text="Start date:").grid(row=0, column=2, padx=5)
        ttk.Label(self.filter_frame, text="End date:").grid(row=0, column=4, padx=5)

        self.start_entry = DateEntry(self.filter_frame, width=12)
        self.end_entry = DateEntry(self.filter_frame, width=12)
        self.start_entry.set_date("01/01/2000")
        self.start_entry.grid(row=0, column=3, padx=5)
        self.end_entry.grid(row=0, column=5, padx=5)

        # Amount found label
        self.amount_label = ttk.Label(self.filter_frame, text="")
        self.amount_label.grid(row=1, column=0, padx=5, columnspan=3, sticky="w")



    def create_treeview(self):
        self.tree = ttk.Treeview(self.tree_frame, columns=("Invoice", "Date", "Amount", "HasFilepath", "Filepath"), show='headings')
        self.tree.column("Invoice", width=200, anchor="w")
        self.tree.column("Date", width=50, anchor="w")
        self.tree.column("Amount", width=50, anchor="w")
        self.tree.column("HasFilepath", width=20, anchor="w")
        self.tree.column("Filepath", width=0, stretch=False)
        self.tree.heading("Invoice", text="Invoice")
        self.tree.heading("Date", text="Date")
        self.tree.heading("Amount", text="Amount")
        self.tree.heading("HasFilepath", text="Has File")
        self.tree.heading("Filepath", text="")
        self.tree.pack(side="left", fill="both", expand=True)
        ttk.Scrollbar(self.tree_frame, command=self.tree.yview).pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=lambda f, l: self.tree.yview_moveto(f))


class AutoCompleteEntry(tk.Entry):
    def __init__(self, root, *a, **kw):
        super().__init__(root.filter_frame, *a, **kw)
        self.companies = root.companies
        self.company_ids = root.company_ids
        self.tree = root.tree
        self.root = root
        self.listbox = None
        self.text = tk.StringVar()
        self["textvariable"] = self.text 
        self.text.trace_add("write", self.show_suggestions)
        self.bind("<Return>", self.close_listbox)
        self.tree.bind("<Double-1>", self.open_file)


    def show_suggestions(self, *_):
        # Get current text and find matches
        text = self.text.get()
        if not text:
            self.close_listbox()
            return

        matches = [w for w in self.company_ids if w.lower().startswith(text.lower())]
        if not matches:
            self.close_listbox()
            return

        if self.listbox is None:
            self.listbox = tk.Listbox(self.root, height=8)
            self.listbox.bind("<<ListboxSelect>>", self.on_select)
        self.listbox.delete(0, tk.END)
        for w in matches:
            self.listbox.insert(tk.END, w)

        # position the listbox just under the entry widget
        x = self.winfo_x() + 10
        y = self.winfo_y() + self.winfo_height() + 10
        self.listbox.place(x=x, y=y, width=self.winfo_width())


    def on_select(self, *_):
        # Get selected company and update treeview
        if not self.listbox:
            return
        
        selection = self.listbox.curselection()
        if not selection:
            return
        
        self.text.set(self.listbox.get(selection[0]))
        company = self.text.get()
        if company not in self.company_ids:
            return
        
        # Update treeview with invoices for selected company
        invoice_count = 0
        for row in self.tree.get_children():
            self.tree.delete(row)
        for entry in self.companies:
            if entry["VendorID"] != company:
                continue
            
            # Check if invoice date is between start and end dates
            date = entry["InvoiceDate"]
            if not date:
                date = ""
            else:
                date = date.strftime("%m/%d/%Y")
                if date and (date < self.root.start_entry.get_date().strftime("%m/%d/%Y") or date > self.root.end_entry.get_date().strftime("%m/%d/%Y")):
                    continue

            invoice = entry["InvoiceNum"]
            amount = entry["Subtotal"]
            if not amount:
                amount = ""
            else:
                amount = f"${amount:,.2f}"
            filepath = entry.get("Filepath", "")
            has_filepath = "Yes"
            if not filepath:
                has_filepath = "No"

                
            # If duplicate invoice, add random int to end of invoice
            try:
                self.tree.insert("", "end", iid=invoice, values=(invoice, date, amount, has_filepath, filepath))
            except Exception as e:
                invoice = f"{invoice}_{os.urandom(4).hex()}"
                self.tree.insert("", "end", iid=invoice, values=(invoice, date, amount, has_filepath, filepath))
            invoice_count += 1

        if invoice_count == 0:
            self.root.amount_label.config(text="No invoices found.")
        elif invoice_count == 1:
            self.root.amount_label.config(text="1 invoice found.")
        else:
            self.root.amount_label.config(text=f"{invoice_count} invoices found.")
        self.close_listbox()


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