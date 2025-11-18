import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
import csv
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from decimal import Decimal
import os
import sys
import shutil  # Added for database backup functionality

# --- GUI Application Class ---

class ImporterApp:
    
    def __init__(self, root):
        self.root = root
        root.title("Frappe Books Importer")

        # --- Style for disclaimer ---
        self.style = ttk.Style()
        self.style.configure("Disclaimer.TLabel", foreground="#D00000", font=('Helvetica', 9, 'bold'))
        
        self.db_path = tk.StringVar()
        self.statement_path = tk.StringVar()
        self.bank_account = tk.StringVar()
        self.suspense_account = tk.StringVar()
        
        self.all_accounts = []
        self.csv_headers = []
        self.csv_guesses = {}
        
        # --- Store correct table names ---
        self.account_table_name = None
        self.ledger_table_name = None
        
        # --- CSV Mapping Vars ---
        self.csv_date_var = tk.StringVar()
        self.csv_desc_var = tk.StringVar()
        self.csv_amt_var = tk.StringVar()
        self.csv_debit_var = tk.StringVar()
        self.csv_credit_var = tk.StringVar()

        # --- Main Frame ---
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # --- 1. Database Selection ---
        db_frame = ttk.LabelFrame(main_frame, text="1. Database", padding="10")
        db_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(db_frame, text="Database File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(db_frame, textvariable=self.db_path, width=60, state='readonly').grid(row=1, column=0, padx=5)
        ttk.Button(db_frame, text="Browse...", command=self.load_db).grid(row=1, column=1, padx=5)
        
        # --- 2. Statement File Selection ---
        file_frame = ttk.LabelFrame(main_frame, text="2. Bank Statement", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(file_frame, text="Statement File (QIF, OFX, CSV):").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.statement_path, width=60, state='readonly').grid(row=1, column=0, padx=5)
        ttk.Button(file_frame, text="Browse...", command=self.load_statement).grid(row=1, column=1, padx=5)

        # --- 3. Account Mapping ---
        map_frame = ttk.LabelFrame(main_frame, text="3. Account Mapping", padding="10")
        map_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(map_frame, text="Bank/Loan Account (Debit/Credit):").grid(row=0, column=0, sticky=tk.E, padx=5)
        self.bank_menu = ttk.OptionMenu(map_frame, self.bank_account, "Load Database First", *[])
        self.bank_menu.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        ttk.Label(map_frame, text="Suspense Account:").grid(row=1, column=0, sticky=tk.E, padx=5)
        self.suspense_menu = ttk.OptionMenu(map_frame, self.suspense_account, "Load Database First", *[])
        self.suspense_menu.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # --- 4. CSV Options (Initially hidden) ---
        self.csv_frame = ttk.LabelFrame(main_frame, text="4. CSV Column Mapping", padding="10")
        # self.csv_frame will be grid()'ed later if a CSV is loaded
        
        self.csv_option_menus = {}
        csv_labels = ["Date", "Description", "Amount (Single)", "Debit (Two-Col)", "Credit (Two-Col)"]
        self.csv_vars = [self.csv_date_var, self.csv_desc_var, self.csv_amt_var, self.csv_debit_var, self.csv_credit_var]
        
        for i, label in enumerate(csv_labels):
            ttk.Label(self.csv_frame, text=f"{label} Column:").grid(row=i, column=0, sticky=tk.E, padx=5, pady=2)
            # Initialize with a valid default option
            menu = ttk.OptionMenu(self.csv_frame, self.csv_vars[i], "N/A", *["N/A"])
            menu.grid(row=i, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
            self.csv_option_menus[label] = menu
            
        # --- 5. Disclaimer Label (NEW) ---
        disclaimer_label = ttk.Label(main_frame, 
                                     text="WARNING: This tool directly modifies your database. A backup is created on load. Use at your own risk.", 
                                     style="Disclaimer.TLabel",
                                     anchor=tk.CENTER)
        disclaimer_label.grid(row=5, column=0, columnspan=2, pady=(10, 0), sticky=(tk.W, tk.E))

        # --- 6. Import (Row index changed) ---
        self.import_button = ttk.Button(main_frame, text="Import Transactions", command=self.run_import, state='disabled')
        self.import_button.grid(row=6, column=0, columnspan=2, pady=10)
        
        # --- Status Bar ---
        self.status_var = tk.StringVar(value="Ready. Load your database file.")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding=5)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Configure resizing
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1) # Allow entry/menus to expand

    # --- Logging Methods ---
    def log_error(self, message):
        self.status_var.set(f"ERROR: {message}")
        print(f"ERROR: {message}")
        messagebox.showerror("Error", message)
        
    def log_status(self, message):
        self.status_var.set(message)
        print(f"STATUS: {message}")

    # --- Database & Schema Logic ---
    
    def _find_table_name(self, conn, potential_names):
        """Helper function to find the correct, case-insensitive table name."""
        cursor = conn.cursor()
        for name in potential_names:
            try:
                # Check if table exists by querying it
                cursor.execute(f"SELECT name FROM {name} LIMIT 1")
                return name # Found it
            except sqlite3.Error:
                continue # Doesn't exist, try next
        return None

    def connect_db(self, db_path):
        """Establishes a connection to the SQLite database."""
        try:
            conn = sqlite3.connect(db_path)
            return conn
        except Exception as e:
            self.log_error(f"Error connecting to database: {e}")
            return None

    def get_accounts(self, conn):
        """
        Fetches the list of accounts from the database.
        Uses the dynamically found self.account_table_name.
        """
        accounts = []
        if not self.account_table_name:
            self.log_error("Account table name not found.")
            return []

        queries_to_try = [
            f"SELECT name FROM {self.account_table_name} WHERE type NOT IN ('Group') ORDER BY name",
            f"SELECT name FROM {self.account_table_name} ORDER BY name"
        ]
        
        cursor = conn.cursor()

        # Try to get accounts, gracefully handling if 'type' column is missing
        for query in queries_to_try:
            try:
                cursor.execute(query)
                accounts = [row[0] for row in cursor.fetchall()]
                if accounts:
                    break
            except sqlite3.Error as e:
                # This will happen if 'type' column is missing on the first query
                continue
                
        if not accounts:
            # This might happen if even the second query fails
            self.log_error("Could not read from Account table. Schema may be incorrect.")
            return []

        # Ensure a clearing account exists, create if not
        # Renamed to "Suspense Clearing" for clarity
        if "Suspense Clearing" not in accounts:
            try:
                # --- FIX: Dynamically build query based on existing columns ---
                cursor.execute(f"PRAGMA table_info({self.account_table_name});")
                account_columns = [info[1] for info in cursor.fetchall()]
                
                # Default values for a basic schema
                sql_cols_dict = {
                    "name": "Suspense Clearing",
                    "isGroup": 0,
                    "createdBy": "FrappeBooksGUIImporter",
                    "modifiedBy": "FrappeBooksGUIImporter",
                    "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "modified": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "lft": 0, # Placeholder
                    "rgt": 0  # Placeholder
                }
                
                # Add columns based on schema detection
                if 'parent' in account_columns:
                    sql_cols_dict['parent'] = "Assets" # Default guess
                if 'type' in account_columns:
                    sql_cols_dict['type'] = "Expense"
                if 'rootType' in account_columns:
                    sql_cols_dict['rootType'] = "Expense"
                if 'accountType' in account_columns:
                    sql_cols_dict['accountType'] = "Suspense"
                if 'parentAccount' in account_columns:
                    sql_cols_dict['parentAccount'] = "Current Assets" # Common default
                    
                
                # Build the final query
                col_names = ", ".join(sql_cols_dict.keys())
                q_marks = ", ".join(["?"] * len(sql_cols_dict))
                sql_vals = list(sql_cols_dict.values())
                
                sql = f"INSERT INTO {self.account_table_name} ({col_names}) VALUES ({q_marks})"
                
                cursor.execute(sql, sql_vals)
                # --- End Fix ---

                conn.commit()
                accounts.append("Suspense Clearing")
                accounts.sort()
                self.log_status("Created 'Suspense Clearing' account.")
                
            except sqlite3.Error as e:
                self.log_error(f"Failed to create Suspense Clearing account: {e}. Please create it manually in Frappe Books.")
                
        return accounts

    def check_and_fix_schema(self, conn):
        """
        Checks for 'remark', 'voucherType', 'voucherNo' columns and adds them if missing.
        Uses the dynamically found self.ledger_table_name.
        """
        if not self.ledger_table_name:
            self.log_error("Ledger table name not found.")
            return False
            
        try:
            cursor = conn.cursor()
            # Use PRAGMA to check table info
            cursor.execute(f"PRAGMA table_info({self.ledger_table_name});")
            columns = [info[1] for info in cursor.fetchall()]
            
            added_cols = []
            if 'remark' not in columns:
                cursor.execute(f"ALTER TABLE {self.ledger_table_name} ADD COLUMN remark TEXT;")
                added_cols.append('remark')
            
            if 'voucherType' not in columns:
                cursor.execute(f"ALTER TABLE {self.ledger_table_name} ADD COLUMN voucherType TEXT;")
                added_cols.append('voucherType')

            if 'voucherNo' not in columns:
                cursor.execute(f"ALTER TABLE {self.ledger_table_name} ADD COLUMN voucherNo TEXT;")
                added_cols.append('voucherNo')
            
            if added_cols:
                conn.commit()
                self.log_status(f"Database schema updated: Added column(s) {', '.join(added_cols)}.")
            else:
                self.log_status("Database schema is OK.")
            
            return True

        except sqlite3.Error as e:
            self.log_error(f"Database schema error: {e}. Could not check/fix columns.")
            return False

    # --- File Parsing Logic ---
    def parse_date(self, date_str):
        """
        Robust date parser, prioritizing Australian/European DD/MM/YYYY.
        """
        if not date_str:
            return None
        
        # Remove any extra characters (like from QIF)
        date_str = date_str.strip().replace("D", "")

        # Common formats to try
        formats_to_try = [
            '%d/%m/%Y',  # DD/MM/YYYY
            '%d/%m/%y',  # DD/MM/YY
            '%Y-%m-%d',  # YYYY-MM-DD
            '%m/%d/%Y',  # MM/DD/YYYY (US)
            '%m/%d/%y',  # MM/DD/YY (US)
            '%d %b %Y',  # 07 Sep 2025
            '%d-%b-%Y',  # 07-Sep-2025
            '%d-%b-%y',  # 07-Sep-05
            '%Y%m%d',    # YYYYMMDD (OFX)
        ]
        
        parsed_date = None
        for fmt in formats_to_try:
            try:
                # strptime converts string to datetime object
                parsed_date = datetime.strptime(date_str, fmt)
                break # Success
            except ValueError:
                continue
        
        if parsed_date:
            return parsed_date

        # --- FIX: Handle non-English locales like 'ec' for 'Dec' ---
        # This is a basic substitution. A full locale library would be overkill.
        date_str_lower = date_str.lower()
        replacements = {
            ' ec ': ' dec ', # Example: 31 ec 2024 -> 31 dec 2024
            # Add other common non-English month abbreviations if needed
        }
        for k, v in replacements.items():
            if k in date_str_lower:
                date_str = date_str_lower.replace(k, v)
                # Try parsing again with the new string
                try:
                    return datetime.strptime(date_str, '%d %b %Y')
                except ValueError:
                    pass # Failed again, will fall through to error
        # --- End Fix ---

        # Log error only if all attempts fail
        self.log_error(f"Could not parse date: {date_str}")
        return None

    def parse_qif(self, file_path):
        """
        Parses a QIF file, designed to be robust for non-standard files like myob.qif.
        """
        transactions = []
        try:
            # Use 'latin-1' encoding as a fallback, common in bank files
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()
                
        # Split by the end-of-transaction marker '^'
        raw_transactions = content.split('^')
        
        for raw_tx in raw_transactions:
            lines = raw_tx.strip().split('\n')
            if not lines or len(lines) < 2:
                continue

            current = {}
            description_parts = []
            
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                prefix = ""
                data = ""
                if line:
                   prefix = line[0].upper()
                   data = line[1:].strip()
                
                if prefix == 'D':
                    current['date'] = self.parse_date(data)
                elif prefix == 'T':
                    try:
                        current['amount'] = Decimal(data.replace(',', ''))
                    except Exception:
                        current['amount'] = Decimal(0)
                elif prefix == 'P':
                    description_parts.append(data)
                elif prefix == 'M':
                    description_parts.append(data)
                elif prefix == 'L':
                    # Handle split categories, just take the first one
                    if 'S' in data: 
                        data = data.split('S')[-1].split('E')[0]
                    description_parts.append(data)
                elif prefix in ('!', 'N'): # Type or Check Number
                    pass # Ignore
                else:
                    # Default case: Handle description lines with no prefix
                    description_parts.append(line) 

            # A valid transaction must have a date and an amount
            if current.get('date') and 'amount' in current:
                current['description'] = ' / '.join(filter(None, description_parts))
                transactions.append(current)
                
        return transactions

    def parse_ofx(self, file_path):
        """
        Parses an OFX file (v1.0 or v2.0 XML).
        """
        transactions = []
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()

        # Clean content: remove whitespace between tags
        content = re.sub(r'>\s+<', '><', content, flags=re.DOTALL)

        # --- Find the main transaction block ---
        tran_list_match = re.search(r'<BANKTRANLIST>(.*?)</BANKTRANLIST>', content, re.DOTALL | re.IGNORECASE)
        if not tran_list_match:
            self.log_status("Could not find <BANKTRANLIST> block in OFX file.")
            return transactions # No transactions found
            
        tran_list_content = tran_list_match.group(1)
        
        # Split into individual transactions
        tx_matches = re.finditer(r'<STMTTRN>(.*?)</STMTTRN>', tran_list_content, re.DOTALL | re.IGNORECASE)
        
        for tx_match in tx_matches:
            tx_content = tx_match.group(1)
            current = {}
            
            # Date
            date_match = re.search(r'<DTPOSTED>(\d{8})', tx_content, re.IGNORECASE)
            if date_match:
                current['date'] = self.parse_date(date_match.group(1))
            
            # Amount
            amt_match = re.search(r'<TRNAMT>([-\d.]+)', tx_content, re.IGNORECASE)
            if amt_match:
                current['amount'] = Decimal(amt_match.group(1))
                
            # Description (Combine NAME and MEMO)
            name_match = re.search(r'<NAME>(.*?)</NAME>', tx_content, re.IGNORECASE)
            memo_match = re.search(r'<MEMO>(.*?)</MEMO>', tx_content, re.IGNORECASE)
            
            desc_parts = []
            if name_match and name_match.group(1) and name_match.group(1).strip():
                desc_parts.append(name_match.group(1).strip().replace('&amp;', '&'))
            if memo_match and memo_match.group(1) and memo_match.group(1).strip():
                desc_parts.append(memo_match.group(1).strip().replace('&amp;', '&'))
                
            current['description'] = ' / '.join(desc_parts)

            if current.get('date') and 'amount' in current:
                transactions.append(current)
                
        return transactions

    def guess_csv_headers(self, file_path):
        """
        Reads the first 5 rows of a CSV and guesses the columns.
        """
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                # Sniff for dialect (commas, tabs, etc.)
                dialect = csv.Sniffer().sniff(f.read(1024))
                f.seek(0)
                reader = csv.reader(f, dialect)
                
                headers = next(reader)
                headers_lower = [h.lower().strip() for h in headers]
                
                guesses = {
                    'date': None,
                    'desc': None,
                    'amt': None,
                    'debit': None,
                    'credit': None
                }
                
                for i, h in enumerate(headers_lower):
                    if 'date' in h:
                        guesses['date'] = headers[i]
                    if 'desc' in h or 'narr' in h or 'payee' in h or 'memo' in h or 'particulars' in h:
                        guesses['desc'] = headers[i]
                    if 'amount' in h or 'total' in h:
                        guesses['amt'] = headers[i]
                    if 'debit' in h or 'withdr' in h or 'payment' in h or 'paid out' in h:
                        guesses['debit'] = headers[i]
                    if 'credit' in h or 'deposit' in h or 'paid in' in h:
                        guesses['credit'] = headers[i]
                
                # If we found debit/credit, we probably don't have a single amount column
                if guesses['debit'] and guesses['credit']:
                    guesses['amt'] = None
                
                return headers, guesses
                
        except Exception as e:
            self.log_error(f"Error reading CSV headers: {e}")
            return [], {}

    def parse_csv(self, file_path, mapping):
        """
        Parses a CSV file based on the user's column mapping.
        """
        transactions = []
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                dialect = csv.Sniffer().sniff(f.read(1024))
                f.seek(0)
                reader = csv.DictReader(f, dialect=dialect)
                
                for row in reader:
                    current = {}
                    
                    # Get Date
                    date_str = row.get(mapping['date'])
                    current['date'] = self.parse_date(date_str)
                    
                    # Get Description
                    current['description'] = row.get(mapping['desc'], '')
                    
                    # Get Amount
                    if mapping.get('amt'):
                        # Single-column amount
                        amt_str = row.get(mapping['amt'], '0').replace(',', '').replace('$', '')
                        try:
                            current['amount'] = Decimal(amt_str)
                        except Exception:
                            current['amount'] = Decimal(0)
                    
                    elif mapping.get('debit') and mapping.get('credit'):
                        # Two-column amount (Debit/Credit)
                        debit_str = row.get(mapping['debit'], '0').replace(',', '').replace('$', '')
                        credit_str = row.get(mapping['credit'], '0').replace(',', '').replace('$', '')
                        
                        try:
                            debit = Decimal(debit_str)
                        except Exception:
                            debit = Decimal(0)
                        try:
                            credit = Decimal(credit_str)
                        except Exception:
                            credit = Decimal(0)
                            
                        # Amount is credit (inflow) minus debit (outflow)
                        current['amount'] = credit - debit
                    
                    else:
                        # No amount found
                        current['amount'] = Decimal(0)

                    if current.get('date') and 'amount' in current:
                        transactions.append(current)
                
                return transactions
                
        except Exception as e:
            self.log_error(f"Error parsing CSV: {e}")
            return []

    # --- GUI Top-Level Methods ---
    def load_db(self):
        """Opens file dialog to select DB, creates a backup, and loads accounts."""
        path = filedialog.askopenfilename(
            title="Select Frappe Books Database",
            filetypes=[("Database files", "*.db"), ("All files", "*.*")]
        )
        if not path:
            return
            
        # --- NEW: Create Backup ---
        try:
            original_dir = os.path.dirname(path)
            original_name = os.path.basename(path)
            timestamp = datetime.now().strftime("%Y-%m-%d %H%M")
            backup_name = f"{timestamp} - {original_name}"
            backup_path = os.path.join(original_dir, backup_name)
            
            shutil.copy2(path, backup_path)
            self.log_status(f"Backup created: {backup_name}")
            
        except Exception as e:
            self.log_error(f"Could not create backup: {e}")
            # Ask user if they want to proceed without a backup
            if not messagebox.askyesno("Backup Failed", 
                                       f"Failed to create database backup:\n{e}\n\nDo you want to continue loading the database anyway? (NOT RECOMMENDED)"):
                return # User cancelled
        # --- End Backup ---

        self.db_path.set(path)
        
        try:
            conn = self.connect_db(path)
            if not conn:
                self.log_error("Failed to connect to database.")
                return

            # --- Find and set the correct table names ---
            self.account_table_name = self._find_table_name(conn, ['Account', 'account'])
            self.ledger_table_name = self._find_table_name(conn, ['AccountingLedgerEntry', 'accountingledgerentry'])
            
            if not self.account_table_name:
                self.log_error("Failed to find Account table (tried 'Account', 'account').")
                conn.close()
                return
            if not self.ledger_table_name:
                self.log_error("Failed to find Ledger table (tried 'AccountingLedgerEntry', 'accountingledgerentry').")
                conn.close()
                return
            
            self.log_status(f"Found tables: '{self.account_table_name}' and '{self.ledger_table_name}'")

            # Check/fix schema *before* loading accounts, in case we need to add Suspense
            if not self.check_and_fix_schema(conn):
                 conn.close()
                 return # Error already logged

            self.all_accounts = self.get_accounts(conn)
            conn.close()
            
            if not self.all_accounts:
                self.log_error("No accounts found in database.")
                return

            # Update option menus
            self.bank_menu['menu'].delete(0, 'end')
            self.suspense_menu['menu'].delete(0, 'end')
            
            for acc in self.all_accounts:
                self.bank_menu['menu'].add_command(label=acc, command=tk._setit(self.bank_account, acc))
                self.suspense_menu['menu'].add_command(label=acc, command=tk._setit(self.suspense_account, acc))
                
            self.bank_account.set(self.all_accounts[0]) # Set default
            
            # Set default suspense account
            if "Suspense Clearing" in self.all_accounts:
                self.suspense_account.set("Suspense Clearing")
            elif "Suspense Account" in self.all_accounts:
                self.suspense_account.set("Suspense Account")
            else:
                self.suspense_account.set(self.all_accounts[0])
            
            self.log_status("Database loaded. Ready to load statement.")
            self.check_ready_to_import()
            
        except Exception as e:
            self.log_error(f"Error loading DB: {e}")

    def load_statement(self):
        """Opens file dialog to select statement file."""
        path = filedialog.askopenfilename(
            title="Select Bank Statement File",
            filetypes=[
                ("All statement files", "*.qif *.ofx *.csv"),
                ("QIF files", "*.qif"),
                ("OFX files", "*.ofx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if not path:
            return
            
        self.statement_path.set(path)
        file_ext = os.path.splitext(path)[1].lower()
        
        # Hide CSV frame by default
        self.csv_frame.grid_forget()
        
        if file_ext == '.csv':
            # --- Handle CSV ---
            self.log_status("CSV file detected. Guessing headers...")
            self.csv_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
            
            headers, guesses = self.guess_csv_headers(path)
            self.csv_headers = [""] + headers # Add blank option
            self.csv_guesses = guesses
            
            # Update all CSV option menus
            option_keys = ["Date", "Description", "Amount (Single)", "Debit (Two-Col)", "Credit (Two-Col)"]
            guess_keys = ['date', 'desc', 'amt', 'debit', 'credit']
            
            for i, key in enumerate(option_keys):
                menu = self.csv_option_menus[key]['menu']
                menu.delete(0, 'end')
                
                var = self.csv_vars[i]
                guess_val = self.csv_guesses.get(guess_keys[i])
                
                # Add "N/A" (or blank) as the first option
                menu.add_command(label="", command=tk._setit(var, ""))
                
                for header in headers: # Use original headers, not self.csv_headers
                    menu.add_command(label=header, command=tk._setit(var, header))
                
                if guess_val in headers:
                    var.set(guess_val)
                else:
                    var.set("") # Set to blank

        else:
            self.log_status(f"{file_ext.upper()} file loaded.")
            # --- FIX: Reset CSV vars to prevent "N/A" bleed-through ---
            for var in self.csv_vars:
                var.set("")
            # --- End Fix ---
            
        self.check_ready_to_import()

    def check_ready_to_import(self):
        """Enables import button if all fields are set."""
        if self.db_path.get() and self.statement_path.get() and self.bank_account.get() and self.suspense_account.get():
            self.import_button.config(state='normal')
            self.log_status("Ready to import.")
        else:
            self.import_button.config(state='disabled')

    def run_import(self):
        """Main function to parse the file and import to DB."""

        # --- NEW: Add Disclaimer Popup ---
        disclaimer_text = """** !! LEGAL DISCLAIMER & WARNING !! **

This software is provided "AS IS". Use at your ABSOLUTE OWN RISK.

This tool performs DIRECT MODIFICATION of your accounting database. The developer assumes NO LIABILITY for any data loss, data corruption, incorrect financial entries, or any other damage.

A backup of your database was created when you loaded it.

Do you understand the risks and wish to proceed with the import?"""
        
        if not messagebox.askyesno("!! WARNING & DISCLAIMER !!", disclaimer_text):
            self.log_status("Import cancelled by user.")
            return
        # --- End Disclaimer ---

        if not self.ledger_table_name:
            self.log_error("Cannot import: Ledger table name is not set.")
            return

        self.log_status("Starting import...")
        
        # --- 1. Get all config ---
        db_path = self.db_path.get()
        file_path = self.statement_path.get()
        bank_acc = self.bank_account.get()
        suspense_acc = self.suspense_account.get()
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if not all([db_path, file_path, bank_acc, suspense_acc]):
            self.log_error("Missing required fields.")
            return

        # --- 2. Parse the file ---
        transactions = []
        try:
            if file_ext == '.csv':
                mapping = {
                    'date': self.csv_date_var.get(),
                    'desc': self.csv_desc_var.get(),
                    'amt': self.csv_amt_var.get(),
                    'debit': self.csv_debit_var.get(),
                    'credit': self.csv_credit_var.get(),
                }
                # Validation for CSV mapping
                if not mapping['date'] or not mapping['desc']:
                    self.log_error("CSV must have Date and Description mapped.")
                    return
                if not mapping['amt'] and not (mapping['debit'] and mapping['credit']):
                    self.log_error("CSV must have either Amount (Single) or BOTH Debit/Credit mapped.")
                    return
                
                transactions = self.parse_csv(file_path, mapping)
                
            elif file_ext == '.qif':
                transactions = self.parse_qif(file_path)
            elif file_ext == '.ofx':
                transactions = self.parse_ofx(file_path)
            else:
                self.log_error(f"Unsupported file type: {file_ext}")
                return
                
        except Exception as e:
            self.log_error(f"Failed to parse file: {e}")
            return
            
        if not transactions:
            self.log_error("No valid transactions found in file.")
            return
            
        self.log_status(f"Parsed {len(transactions)} transactions. Importing to database...")
        
        # --- 3. Import to Database ---
        conn = self.connect_db(db_path)
        if not conn:
            self.log_error("Failed to connect to database.")
            return
            
        # Schema should already be fixed, but we re-check table names
        if not self.account_table_name or not self.ledger_table_name:
             self.log_error("Table names lost. Please reload database.")
             conn.close()
             return

        cursor = conn.cursor()
        
        # Get next ID for 'name'
        start_name = 1
        try:
            # Try to get max numeric name
            cursor.execute(f"SELECT MAX(CAST(name AS INTEGER)) FROM {self.ledger_table_name} WHERE name GLOB '[0-9]*'")
            result = cursor.fetchone()
            if result and result[0]:
                start_name = int(result[0]) + 1
        except Exception as e:
            self.log_status(f"Could not find max ID, starting from 1. (Error: {e})")

        # Prepare INSERT statement
        sql = f"""
            INSERT INTO {self.ledger_table_name}
            (name, date, party, account, debit, credit, remark, voucherType, voucherNo, createdBy, modifiedBy, created, modified)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        import_count = 0
        
        try:
            for tx in transactions:
                if not tx.get('date'):
                    self.log_status(f"Skipping transaction, invalid date: {tx.get('description')}")
                    continue

                # Get data
                tx_date = tx['date'].strftime("%Y-%m-%d") # Format as YYYY-MM-DD
                tx_desc = tx.get('description', '')[:280] # Truncate description if too long
                tx_amt = tx['amount'].quantize(Decimal('0.01'))
                
                # --- Double-Entry Logic ---
                # amount > 0 is a Deposit (Inflow) -> Debit Bank, Credit Suspense
                # amount < 0 is a Withdrawal (Outflow) -> Credit Bank, Debit Suspense
                
                if tx_amt > 0:
                    # Deposit
                    debit_acc = bank_acc
                    credit_acc = suspense_acc
                    amt = tx_amt
                elif tx_amt < 0:
                    # Withdrawal
                    debit_acc = suspense_acc
                    credit_acc = bank_acc
                    amt = -tx_amt # Amount is positive
                else:
                    continue # Skip zero-amount transactions

                # Use a common voucher number for both entries
                voucher_no = str(start_name)

                # Entry 1: Debit Side
                name1 = str(start_name)
                data1 = (name1, tx_date, None, debit_acc, str(amt), "0", tx_desc, "Bank Import", voucher_no, "system", "system", now, now)
                cursor.execute(sql, data1)
                
                # Entry 2: Credit Side
                name2 = str(start_name + 1)
                data2 = (name2, tx_date, None, credit_acc, "0", str(amt), tx_desc, "Bank Import", voucher_no, "system", "system", now, now)
                cursor.execute(sql, data2)

                start_name += 2
                import_count += 1
            
            # Commit all transactions at once
            conn.commit()
            conn.close()
            
            self.log_status(f"Successfully imported {import_count} transactions.")
            messagebox.showinfo("Success", f"Successfully imported {import_count} transactions.")

        except Exception as e:
            conn.rollback() # Roll back any changes if an error occurs
            conn.close()
            self.log_error(f"Error during import: {e}")


# --- Main execution ---
if __name__ == "__main__":
    try:
        # Fix blurry fonts on Windows
        if sys.platform == "win32":
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception as e:
                print(f"Could not set DPI awareness: {e}")
            
        root = tk.Tk()
        app = ImporterApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Failed to start application: {e}")
        # Use a simple tk messagebox to show the startup error
        root = tk.Tk()
        root.withdraw() # Hide the main window
        messagebox.showerror("Application Startup Error", str(e))
        try:
            input("Press Enter to exit...") # For console
        except:
            pass # In case console is not available