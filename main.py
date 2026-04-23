import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import random
from datetime import datetime
import database
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import os
from tkinter import filedialog
import pandas as pd

# --- Database Helper ---
def run_query(query, parameters=()):
    conn = database.get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, parameters)
    conn.commit()
    conn.close()

def get_data(query, parameters=()):
    conn = database.get_db_connection()
    cursor = conn.cursor()
    cursor.execute(query, parameters)
    rows = cursor.fetchall()
    conn.close()
    return rows

# --- Main Application Class ---
class ExamApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("College Exam Hall Allotment System")
        self.geometry("900x650")
        self.configure(bg="#f0f0f0")
        
        # Initialize Database
        database.init_db()
        
        # Container for frames
        self.container = tk.Frame(self)
        self.container.pack(side="top", fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)
        
        self.frames = {}
        
        # Define all pages
        for F in (LoginPage, Dashboard, StaffPage, HallPage, AllotmentPage):
            page_name = F.__name__
            frame = F(parent=self.container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")
            
        self.show_frame("LoginPage")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()
        # Refresh data if the page has a refresh method
        if hasattr(frame, "refresh_data"):
            frame.refresh_data()

# --- Login Page ---
class LoginPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(bg="#e1e1e1")
        
        frame_login = tk.Frame(self, bg="white", padx=20, pady=20, relief=tk.RAISED)
        frame_login.place(relx=0.5, rely=0.5, anchor="center")
        
        tk.Label(frame_login, text="Admin Login", font=("Arial", 18, "bold"), bg="white").pack(pady=10)
        
        tk.Label(frame_login, text="User ID:", bg="white").pack(anchor="w")
        self.entry_user = tk.Entry(frame_login)
        self.entry_user.pack(fill="x", pady=5)
        
        tk.Label(frame_login, text="Password:", bg="white").pack(anchor="w")
        self.entry_pass = tk.Entry(frame_login, show="*")
        self.entry_pass.pack(fill="x", pady=5)
        
        tk.Button(frame_login, text="Login", command=self.check_login, bg="#4CAF50", fg="white").pack(pady=15, fill="x")

    def check_login(self):
        user = self.entry_user.get()
        pwd = self.entry_pass.get()
        
        # Simple check against DB
        rows = get_data("SELECT * FROM admin WHERE user_id=? AND password=?", (user, pwd))
        if rows:
            self.entry_user.delete(0, tk.END)
            self.entry_pass.delete(0, tk.END)
            self.controller.show_frame("Dashboard")
        else:
            messagebox.showerror("Error", "Invalid Credentials")

# --- Dashboard ---
class Dashboard(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        lbl_title = tk.Label(self, text="Dashboard", font=("Arial", 20, "bold"))
        lbl_title.pack(pady=20)
        
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        
        btn_staff = tk.Button(btn_frame, text="1. Staff Registration", width=25, height=2, 
                              command=lambda: controller.show_frame("StaffPage"))
        btn_staff.grid(row=0, column=0, padx=10, pady=10)
        
        btn_halls = tk.Button(btn_frame, text="2. Exam Hall Details", width=25, height=2,
                              command=lambda: controller.show_frame("HallPage"))
        btn_halls.grid(row=0, column=1, padx=10, pady=10)
        
        btn_allot = tk.Button(btn_frame, text="3. Hall Allotment", width=25, height=2,
                              command=lambda: controller.show_frame("AllotmentPage"))
        btn_allot.grid(row=1, column=0, columnspan=2, pady=10)
        
        btn_logout = tk.Button(self, text="Logout", command=lambda: controller.show_frame("LoginPage"), bg="#f44336", fg="white")
        btn_logout.pack(pady=20)

# --- Staff Registration Page ---
class StaffPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        # Header
        header = tk.Frame(self, bg="#ddd")
        header.pack(fill="x")
        tk.Button(header, text="< Back", command=lambda: controller.show_frame("Dashboard")).pack(side="left", padx=5, pady=5)
        tk.Label(header, text="Staff Registration", font=("Arial", 14, "bold"), bg="#ddd").pack(side="left", padx=20)
        
        # Content
        content = tk.Frame(self, padx=10, pady=10)
        content.pack(fill="both", expand=True)
        
        self.lbl_total_staff = tk.Label(content, text="Total Staff: 0", font=("Arial", 12, "bold"), fg="blue")
        self.lbl_total_staff.grid(row=0, column=0, columnspan=2, sticky="e", pady=(0, 5))
        
        # Form
        form_frame = tk.LabelFrame(content, text="Staff Details")
        form_frame.grid(row=1, column=0, sticky="nsew", padx=5)
        
        tk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky="w", pady=2)
        self.ent_name = tk.Entry(form_frame)
        self.ent_name.grid(row=0, column=1, sticky="ew", pady=2)
        
        tk.Label(form_frame, text="Department:").grid(row=1, column=0, sticky="w", pady=2)
        self.ent_dept = tk.Entry(form_frame)
        self.ent_dept.grid(row=1, column=1, sticky="ew", pady=2)
        
        tk.Label(form_frame, text="Designation:").grid(row=2, column=0, sticky="w", pady=2)
        self.cb_desig = ttk.Combobox(form_frame, values=["Assistant Professor", "Associate Professor", "Head of Department"])
        self.cb_desig.grid(row=2, column=1, sticky="ew", pady=2)
        
        tk.Label(form_frame, text="Date of Joining:").grid(row=3, column=0, sticky="w", pady=2)
        self.ent_doj = tk.Entry(form_frame)
        self.ent_doj.grid(row=3, column=1, sticky="ew", pady=2)
        
        # tk.Label(form_frame, text="Staff Level (1-550):").grid(row=4, column=0, sticky="w", pady=2)
        # self.ent_level = tk.Entry(form_frame)
        # self.ent_level.grid(row=4, column=1, sticky="ew", pady=2)
        
        # Buttons
        btn_box = tk.Frame(form_frame)
        btn_box.grid(row=5, column=0, columnspan=2, pady=10)
        
        tk.Button(btn_box, text="Save", command=self.save_staff).pack(side="left", padx=2)
        tk.Button(btn_box, text="Update", command=self.update_staff).pack(side="left", padx=2)
        tk.Button(btn_box, text="Delete", command=self.delete_staff).pack(side="left", padx=2)
        tk.Button(btn_box, text="Clear", command=self.clear_form).pack(side="left", padx=2)
        tk.Button(btn_box, text="Upload Excel", command=self.upload_staff, bg="#FFC107").pack(side="left", padx=2)
        
        # Treeview
        tree_frame = tk.Frame(content)
        tree_frame.grid(row=1, column=1, sticky="nsew", padx=5)
        
        cols = ("ID", "Name", "Dept", "Desig", "DOJ")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

    def refresh_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        rows = get_data("SELECT * FROM staff")
        for row in rows:
            self.tree.insert("", "end", values=(row['id'], row['name'], row['department'], row['designation'], row['joining_date']))
        self.lbl_total_staff.config(text=f"Total Staff: {len(rows)}")

    def clear_form(self):
        self.ent_name.delete(0, tk.END)
        self.ent_dept.delete(0, tk.END)
        self.cb_desig.set('')
        self.ent_doj.delete(0, tk.END)
        self.selected_id = None

    def save_staff(self):
        try:
            run_query("INSERT INTO staff (name, department, designation, joining_date) VALUES (?,?,?,?)",
                      (self.ent_name.get(), self.ent_dept.get(), self.cb_desig.get(), self.ent_doj.get()))
            self.refresh_data()
            self.clear_form()
            messagebox.showinfo("Success", "Staff Saved")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_select(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            vals = item['values']
            self.selected_id = vals[0]
            self.clear_form()
            self.ent_name.insert(0, vals[1])
            self.ent_dept.insert(0, vals[2])
            self.cb_desig.set(vals[3])
            self.ent_doj.insert(0, vals[4])
            # Restore ID since clear removed it
            self.selected_id = vals[0]

    def update_staff(self):
        if not hasattr(self, 'selected_id') or not self.selected_id:
            return
        run_query("UPDATE staff SET name=?, department=?, designation=?, joining_date=? WHERE id=?",
                  (self.ent_name.get(), self.ent_dept.get(), self.cb_desig.get(), self.ent_doj.get(), self.selected_id))
        self.refresh_data()
        self.clear_form()
        messagebox.showinfo("Success", "Staff Updated")

    def delete_staff(self):
        if not hasattr(self, 'selected_id') or not self.selected_id:
            return
        if messagebox.askyesno("Confirm", "Delete this staff?"):
            run_query("DELETE FROM staff WHERE id=?", (self.selected_id,))
            self.refresh_data()
            self.clear_form()

    def upload_staff(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            return
            
        try:
            df = pd.read_excel(file_path)
            
            # Normalize column names: strip whitespace and convert to lower case
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            # Define aliases for each required DB column
            # schema_col: [list of possible excel headers (normalized)]
            column_aliases = {
                "name": ["name", "staff name", "staff_name"],
                "department": ["department", "dept", "depatment"],
                "designation": ["designation", "desig", "role", "position"],
                "joining_date": ["joining date", "date of joining", "doj", "joining_date", "join date", "d.o.j", "d.o.j."]
            }
            
            final_col_map = {} # Maps found excel col -> db col
            missing_db_cols = []
            
            for db_col, aliases in column_aliases.items():
                found = False
                for alias in aliases:
                    if alias in df.columns:
                        final_col_map[alias] = db_col
                        found = True
                        break # Found a match for this db_col
                if not found:
                    missing_db_cols.append(db_col)
            
            if missing_db_cols:
                # Construct helpful error message showing what we looked for
                msg = "Could not find columns for:\n"
                for col in missing_db_cols:
                    msg += f"- {col} (checked: {', '.join(column_aliases[col])})\n"
                messagebox.showerror("Column Mismatch", msg)
                return
                
            # Filter and rename
            # We only keep the columns we found in the map
            df_clean = df[list(final_col_map.keys())].copy()
            df_clean.rename(columns=final_col_map, inplace=True)
            
            # Handle empty values
            df_clean.fillna("", inplace=True)
            
            conn = database.get_db_connection()
            cursor = conn.cursor()
            
            count = 0
            for _, row in df_clean.iterrows():
                cursor.execute("INSERT INTO staff (name, department, designation, joining_date) VALUES (?,?,?,?)",
                              (str(row['name']), str(row['department']), str(row['designation']), str(row['joining_date'])))
                count += 1
                
            conn.commit()
            conn.close()
            
            self.refresh_data()
            messagebox.showinfo("Success", f"Uploaded {count} staff records successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to upload: {str(e)}")

# --- Hall Details Page ---
class HallPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        
        header = tk.Frame(self, bg="#ddd")
        header.pack(fill="x")
        tk.Button(header, text="< Back", command=lambda: controller.show_frame("Dashboard")).pack(side="left", padx=5, pady=5)
        tk.Label(header, text="Exam Hall Details", font=("Arial", 14, "bold"), bg="#ddd").pack(side="left", padx=20)
        
        content = tk.Frame(self, padx=10, pady=10)
        content.pack(fill="both", expand=True)
        
        form_frame = tk.LabelFrame(content, text="Hall Details")
        form_frame.grid(row=0, column=0, sticky="nsew", padx=5)
        
        tk.Label(form_frame, text="Hall Number:").grid(row=0, column=0, sticky="w", pady=2)
        self.ent_hall = tk.Entry(form_frame)
        self.ent_hall.grid(row=0, column=1, sticky="ew", pady=2)
        
        btn_box = tk.Frame(form_frame)
        btn_box.grid(row=2, column=0, columnspan=2, pady=10)
        
        tk.Button(btn_box, text="Save", command=self.save_hall).pack(side="left", padx=2)
        tk.Button(btn_box, text="Update", command=self.update_hall).pack(side="left", padx=2)
        tk.Button(btn_box, text="Delete", command=self.delete_hall).pack(side="left", padx=2)
        tk.Button(btn_box, text="Clear", command=self.clear_form).pack(side="left", padx=2)
        tk.Button(btn_box, text="Upload Excel", command=self.upload_hall, bg="#FFC107").pack(side="left", padx=2)
        
        tree_frame = tk.Frame(content)
        tree_frame.grid(row=0, column=1, sticky="nsew", padx=5)
        
        cols = ("ID", "Hall No")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        
        content.columnconfigure(1, weight=1)
        content.rowconfigure(0, weight=1)

    def refresh_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        rows = get_data("SELECT * FROM halls")
        for row in rows:
            self.tree.insert("", "end", values=(row['id'], row['hall_number']))

    def clear_form(self):
        self.ent_hall.delete(0, tk.END)
        self.selected_id = None

    def save_hall(self):
        try:
            run_query("INSERT INTO halls (hall_number) VALUES (?)",
                      (self.ent_hall.get(),))
            self.refresh_data()
            self.clear_form()
            messagebox.showinfo("Success", "Hall Saved")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_select(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            vals = item['values']
            self.selected_id = vals[0]
            self.clear_form()
            self.ent_hall.insert(0, vals[1])
            self.selected_id = vals[0]

    def update_hall(self):
        if not hasattr(self, 'selected_id') or not self.selected_id:
            return
        run_query("UPDATE halls SET hall_number=? WHERE id=?",
                  (self.ent_hall.get(), self.selected_id))
        self.refresh_data()
        self.clear_form()
        messagebox.showinfo("Success", "Hall Updated")

    def delete_hall(self):
        if not hasattr(self, 'selected_id') or not self.selected_id:
            return
        if messagebox.askyesno("Confirm", "Delete this hall?"):
            run_query("DELETE FROM halls WHERE id=?", (self.selected_id,))
            self.refresh_data()
            self.clear_form()

    def upload_hall(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path:
            return
            
        try:
            df = pd.read_excel(file_path)
            
            # Normalize column names
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            # Aliases
            column_aliases = {
                "hall_number": ["hall number", "hall no", "room no", "hall_no", "room number"]
            }
            
            final_col_map = {}
            missing_db_cols = []
            
            for db_col, aliases in column_aliases.items():
                found = False
                for alias in aliases:
                    if alias in df.columns:
                        final_col_map[alias] = db_col
                        found = True
                        break
                if not found:
                    missing_db_cols.append(db_col)
            
            if missing_db_cols:
                msg = "Could not find columns for:\n"
                for col in missing_db_cols:
                    msg += f"- {col} (checked: {', '.join(column_aliases[col])})\n"
                messagebox.showerror("Column Mismatch", msg)
                return
                
            df_clean = df[list(final_col_map.keys())].copy()
            df_clean.rename(columns=final_col_map, inplace=True)
            
            df_clean.fillna({"hall_number": ""}, inplace=True)
            
            conn = database.get_db_connection()
            cursor = conn.cursor()
            
            count = 0
            for _, row in df_clean.iterrows():
                try:
                    cursor.execute("INSERT INTO halls (hall_number) VALUES (?)",
                                  (str(row['hall_number']),))
                    count += 1
                except sqlite3.IntegrityError:
                    continue
                
            conn.commit()
            conn.close()
            
            self.refresh_data()
            messagebox.showinfo("Success", f"Uploaded {count} halls successfully.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to upload: {str(e)}")

# --- Allotment Page ---
class AllotmentPage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.last_allotment_data = [] # To store data for PDF
        
        # Main layout: Header, Top Controls, Selection Lists, Bottom Controls, Results
        self.grid_rowconfigure(2, weight=1) # Selection area expands
        self.grid_columnconfigure(0, weight=1)

        # 1. Header
        header = tk.Frame(self, bg="#ddd")
        header.pack(fill="x")
        tk.Button(header, text="< Back", command=lambda: controller.show_frame("Dashboard")).pack(side="left", padx=5, pady=5)
        tk.Label(header, text="Hall Allotment", font=("Arial", 14, "bold"), bg="#ddd").pack(side="left", padx=20)
        
        # 2. Controls (Dates, CIA)
        control_frame = tk.Frame(self, pady=10)
        control_frame.pack(fill="x")
        
        tk.Label(control_frame, text="CIA Num:").pack(side="left", padx=5)
        self.ent_cia = tk.Entry(control_frame, width=5)
        self.ent_cia.pack(side="left", padx=5)
        
        tk.Label(control_frame, text="Category of Dates:").pack(side="left", padx=5)
        self.cb_category = ttk.Combobox(control_frame, values=[str(i) for i in range(1, 15)], width=5, state="readonly")
        self.cb_category.pack(side="left", padx=5)
        self.cb_category.bind("<<ComboboxSelected>>", self.on_category_change)
        
        tk.Button(control_frame, text="Calculate Totals", command=self.calculate_totals, bg="#9E9E9E", fg="white").pack(side="left", padx=15)
        self.lbl_totals = tk.Label(control_frame, text="Days: 0 | Total Duties: 0", font=("Arial", 10, "bold"), fg="#D84315")
        self.lbl_totals.pack(side="left", padx=5)
        
        # 2.1 Dynamic Dates Frame
        self.dates_frame = tk.Frame(self, pady=5)
        self.dates_frame.pack(fill="x")
        self.date_entries = []
        
        # 3. Selection Lists Frame (2 Columns now instead of 3)
        sel_frame = tk.Frame(self, padx=10, pady=5)
        sel_frame.pack(fill="both", expand=True)
        sel_frame.columnconfigure(0, weight=1)
        sel_frame.columnconfigure(1, weight=1)
        
        # --- Column 1: Exclude Staff ---
        f2 = tk.LabelFrame(sel_frame, text="1. Exclude Staff")
        f2.grid(row=0, column=0, sticky="nsew", padx=5)
        
        tk.Label(f2, text="Select staff to OMIT", fg="red").pack(pady=2)
        
        self.lb_exclude = tk.Listbox(f2, selectmode=tk.MULTIPLE, exportselection=False)
        self.lb_exclude.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        sb2 = tk.Scrollbar(f2, orient="vertical", command=self.lb_exclude.yview)
        sb2.pack(side="right", fill="y")
        self.lb_exclude.config(yscrollcommand=sb2.set)
        
        # --- Column 2: Single Duty ---
        f3 = tk.LabelFrame(sel_frame, text="2. Single Duty Staff")
        f3.grid(row=0, column=1, sticky="nsew", padx=5)
        
        tk.Label(f3, text="Select staff for 1 SLOT TOTAL", fg="green").pack(pady=2)
        
        self.lb_single = tk.Listbox(f3, selectmode=tk.MULTIPLE, exportselection=False)
        self.lb_single.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        sb3 = tk.Scrollbar(f3, orient="vertical", command=self.lb_single.yview)
        sb3.pack(side="right", fill="y")
        self.lb_single.config(yscrollcommand=sb3.set)
        
        # 4. Action Buttons
        btn_frame = tk.Frame(self, pady=10)
        btn_frame.pack(fill="x")
        
        tk.Button(btn_frame, text="Generate Bulk Allotment", command=self.generate_allotment, bg="#2196F3", fg="white", font=("Arial", 11)).pack(side="left", padx=20)
        tk.Button(btn_frame, text="Export PDF", command=self.export_to_pdf, bg="#FF5722", fg="white", font=("Arial", 11)).pack(side="left", padx=20)
        
        # 5. Results
        self.result_text = tk.Text(self, height=10)
        self.result_text.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Data caches
        self.all_halls = []
        self.all_staff = []

    def refresh_data(self):
        # Clear Lists
        self.lb_exclude.delete(0, tk.END)
        self.lb_single.delete(0, tk.END)
        
        # Load Halls
        self.all_halls = get_data("SELECT * FROM halls")
        
        # Load Staff
        self.all_staff = get_data("SELECT * FROM staff")
        for staff in self.all_staff:
            display = f"{staff['name']} ({staff['department']})"
            self.lb_exclude.insert(tk.END, display)
            self.lb_single.insert(tk.END, display)

    def on_category_change(self, event=None):
        # Clear existing date entries
        for widget in self.dates_frame.winfo_children():
            widget.destroy()
        self.date_entries.clear()
        
        # Dictionary to store selected halls per day. Key: Entry widget reference, Value: List of hall dictionaries
        self.daily_halls = {}

        try:
            num_dates = int(self.cb_category.get())
        except ValueError:
            return

        from datetime import timedelta
        base_date = datetime.now()

        # Maximum dates per row to avoid expanding window too wide
        max_per_row = 3 # Reduced max per row to fit the buttons
        current_row = tk.Frame(self.dates_frame)
        current_row.pack(fill="x", pady=2)

        for i in range(num_dates):
            if i > 0 and i % max_per_row == 0:
                current_row = tk.Frame(self.dates_frame)
                current_row.pack(fill="x", pady=2)

            day_frame = tk.Frame(current_row, borderwidth=1, relief="ridge")
            day_frame.pack(side="left", padx=5, pady=2)

            tk.Label(day_frame, text=f"D{i+1}:").pack(side="left", padx=2)
            ent_date = tk.Entry(day_frame, width=11)
            ent_date.pack(side="left", padx=2)
            # Default dates: today, tomorrow, etc.
            dt = base_date + timedelta(days=i)
            ent_date.insert(0, dt.strftime("%Y-%m-%d"))
            self.date_entries.append(ent_date)
            
            # Default to all halls if no halls are explicitly selected yet to avoid immediate zero errors
            self.daily_halls[ent_date] = list(self.all_halls)
            
            lbl_count = tk.Label(day_frame, text=f"({len(self.all_halls)})", fg="blue")
            lbl_count.pack(side="left")
            
            btn_halls = tk.Button(day_frame, text="Select Halls", 
                                  command=lambda e=ent_date, l=lbl_count: self.open_hall_selector(e, l))
            btn_halls.pack(side="left", padx=2)
            
        self.calculate_totals()
        
    def open_hall_selector(self, date_entry, lbl_count):
        top = tk.Toplevel(self)
        top.title(f"Select Halls for {date_entry.get()}")
        top.geometry("300x450")
        top.transient(self)
        top.grab_set()

        tk.Label(top, text="Select Halls", font=("Arial", 12, "bold")).pack(pady=5)
        
        lbl_stats = tk.Label(top, text=f"Total: {len(self.all_halls)} | Selected: 0 | Remaining: {len(self.all_halls)}", fg="blue")
        lbl_stats.pack(pady=2)
        
        # Use a listbox with multiple selection for the popup
        lb = tk.Listbox(top, selectmode=tk.MULTIPLE, exportselection=False)
        lb.pack(fill="both", expand=True, padx=10, pady=5)
        
        def update_stats(event=None):
            selected = len(lb.curselection())
            remaining = len(self.all_halls) - selected
            lbl_stats.config(text=f"Total: {len(self.all_halls)} | Selected: {selected} | Remaining: {remaining}")
            
        lb.bind('<<ListboxSelect>>', update_stats)
        
        # Pre-populate list and select currently active ones
        current_selected = self.daily_halls.get(date_entry, [])
        current_ids = {h['id'] for h in current_selected}
        
        for idx, hall in enumerate(self.all_halls):
            lb.insert(tk.END, hall['hall_number'])
            if hall['id'] in current_ids:
                lb.selection_set(idx)
        
        # Initial stats update
        update_stats()
                
        def save_selection():
            indices = lb.curselection()
            selected_halls = [self.all_halls[i] for i in indices]
            self.daily_halls[date_entry] = selected_halls
            lbl_count.config(text=f"({len(selected_halls)})")
            self.calculate_totals()
            top.destroy()
            
        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Save", command=save_selection, bg="#4CAF50", fg="white").pack(side="left", padx=5)
        
        # Quick actions
        def select_all():
            lb.selection_set(0, tk.END)
            update_stats()
            
        def clear_all():
            lb.selection_clear(0, tk.END)
            update_stats()
            
        tk.Button(btn_frame, text="All", command=select_all).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Clear", command=clear_all).pack(side="left", padx=5)

    def _get_dates(self):
        if not hasattr(self, 'date_entries') or not self.date_entries:
            return None, 0, []

        dates = []
        display_dates = []
        for ent in self.date_entries:
            date_str = ent.get()
            try:
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                dates.append(date_str)
                display_dates.append(dt.strftime("%d/%m/%Y"))
            except ValueError:
                return None, 0, []

        return dates, len(dates), display_dates

    def calculate_totals(self):
        dates_info = self._get_dates()
        if not dates_info[0]:
            self.lbl_totals.config(text="Invalid Dates", fg="red")
            return None, 0, 0, []
            
        dates, num_days, display_dates = dates_info
        
        # Calculate total duties needed by summing up daily hall selections
        total_duties = 0
        for ent in self.date_entries:
            daily_list = self.daily_halls.get(ent, [])
            total_duties += len(daily_list)
        
        self.lbl_totals.config(text=f"Days: {num_days} | Total Duties Needed: {total_duties}", fg="#D84315")
        return dates, num_days, total_duties, display_dates

    def generate_allotment(self):
        try:
            self.result_text.delete(1.0, tk.END)
            
            calc_result = self.calculate_totals()
            if not calc_result or not calc_result[0]:
                messagebox.showwarning("Warning", "Please enter valid Start and End dates (YYYY-MM-DD)")
                return
                
            dates, num_days, total_duties_needed, display_dates = calc_result
                
            cia_num = self.ent_cia.get()
            if not cia_num:
                messagebox.showwarning("Warning", "Please enter a CIA Number")
                return
            
            # 1. No global hall indices check needed, we check daily halls during the loop.
            # But we should ensure at least SOME halls are selected in total.
            if total_duties_needed == 0:
                messagebox.showwarning("Warning", "No halls selected for any date or no dates generated!")
                return
            
            # 2. Identify Staff constraints
            exclude_indices = self.lb_exclude.curselection()
            single_indices = self.lb_single.curselection()
            
            excluded_ids = {self.all_staff[i]['id'] for i in exclude_indices}
            single_duty_ids = {self.all_staff[i]['id'] for i in single_indices}
            
            master_pool = []
            
            total_staff_capacity = 0
            
            # 3. Build Pool based on Designation rules
            from datetime import datetime
            
            # Map original department order
            dept_order = {}
            for s in self.all_staff:
                dept = str(s['department']).strip()
                if dept and dept not in dept_order:
                    dept_order[dept] = len(dept_order)

            for staff in self.all_staff:
                s_id = staff['id']
                if s_id in excluded_ids:
                    continue
                    
                staff_obj = dict(staff) # Convert Row to dict
                
                desig = staff_obj.get('designation', '').lower()
                dept = str(staff_obj.get('department', '')).strip()
                
                if s_id in single_duty_ids:
                    target_duties = 1
                elif 'head' in desig or 'hod' in desig:
                    target_duties = 2
                else:
                    # Default capacity to 3 for all other staff (Assistant, Associate, Professors, etc.)
                    target_duties = 3
                    
                total_staff_capacity += target_duties
                
                # Parse Date of Joining for Seniority
                doj_str = str(staff_obj.get('joining_date', '')).strip()
                parsed_doj = datetime.max # Default unparsed dates to Ultra-Junior so they don't lose duties due to bad strings
                try:
                    # Try common formats
                    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d"):
                        try:
                            parsed_doj = datetime.strptime(doj_str, fmt)
                            break
                        except ValueError:
                            pass
                except Exception:
                    pass
                master_pool.append({
                    'id': s_id,
                    'staff_name': staff['name'],
                    'department': dept,
                    'designation': desig,
                    'target_duties': target_duties,
                    'assigned_duties': 0,
                    'seniority_date': parsed_doj
                })
            
            if total_staff_capacity < total_duties_needed:
                msg = f"WARNING! Insufficient Staff Capacity.\n\n"
                msg += f"Duties Needed: {total_duties_needed}\n"
                msg += f"Max Duties Staff Can Handle (Based on Designation rules): {total_staff_capacity}\n\n"
                msg += "Allotment will have Unstaffed halls (--- NO STAFF ---).\nDo you want to proceed?"
                if not messagebox.askyesno("Capacity Warning", msg):
                    return
            
            # 4. Pre-Calculate Exact Duties Quota
            # To ensure Seniors get FEWER overall duties when there's a remainder, 
            # we distribute quotas prior to day-by-day assignment using hierarchy tiers.
            for s in master_pool:
                s['exact_quota'] = 0
                
            def get_priority(s):
                if s['exact_quota'] == 0:
                    return 1 # Priority 1: Ensure everyone gets at least 1
                desig_str = s['designation']
                if 'assistant' in desig_str:
                    return 2 # Priority 2: Assistants strictly hit max
                if 'head' in desig_str or 'hod' in desig_str:
                    return 3 # Priority 3: HODs
                if 'associate' in desig_str:
                    return 4 # Priority 4: Associates fill remaining
                return 5 # Priority 5: Default staff
                
            unassigned_duties = total_duties_needed
            while unassigned_duties > 0:
                eligible_for_quota = [s for s in master_pool if s['exact_quota'] < s['target_duties']]
                if not eligible_for_quota:
                    break
                    
                # We use a stable sort sequence:
                # 1. Alphabetical tie-breaker
                eligible_for_quota.sort(key=lambda s: s['staff_name'])
                # 2. Juniors first
                eligible_for_quota.sort(key=lambda s: s['seniority_date'], reverse=True)
                # 3. Fair share counting (lowest exact_quota first)
                eligible_for_quota.sort(key=lambda s: s['exact_quota'])
                # 4. Designation hierarchy (highest priority first)
                eligible_for_quota.sort(key=lambda s: get_priority(s))
                
                best_s = eligible_for_quota[0]
                best_s['exact_quota'] += 1
                unassigned_duties -= 1
                    
            # 5. Assignment Logic - Loop through Days
            raw_assignments = []  # List of tuples (date_idx, date_str, display_date, hall_no, staff_name)
            
            num_assigned = 0
            
            for d_idx, date in enumerate(dates):
                disp_date = display_dates[d_idx]
                
                # Retrieve the actual entry widget to get the daily halls list
                ent_widget = self.date_entries[d_idx]
                daily_selected_halls = self.daily_halls.get(ent_widget, [])
                
                if not daily_selected_halls:
                    continue # No halls for this day, skip
                
                random.shuffle(daily_selected_halls) # Shuffle halls so the first hall isn't always the same
                assigned_staff_ids_today = set() # Prevent duplicate assignments on the SAME day
                
                for hall in daily_selected_halls:
                    staff_assigned_name = "--- NO STAFF ---"
                    
                    # Filter eligible staff: not assigned today, haven't reached their EXACT QUOTA
                    eligible_staff = [s for s in master_pool if s['id'] not in assigned_staff_ids_today and s['assigned_duties'] < s['exact_quota']]
                    
                    if eligible_staff:
                        # Sort by Name (stable tie-breaker)
                        eligible_staff.sort(key=lambda s: s['staff_name'])
                        # "first junior staff a eduthu allot pannika apro senior staff ku"
                        # Sort by Seniority descending (Junior staff are picked first)
                        eligible_staff.sort(key=lambda s: s['seniority_date'], reverse=True)
                        # Finally sort by assigned_duties ascending (Fairness per day)
                        eligible_staff.sort(key=lambda s: s['assigned_duties'])
                        
                        best_candidate = eligible_staff[0] # Pick the first one
                        staff_assigned_name = best_candidate['staff_name']
                        assigned_staff_ids_today.add(best_candidate['id'])
                        
                        # Update assignment count in master pool directly
                        best_candidate['assigned_duties'] += 1
                            
                    raw_assignments.append((d_idx, date, disp_date, hall['hall_number'], staff_assigned_name))
                    if staff_assigned_name != "--- NO STAFF ---":
                        num_assigned += 1
            
            # Group assignments by Staff Name
            # staff_assignments: dict[staff_name] -> list of strings "DD/MM/YYYY - HallNo"
            staff_assignments_dict = {}
            for d_idx, date_str, disp_date, hall_no, staff_name in raw_assignments:
                if staff_name == "--- NO STAFF ---":
                    continue
                if staff_name not in staff_assignments_dict:
                    staff_assignments_dict[staff_name] = []
                
                # Format: 24/02/2026 - 50
                duty_str = f"{disp_date} - {hall_no}"
                staff_assignments_dict[staff_name].append((d_idx, duty_str))
            
            # Output generation
            output = f"GOBI ARTS & SCIENCE COLLEGE\n"
            output += f"STAFF DUTY ALLOTMENT FOR CIA {cia_num}\n"
            output += f"Duration: {display_dates[0]} to {display_dates[-1]} ({num_days} Days)\n"
            output += "-"*80 + "\n"
            output += f"{'Staff Name (Dept)':<40} Duties\n"
            output += "-"*80 + "\n"
            
            # Create a lookup for final sorting
            staff_lookup = { s['staff_name']: s for s in master_pool }
            
            # Group data for sorting
            final_report_data_raw = []
            for staff_name, duties_data in staff_assignments_dict.items():
                duties_data.sort(key=lambda x: x[0])
                combined_duties_string = ", ".join([d[1] for d in duties_data]) + ","
                staff_obj = staff_lookup[staff_name]
                final_report_data_raw.append((staff_obj, combined_duties_string))

            # Sort final output: Department Order -> Seniority (Senior first/Ascending) -> Name
            final_report_data_raw.sort(key=lambda x: (
                dept_order.get(x[0]['department'], 999), 
                x[0]['seniority_date'], 
                x[0]['staff_name']
            ))
            
            final_report_data = [] # List of tuples: (display_name, ["date - hall, ", ...])
            
            for staff_obj, combined_duties_string in final_report_data_raw:
                display_name = f"{staff_obj['staff_name']} ({staff_obj['department']})"
                output += f"{display_name:<40} {combined_duties_string}\n"
                final_report_data.append((display_name, combined_duties_string))
            
            output += "-"*80 + "\n"
            output += f"Total Slots Filled: {num_assigned} / {total_duties_needed}\n"
            
            self.result_text.insert(tk.END, output)
            
            self.last_allotment_data = final_report_data
            self.last_summary = {
                "cia": cia_num,
                "start": display_dates[0],
                "end": display_dates[-1],
                "num_days": num_days,
                "slots_filled": num_assigned,
                "total_duties": total_duties_needed
            }
            self.master_pool_debug = master_pool
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            messagebox.showerror("Allotment Error", f"An error occurred:\n{error_msg}")
            print(error_msg)

    def export_to_pdf(self):
        if not hasattr(self, 'last_allotment_data') or not self.last_allotment_data:
            messagebox.showerror("Error", "Please generate allotment first!")
            return
            
        try:
            filename = f"CIA_{self.last_summary['cia']}_Allotment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            import sys
            import os
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            filepath = os.path.join(base_path, filename)
            
            from reportlab.lib.pagesizes import A4
            from reportlab.lib import colors
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            
            doc = SimpleDocTemplate(filepath, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=14,
                alignment=1, # Center
                spaceAfter=6
            )
            subtitle_style = ParagraphStyle(
                'CustomSubtitle',
                parent=styles['Heading2'],
                fontSize=12,
                alignment=1, # Center
                spaceAfter=12
            )
            
            elements.append(Paragraph("GOBI ARTS & SCIENCE COLLEGE", title_style))
            elements.append(Paragraph(f"STAFF DUTY ALLOTMENT FOR CIA {self.last_summary['cia']}", subtitle_style))
            
            # Table Data Preparation
            table_data = []
            
            header_style = ParagraphStyle(
                'TableHeader',
                parent=styles['Normal'],
                fontName='Helvetica-Bold',
                alignment=1 # Center
            )
            
            table_data.append([
                Paragraph("<b>S.No</b>", header_style),
                Paragraph("<b>Staff Name</b>", header_style),
                Paragraph("<b>Allotted Duties (Date - Hall)</b>", header_style)
            ])
            
            s_no = 1
            for staff_name, duty_string in self.last_allotment_data:
                # Wrap the duty string in a Paragraph so it wraps inside the table cell if needed
                sno_p = Paragraph(str(s_no), styles['Normal'])
                duty_p = Paragraph(duty_string, styles['Normal'])
                name_p = Paragraph(staff_name, styles['Normal'])
                table_data.append([sno_p, name_p, duty_p])
                s_no += 1
                
            if len(table_data) == 1:
                table_data.append(["", "No assignments", ""])
                
            # Create Table
            # A4 width is ~595. Margins are 30+30=60. Available width = 535
            col_widths = [40, 160, 335]
            t = Table(table_data, colWidths=col_widths)
            
            # Table Style
            # Add grid lines for row and column boxes
            t.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), # Gray header
                ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ('ALIGN', (0,1), (0,-1), 'CENTER'), # Center S.No
                ('GRID', (0,0), (-1,-1), 1, colors.black), # solid black borders
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('TOPPADDING', (0,0), (-1,-1), 8),
            ]))
            
            elements.append(t)
            
            elements.append(Spacer(1, 15))
            elements.append(Paragraph(f"<b>Total Duties Filled:</b> {self.last_summary['slots_filled']} / {self.last_summary['total_duties']}", styles['Normal']))
            
            doc.build(elements)
            
            messagebox.showinfo("Success", f"PDF Exported to {filename}")
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            messagebox.showerror("Error", f"Failed to export PDF:\n{error_msg}")
            print(error_msg)
        
        # Optional: Save to DB or CSV could go here

if __name__ == "__main__":
    app = ExamApp()
    app.mainloop()
