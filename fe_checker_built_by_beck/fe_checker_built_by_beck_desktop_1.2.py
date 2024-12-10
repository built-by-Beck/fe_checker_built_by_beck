import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
from datetime import datetime
import openpyxl  # Ensure this is installed for Excel file handling


class FireExtinguisherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Fire Extinguisher Checker")

        self.extinguishers = []
        self.current_section = "All"
        self.filtered_extinguishers = []

        # Title Label
        self.title_label = tk.Label(root, text="Fire Extinguisher Checker\nBuilt by Beck", font=("Helvetica", 16, "bold"), fg="blue")
        self.title_label.pack(pady=10)

        # Buttons for Status Update (Moved to the Top)
        status_button_frame = tk.Frame(root)
        status_button_frame.pack(pady=5)

        tk.Button(status_button_frame, text="Mark Pass", command=lambda: self.update_status("Pass")).grid(row=0, column=0, padx=5)
        tk.Button(status_button_frame, text="Mark Fail", command=lambda: self.update_status("Fail")).grid(row=0, column=1, padx=5)

        # Load, Save, and Reset Buttons
        button_frame = tk.Frame(root)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="Load File", command=self.load_file).grid(row=0, column=0, padx=5)
        tk.Button(button_frame, text="Save Progress", command=self.save_progress).grid(row=0, column=1, padx=5)
        tk.Button(button_frame, text="Monthly Reset", command=self.monthly_reset).grid(row=0, column=2, padx=5)
        tk.Button(button_frame, text="Save Log File", command=self.save_log_file).grid(row=0, column=3, padx=5)

        # Section Filter
        tk.Label(root, text="Select Section:").pack(pady=5)
        self.section_var = tk.StringVar(value="All")
        self.section_dropdown = ttk.Combobox(root, textvariable=self.section_var, state="readonly")
        self.section_dropdown.pack(pady=5)
        self.section_dropdown.bind("<<ComboboxSelected>>", self.filter_by_section)

        # Search Bar
        self.search_var = tk.StringVar()
        tk.Label(root, text="Search by Barcode:").pack(pady=5)
        self.search_entry = tk.Entry(root, textvariable=self.search_var, width=30)
        self.search_entry.pack(pady=5)
        self.search_entry.bind("<KeyRelease>", self.search_extinguishers)

        # Treeview with Scrollbar
        tree_frame = tk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll.pack(side="right", fill="y")

        self.tree = ttk.Treeview(tree_frame, columns=("Section", "Location", "Barcode", "Serial Number", "Status"), show="headings", height=20, yscrollcommand=tree_scroll.set)
        tree_scroll.config(command=self.tree.yview)

        for col, width in zip(
            ("Section", "Location", "Barcode", "Serial Number", "Status"),
            (150, 400, 200, 200, 150)
        ):
            self.tree.heading(col, text=col, anchor="center")
            self.tree.column(col, anchor="center", width=width)
        self.tree.pack(fill="both", expand=True)

        # Style adjustments for better readability
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))
        style.configure("Treeview", font=("Helvetica", 11), rowheight=55)
        self.tree.tag_configure("pass", background="lightgreen")
        self.tree.tag_configure("fail", background="lightcoral")

        # Set focus to search bar on startup
        self.root.bind("<<FocusIn>>", lambda e: self.search_entry.focus_set())
        self.search_entry.focus_set()

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("JSON Files", "*.json")])
        if file_path:
            try:
                if file_path.endswith(".json"):
                    with open(file_path, "r") as file:
                        self.extinguishers = json.load(file)
                elif file_path.endswith(".xlsx"):
                    wb = openpyxl.load_workbook(file_path)
                    sheet = wb.active

                    self.extinguishers = []
                    for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
                        # Safely handle invalid numeric fields
                        barcode = row[2]
                        try:
                            barcode = int(barcode) if barcode else 0
                        except ValueError:
                            barcode = 0  # Default value for invalid barcodes

                        self.extinguishers.append({
                            "Section": row[0] or "Unknown",
                            "Location": row[1] or "Unknown",
                            "Barcode": barcode,
                            "Serial Number": row[3] or "Unknown",
                            "Status": row[4] or "Not Checked"
                        })
                else:
                    messagebox.showerror("Error", "Unsupported file type selected!")
                    return

                self.update_section_dropdown()
                self.filter_by_section()
                messagebox.showinfo("Success", "File loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")

    def save_progress(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if file_path:
            try:
                with open(file_path, "w") as file:
                    json.dump(self.extinguishers, file, indent=4)
                messagebox.showinfo("Success", "Progress saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save progress: {e}")

    def update_section_dropdown(self):
        sections = sorted(set(ext["Section"] for ext in self.extinguishers))
        self.section_dropdown["values"] = ["All"] + sections

    def filter_by_section(self, event=None):
        self.current_section = self.section_var.get()
        if self.current_section == "All":
            self.filtered_extinguishers = self.extinguishers
        else:
            self.filtered_extinguishers = [ext for ext in self.extinguishers if ext["Section"] == self.current_section]
        self.update_tree()

    def search_extinguishers(self, event=None):
        query = self.search_var.get().lower()
        filtered = [
            ext for ext in self.extinguishers if query in str(ext["Barcode"]).lower()
        ]
        self.update_tree(filtered)

    def update_tree(self, data=None):
        self.tree.delete(*self.tree.get_children())
        data = data if data else self.filtered_extinguishers
        for ext in data:
            tag = "pass" if ext["Status"] == "Pass" else "fail" if ext["Status"] == "Fail" else ""
            self.tree.insert("", "end", values=(ext["Section"], ext["Location"], ext.get("Barcode", "Unknown"), ext.get("Serial Number", "Unknown"), ext["Status"]), tags=(tag,))

    def update_status(self, status):
        selected_item = self.tree.focus()
        if selected_item:
            values = self.tree.item(selected_item, "values")
            barcode = int(values[2])
            for extinguisher in self.extinguishers:
                if extinguisher.get("Barcode") == barcode:
                    extinguisher["Status"] = status
                    self.update_tree()
                    break
        self.search_entry.focus_set()

    def save_log_file(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            date_str = datetime.now().strftime("%Y-%m-%d")
            section_name = self.current_section if self.current_section != "All" else "All_Sections"
            file_name = f"Extinguisher_Check_Log_{section_name}_{date_str}.csv"
            file_path = f"{folder_path}/{file_name}"

            try:
                with open(file_path, "w", newline="") as file:
                    writer = csv.writer(file)
                    writer.writerow(["Section", "Location", "Barcode", "Serial Number", "Status"])
                    for extinguisher in self.filtered_extinguishers:
                        writer.writerow([extinguisher.get("Section", "Unknown"), extinguisher.get("Location", "Unknown"), extinguisher.get("Barcode", "Unknown"), extinguisher.get("Serial Number", "Unknown"), extinguisher.get("Status", "Not Checked")])
                messagebox.showinfo("Success", f"Log file saved successfully at: {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save log file: {e}")

    def monthly_reset(self):
        for extinguisher in self.extinguishers:
            extinguisher["Status"] = "Not Checked"
        self.update_tree()
        messagebox.showinfo("Success", "Monthly reset completed!")
        self.search_entry.focus_set()


root = tk.Tk()
app = FireExtinguisherApp(root)
root.mainloop()
