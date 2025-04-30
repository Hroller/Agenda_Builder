# Agenda Maker 2 - Tkinter GUI Version

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime
import os
import re

CATEGORIES = [
    "Unapproved minutes from the last meeting",
    "Petitions",
    "Reports from Committees",
    "Financial Report",
    "Sickness & Distress",
    "Communication",
    "Unfinished Business",
    "New Business"
]

def parse_flexible_date(date_str):
    date_str = date_str.strip().lower()
    date_str = re.sub(r'(\d{1,2})(st|nd|rd|th)', r'\1', date_str)
    date_str = date_str.title()
    formats = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%B %d, %Y",
        "%b %d, %Y"
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError("Unrecognized or incomplete date. Please include a year (YYYY or YY).")

def create_text_documents(agenda_items, agenda_filename, minutes_filename, meeting_date, group_name):
    def write_text_file(filename, title, is_minutes):
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(title + "\n\n")
                if is_minutes:
                    f.write("Opening Time: __________________\n")
                    f.write("Number of Members Present: __________    Number of Visitors Present: __________    Total: _____\n")
                    f.write("MEC/VEC/ECs Present:\n")
                    f.write("    ____________________________________________\n" * 3)
                else:
                    f.write("Greet all MEC/VEC/ECs\n")
                    f.write("    ____________________________________________\n" * 3)

                for index, category in enumerate(CATEGORIES, 1):
                    f.write(f"{index}. {category}\n")
                    for item in agenda_items.get(category, []):
                        f.write(f"    - {item}\n")

                if is_minutes:
                    f.write("Closing Time: __________________\n")
        except Exception as e:
            messagebox.showerror("Error", f"Error writing file: {e}")

    create_title = f"{group_name} Agenda for {meeting_date}"
    minutes_title = f"Meeting minutes for {group_name} for {meeting_date}"

    write_text_file(agenda_filename, create_title, is_minutes=False)
    write_text_file(minutes_filename, minutes_title, is_minutes=True)

    messagebox.showinfo("Success", f"Agenda and minutes created:\n{agenda_filename}\n{minutes_filename}")

# Tkinter GUI
class AgendaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Agenda Maker 2 - GUI Version")

        self.group_var = tk.StringVar()
        self.date_var = tk.StringVar()
        self.save_path = tk.StringVar(value=os.getcwd())
        self.entries = {category: tk.StringVar() for category in CATEGORIES}

        self.build_interface()

    def build_interface(self):
        groups = [
            "Oriental-Rabboni Chapter No. 78 RAM",
            "Jeremiah Council No. 48",
            "Other"
        ]

        ttk.Label(self.root, text="Select Group:").grid(row=0, column=0, sticky="w")
        self.group_combo = ttk.Combobox(self.root, textvariable=self.group_var, values=groups, state="readonly")
        self.group_combo.grid(row=0, column=1, padx=5, pady=5)
        self.group_combo.current(0)

        ttk.Label(self.root, text="Meeting Date:").grid(row=1, column=0, sticky="w")
        self.date_picker = DateEntry(self.root, textvariable=self.date_var, date_pattern='y-mm-dd')
        self.date_picker.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.root, text="Save Folder:").grid(row=2, column=0, sticky="w")
        ttk.Entry(self.root, textvariable=self.save_path, width=40).grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(self.root, text="Browse...", command=self.select_folder).grid(row=2, column=2, padx=5)

        self.fields_frame = ttk.LabelFrame(self.root, text="Agenda Items")
        self.fields_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        for i, category in enumerate(CATEGORIES):
            ttk.Label(self.fields_frame, text=category).grid(row=i, column=0, sticky="w")
            ttk.Entry(self.fields_frame, textvariable=self.entries[category], width=60).grid(row=i, column=1, padx=5, pady=2)
            ttk.Label(self.fields_frame, text="(No commas in dates. Use YYYY-MM-DD or MM/DD/YYYY.)", foreground="gray").grid(row=i, column=2, sticky="w") if 'Unapproved' in category else None

        ttk.Button(self.root, text="Generate Agenda", command=self.generate).grid(row=4, column=0, columnspan=3, pady=10)

    def select_folder(self):
        path = filedialog.askdirectory(initialdir=self.save_path.get(), title="Select Folder")
        if path:
            self.save_path.set(path)

    def generate(self):
        group = self.group_var.get()
        date_input = self.date_var.get().strip()

        if group == "Other":
            group = tk.simpledialog.askstring("Custom Group", "Enter the custom group name:")
            if not group:
                return

        try:
            dt = parse_flexible_date(date_input)
        except ValueError as e:
            messagebox.showerror("Invalid Date", str(e))
            return

        formatted_date = dt.strftime("%B %d, %Y")
        date_filename = dt.strftime("%Y-%m-%d")

        known_prefixes = {
            "Oriental-Rabboni Chapter No. 78 RAM": "OR78",
            "Jeremiah Council No. 48": "J48"
        }
        prefix = known_prefixes.get(group, group.replace(" ", "")[:4].upper())

        agenda_filename = os.path.join(self.save_path.get(), f"{prefix}_agenda_{date_filename}.txt")
        minutes_filename = os.path.join(self.save_path.get(), f"{prefix}_minutes_{date_filename}.txt")

        agenda_items = {
            category: [item.strip() for item in self.entries[category].get().split(',') if item.strip()]
            for category in CATEGORIES
        }

        create_text_documents(agenda_items, agenda_filename, minutes_filename, formatted_date, group)

if __name__ == "__main__":
    root = tk.Tk()
    app = AgendaApp(root)
    root.mainloop()
