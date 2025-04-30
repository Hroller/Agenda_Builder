import sys
import re
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit,
    QVBoxLayout, QHBoxLayout, QComboBox, QScrollArea, QMessageBox, QDateEdit, QFileDialog, QCalendarWidget
)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QIcon

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
        "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y",
        "%B %d, %Y", "%b %d, %Y"
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    raise ValueError("Unrecognized or incomplete date format.")

class TabFocusTextEdit(QTextEdit):
    def __init__(self):
        super().__init__()
        self.setTabChangesFocus(True)

class AgendaMaker(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Agenda Builder")
        self.setGeometry(200, 200, 600, 700)
        self.setWindowIcon(QIcon("agenda_Icon.ico"))

        self.layout = QVBoxLayout()

        self.group_dropdown = QComboBox()
        self.group_dropdown.addItems([
            "Select group...",
            "Oriental-Rabboni Chapter No. 78 RAM",
            "Jeremiah Council No. 48",
            "Other"
        ])
        self.group_input = QLineEdit()
        self.group_input.setPlaceholderText("Enter custom group name if 'Other'")
        self.layout.addWidget(QLabel("Select Group:"))
        self.layout.addWidget(self.group_dropdown)
        self.layout.addWidget(self.group_input)

        self.date_picker = QCalendarWidget()
        self.date_picker.setSelectedDate(QDate.currentDate())
        self.layout.addWidget(QLabel("Meeting Date:"))
        self.layout.addWidget(self.date_picker)

        self.import_button = QPushButton("Import 'New Business' from Previous Minutes")
        self.import_button.clicked.connect(self.import_old_new_business)
        self.layout.addWidget(self.import_button)

        self.category_inputs = {}
        scroll = QScrollArea()
        inner_widget = QWidget()
        inner_layout = QVBoxLayout()
        for cat in CATEGORIES:
            label = QLabel(cat)
            text_input = TabFocusTextEdit()
            text_input.setFixedHeight(60)
            inner_layout.addWidget(label)
            inner_layout.addWidget(text_input)
            self.category_inputs[cat] = text_input
        inner_widget.setLayout(inner_layout)
        scroll.setWidget(inner_widget)
        scroll.setWidgetResizable(True)
        self.layout.addWidget(scroll)

        self.submit_button = QPushButton("Generate Agenda & Minutes")
        self.submit_button.clicked.connect(self.submit)
        self.layout.addWidget(self.submit_button)

        self.setLayout(self.layout)

    def import_old_new_business(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Previous Minutes File", "", "Text Files (*.txt)")
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    lines = file.readlines()

                new_business_section = []
                in_section = False

                for line in lines:
                    if re.match(r"\d+\.\s*New Business", line):
                        in_section = True
                        continue
                    elif re.match(r"\d+\.\s*", line) and in_section:
                        break
                    elif in_section and line.strip().startswith("-"):
                        new_business_section.append(line.strip()[1:].strip())

                if new_business_section:
                    existing_text = self.category_inputs["Unfinished Business"].toPlainText()
                    if existing_text:
                        existing_text += ", "
                    existing_text += ", ".join(new_business_section)
                    self.category_inputs["Unfinished Business"].setPlainText(existing_text)
                    QMessageBox.information(self, "Import Successful", "Imported 'New Business' items into 'Unfinished Business'.")
                else:
                    QMessageBox.warning(self, "No Data Found", "Could not find a 'New Business' section in the file.")

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to read file: {e}")

    def submit(self):
        group_sel = self.group_dropdown.currentText()
        group_name = self.group_input.text() if group_sel == "Other" else group_sel
        if not group_name or group_sel == "Select group...":
            QMessageBox.warning(self, "Input Error", "Please select or enter a group name.")
            return

        qdate = self.date_picker.selectedDate()
        dt = datetime(qdate.year(), qdate.month(), qdate.day())
        formatted_date = dt.strftime("%B %d, %Y")
        short_date = dt.strftime("%Y-%m-%d")

        prefix_map = {
            "Oriental-Rabboni Chapter No. 78 RAM": "OR78",
            "Jeremiah Council No. 48": "J48"
        }
        prefix = prefix_map.get(group_name, group_name.replace(" ", "")[:4].upper())
        agenda_filename = f"{prefix}_agenda_{short_date}.txt"
        minutes_filename = f"{prefix}_minutes_{short_date}.txt"

        agenda_items = {}
        for cat, widget in self.category_inputs.items():
            text = widget.toPlainText().strip()
            items = [i.strip() for i in text.split(',') if i.strip()]
            if cat == "Unapproved minutes from the last meeting" and not items:
                items = ["None"]
            agenda_items[cat] = items

        self.write_file(agenda_filename, f"{group_name} Agenda for {formatted_date}", agenda_items, is_minutes=False)
        self.write_file(minutes_filename, f"Meeting minutes for {group_name} for {formatted_date}", agenda_items, is_minutes=True)

        QMessageBox.information(self, "Success", f"Files saved:\n{agenda_filename}\n{minutes_filename}")

    def write_file(self, filename, title, agenda_items, is_minutes=False):
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

            for i, cat in enumerate(CATEGORIES, 1):
                f.write(f"{i}. {cat}\n")
                for item in agenda_items[cat]:
                    f.write(f"    - {item}\n")

            if is_minutes:
                f.write("Closing Time: __________________\n")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AgendaMaker()
    window.show()
    sys.exit(app.exec_())
