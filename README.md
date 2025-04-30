# Agenda Builder

**Agenda_Builder.exe** is a standalone Windows application built with Python and PyQt5 that helps Masonic secretaries, scribes, and recorders create formatted meeting agendas and minutes quickly and consistently.

---

## 🚀 Features

- Point-and-click calendar to select meeting date
- Pre-filled agenda categories based on Masonic meeting standards
- Custom group support and file naming based on body name and date
- Imports previous “New Business” items into “Unfinished Business”
- Auto-generates plain-text `.txt` files for:
  - Agenda
  - Minutes (with placeholders)

---

## 📦 Installation

1. Download the latest version of **Agenda_Builder.exe** from the [Releases](https://github.com/Hroller/Agenda_Builder/releases) tab.
2. Place it in a working folder (no install required).
3. Double-click to launch. That’s it!

> ⚠️ Requires Windows. Python is **not** needed on the target machine.

---

## 🖥️ Usage

1. Launch `Agenda_Builder.exe`.
2. Select your Masonic body (or enter a custom name).
3. Pick the meeting date from the calendar.
4. Enter agenda items in each category.
5. (Optional) Import “New Business” from a previous minutes file.
6. Click **Generate Agenda & Minutes**.
7. Files will be saved in the same folder with filenames like:
   - `OR78_agenda_2024-05-29.txt`
   - `OR78_minutes_2024-05-29.txt`

---

## 📂 File Naming Conventions

| Body                                 | Prefix |
|--------------------------------------|--------|
| Oriental-Rabboni Chapter No. 78 RAM | OR78   |
| Jeremiah Council No. 48             | J48    |
| Custom names                        | First 4 capital letters (no spaces) |

---

## 🛠️ Development

This project is written in Python 3.11+ and PyQt5. To build the `.exe`:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=agenda_Icon.ico agenda_gui_qt.py
```

---

## 📄 License

GPL v3 License. See `LICENSE` file for details.

---

## 🤝 Contributing

Pull requests are welcome! If you have suggestions or feature requests, feel free to open an issue.













Previous and TXT ONLY versions below.

# Agenda_Builder2

This script generates Masonic meeting agendas and minutes in plain `.txt` format, customized by group and date.

---

## ✅ Features
- Supports three group types:
  - Oriental-Rabboni Chapter No. 78 RAM
  - Jeremiah Council No. 48
  - Custom-named group
- Requires a valid date (must include a year, in 2- or 4-digit format)
- Automatically displays the day of the week for the meeting date
- Outputs two files:
  - `<prefix>_agenda_<date>.txt`
  - `<prefix>_minutes_<date>.txt`

---

## 📦 Dependencies

This script only uses Python's built-in libraries:
- `datetime`
- `os`
- `re`

No external packages need to be installed.

---

## ▶️ How to Run

### Prerequisites
- Python 3.6 or newer installed on your system

### Steps
1. Open a terminal (Command Prompt or PowerShell on Windows)
2. Navigate to the directory where the script is saved:
   ```
   cd path\to\your\script
   ```
3. Run the script:
   ```
   python Agenda_Builder2.py
   ```

4. Follow the on-screen prompts:
   - Select a group (1, 2, or 3)
   - Enter the meeting date in a valid format (`YYYY-MM-DD`, `5/29/24`, or `May 29, 2024`)
   - Enter agenda items for each category, seperate multiple agenda items in the same catagory with commas (,)

---

## 📁 Output

Files are created in the same directory as the script:
- Example:
  - `OR78_agenda_2024-06-20.txt`
  - `OR78_minutes_2024-06-20.txt`

---

## 🔒 Notes

- Dates without a year are rejected to avoid ambiguity
- You may edit the generated `.txt` files after creation if needed

---

Created by Mike Martin, customized with flexibility and Masonic structure in mind.



BELOW IS DEPRECATED: It still works, but I've made the new script (Agenda_Builder2.0) MUCH simpler and easier to use:
I wanted a tool to build an agenda for Masonic meetings and also create a Meeting Minutes outline.

You do need to install the python-docx library to use this, pip install python-docx

How It Works:

  Step-by-Step Instructions:

1. Launch the Script

Run: python Agenda_Maker2.py

2. Select the Group

Choose from:

1: Oriental-Rabboni Chapter No. 78 RAM

2: Jeremiah Council No. 48

3: Webster Groves Lodge No. 84

4: Ivanhoe Commandery No. 8

5: Other (enter custom name)

3. Enter the Meeting Date

Format: YYYY-MM-DD

Example: 2025-05-05

4. Choose File Format

Type y for Word (.docx)

Type n for Text (.txt)

5. Choose Mode

(C)reate New Agenda/Minutes

(E)dit Existing Agenda/Minutes

6. (Create Mode)

Enter unapproved minutes dates or type 0/none.

Enter items for each category when prompted.

7. (Edit Mode)

Script automatically creates timestamped backups (e.g., agenda_2025-05-05_20240501-1530.docx).

Edit Menu Options:

1: Add item

2: Remove item

3: View current agenda

4: Reorder items

5: Undo last change

6: Save and Exit

8. Successful Save

Agenda and Minutes files are saved with the entered date.

You will see: All tasks completed successfully. Goodbye!

Notes:

Backup copies are timestamped to prevent overwriting.

Word documents use Times New Roman font.

Reordering accepts comma-separated numbers.

If existing files are found in Create mode, you must confirm overwrite.

Common Errors and Solutions:

Invalid date format: Must use YYYY-MM-DD.

Typing non-numeric input when selecting category/item number: prompts a retry.

End of Guide


