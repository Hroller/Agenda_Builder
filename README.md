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
