from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
import os
import shutil
import copy

# Define the agenda categories
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

# Set document margins
def set_margins(section):
    section.left_margin = Inches(0.2)
    section.right_margin = Inches(0.2)

# Ask for date(s) of unapproved minutes
def get_date_input(prompt):
    while True:
        date_input = input(prompt)
        if date_input.strip().lower() in ("0", "o", "none", ""):
            return ["None"]
        dates = [d.strip() for d in date_input.split(',')]
        try:
            formatted = [datetime.strptime(d, '%Y-%m-%d').strftime('%B %d, %Y') for d in dates]
            return formatted
        except ValueError:
            print("Please enter the date(s) in YYYY-MM-DD format, or enter 0 (or O/none) if there are none.")

# Ask for agenda items per category
def collect_agenda_items():
    agenda_items = {}
    for category in CATEGORIES:
        if category == "Unapproved minutes from the last meeting":
            dates = get_date_input("Enter the date(s) of the unapproved minutes (comma-separated, YYYY-MM-DD, or 0/O/none if none): ")
            agenda_items[category] = dates
        else:
            items = input(f"Enter items for '{category}' (comma-separated): ").split(',')
            agenda_items[category] = [item.strip() for item in items if item.strip()]
    return agenda_items

# Create .docx agenda and minutes (combined)
def create_documents(agenda_items, agenda_filename, minutes_filename, meeting_date, group_name):
    def create_doc(title, is_minutes):
        doc = Document()
        set_margins(doc.sections[0])

        heading = doc.add_paragraph(title)
        heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = heading.runs[0]
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

        if is_minutes:
            doc.add_paragraph("Opening Time: __________________").runs[0].font.name = 'Times New Roman'
            doc.add_paragraph("Number of Members Present: __________    Number of Visitors Present: __________    Total: _____").runs[0].font.name = 'Times New Roman'
            doc.add_paragraph("MEC/VEC/ECs Present:").runs[0].font.name = 'Times New Roman'
            for _ in range(3):
                doc.add_paragraph("    ____________________________________________").runs[0].font.name = 'Times New Roman'

        else:
            doc.add_paragraph("Greet all MEC/VEC/ECs").runs[0].font.name = 'Times New Roman'
            for _ in range(3):
                doc.add_paragraph("    ____________________________________________").runs[0].font.name = 'Times New Roman'

        for index, category in enumerate(CATEGORIES, 1):
            p = doc.add_paragraph(f"{index}. {category}")
            p.runs[0].font.name = 'Times New Roman'
            p.runs[0].font.size = Pt(9)
            for item in agenda_items[category]:
                line = doc.add_paragraph(f"    - {item}")
                line.runs[0].font.name = 'Times New Roman'
                line.runs[0].font.size = Pt(9)

        if is_minutes:
            doc.add_paragraph("Closing Time: __________________").runs[0].font.name = 'Times New Roman'

        return doc

    agenda_doc = create_doc(f"{group_name} Agenda for {meeting_date}", False)
    minutes_doc = create_doc(f"Meeting minutes for {group_name} for {meeting_date}", True)

    agenda_doc.save(agenda_filename)
    minutes_doc.save(minutes_filename)

    print(f"Agenda created: {agenda_filename}")
    print(f"Minutes created: {minutes_filename}")

# Load existing .docx agenda
def load_agenda_docx(filename):
    doc = Document(filename)
    agenda_items = {category: [] for category in CATEGORIES}

    current_category = None
    for para in doc.paragraphs:
        text = para.text.strip()
        if text and any(text.startswith(str(i) + ".") for i in range(1, len(CATEGORIES) + 1)):
            for i, category in enumerate(CATEGORIES, 1):
                if text.startswith(f"{i}. {category}"):
                    current_category = category
                    break
        elif text.startswith("-") and current_category:
            item_text = text.lstrip("-").strip()
            agenda_items[current_category].append(item_text)

    return agenda_items

# Edit agenda interactively with undo and preview
def edit_agenda(agenda_items):
    history = []
    while True:
        print("\nEdit Menu:")
        print("1. Add item")
        print("2. Remove item")
        print("3. View current agenda")
        print("4. Reorder items in a category")
        print("5. Undo last change")
        print("6. Save and Exit")
        choice = input("Choose an option: ").strip()

        if choice == "1":
            history.append(copy.deepcopy(agenda_items))
            category = auto_select_category()
            item = input("Enter the item to add: ").strip()
            agenda_items[category].append(item)
            print(f"Added '{item}' to '{category}'.")
        elif choice == "2":
            history.append(copy.deepcopy(agenda_items))
            category = auto_select_category()
            if not agenda_items[category]:
                print("No items to remove.")
                continue
            for idx, item in enumerate(agenda_items[category], 1):
                print(f"{idx}. {item}")
            item_idx = int(input("Enter item number to remove: ")) - 1
            if 0 <= item_idx < len(agenda_items[category]):
                removed = agenda_items[category].pop(item_idx)
                print(f"Removed '{removed}'.")
            else:
                print("Invalid item number.")
        elif choice == "3":
            for cat, items in agenda_items.items():
                print(f"\n{cat}:")
                for item in items:
                    print(f"  - {item}")
        elif choice == "4":
            history.append(copy.deepcopy(agenda_items))
            category = auto_select_category()
            if not agenda_items[category]:
                print("No items to reorder.")
                continue
            print(f"Current order in {category}:")
            for idx, item in enumerate(agenda_items[category], 1):
                print(f"{idx}. {item}")
            new_order = input("Enter new order (comma-separated numbers): ").split(',')
            try:
                new_order = [agenda_items[category][int(i.strip()) - 1] for i in new_order]
                agenda_items[category] = new_order
                print(f"Reordered items in '{category}'.")
            except (ValueError, IndexError):
                print("Invalid reorder input.")
        elif choice == "5":
            if history:
                agenda_items.clear()
                agenda_items.update(history.pop())
                print("Undid last change.")
            else:
                print("No actions to undo.")
        elif choice == "6":
            print("Preview of final agenda:")
            for cat, items in agenda_items.items():
                print(f"\n{cat}:")
                for item in items:
                    print(f"  - {item}")
            confirm = input("Save and exit? (Y/N): ").strip().lower()
            if confirm == 'y':
                break
        else:
            print("Invalid choice.")

def auto_select_category():
    print("Select a category:")
    for idx, cat in enumerate(CATEGORIES, 1):
        print(f"{idx}. {cat}")
    while True:
        try:
            selection = int(input("Enter category number: "))
            if 1 <= selection <= len(CATEGORIES):
                return CATEGORIES[selection - 1]
            else:
                print("Invalid number.")
        except ValueError:
            print("Please enter a valid number.")

# --- Main Execution ---

print("Select the group for the meeting:")
print("1. Oriental-Rabboni Chapter No. 78 RAM")
print("2. Jeremiah Council No. 48")
print("3. Webster Groves Lodge No. 84")
print("4. Ivanhoe Commandery No. 8")
print("5. Other (Enter your own group name)")
group_dict = {
    "1": "Oriental-Rabboni Chapter No. 78 RAM",
    "2": "Jeremiah Council No. 48",
    "3": "Webster Groves Lodge No. 84",
    "4": "Ivanhoe Commandery No. 8"
}
group_choice = input("Enter the number of the group: ").strip()
if group_choice == "5":
    group_name = input("Enter your group name: ").strip()
else:
    group_name = group_dict.get(group_choice, "Unknown Group")

meeting_date_input = input("Enter the date of the meeting (YYYY-MM-DD): ")
formatted_date = datetime.strptime(meeting_date_input, '%Y-%m-%d').strftime('%B %d, %Y')

word_choice = input("Use Microsoft Word format? (y/N): ").strip().lower()
file_ext = ".docx" if word_choice == "y" else ".txt"
agenda_filename = f"agenda_{meeting_date_input}{file_ext}"
minutes_filename = f"minutes_{meeting_date_input}{file_ext}"

mode_choice = input("Do you want to (C)reate new or (E)dit existing agenda/minutes? (C/E): ").strip().lower()

if mode_choice == "e" and os.path.exists(agenda_filename):
    shutil.copy(agenda_filename, agenda_filename + ".bak")
    print(f"Backup created: {agenda_filename}.bak")
    agenda_items = load_agenda_docx(agenda_filename)
    edit_agenda(agenda_items)
else:
    agenda_items = collect_agenda_items()

if file_ext == ".docx":
    create_documents(agenda_items, agenda_filename, minutes_filename, formatted_date, group_name)
else:
    print("Text file editing is not yet supported in Edit mode. Proceeding with fresh collection.")
