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
    raise ValueError(f"Unrecognized or incomplete date: '{date_str}'. Please include a year (YYYY or YY).")

def get_date_input(prompt):
    while True:
        date_input = input(prompt)
        if date_input.strip().lower() in ('0', 'o', 'none', ''):
            return ['None']
        dates = [d.strip() for d in date_input.split(',')]
        formatted = []
        for d in dates:
            try:
                dt = parse_flexible_date(d)
                formatted.append(dt.strftime('%B %d, %Y'))
            except ValueError as e:
                print(e)
        if formatted:
            return formatted
        print("No valid dates entered. Try again.")

def collect_agenda_items():
    agenda_items = {}
    for category in CATEGORIES:
        if category == "Unapproved minutes from the last meeting":
            dates = get_date_input("Enter the date(s) of the unapproved minutes (comma-separated, or 0/O/none if none): ")
            agenda_items[category] = dates
        else:
            items = input(f"Enter items for '{category}' (comma-separated): ").split(',')
            agenda_items[category] = [item.strip() for item in items if item.strip()]
    return agenda_items

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
                    for item in agenda_items[category]:
                        f.write(f"    - {item}\n")

                if is_minutes:
                    f.write("Closing Time: __________________\n")
        except Exception as e:
            print(f"Error writing file {filename}: {e}")

    write_text_file(agenda_filename, f"{group_name} Agenda for {meeting_date}", False)
    write_text_file(minutes_filename, f"Meeting minutes for {group_name} for {meeting_date}", True)
    print(f"Agenda created: {agenda_filename}")
    print(f"Minutes created: {minutes_filename}")

# --- Main Execution ---

print("Select the group for the meeting:")
print("1. Oriental-Rabboni Chapter No. 78 RAM")
print("2. Jeremiah Council No. 48")
print("3. Other (Enter your own group name)")

group_dict = {
    "1": "Oriental-Rabboni Chapter No. 78 RAM",
    "2": "Jeremiah Council No. 48"
}
group_choice = input("Enter the number of the group: ").strip()
group_name = group_dict.get(group_choice, input("Enter your group name: ").strip() if group_choice == "3" else "Unknown Group")

while True:
    meeting_date_input = input("Enter the date of the meeting (formats: YYYY-MM-DD, 5/29/24, May 29, 2024): ")
    try:
        dt = parse_flexible_date(meeting_date_input)
        formatted_date = dt.strftime("%B %d, %Y")
        meeting_date_input = dt.strftime("%Y-%m-%d")
        print("That date falls on a", dt.strftime("%A"))
        break
    except ValueError as e:
        print(e)

known_prefixes = {
    "Oriental-Rabboni Chapter No. 78 RAM": "OR78",
    "Jeremiah Council No. 48": "J48"
}
prefix = known_prefixes.get(group_name, group_name.replace(" ", "")[:4].upper())
agenda_filename = f"{prefix}_agenda_{meeting_date_input}.txt"
minutes_filename = f"{prefix}_minutes_{meeting_date_input}.txt"

agenda_items = collect_agenda_items()
create_text_documents(agenda_items, agenda_filename, minutes_filename, formatted_date, group_name)
