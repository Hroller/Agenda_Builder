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


