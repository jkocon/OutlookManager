# Email Export and Archiving Script Documentation

## Overview
This script provides functionality to:
1. Export emails from Microsoft Outlook to a CSV file.
2. Move old emails (older than 12 months) to an Online Archive within Outlook.

The script is divided into three main components:
- `main.py`: The entry point where the user selects between exporting emails or archiving old emails.
- `export_emails.py`: Handles exporting emails to a CSV file.
- `archive_emails.py`: Manages moving old emails to an Online Archive.

---
## Prerequisites
### Required Python Packages:
- `pywin32` (for interacting with Outlook)

To install the required package, run:
```sh
pip install pywin32
```

### Permissions:
- The script must be run with an Outlook profile configured.
- The user must have access to the mailboxes they want to process.

---
## `main.py`
This script acts as the control panel for selecting different operations:
1. **Export emails to CSV**
2. **Move old emails to Online Archive**

### Usage:
Run the script and select the desired option:
```sh
python main.py
```
Example output:
```
Choose an action:
1. Export emails to CSV
2. Move old emails to Online Archive
Select an option (1 or 2):
```

---
## `export_emails.py`
### Functionality:
- Lists available Outlook mailboxes.
- Exports email details (subject, date received, size, folder) to a CSV file.
- Processes both main folders and subfolders recursively.
- Displays statistics after export.

### Key Functions:
- `list_outlook_mailboxes()`: Lists available Outlook mailboxes.
- `export_outlook_emails(output_file, root_folder_name)`: Exports emails from a selected folder.
- `process_folder(folder, writer, base_folder, counter, last_log_time)`: Recursively processes folders and saves email details.
- `print_exported_folder_stats(output_file)`: Displays statistics about exported emails.

### Example Output:
```
Available mailboxes:
1. user@example.com
2. shared@example.com
Select mailbox number: 1
Export completed. Total emails processed: 1200. File saved as: ./export/user@example.com/2024_06_01_12_30.csv
```

---
## `archive_emails.py`
### Functionality:
- Lists Outlook mailboxes.
- Identifies corresponding Online Archive for selected mailboxes.
- Moves emails older than 12 months from selected folders (including subfolders) to an Online Archive.
- Provides a **dry-run** mode to preview how many emails and total size would be moved.

### Key Functions:
- `find_online_archive(mailboxes, selected_mailbox)`: Finds the corresponding Online Archive for a mailbox.
- `list_main_folders(mailbox_name)`: Lists main folders for a selected mailbox.
- `process_folder(folder, archive_folder, run_type, depth)`: Recursively processes a folder and its subfolders, counting or moving emails.
- `main_archive()`: The main function that orchestrates mailbox selection and processing.

### Example Output:
```
Available mailboxes:
1. user@example.com
2. shared@example.com
Select mailbox number: 1
Main folders:
1. Inbox
2. Sent Items
3. Archive
Select folders by number (comma-separated, e.g., 1,3,6): 2
Do you want to perform a dry-run or actually move emails? (dry-run/move): dry-run
Processing folder: Sent Items
Folder 'Sent Items': 520 emails (105.4 MB) would be moved.
Dry-run complete. No emails were moved.
```

---
## Error Handling
### Common Issues:
- **Missing Online Archive:** If the script cannot locate the Online Archive, it exits with an error.
- **Folder Access Issues:** If a folder cannot be accessed or created, the script logs an error and continues.
- **Invalid Selections:** If an invalid mailbox or folder is selected, the script prompts the user and exits.

### Debugging:
- Enable logging by adding `print()` statements in the script.
- Run the script in **dry-run** mode before actual execution.

---
## Future Improvements
- **Logging System**: Implement a structured logging system instead of `print()`.
- **GUI Version**: Create a graphical interface for easier navigation.
- **Selective Email Archiving**: Allow users to define custom retention periods.

---
## Conclusion
This script automates the process of exporting and archiving emails in Outlook. It provides a simple way to manage large volumes of emails while maintaining efficiency.

For any issues or improvements, feel free to contribute or modify the script!

