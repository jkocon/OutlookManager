import sys
import csv
import os
import time
import win32com.client


def list_outlook_mailboxes():
    """Returns a list of available mailboxes in Outlook."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    mailboxes = [folder.Name for folder in namespace.Folders]
    return mailboxes

def get_main_folder_name(base_folder, full_path):
    """Returns the main folder name from the full path."""
    full_path = full_path.replace(base_folder + "\\", "")
    parts = full_path.split("\\")
    if len(parts) > 0:
        return parts[0]  # Zwraca główny folder bez adresu e-mail
    return full_path[:first_slash_pos] if first_slash_pos > 0 else full_path # type: ignore

def process_folder(folder, writer, base_folder, counter, last_log_time):
    """Recursively scans folders in Outlook and saves email data."""
    for item in folder.Items:
        if hasattr(item, 'Class') and item.Class == 43:  # Checks if the item is a MailItem
            try:
                writer.writerow([
                    item.Subject,
                    item.ReceivedTime.strftime("%Y-%m-%d %H:%M:%S"),
                    item.Size / 1024,  # Size in KB
                    item.Size / (1024 * 1024),  # Size in MB
                    folder.FolderPath.replace(base_folder + "\\", ""),
                    get_main_folder_name(base_folder, folder.FolderPath)
                ])
                counter[0] += 1
                
                current_time = time.time()
                if current_time - last_log_time[0] >= 5:
                    print(f"Processed {counter[0]} emails...")
                    last_log_time[0] = current_time
            except Exception as e:
                print(f"Error processing email: {e}")
    
    for subfolder in folder.Folders:
        if "PersonMetadata" not in subfolder.FolderPath:
            process_folder(subfolder, writer, base_folder, counter, last_log_time)

def export_outlook_emails(output_file, root_folder_name):
    """Exports emails from Outlook to a CSV file."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    try:
        root_folder = namespace.Folders[root_folder_name]
    except Exception as e:
        print(f"Cannot find folder: {root_folder_name}. Check the name in Outlook.")
        return
    
    os.makedirs(os.path.dirname(output_file), exist_ok=True)  # Tworzy katalog jeśli nie istnieje
    counter = [0]
    last_log_time = [time.time()]
    
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Subject", "Date Received", "Size (KB)", "Size (MB)", "Folder", "Main Folder"])
        process_folder(root_folder, writer, root_folder.FolderPath, counter, last_log_time)
    
    print(f"Export completed. Total emails processed: {counter[0]}. File saved as: {output_file}")
    print_exported_folder_stats(output_file)

def print_folder_stats(base_folder):
    """Prints the number of files and total size in main folders."""
    if not os.path.exists(base_folder):
        print("No files found in export folder.")
        return
    
    for folder in os.listdir(base_folder):
        folder_path = os.path.join(base_folder, folder)
        if os.path.isdir(folder_path):
            num_files = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
            total_size = sum(os.path.getsize(os.path.join(folder_path, f)) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f)))
            print(f"Folder: {folder} | Files: {num_files} | Total size: {total_size / (1024 * 1024):.2f} MB")

def print_exported_folder_stats(output_file):
    """Prints the number of emails and total size in exported main folders."""
    folder_sizes = {}
    num_files = 0
    total_size = 0
    
    with open(output_file, newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # Skip header
        for row in reader:
            main_folder = row[5]  # Column for Main Folder
            size_kb = float(row[2])
            folder_sizes[main_folder] = folder_sizes.get(main_folder, 0) + size_kb
            num_files += 1
            total_size += size_kb
    
    print("Exported folder statistics:")
    for folder, size in folder_sizes.items():
        print(f"Main Folder: {folder} | Emails: {num_files} | Total size: {size / 1024:.2f} MB")
    print(f"Total exported emails: {num_files} | Overall size: {total_size / 1024:.2f} MB")

def main_export():
    mailboxes = list_outlook_mailboxes()
    if not mailboxes:
        print("No mailboxes found in Outlook.")
        sys.exit(1)

    for idx, mailbox in enumerate(mailboxes, 1):
        print(f"{idx}. {mailbox}")

    try:
        choice = int(input("Select mailbox number: ")) - 1
        if 0 <= choice < len(mailboxes):
            folder_name = mailboxes[choice]
        else:
            print("Invalid mailbox selection.")
            sys.exit(1)
    except (ValueError, IndexError):
        print("Invalid mailbox selection.")
        sys.exit(1)

    output_folder = os.path.join(os.getcwd(), "export", folder_name)
    os.makedirs(output_folder, exist_ok=True)
    output_csv = os.path.join(output_folder, f"{time.strftime('%Y_%m_%d_%H_%M')}.csv")

    export_outlook_emails(output_csv, folder_name)

if __name__ == "__main__":
    main_export()