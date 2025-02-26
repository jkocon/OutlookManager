import win32com.client
import sys
import datetime

def list_outlook_mailboxes():
    """Returns a list of available mailboxes in Outlook."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    mailboxes = [folder.Name for folder in namespace.Folders]
    return mailboxes

def find_online_archive(mailboxes, selected_mailbox):
    """Finds the corresponding Online Archive for the selected mailbox."""
    archive_name = f"Online Archive - {selected_mailbox}"
    for mailbox in mailboxes:
        if mailbox == archive_name:
            print(f"Mapped Online Archive: {mailbox}")
            return mailbox
    print("No Online Archive found for the selected mailbox.")
    return None

def list_main_folders(mailbox_name):
    """Lists main folders of the selected mailbox."""
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    
    try:
        mailbox = namespace.Folders[mailbox_name]
    except Exception:
        print(f"Cannot find mailbox: {mailbox_name}. Check the name in Outlook.")
        return []
    
    return [folder.Name for folder in mailbox.Folders]

def main_archive():
    """Main function to select a mailbox and folders for archiving."""
    total_emails_moved = 0
    total_size_moved_kb = 0

    mailboxes = list_outlook_mailboxes()
    if not mailboxes:
        print("No mailboxes found in Outlook.")
        sys.exit(1)

    print("Available mailboxes:")
    for idx, mailbox in enumerate(mailboxes, 1):
        print(f"{idx}. {mailbox}")

    try:
        choice = int(input("Select mailbox number: ")) - 1
        if 0 <= choice < len(mailboxes):
            mailbox_name = mailboxes[choice]
        else:
            raise ValueError
    except ValueError:
        print("Invalid mailbox selection.")
        sys.exit(1)

    archive_mailbox = find_online_archive(mailboxes, mailbox_name)
    if not archive_mailbox:
        print("No corresponding Online Archive found. Exiting.")
        sys.exit(1)

    print(f"Selected mailbox: {mailbox_name}")

    main_folders = list_main_folders(mailbox_name)
    if not main_folders:
        print("No main folders found in the selected mailbox.")
        sys.exit(1)

    print(f"Main folders in '{mailbox_name}':")
    for idx, folder in enumerate(main_folders, 1):
        print(f"{idx}. {folder}")

    selected_folders = input("Select folders by number (comma-separated, e.g., 1,3,6): ")
    try:
        selected_indices = [int(i.strip()) - 1 for i in selected_folders.split(",")]
        selected_folder_names = [main_folders[i] for i in selected_indices if 0 <= i < len(main_folders)]
        print(f"Selected folders: {', '.join(selected_folder_names)}")

        run_type = input("Do you want to perform a dry-run or actually move emails? (dry-run/move): ").strip().lower()
        if run_type not in ['dry-run', 'move']:
            print("Invalid selection. Please restart and choose 'dry-run' or 'move'.")
            sys.exit(1)

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Referencja do wybranego mailboxa i archiwum
        mailbox = namespace.Folders[mailbox_name]
        archive_root = namespace.Folders[archive_mailbox]

        for folder_name in selected_folder_names:
            try:
                print(f"Accessing folder: {folder_name}")
                folder = mailbox.Folders[folder_name]

                # [ZMIANA] - Tworzymy (lub pobieramy) identyczny folder w archiwum
                archive_folder = get_or_create_subfolder(archive_root, folder_name)

                emails_moved, size_moved_kb = process_folder(folder, archive_folder, run_type)
                total_emails_moved += emails_moved
                total_size_moved_kb += size_moved_kb

            except Exception as e:
                print(f"Error accessing folder {folder_name}: {e}")

        print("Dry-run complete. No emails were moved." if run_type == 'dry-run' else "Email move completed.")
        print(f"Total emails processed: {total_emails_moved}")
        print(f"Total size processed: {total_size_moved_kb / 1024:.2f} MB")

    except (ValueError, IndexError):
        print("Invalid folder selection.")
        sys.exit(1)

def get_or_create_subfolder(parent_folder, subfolder_name):
    """
    Sprawdza, czy w folderze 'parent_folder' istnieje podfolder 'subfolder_name'.
    Jeśli nie istnieje — tworzy go. Zwraca referencję do podfolderu.
    """
    try:
        return parent_folder.Folders[subfolder_name]
    except:
        print(f"Archive folder '{subfolder_name}' does not exist in '{parent_folder.Name}'. Attempting to create it...")
        try:
            new_folder = parent_folder.Folders.Add(subfolder_name)
            return new_folder
        except Exception as e:
            print(f"Failed to create subfolder '{subfolder_name}' in '{parent_folder.Name}'. Error: {e}")
            # Jeśli nie możemy utworzyć podfolderu, zwracamy sam parent_folder,
            # ale w praktyce powinniśmy przerywać, bo maile nie trafią tam, gdzie chcemy.
            return parent_folder

def process_folder(folder, archive_folder, run_type, depth=0):
    """
    Processes a folder and its subfolders, moving or counting old emails.
    Zwraca liczbę przeniesionych maili i sumaryczny rozmiar w KB.
    """
    old_email_count = 0
    total_size_kb = 0

    try:
        #print(f"{'  ' * depth}Processing folder: {folder.Name}")
        # Dla debugowania, możemy podejrzeć obie ścieżki:
        #print(f"{'  ' * depth}Source folder path: {folder.FolderPath}")
        #print(f"{'  ' * depth}Archive folder path: {archive_folder.FolderPath}")

        for item in folder.Items:
            # Sprawdzamy np. 12 miesięcy wstecz
            if hasattr(item, 'ReceivedTime') and item.ReceivedTime < (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=365)):
                old_email_count += 1
                total_size_kb += item.Size / 1024
                #print(item.Subject)
                if run_type == 'move':
                    try:
                        item.Move(archive_folder)
                    except Exception as move_err:
                        print(f"Error moving item: {move_err}")
                        print(f"{'  ' * depth}Source folder path: {folder.FolderPath}")
                        print(f"{'  ' * depth}Archive folder path: {archive_folder.FolderPath}")
                        
        if old_email_count > 0:
            status_text = "would be moved" if run_type == "dry-run" else "moved"
            print(f"{'  ' * depth}Folder '{folder.Name}': {old_email_count} emails "
                f"({total_size_kb / 1024:.2f} MB) {status_text}.")

    except Exception as e:
        print(f"Error processing folder {folder.Name}: {e}")

    # [ZMIANA] Dla każdego podfolderu w skrzynce źródłowej odtwarzamy podfolder w archiwum
    for subfolder in folder.Folders:
        sub_archive_folder = get_or_create_subfolder(archive_folder, subfolder.Name)
        sub_emails, sub_size = process_folder(subfolder, sub_archive_folder, run_type, depth + 1)
        old_email_count += sub_emails
        total_size_kb += sub_size

    return old_email_count, total_size_kb

if __name__ == "__main__":
    main_archive()
