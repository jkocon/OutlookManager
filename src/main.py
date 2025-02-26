import sys
from export_emails import main_export
from archive_emails import main_archive  # jeśli dodasz archiwizację

try:
    import win32com.client
except ImportError:
    print("Error: Missing module 'pywin32'. Install it using: pip install pywin32")
    sys.exit(1)

def main():
    print("Choose an action:")
    print("1. Export emails to CSV")
    print("2. Move old emails to Online Archive")
    
    try:
        choice = int(input("Select an option (1 or 2): "))
        if choice == 1:
            main_export()
        elif choice == 2:
            main_archive()  # jeśli zaimplementujesz archiwizację
        else:
            raise ValueError("Invalid selection.")
    except ValueError:
        print("Invalid input. Please select 1 or 2.")
        sys.exit(1)

if __name__ == "__main__":
    main()
