from tda_api import process_tda_spreadsheets
from etrade_api import process_etrade_spreadsheets
from dotenv import load_dotenv
from spreadsheet_formatter import clear_screen

def display_menu():
    """Display the main menu."""
    clear_screen()
    print("\nOptions Trading Spreadsheet Updater")
    print("===================================")
    print("1. Update TDA Spreadsheets")
    print("2. Update E*TRADE Spreadsheets")
    print("3. Exit")
    print("===================================")

def display_etrade_submenu(account_keys):
    """Display the E*TRADE account selection menu."""
    clear_screen()
    print("\nE*TRADE Account Selection")
    print("===================================")
    print("0. Update All Accounts")
    for i, account_id in enumerate(account_keys.keys(), 1):
        print(f"{i}. Update Account ending in {account_id[-4:]}")
    print(f"{len(account_keys) + 1}. Back to Main Menu")
    print("===================================")

def main():
    # Load environment variables at startup
    load_dotenv()
    
    while True:
        display_menu()
        choice = input("\nEnter your choice (1-3): ")
        
        if choice == "1":
            clear_screen()
            print("\nUpdating TDA Spreadsheets...")
            if process_tda_spreadsheets():
                print("\nTDA spreadsheet update completed successfully!")
            input("\nPress Enter to continue...")
            
        elif choice == "2":
            clear_screen()
            print("\nInitializing E*TRADE connection...")
            if process_etrade_spreadsheets():
                input("\nPress Enter to continue...")
            
        elif choice == "3":
            print("\nExiting program. Goodbye!")
            break
            
        else:
            print("\nInvalid choice. Please enter 1, 2, or 3.")
            input("\nPress Enter to continue...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nProgram terminated by user.")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {str(e)}")
        input("\nPress Enter to exit...")