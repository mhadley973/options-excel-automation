from requests_oauthlib import OAuth1Session
import xml.etree.ElementTree as ET
import pandas as pd
from dotenv import load_dotenv
import os
import webbrowser
from spreadsheet_formatter import (
    clear_screen, 
    open_file, 
    sanitize_sheet_name,
    format_sheet,
    populate_template
)

# Load environment variables
load_dotenv()
CONSUMER_KEY = os.getenv("CONSUMER_KEY")
CONSUMER_SECRET = os.getenv("CONSUMER_SECRET")
PROD_BASE_URL = os.getenv("PROD_BASE_URL")
REQUEST_TOKEN_URL = f"{PROD_BASE_URL}/oauth/request_token"
AUTHORIZE_URL = "https://us.etrade.com/e/t/etws/authorize"
ACCESS_TOKEN_URL = f"{PROD_BASE_URL}/oauth/access_token"
PORTFOLIO_URL_TEMPLATE = f"{PROD_BASE_URL}/v1/accounts/{{account_key}}/portfolio"
FILTERED_ACCOUNTS = ["example", "example", "example", "example"] #add your account hashes here

def authenticate():
    """Authenticate with E*TRADE API."""
    oauth = OAuth1Session(CONSUMER_KEY, client_secret=CONSUMER_SECRET, callback_uri="oob")
    response = oauth.fetch_request_token(REQUEST_TOKEN_URL)
    resource_owner_key = response.get('oauth_token')
    resource_owner_secret = response.get('oauth_token_secret')

    auth_url = f"{AUTHORIZE_URL}?key={CONSUMER_KEY}&token={resource_owner_key}"
    print(f"Please go to this URL for authorization: {auth_url}")
    webbrowser.open(auth_url)
    verifier = input("Enter the verification code: ")

    oauth = OAuth1Session(
        CONSUMER_KEY,
        client_secret=CONSUMER_SECRET,
        resource_owner_key=resource_owner_key,
        resource_owner_secret=resource_owner_secret,
        verifier=verifier
    )
    oauth_tokens = oauth.fetch_access_token(ACCESS_TOKEN_URL)
    access_token = oauth_tokens.get('oauth_token')
    access_token_secret = oauth_tokens.get('oauth_token_secret')

    return OAuth1Session(
        CONSUMER_KEY,
        client_secret=CONSUMER_SECRET,
        resource_owner_key=access_token,
        resource_owner_secret=access_token_secret
    )

def fetch_accounts(session):
    """Fetch account list from E*TRADE."""
    url = f"{PROD_BASE_URL}/v1/accounts/list"
    response = session.get(url)
    response.raise_for_status()

    try:
        root = ET.fromstring(response.text)
        account_keys = {}
        for account in root.findall(".//Account"):
            account_id = account.find("accountId").text
            account_key = account.find("accountIdKey").text
            if account_id in FILTERED_ACCOUNTS:
                account_keys[account_id] = account_key
        return account_keys
    except ET.ParseError as e:
        print(f"Error parsing XML response: {e}")
        return {}

def fetch_portfolio(session, account_key):
    """Fetch portfolio data from E*TRADE."""
    url = PORTFOLIO_URL_TEMPLATE.format(account_key=account_key)
    response = session.get(url)
    response.raise_for_status()

    try:
        root = ET.fromstring(response.text)
        positions = []
        current_prices = {}  # Dictionary to store current prices
        
        for position in root.findall(".//Position"):
            security_type = position.find(".//securityType")
            
            if security_type is not None:
                symbol = position.find(".//symbol").text
                quantity = float(position.find("quantity").text)
                price_paid = float(position.find("pricePaid").text)
                current_price = float(position.find(".//lastTrade").text)
                description = position.find("symbolDescription").text
                
                # Store current price for equity positions
                if security_type.text in ['EQUITY', 'EQ']:
                    current_prices[symbol] = current_price
                
                position_data = {
                    'Symbol': symbol,
                    'Description': description,
                    'Asset Type': security_type.text,
                    'Quantity': quantity,
                    'Trade Price': current_price,
                    'Average Price': price_paid,
                    'Average Long Price': price_paid if quantity > 0 else 0,
                    'Average Short Price': price_paid if quantity < 0 else 0,
                    'Market Value': current_price * quantity,
                    'Call/Put Price': None  # Initialize with None
                }
                
                if security_type.text == "OPTN":
                    product = position.find("Product")
                    expiry_date = f"{product.find('expiryYear').text}-{product.find('expiryMonth').text}-{product.find('expiryDay').text}"
                    strike_price = float(product.find("strikePrice").text)
                    position_data.update({
                        'Put/Call': product.find("callPut").text,
                        'Strike Price': strike_price,
                        'Expiration Date': expiry_date,
                        'Call/Put Price': strike_price  # Use strike price for Call/Put Price
                    })
                else:
                    # For non-option securities, set these fields to empty or None
                    position_data.update({
                        'Put/Call': '',
                        'Strike Price': None,
                        'Expiration Date': '',
                        'Call/Put Price': None
                    })
                
                positions.append(position_data)
        
        return pd.DataFrame(positions) if positions else pd.DataFrame(), current_prices
    
    except ET.ParseError as e:
        print(f"Error parsing XML response: {e}")
        return pd.DataFrame(), {}

def process_etrade_spreadsheets(selected_account=None):
    """Process and create E*TRADE spreadsheets."""
    try:
        print("Starting E*TRADE spreadsheet update...")
        
        # Initial authentication attempt
        session = None
        while session is None:
            try:
                session = authenticate()
                account_keys = fetch_accounts(session)
                if not account_keys:
                    print("No accounts found")
                    return False
            except Exception as auth_error:
                print(f"\nAuthentication error: {str(auth_error)}")
                retry = input("\nWould you like to try authenticating again? (y/n): ")
                if retry.lower() != 'y':
                    return False
                print("\nRetrying authentication...")
                continue
        
        while True:  # Main account processing loop
            if selected_account is None:
                clear_screen()
                print("\nE*TRADE Account Selection")
                print("===================================")
                print("0. Update All Accounts")
                for i, account_id in enumerate(account_keys.keys(), 1):
                    print(f"{i}. Update Account ending in {account_id[-4:]}")
                print(f"{len(account_keys) + 1}. Back to Main Menu")
                print("===================================")
                
                try:
                    choice = int(input("\nEnter your choice: "))
                    if choice == 0:
                        accounts_to_process = account_keys
                    elif choice == len(account_keys) + 1:
                        return True  # Return to main menu
                    elif 1 <= choice <= len(account_keys):
                        selected_account = list(account_keys.keys())[choice - 1]
                        accounts_to_process = {selected_account: account_keys[selected_account]}
                    else:
                        print("Invalid choice. Please try again.")
                        input("\nPress Enter to continue...")
                        continue
                except ValueError:
                    print("Invalid input. Please enter a number.")
                    input("\nPress Enter to continue...")
                    continue
            else:
                accounts_to_process = {selected_account: account_keys[selected_account]}
            
            # Process the selected account(s)
            for account_id, account_key in accounts_to_process.items():
                print(f"\nProcessing account: {account_id}")
                output_file = f"ETRADE{account_id[-4:]}.xlsx"
                
                try:
                    portfolio_data, current_prices = fetch_portfolio(session, account_key)
                    if portfolio_data.empty:
                        print(f"No data available for account {account_id}")
                        continue
                        
                except Exception as api_error:
                    print(f"\nAPI error occurred: {str(api_error)}")
                    print("Attempting to re-authenticate...")
                    try:
                        session = authenticate()
                        portfolio_data, current_prices = fetch_portfolio(session, account_key)
                        if portfolio_data.empty:
                            print(f"No data available for account {account_id}")
                            continue
                    except Exception as retry_error:
                        print(f"Failed to re-authenticate: {str(retry_error)}")
                        input("\nPress Enter to continue...")
                        continue
                    
                try:
                    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                        grouped = portfolio_data.groupby('Symbol')
                        sorted_symbols = sorted(grouped.groups.keys(), 
                                             key=lambda x: (x[0].isdigit(), x))
                        
                        for symbol in sorted_symbols:
                            sanitized_symbol = sanitize_sheet_name(symbol)
                            symbol_data = grouped.get_group(symbol)
                            
                            avg_price = symbol_data[symbol_data['Asset Type'] == 'EQ']['Trade Price'].mean()
                            if pd.isna(avg_price):
                                avg_price = symbol_data['Trade Price'].mean()
                            
                            writer.book.add_worksheet(sanitized_symbol)
                            format_sheet(writer, sanitized_symbol, avg_price, account_id)
                            
                            # Write current price to B3
                            current_price = current_prices.get(symbol, 0.0)
                            if current_price:
                                worksheet = writer.sheets[sanitized_symbol]
                                worksheet.write_number('B3', current_price)
                            
                            populate_template(writer, sanitized_symbol, symbol_data)
                    
                    print(f"\nSuccessfully created {output_file}")
                    open_file(output_file)
                    
                except PermissionError:
                    print(f"Error: The file '{output_file}' is open. Please close it and try again.")
                    input("\nPress Enter to continue...")
                    continue
            
            print("\nE*TRADE spreadsheet update completed!")
            input("\nPress Enter to return to account selection...")
            selected_account = None  # Reset selected_account to show menu again
            
    except Exception as e:
        print(f"An error occurred while processing E*TRADE spreadsheets: {str(e)}")
        input("\nPress Enter to continue...")
        return False

if __name__ == "__main__":
    process_etrade_spreadsheets()