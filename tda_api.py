import pandas as pd
import json
from dotenv import load_dotenv
import schwabdev
import os
from spreadsheet_formatter import *
from datetime import datetime
import re

def delete_token_file():
    """Delete the token.json file if it exists."""
    token_file_path = 'tokens.json'
    try:
        if os.path.exists(token_file_path):
            os.remove(token_file_path)
            print("Deleted old tokens.json file.")
        else:
            print("No tokens.json file found to delete.")
    except Exception as e:
        print(f"Error deleting tokens.json: {str(e)}")

def attempt_authentication(use_existing_tokens=True):
    """Helper function to attempt authentication and return client."""
    load_dotenv()
    app_key = os.getenv('app_key')
    app_secret = os.getenv('app_secret')
    callback_url = os.getenv('callback_url')

    if not all([app_key, app_secret, callback_url]):
        raise ValueError("Missing required environment variables. Please check your .env file.")

    try:
        if use_existing_tokens and os.path.exists('tokens.json'):
            print("Attempting to use existing tokens...")
            return schwabdev.Client(app_key, app_secret, callback_url)
    except Exception as e:
        print(f"Error using existing tokens: {e}")

    print("\nStarting fresh authentication process...")
    print("Please complete the authentication in your browser when it opens...")

    delete_token_file()  # Clear tokens before starting new authentication
    return schwabdev.Client(app_key, app_secret, callback_url)

def extract_expiration_and_call_price_tda(description):
    """Extract expiration date and call/put price from the description."""
    expiration_date = ""
    call_price = ""
    match = re.search(r'(\d{2}/\d{2}/\d{4})', description)
    if match:
        expiration_date = match.group(1)[:5]  # Remove the year from the expiration date
    match = re.search(r'\$(\d+\.?\d*)', description)
    if match:
        call_price = float(match.group(1))
    return expiration_date, call_price

def fetch_and_format_positions():
    """Fetch and format positions using the TDA API."""
    try:
        client = None
        while client is None:
            try:
                client = attempt_authentication()
                # Test the client by making a sample request
                response = client.account_details(
                    'example', #add your account hashes here
                    fields="positions"
                )
                if not response.ok:
                    raise Exception("Authentication failed.")
                data = response.json()
                if not data:
                    raise Exception("No data received.")
                break  # Successful authentication
            except Exception as auth_error:
                error_str = str(auth_error).lower()
                print(f"\nAuthentication error: {str(auth_error)}")
                if "refresh_token_authentication_error" in error_str or "unsupported_token_type" in error_str:
                    print("Refresh token authentication failed. Deleting old tokens...")
                    delete_token_file()
                    client = attempt_authentication(use_existing_tokens=False)
                else:
                    retry = input("Would you like to try authenticating again? (y/n): ")
                    if retry.lower() != 'y':
                        return pd.DataFrame(), {}
        
        # Process positions after successful authentication
        positions = data.get('securitiesAccount', {}).get('positions', [])
        formatted_data = []
        current_prices = {}
        
        for position in positions:
            try:
                instrument = position.get('instrument', {})
                description = instrument.get('description', '')
                symbol = instrument.get('underlyingSymbol', instrument.get('symbol', ''))

                asset_type = instrument.get('assetType', '')
                put_call = instrument.get('putCall', '')  # Default to an empty string if missing
                expiration_date, call_price = extract_expiration_and_call_price_tda(description)
                strike_price = instrument.get('strikePrice', call_price)
                average_price = position.get('averagePrice', None)  # Handle missing Average Price
                average_long_price = position.get('averageLongPrice', None)  # Handle Average Long Price

                quantity = position.get('longQuantity', 0.0) - position.get('shortQuantity', 0.0)
                market_value = position.get('marketValue', 0.0)

                # Fallback for Average Price if missing
                if average_price is None and quantity > 0:
                    average_price = market_value / quantity

                # Fallback for Call/Put Price
                call_put_price = strike_price if asset_type == 'OPTION' else None

                formatted_data.append({
                    'Symbol': symbol,
                    'Description': description,
                    'Asset Type': asset_type,
                    'Put/Call': put_call,
                    'Quantity': quantity,
                    'Market Value': market_value,
                    'Average Price': average_price,
                    'Average Long Price': average_long_price,
                    'Expiration Date': expiration_date,
                    'Call/Put Price': call_put_price,
                })
            except KeyError as e:
                print(f"KeyError for position: {position}")
                print(f"Missing key: {e}")
            except Exception as e:
                print(f"Unexpected error processing position: {e}")
        
        return pd.DataFrame(formatted_data), current_prices
    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")
        return pd.DataFrame(), {}

def process_tda_spreadsheets():
    """Process and create TDA spreadsheets."""
    try:
        print("Starting TDA spreadsheet update...")
        portfolio_data, current_prices = fetch_and_format_positions()
        
        if portfolio_data.empty:
            print("No data available")
            return False
        
        output_file = "TDA.xlsx"
        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                grouped = portfolio_data.groupby('Symbol')
                sorted_symbols = sorted(grouped.groups.keys(), key=lambda x: (x[0].isdigit(), x))
                
                for symbol in sorted_symbols:
                    sanitized_symbol = sanitize_sheet_name(symbol)
                    symbol_data = grouped.get_group(symbol)
                    
                    avg_price = symbol_data[symbol_data['Asset Type'] == 'EQUITY']['Average Price'].mean()
                    if pd.isna(avg_price):
                        avg_price = symbol_data['Average Price'].mean()

                    avg_long_price = symbol_data[symbol_data['Asset Type'] == 'EQUITY']['Average Long Price'].mean()
                    if pd.isna(avg_long_price):
                        avg_long_price = 0.0
                    
                    writer.book.add_worksheet(sanitized_symbol)
                    format_sheet(writer, sanitized_symbol, avg_price, 'TDA')
                    
                    worksheet = writer.sheets[sanitized_symbol]
                    worksheet.write_number('B9', avg_long_price)
                    
                    current_price = current_prices.get(symbol, 0.0)
                    if current_price:
                        worksheet.write_number('B3', current_price)
                    
                    populate_template_tda(writer, sanitized_symbol, symbol_data)
            
            print(f"Successfully created {output_file}")
            open_file(output_file)
            return True
        
        except PermissionError:
            print(f"Error: The file '{output_file}' is open. Please close it and try again.")
            return False
        
    except Exception as e:
        print(f"An error occurred while processing TDA spreadsheets: {str(e)}")
        return False

if __name__ == "__main__":
    process_tda_spreadsheets()
