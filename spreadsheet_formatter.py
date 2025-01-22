import pandas as pd
import xlsxwriter
import re
import platform
import subprocess
import os
from datetime import datetime

def clear_screen():
    """Clear the console screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def open_file(filepath):
    """Open the file using the default program."""
    if platform.system() == "Windows":
        os.startfile(filepath)
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", filepath])
    else:  # Linux and other systems
        subprocess.run(["xdg-open", filepath])

def sanitize_sheet_name(name):
    """Replace invalid Excel characters with '_'."""
    return re.sub(r'[\\/*?:\[\]]', '_', name)

def extract_expiration_and_call_price(description):
    """Extract expiration date and call/put price from TDA option description."""
    expiration_date = ""
    strike_price = ""
    
    # Example TDA description: "TSLA Jan 19 '24 $240 Call"
    # or "TSLA 240 Call 01/19/24 Equity"
    try:
        # First try the MM/DD/YY format
        date_match = re.search(r'(\d{2}/\d{2}/\d{2})', description)
        if date_match:
            date_str = date_match.group(1)
            date_obj = datetime.strptime(date_str, "%m/%d/%y")
            expiration_date = date_obj.strftime("%m/%d")
        else:
            # Try the "MMM DD 'YY" format
            date_match = re.search(r'([A-Za-z]{3}\s+\d{1,2}\s+\'\d{2})', description)
            if date_match:
                date_str = date_match.group(1)
                date_obj = datetime.strptime(date_str, "%b %d '%y")
                expiration_date = date_obj.strftime("%m/%d")
        
        # Extract strike price
        price_match = re.search(r'\$?(\d+(?:\.\d+)?)', description)
        if price_match:
            strike_price = float(price_match.group(1))
            
    except (ValueError, AttributeError) as e:
        print(f"Error parsing date from description: {description}")
        print(f"Error details: {str(e)}")
        
    return expiration_date, strike_price

def adjust_empty_columns_width(worksheet, used_columns, start_col, end_col):
    """Adjust the width of empty columns to 1/3 of their current width."""
    for col in range(start_col, end_col + 1):
        # Skip columns N (14th column), O (15th column), and Z (26th column)
        if col in {13, 14, 25, 26}:
            continue
        if col not in used_columns:
            current_width = 8 * 0.75  
            worksheet.set_column(col, col, current_width / 3)

def format_sheet(writer, symbol, average_price, account_id):
    """Apply the specified formatting to each sheet."""
    worksheet = writer.sheets[symbol]

    # Add formatting
    workbook = writer.book
    worksheet.set_default_row(15 * 0.75)
    for col_num in range(0, 256):
        worksheet.set_column(col_num, col_num, 8 * 0.75)  

    dark_line_format = workbook.add_format({'bottom': 2, 'border_color': 'black'})
    worksheet.set_row(6, None, dark_line_format)
    worksheet.set_row(8, None, dark_line_format)
    worksheet.set_row(3, None, dark_line_format)
    worksheet.set_row(43, None, dark_line_format)
    worksheet.write('A2', 'Stock')
    worksheet.write('A3', symbol)
    worksheet.write('A4', account_id[-4:])  # Show last 4 digits of account number
    worksheet.write_formula('A10', '=ROUND(B9/2, 0)')
    worksheet.write('A8', '# of options')
    worksheet.write('A47', 'total cost')
    worksheet.write('A48', 'total cost')
    worksheet.write('A49', 'cost per')
    for i in range(11, 45):
        worksheet.write(f'A{i}', f'=A{i-1}+$I$3')
    worksheet.write('B5', 'Net')
    worksheet.write('B6', 'Underlying')
    worksheet.write('B7', 'Position')
    for i in range(10, 45):
        worksheet.write(f'B{i}', f'=A{i}*$B$8')
    worksheet.write('B47', '=SUM(B8*B9)')
    worksheet.write('C4', 'Calls')
    worksheet.write('I2', 'Increment')
    worksheet.write_number('I3', 1)  
    worksheet.write('N5', 'CALLS')
    worksheet.write('N6', 'Total')
    for i in range(10, 45):
        worksheet.write(f'N{i}', f'=SUM(C{i}:M{i})')
    worksheet.write('O4', 'Puts')
    worksheet.write('Z5', 'PUTS')
    worksheet.write('Z6', 'Total')
    for i in range(10, 45):
        worksheet.write(f'Z{i}', f'=SUM(O{i}:Y{i})')
    worksheet.write('AA5', 'Grand')
    worksheet.write('AA6', 'Total')
    for i in range(10, 45):
        worksheet.write(f'AA{i}', f'=SUM(N{i},Z{i},B{i})')
    worksheet.write('AB7', '/')
    worksheet.write('AB8', '# of Options')
    worksheet.write('AB9', 'Strike Price')
    for i in range(10, 45):
        worksheet.write(f'AB{i}', f'=A{i}')
    for i in range(67, 78):  # C through M columns for calls
        for j in range(10, 45):
            worksheet.write(f'{chr(i)}{j}', f'=IF($A{j}<${chr(i)}$9, 0, ${chr(i)}$8*($A{j}-${chr(i)}$9)*100 )')
    for i in range(79, 90):  # O through Y columns for puts
        for j in range(10, 45):
            worksheet.write(f'{chr(i)}{j}', f'=IF( A{j}>${chr(i)}$9, 0,  ${chr(i)}$8*(${chr(i)}$9-A10 )*100)')
    worksheet.write('AA47', '=SUM(B47)')
    worksheet.write('AA48', '=SUM(A48:Z48)')
    # Write formulas for row 48 (includes K, L, M columns now)
    for i in range(67, 78):  # C through M
        worksheet.write(f'{chr(i)}48', f'={chr(i)}49*-{chr(i)}8*100')
    for i in range(79, 90):  # O through Y
        worksheet.write(f'{chr(i)}48', f'={chr(i)}49*-{chr(i)}8*100')

def populate_template_tda(writer, symbol, data):
    """Populate the template for a given symbol with its positions."""
    worksheet = writer.sheets[symbol]
    used_columns = set()

    # Filter EQUITY and COLLECTIVE_INVESTMENT data
    equity_data = data[data['Asset Type'].isin(['EQUITY', 'COLLECTIVE_INVESTMENT', 'EQ'])]
    if not equity_data.empty:
        equity_row = equity_data.iloc[0]
        average_long_price = equity_row.get('Average Long Price', 0)
        average_short_price = equity_row.get('Average Short Price', 0)
        worksheet.write_number('B8', equity_row['Quantity'])
        if equity_row['Quantity'] < 0:
            worksheet.write_number('B9', average_short_price)
        else:
            worksheet.write_number('B9', average_long_price)
        used_columns.add(2)

    # Filter PUT and CALL data
    put_data = data[data['Put/Call'] == 'PUT'].sort_values(by=['Expiration Date', 'Call/Put Price'])
    call_data = data[data['Put/Call'] == 'CALL'].sort_values(by=['Expiration Date', 'Call/Put Price'])

    # Populate PUTs (columns O-Y)
    put_columns = list(range(15, 25))  # Columns O-Y
    for idx, (_, put) in enumerate(put_data.iterrows()):
        if idx >= len(put_columns):
            break
        col_letter = chr(put_columns[idx] + 64)
        worksheet.write(f'{col_letter}7', put['Expiration Date'])  # Write expiration date
        worksheet.write_number(f'{col_letter}8', put['Quantity'])
        try:
            worksheet.write_number(f'{col_letter}9', put['Call/Put Price'])
            worksheet.write_number(f'{col_letter}49', put['Average Price'])
        except (ValueError, TypeError):
            worksheet.write(f'{col_letter}9', 'N/A')
            worksheet.write(f'{col_letter}49', 'N/A')
        used_columns.add(put_columns[idx])

    # Populate CALLs (columns C-M)
    call_columns = list(range(3, 14))  # Columns C-M
    for idx, (_, call) in enumerate(call_data.iterrows()):
        if idx >= len(call_columns):
            break
        col_letter = chr(call_columns[idx] + 64)
        worksheet.write(f'{col_letter}7', call['Expiration Date'])  # Write expiration date
        worksheet.write_number(f'{col_letter}8', call['Quantity'])
        try:
            worksheet.write_number(f'{col_letter}9', call['Call/Put Price'])
            worksheet.write_number(f'{col_letter}49', call['Average Price'])
        except (ValueError, TypeError):
            worksheet.write(f'{col_letter}9', 'N/A')
            worksheet.write(f'{col_letter}49', 'N/A')
        used_columns.add(call_columns[idx])

    # Adjust empty column widths
    adjust_empty_columns_width(worksheet, used_columns, 15, 25)  # Adjust put columns
    adjust_empty_columns_width(worksheet, used_columns, 3, 14)   # Adjust call columns

def populate_template(writer, symbol, data):
    """Populate the template for a given symbol with its positions."""
    worksheet = writer.sheets[symbol]
    used_columns = set()

    # Filter EQUITY and COLLECTIVE_INVESTMENT data
    equity_data = data[data['Asset Type'].isin(['EQUITY', 'COLLECTIVE_INVESTMENT', 'EQ'])]
    if not equity_data.empty:
        equity_row = equity_data.iloc[0]
        average_long_price = equity_row.get('Average Long Price', 0)
        average_short_price = equity_row.get('Average Short Price', 0)
        worksheet.write_number('B8', equity_row['Quantity'])
        if equity_row['Quantity'] < 0:
            worksheet.write_number('B9', average_short_price)
        else:
            worksheet.write_number('B9', average_long_price)
        used_columns.add(2)

    # Filter PUT and CALL data
    put_data = data[data['Put/Call'] == 'PUT'].sort_values(by=['Expiration Date', 'Call/Put Price'])
    call_data = data[data['Put/Call'] == 'CALL'].sort_values(by=['Expiration Date', 'Call/Put Price'])

    # Populate PUTs (columns O-Y)
    put_columns = list(range(15, 25))  # Columns O-Y
    for idx, (_, put) in enumerate(put_data.iterrows()):
        if idx >= len(put_columns):
            break
        col_letter = chr(put_columns[idx] + 64)
        expiration_date, call_price = extract_expiration_and_call_price(put['Description'])
        worksheet.write(f'{col_letter}7', expiration_date)
        worksheet.write_number(f'{col_letter}8', put['Quantity'])
        try:
            worksheet.write_number(f'{col_letter}9', call_price)
            worksheet.write_number(f'{col_letter}49', put['Average Price'])
        except (ValueError, TypeError):
            worksheet.write(f'{col_letter}9', 'N/A')
            worksheet.write(f'{col_letter}49', 'N/A')
        used_columns.add(put_columns[idx])

    # Populate CALLs (columns C-M)
    call_columns = list(range(3, 14))  # Columns C-M
    for idx, (_, call) in enumerate(call_data.iterrows()):
        if idx >= len(call_columns):
            break
        col_letter = chr(call_columns[idx] + 64)
        expiration_date, call_price = extract_expiration_and_call_price(call['Description'])
        worksheet.write(f'{col_letter}7', expiration_date)
        worksheet.write_number(f'{col_letter}8', call['Quantity'])
        try:
            worksheet.write_number(f'{col_letter}9', call_price)
            worksheet.write_number(f'{col_letter}49', call['Average Price'])
        except (ValueError, TypeError):
            worksheet.write(f'{col_letter}9', 'N/A')
            worksheet.write(f'{col_letter}49', 'N/A')
        used_columns.add(call_columns[idx])

    # Adjust empty column widths
    adjust_empty_columns_width(worksheet, used_columns, 15, 25)  # Adjust put columns
    adjust_empty_columns_width(worksheet, used_columns, 3, 14)   # Adjust call columns