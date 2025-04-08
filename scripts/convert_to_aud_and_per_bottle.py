#!/usr/bin/env python3
"""
Google Sheet Price Updater

This script processes a Google Sheet with wine data and updates pricing information:
1. Converts non-AUD prices to AUD in RRP and Discount RRP columns
2. For prices with " / Case", divides by 12 to get per-bottle prices

Features:
- Processes a specific sheet in the Google Sheet (specified by URL or sheet name)
- Handles various currency formats (R, $, €, £)
- Converts case prices to per-bottle prices (dividing by the number of bottles)
- Supports multiple number formats with commas and decimal points

Usage:
    # Using a full Google Sheet URL (preferred):
    python scripts/sheet_price_updater.py --url "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit?gid=SHEET_GID"
    
    # Using just a sheet name (requires SPREADSHEET_ID to be set in the script):
    python scripts/sheet_price_updater.py --sheet "Sheet Name"
    
    # To list available sheets:
    python scripts/sheet_price_updater.py --list [--url "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit"]

Prerequisites:
- The Google Sheet must be shared with the service account email address
  with Editor permissions
- The service account JSON key file must be in the 'keys' directory

Currency Conversion:
The script uses the following conversion rates to AUD:
- South African Rand (R): 0.086 AUD
- US Dollar ($): 1.66 AUD
- Euro (€): 1.81 AUD
- British Pound (£): 2.13 AUD

To update these rates, modify the CURRENCY_RATES dictionary below.
"""

import os
import re
import argparse
import urllib.parse
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Constants
DEFAULT_SPREADSHEET_ID = '1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k'
SERVICE_ACCOUNT_FILE = 'keys/firm-harbor-456101-q0-5deb4bff67ac.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Currency conversion rates to AUD (as of script creation)
# You may want to replace this with an API call to get real-time rates
CURRENCY_RATES = {
    'R': 0.086,  # South African Rand to AUD
    '$': 1.66,    # USD to AUD
    '€': 1.81,   # EUR to AUD
    '£': 2.13,   # GBP to AUD
}

def extract_spreadsheet_info(url):
    """Extract spreadsheet ID and sheet GID from a Google Sheets URL."""
    if not url:
        return None, None
    
    # Extract spreadsheet ID
    spreadsheet_id_match = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', url)
    if not spreadsheet_id_match:
        return None, None
    
    spreadsheet_id = spreadsheet_id_match.group(1)
    
    # Extract sheet GID
    gid_match = re.search(r'[?&#]gid=(\d+)', url)
    sheet_gid = gid_match.group(1) if gid_match else None
    
    return spreadsheet_id, sheet_gid

def authenticate_google_sheets():
    """Authenticate with Google Sheets API using service account."""
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    return service.spreadsheets()

def get_sheet_names_and_gids(sheets, spreadsheet_id):
    """Get all sheet names and their GIDs from the spreadsheet."""
    sheet_metadata = sheets.get(spreadsheetId=spreadsheet_id).execute()
    sheets_data = sheet_metadata.get('sheets', [])
    
    sheets_info = []
    for sheet in sheets_data:
        properties = sheet.get('properties', {})
        sheets_info.append({
            'title': properties.get('title', ''),
            'gid': properties.get('sheetId', '')
        })
    
    return sheets_info

def get_sheet_data(sheets, spreadsheet_id, sheet_name):
    """Retrieve all data from the specified sheet."""
    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:P"  # Adjust range as needed
    ).execute()
    return result.get('values', [])

def convert_currency_to_aud(price_str, wine_name=None):
    """Convert a price string to AUD numeric value."""
    if not price_str or price_str == "N/A" or price_str == "Not available" or price_str == "Not specified":
        return price_str
    
    # Convert to string if necessary
    price_str = str(price_str).strip()
    
    # Check for case pricing and bottles per case
    case_price = False
    bottles_per_case = 12  # Default bottles per case
    
    if " / Case" in price_str or " per case" in price_str:
        case_price = True
        price_str = price_str.replace(" / Case", "").replace(" per case", "")
    
    # Check for formats like "R1,620,00 for 6x 750ml"
    case_match = re.search(r'for\s+(\d+)x', price_str)
    if case_match:
        case_price = True
        bottles_per_case = int(case_match.group(1))
        price_str = re.sub(r'for\s+\d+x\s+\d+ml', '', price_str)
    
    # First, extract any currency symbol or code
    currency_symbol = None
    # Look for common currency symbols at the start
    if price_str.startswith(('R', '$', '€', '£')):
        currency_symbol = price_str[0]
        price_str = price_str[1:]
    # Look for currency code at the end
    elif price_str.endswith(('USD', 'EUR', 'GBP', 'ZAR', 'AUD')):
        currency_code = price_str[-3:]
        currency_mapping = {
            "USD": "$", "EUR": "€", "GBP": "£", "ZAR": "R", "AUD": "$"
        }
        currency_symbol = currency_mapping.get(currency_code, "$")
        price_str = price_str[:-3]
    # Look for currency symbols at the end
    elif price_str.endswith(('R', '$', '€', '£')):
        currency_symbol = price_str[-1]
        price_str = price_str[:-1]
    
    # Skip conversion if no currency symbol was found
    if not currency_symbol:
        return price_str
    
    # Clean the price string - remove any non-numeric characters except decimal point
    # First replace comma with dot if it's in decimal position (e.g., R100,50)
    if ',' in price_str and '.' not in price_str:
        # Check if the comma is in decimal position
        parts = price_str.split(',')
        if len(parts) == 2 and len(parts[1]) <= 2:  # e.g. 100,50
            price_str = price_str.replace(',', '.')
        else:
            # Remove commas used as thousand separators
            price_str = price_str.replace(',', '')
    else:
        # Remove commas used as thousand separators
        price_str = price_str.replace(',', '')
    
    # Remove any remaining non-numeric characters except the decimal point
    price_str = ''.join(c for c in price_str if c.isdigit() or c == '.')
    
    try:
        # Convert to float
        amount = float(price_str)
        
        # Check if this is a promo item (based on wine name) and price is above 90
        if wine_name and 'promo' in wine_name.lower() and amount > 90:
            case_price = True
        
        # Convert to AUD if not already AUD
        if currency_symbol in CURRENCY_RATES and currency_symbol != '$AUD':
            amount = amount * CURRENCY_RATES[currency_symbol]
        
        # For case prices, divide by appropriate factor
        if case_price:
            amount = amount / bottles_per_case
        
        # Format as AUD
        return f"${amount:.2f}"
    except ValueError:
        # If we couldn't parse the price, return the original
        return price_str

def update_sheet_data(sheets, updates, spreadsheet_id, sheet_name):
    """Update the sheet with the provided value updates."""
    if not updates:
        print(f"No updates to make in sheet '{sheet_name}'.")
        return
    
    batch_update_body = {
        'valueInputOption': 'USER_ENTERED',
        'data': []
    }
    
    for row_idx, col_idx, new_value in updates:
        cell_range = f"{sheet_name}!{chr(65 + col_idx)}{row_idx+1}"
        batch_update_body['data'].append({
            'range': cell_range,
            'values': [[new_value]]
        })
    
    result = sheets.values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=batch_update_body
    ).execute()
    
    print(f"Updated {len(updates)} cells in sheet '{sheet_name}'.")
    return result

def process_pricing_data(data):
    """Process the sheet data and return list of updates."""
    updates = []
    
    # Find column indices for RRP and Discount RRP
    try:
        headers = data[0]
        rrp_col = headers.index("RRP")
        discount_rrp_col = headers.index("Discount RRP")
        wine_name_col = headers.index("Wine Name")
    except (ValueError, IndexError):
        print("Could not find required column headers (RRP, Discount RRP, Wine Name)")
        try:
            headers = data[0]
            rrp_col = headers.index("RRP")
            discount_rrp_col = headers.index("Discount RRP")
            wine_name_col = None
            print("Warning: Wine Name column not found. Proceeding without promo detection.")
        except (ValueError, IndexError):
            print("Could not find required column headers (RRP, Discount RRP)")
            return updates
    
    # Process each row
    for row_idx, row in enumerate(data[1:], start=1):
        # Skip rows that don't have enough columns
        if len(row) <= max(rrp_col, discount_rrp_col):
            continue
        
        row_updated = False
        product_name = row[wine_name_col] if wine_name_col is not None and wine_name_col < len(row) else (row[3] if len(row) > 3 else f"Row {row_idx+1}")
        
        # Process RRP column
        if rrp_col < len(row) and row[rrp_col]:
            original_rrp = row[rrp_col]
            new_rrp = convert_currency_to_aud(original_rrp, product_name if wine_name_col is not None else None)
            if new_rrp != original_rrp:
                updates.append((row_idx, rrp_col, new_rrp))
                row_updated = True
                print(f"Row {row_idx+1} - {product_name}")
                print(f"  RRP: {original_rrp} → {new_rrp}")
        
        # Process Discount RRP column
        if discount_rrp_col < len(row) and row[discount_rrp_col]:
            original_discount = row[discount_rrp_col]
            new_discount = convert_currency_to_aud(original_discount, product_name if wine_name_col is not None else None)
            if new_discount != original_discount:
                updates.append((row_idx, discount_rrp_col, new_discount))
                if not row_updated:
                    print(f"Row {row_idx+1} - {product_name}")
                print(f"  Discount RRP: {original_discount} → {new_discount}")
                row_updated = True
        
        if row_updated:
            print("")
    
    return updates

def process_sheet(sheets, spreadsheet_id, sheet_name):
    """Process a single sheet."""
    print(f"\nProcessing sheet '{sheet_name}' in spreadsheet {spreadsheet_id}...")
    data = get_sheet_data(sheets, spreadsheet_id, sheet_name)
    
    if not data:
        print(f"No data found in sheet '{sheet_name}'.")
        return 0
    
    updates = process_pricing_data(data)
    
    if updates:
        update_sheet_data(sheets, updates, spreadsheet_id, sheet_name)
        return len(updates)
    else:
        print(f"No price updates necessary in sheet '{sheet_name}'.")
        return 0

def main():
    """Main function to execute the script."""
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description='Update pricing information in a Google Sheet.')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--sheet', type=str, help='Name of the sheet to process')
    group.add_argument('--url', type=str, help='Full Google Sheet URL (including sheet ID)')
    group.add_argument('--list', action='store_true', help='List all available sheets')
    args = parser.parse_args()
    
    # Extract spreadsheet ID and sheet GID from URL if provided
    spreadsheet_id = DEFAULT_SPREADSHEET_ID
    sheet_gid = None
    
    if args.url:
        extracted_id, extracted_gid = extract_spreadsheet_info(args.url)
        if extracted_id:
            spreadsheet_id = extracted_id
            sheet_gid = extracted_gid
        else:
            print("Error: Could not extract spreadsheet ID from the provided URL.")
            return
    
    print("Authenticating with Google Sheets API...")
    sheets = authenticate_google_sheets()
    
    # Get all sheet names and GIDs
    print(f"Getting spreadsheet information for ID: {spreadsheet_id}...")
    sheets_info = get_sheet_names_and_gids(sheets, spreadsheet_id)
    sheet_names = [sheet['title'] for sheet in sheets_info]
    
    # Create a mapping from GID to sheet name
    gid_to_name = {str(sheet['gid']): sheet['title'] for sheet in sheets_info}
    
    # If --list flag is used, just list the available sheets and exit
    if args.list:
        print("\nAvailable sheets:")
        for i, sheet in enumerate(sheets_info, 1):
            print(f"{i}. {sheet['title']} (GID: {sheet['gid']})")
        return
    
    # Determine which sheet to process
    sheet_to_process = None
    
    if args.url and sheet_gid:
        # If URL with GID was provided, look up the sheet name
        if sheet_gid in gid_to_name:
            sheet_to_process = gid_to_name[sheet_gid]
        else:
            print(f"Error: Sheet with GID {sheet_gid} not found in the spreadsheet.")
            print("Available sheets:")
            for i, sheet in enumerate(sheets_info, 1):
                print(f"{i}. {sheet['title']} (GID: {sheet['gid']})")
            return
    elif args.sheet:
        # If sheet name was provided directly
        if args.sheet in sheet_names:
            sheet_to_process = args.sheet
        else:
            print(f"Error: Sheet '{args.sheet}' not found in the spreadsheet.")
            print("Available sheets:")
            for i, sheet in enumerate(sheets_info, 1):
                print(f"{i}. {sheet['title']} (GID: {sheet['gid']})")
            return
    
    # Process the sheet
    try:
        updates = process_sheet(sheets, spreadsheet_id, sheet_to_process)
        print(f"\nProcessing complete!")
        print(f"Made {updates} price updates in sheet '{sheet_to_process}'.")
    except Exception as e:
        print(f"Error processing sheet '{sheet_to_process}': {str(e)}")

if __name__ == "__main__":
    main() 