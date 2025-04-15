#!/usr/bin/env python3
"""
Move To Queue

This script copies data from a given Google Sheet to the queue sheet.
It maps columns from the source sheet to the queue sheet as follows:

Source sheet format (based on Prowein Argentina sheet):
- Column 7: "Winery"
- Column 8: "Website" 
- Column 11: "Supplier Contact Email"

Queue sheet format:
- Country: Added from the --country parameter
- Winery or Supplier Name: Copied from "Winery" column
- Email: Copied from "Supplier Contact Email" column
- Website: Copied from "Website" column

Usage:
    python scripts/move_to_queue.py --url "https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit?gid=SHEET_GID" --country "COUNTRY_CODE"

Prerequisites:
- The Google Sheet must be shared with the service account email address with Editor permissions
- The service account JSON key file must be in the 'keys' directory
"""

import os
import re
import argparse
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Constants
QUEUE_SPREADSHEET_ID = '1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k'  # Replace with your queue spreadsheet ID
QUEUE_SHEET_NAME = '** Queue to Send Out Emails'  # Replace with your queue sheet name
SERVICE_ACCOUNT_FILE = 'keys/firm-harbor-456101-q0-5deb4bff67ac.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Fixed column indices from the source sheet format
WINERY_COL = 7
WEBSITE_COL = 8
EMAIL_COL = 11

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

def get_sheet_data(sheets, spreadsheet_id, sheet_name):
    """Retrieve all data from the specified sheet."""
    print(f"\nFetching data from sheet: {sheet_name}")
    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:Z"  # Get all columns
    ).execute()
    values = result.get('values', [])
    print(f"Found {len(values)} rows in sheet")
    if values and len(values) > 0:
        print("First row (headers):")
        first_row = values[0] if values else []
        for i, cell in enumerate(first_row):
            print(f"Column {i}: '{cell}'")
    return values

def get_queue_column_indices(headers):
    """Get the indices of required columns in the queue sheet."""
    try:
        country_col = headers.index("Country")
        name_col = headers.index("Winery or Supplier Name")
        email_col = headers.index("Email")
        website_col = headers.index("Website")
        return country_col, name_col, email_col, website_col
    except ValueError as e:
        print(f"Error: Required column not found in queue sheet - {str(e)}")
        print("Available columns in queue sheet:")
        for i, header in enumerate(headers):
            print(f"{i}: '{header}'")
        return None, None, None, None

def get_existing_entries(sheets, spreadsheet_id, sheet_name):
    """Get existing entries from the queue sheet to avoid duplicates."""
    data = get_sheet_data(sheets, spreadsheet_id, sheet_name)
    if not data:
        return set()
    
    # Get column indices
    headers = data[0]
    country_col, name_col, email_col, website_col = get_queue_column_indices(headers)
    if None in (country_col, name_col, email_col, website_col):
        return set()
    
    # Skip header row
    existing_entries = set()
    for row in data[1:]:
        if len(row) > name_col:  # We only need the winery name for comparison
            name = row[name_col].strip() if name_col < len(row) else ""
            
            if name:  # Only add non-empty entries
                existing_entries.add(name.lower())  # Case-insensitive comparison
    
    return existing_entries

def process_source_data(sheets, source_spreadsheet_id, source_sheet_name, queue_spreadsheet_id, queue_sheet_name, country):
    """Process the source data and copy unique entries to the queue sheet."""
    # Get source data
    source_data = get_sheet_data(sheets, source_spreadsheet_id, source_sheet_name)
    if not source_data:
        print(f"No data found in source sheet '{source_sheet_name}'.")
        return 0
    
    # Verify that the source data contains the expected columns
    if len(source_data[0]) <= max(WINERY_COL, WEBSITE_COL, EMAIL_COL):
        print(f"Error: Source sheet does not have the expected columns. It has {len(source_data[0])} columns, but we need at least {max(WINERY_COL, WEBSITE_COL, EMAIL_COL) + 1} columns.")
        return 0
    
    # Get existing entries from queue sheet
    existing_entries = get_existing_entries(sheets, queue_spreadsheet_id, queue_sheet_name)
    
    # Get queue sheet headers and column indices
    queue_data = get_sheet_data(sheets, queue_spreadsheet_id, queue_sheet_name)
    if not queue_data:
        print(f"No data found in queue sheet '{queue_sheet_name}'.")
        return 0
    
    queue_headers = queue_data[0]
    country_col, name_col, email_col, website_col = get_queue_column_indices(queue_headers)
    if None in (country_col, name_col, email_col, website_col):
        return 0
    
    # Collect unique entries from source
    unique_entries = []
    seen_wineries = set()  # To track wineries we've already seen in this processing batch
    
    for row_idx, row in enumerate(source_data[1:], start=2):  # Skip header row, 1-indexed
        if len(row) <= WINERY_COL:  # Skip rows that don't have the winery field
            continue
        
        # Extract the data from the fixed column positions
        winery = row[WINERY_COL].strip() if WINERY_COL < len(row) else ""
        website = row[WEBSITE_COL].strip() if WEBSITE_COL < len(row) and len(row) > WEBSITE_COL else ""
        email = row[EMAIL_COL].strip() if EMAIL_COL < len(row) and len(row) > EMAIL_COL else ""
        
        # Skip empty entries
        if not winery:
            continue
        
        # Skip if already in the queue or if we've already seen it in this batch
        if winery.lower() in existing_entries or winery.lower() in seen_wineries:
            continue
        
        # Create a new row with the correct column order
        new_row = [""] * (max(country_col, name_col, email_col, website_col) + 1)
        new_row[country_col] = country
        new_row[name_col] = winery
        new_row[email_col] = email
        new_row[website_col] = website
        unique_entries.append(new_row)
        
        # Remember we've seen this winery
        seen_wineries.add(winery.lower())
    
    if not unique_entries:
        print("No new entries to add to the queue sheet.")
        return 0
    
    # Get the next empty row in the queue sheet
    next_row = len(queue_data) + 1
    
    # Update the queue sheet
    body = {
        'values': unique_entries
    }
    
    result = sheets.values().update(
        spreadsheetId=queue_spreadsheet_id,
        range=f"{queue_sheet_name}!A{next_row}",
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    
    print(f"Added {len(unique_entries)} new entries to the queue sheet:")
    for entry in unique_entries[:5]:  # Print first 5 entries
        print(f"  {entry[country_col]} - {entry[name_col]}")
    if len(unique_entries) > 5:
        print(f"  ... and {len(unique_entries) - 5} more")
    
    return len(unique_entries)

def main():
    """Main function to execute the script."""
    parser = argparse.ArgumentParser(description='Copy entries from a source sheet to the queue sheet.')
    parser.add_argument('--url', type=str, required=True,
                      help='Full Google Sheet URL of the source sheet (including sheet ID)')
    parser.add_argument('--country', type=str, required=True,
                      help='Country code to add to the Country column (e.g., AR for Argentina)')
    args = parser.parse_args()
    
    # Extract spreadsheet ID and sheet GID from URL
    source_spreadsheet_id, source_sheet_gid = extract_spreadsheet_info(args.url)
    if not source_spreadsheet_id:
        print("Error: Could not extract spreadsheet ID from the provided URL.")
        return
    
    print("Authenticating with Google Sheets API...")
    sheets = authenticate_google_sheets()
    
    # Get sheet name from GID
    source_sheet_metadata = sheets.get(spreadsheetId=source_spreadsheet_id).execute()
    source_sheet_name = None
    for sheet in source_sheet_metadata.get('sheets', []):
        if str(sheet['properties']['sheetId']) == source_sheet_gid:
            source_sheet_name = sheet['properties']['title']
            break
    
    if not source_sheet_name:
        print(f"Error: Could not find sheet with GID {source_sheet_gid}.")
        return
    
    # Process the source data and update the queue sheet
    try:
        entries_added = process_source_data(
            sheets,
            source_spreadsheet_id,
            source_sheet_name,
            QUEUE_SPREADSHEET_ID,
            QUEUE_SHEET_NAME,
            args.country
        )
        print(f"\nProcessing complete! Added {entries_added} new entries to the queue sheet.")
    except Exception as e:
        print(f"Error processing sheets: {str(e)}")

if __name__ == "__main__":
    main() 