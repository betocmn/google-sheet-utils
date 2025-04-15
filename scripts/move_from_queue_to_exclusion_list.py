#!/usr/bin/env python3
"""
Move From Queue to Exclusion List

This script processes a Google Sheet queue and moves unique entries to an exclusion list spreadsheet.
It extracts "Winery or Supplier Name", "Email", and "Website" columns from the queue
and adds them to the exclusion list with columns "Winery or Supplier Name", "Email", "Website".
After copying entries to the exclusion list, they are removed from the queue sheet.

Usage:
    python scripts/move_from_queue_to_exclusion_list.py

Prerequisites:
- The Google Sheet must be shared with the service account email address with Editor permissions
- The service account JSON key file must be in the 'keys' directory
"""

import os
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Constants
QUEUE_SPREADSHEET_ID = '1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k'  # Replace with your queue spreadsheet ID
QUEUE_SHEET_NAME = '** Queue to Send Out Emails'  # Replace with your queue sheet name
EXCLUSION_LIST_SPREADSHEET_ID = '1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k'  # Replace with your exclusion list spreadsheet ID
EXCLUSION_LIST_SHEET_NAME = '***Exclusion List (GPD + Already Sent)'  # Replace with your exclusion list sheet name
SERVICE_ACCOUNT_FILE = 'keys/firm-harbor-456101-q0-5deb4bff67ac.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

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
    if values:
        print("First row (headers):")
        for i, header in enumerate(values[0]):
            print(f"Column {i}: '{header}'")
    return values

def get_column_indices(headers, sheet_type):
    """Get the indices of required columns based on sheet type."""
    if sheet_type == 'queue':
        try:
            # Updated column names to match new queue sheet format
            name_col = headers.index("Winery or Supplier Name")
            email_col = headers.index("Email")
            website_col = headers.index("Website")
            return name_col, email_col, website_col
        except ValueError as e:
            print(f"Error: Required column not found in queue sheet - {str(e)}")
            print("Available columns in queue sheet:")
            for i, header in enumerate(headers):
                print(f"{i}: '{header}'")
            return None, None, None
    else:  # exclusion list
        try:
            name_col = headers.index("Winery or Supplier Name")
            email_col = headers.index("Email")
            website_col = headers.index("Website")
            return name_col, email_col, website_col
        except ValueError as e:
            print(f"Error: Required column not found in exclusion list - {str(e)}")
            print("Available columns in exclusion list:")
            for i, header in enumerate(headers):
                print(f"{i}: '{header}'")
            return None, None, None

def get_existing_entries(sheets, spreadsheet_id, sheet_name):
    """Get existing entries from the exclusion list to avoid duplicates."""
    data = get_sheet_data(sheets, spreadsheet_id, sheet_name)
    if not data:
        return set()
    
    # Get column indices
    headers = data[0]
    name_col, email_col, website_col = get_column_indices(headers, 'exclusion')
    if None in (name_col, email_col, website_col):
        return set()
    
    # Skip header row
    existing_entries = set()
    for row in data[1:]:
        if len(row) > max(name_col, email_col, website_col):
            name = row[name_col].strip() if name_col < len(row) else ""
            email = row[email_col].strip() if email_col < len(row) else ""
            website = row[website_col].strip() if website_col < len(row) else ""
            
            if name or email or website:  # Only add non-empty entries
                entry = (name, email, website)
                existing_entries.add(entry)
    
    return existing_entries

def clear_queue_rows(sheets, spreadsheet_id, sheet_name, rows_to_clear):
    """Clear the specified rows from the queue sheet."""
    if not rows_to_clear:
        return
    
    # Sort rows in descending order to avoid index shifting issues
    rows_to_clear.sort(reverse=True)
    
    # Get the sheet ID
    sheet_metadata = sheets.get(spreadsheetId=spreadsheet_id).execute()
    sheet_id = None
    for sheet in sheet_metadata.get('sheets', []):
        if sheet['properties']['title'] == sheet_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    if not sheet_id:
        print(f"Error: Could not find sheet ID for '{sheet_name}'")
        return
    
    # Prepare the batch update request
    requests = []
    for row in rows_to_clear:
        requests.append({
            'deleteDimension': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': row - 1,  # Convert to 0-based index
                    'endIndex': row
                }
            }
        })
    
    # Execute the batch update
    body = {
        'requests': requests
    }
    
    try:
        result = sheets.batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        print(f"Removed {len(rows_to_clear)} rows from the queue sheet.")
    except Exception as e:
        print(f"Error removing rows from queue sheet: {str(e)}")

def process_queue_data(sheets, queue_spreadsheet_id, queue_sheet_name, exclusion_spreadsheet_id, exclusion_sheet_name):
    """Process the queue data and move unique entries to the exclusion list."""
    # Get queue data
    queue_data = get_sheet_data(sheets, queue_spreadsheet_id, queue_sheet_name)
    if not queue_data:
        print(f"No data found in queue sheet '{queue_sheet_name}'.")
        return 0
    
    # Get column indices
    headers = queue_data[0]
    name_col, email_col, website_col = get_column_indices(headers, 'queue')
    if None in (name_col, email_col, website_col):
        return 0
    
    # Get existing entries from exclusion list
    existing_entries = get_existing_entries(sheets, exclusion_spreadsheet_id, exclusion_sheet_name)
    
    # First, collect unique entries from the queue
    unique_queue_entries = set()
    rows_to_remove = []  # Track which rows to remove
    for row_idx, row in enumerate(queue_data[1:], start=2):  # Start from 2 to account for header row
        if len(row) > max(name_col, email_col, website_col):
            name = row[name_col].strip() if name_col < len(row) else ""
            email = row[email_col].strip() if email_col < len(row) else ""
            website = row[website_col].strip() if website_col < len(row) else ""
            
            # Skip empty entries
            if not name and not email and not website:
                continue
                
            entry = (name, email, website)
            if entry not in existing_entries:
                unique_queue_entries.add(entry)
                rows_to_remove.append(row_idx)
    
    if not unique_queue_entries:
        print("No new entries to add to the exclusion list.")
        return 0
    
    # Prepare the data for batch update
    values = [[entry[0], entry[1], entry[2]] for entry in unique_queue_entries]
    
    # Get the next empty row in the exclusion list
    exclusion_data = get_sheet_data(sheets, exclusion_spreadsheet_id, exclusion_sheet_name)
    next_row = len(exclusion_data) + 1 if exclusion_data else 2  # Start after header row
    
    # Update the exclusion list
    body = {
        'values': values
    }
    
    result = sheets.values().update(
        spreadsheetId=exclusion_spreadsheet_id,
        range=f"{exclusion_sheet_name}!A{next_row}",
        valueInputOption='USER_ENTERED',
        body=body
    ).execute()
    
    print(f"Added {len(unique_queue_entries)} new entries to the exclusion list.")
    
    # Remove the processed entries from the queue sheet
    clear_queue_rows(sheets, queue_spreadsheet_id, queue_sheet_name, rows_to_remove)
    
    return len(unique_queue_entries)

def main():
    """Main function to execute the script."""
    print("Authenticating with Google Sheets API...")
    sheets = authenticate_google_sheets()
    
    # Process the queue and update the exclusion list
    try:
        entries_added = process_queue_data(
            sheets,
            QUEUE_SPREADSHEET_ID,
            QUEUE_SHEET_NAME,
            EXCLUSION_LIST_SPREADSHEET_ID,
            EXCLUSION_LIST_SHEET_NAME
        )
        print(f"\nProcessing complete! Added {entries_added} new entries to the exclusion list.")
    except Exception as e:
        print(f"Error processing sheets: {str(e)}")

if __name__ == "__main__":
    main() 