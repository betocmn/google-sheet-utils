#!/usr/bin/env python3
"""
Flag Suppliers in Exclusion List

This script checks a queue sheet against an exclusion list and flags potential matches
by highlighting rows in red. It uses fuzzy matching to identify similar entries across
different columns.

The script compares:
- "Winery or Supplier Name" with "Winery"
- "Email" with "Supplier Contact Email"
- "Website" with "Website"
- Domain names from email addresses and websites (excluding common email providers)

Usage:
    python scripts/flag_suppliers_in_exclusion_list.py

Prerequisites:
- The Google Sheet must be shared with the service account email address with Editor permissions
- The service account JSON key file must be in the 'keys' directory
"""

import os
import re
from fuzzywuzzy import fuzz
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Constants
QUEUE_SPREADSHEET_ID = '1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k'  # Replace with your queue spreadsheet ID
QUEUE_SHEET_NAME = '** Queue to Send Out Emails'  # Replace with your queue sheet name
EXCLUSION_LIST_SPREADSHEET_ID = '1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k'  # Replace with your exclusion list spreadsheet ID
EXCLUSION_LIST_SHEET_NAME = '***Exclusion List (GPD + Already Sent)'  # Replace with your exclusion list sheet name
SERVICE_ACCOUNT_FILE = 'keys/firm-harbor-456101-q0-5deb4bff67ac.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Similarity thresholds (0-100)
NAME_SIMILARITY_THRESHOLD = 80
EMAIL_SIMILARITY_THRESHOLD = 90
WEBSITE_SIMILARITY_THRESHOLD = 90
DOMAIN_SIMILARITY_THRESHOLD = 100  # Exact match for domains

# Common email provider base domains
# Instead of listing every possible country variation, we use base names
COMMON_EMAIL_PROVIDER_BASES = [
    'gmail',
    'yahoo',
    'hotmail',
    'outlook',
    'live',
    'msn',
    'aol',
    'icloud',
    'me.com',
    'mac.com',
    'mail',
    'comcast',
    'verizon',
    'att',
    'protonmail',
    'zoho',
    'gmx',
    'mailinator',
    'yandex',
    'sbcglobal',
    'cox',
    'earthlink',
    'rocketmail',
    'mindspring',
    'fastmail',
    'rediffmail',
    'btinternet',
    'naver',
    'qq.com',
    '126.com',
    '163.com',
    'bellsouth'
]

# Exact domains for specific cases
EXACT_COMMON_DOMAINS = [
    'pm.me',  # ProtonMail short domain
    'ymail.com',  # Yahoo
]

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
            return name_col, website_col, email_col
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

def normalize_text(text):
    """Normalize text for comparison by removing common variations."""
    if not text:
        return ""
    # Convert to lowercase
    text = text.lower()
    # Remove common business suffixes
    text = re.sub(r'\b(pty|ltd|limited|llc|inc|incorporated|corp|corporation)\b', '', text)
    # Remove special characters and extra spaces
    text = re.sub(r'[^\w\s]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_domain_from_email(email):
    """Extract domain name from an email address."""
    if not email:
        return ""
    try:
        return email.split('@')[1].lower()
    except (IndexError, AttributeError):
        return ""

def extract_domain_from_website(website):
    """Extract domain name from a website URL."""
    if not website:
        return ""
    
    # Remove protocol (http://, https://)
    domain = re.sub(r'^https?://', '', website.lower())
    
    # Remove www.
    domain = re.sub(r'^www\.', '', domain)
    
    # Remove everything after the first slash
    domain = domain.split('/')[0]
    
    # Remove port if present
    domain = domain.split(':')[0]
    
    return domain.strip()

def is_proprietary_domain(domain):
    """Check if a domain is proprietary (not a common email provider).
    Uses base domain matching to catch country-specific variants."""
    if not domain:
        return False
    
    domain = domain.lower()
    
    # Check for exact matches first
    if domain in EXACT_COMMON_DOMAINS:
        return False
    
    # Check if any common provider base is in the domain
    # We split the domain by dots and check the first part
    # This handles yahoo.com, yahoo.co.uk, yahoo.fr, etc.
    domain_parts = domain.split('.')
    if domain_parts and domain_parts[0] in COMMON_EMAIL_PROVIDER_BASES:
        return False
    
    # Check for domains that might be included as part of other TLDs
    for base in COMMON_EMAIL_PROVIDER_BASES:
        # Check if the base name is at the start of the domain
        # This handles cases like gmail.com, gmailuser.com.au, etc.
        if domain.startswith(f"{base}."):
            return False
    
    # If we get here, it's a proprietary domain
    return True

def is_similar(text1, text2, threshold):
    """Check if two texts are similar using fuzzy matching."""
    if not text1 or not text2:
        return False
    
    # Normalize both texts
    norm1 = normalize_text(text1)
    norm2 = normalize_text(text2)
    
    # Calculate similarity score
    similarity = fuzz.ratio(norm1, norm2)
    return similarity >= threshold

def find_matches(queue_data, exclusion_data):
    """Find potential matches between queue and exclusion list."""
    matches = []
    
    # Get column indices
    queue_headers = queue_data[0]
    exclusion_headers = exclusion_data[0]
    
    queue_name_col, queue_website_col, queue_email_col = get_column_indices(queue_headers, 'queue')
    excl_name_col, excl_email_col, excl_website_col = get_column_indices(exclusion_headers, 'exclusion')
    
    if None in (queue_name_col, excl_name_col):  # Only require name columns
        return matches
    
    # Process queue data
    for row_idx, row in enumerate(queue_data[1:], start=2):  # Start from 2 to account for header row
        if len(row) <= queue_name_col:  # Only check if we have enough columns for name
            continue
            
        queue_name = row[queue_name_col].strip() if queue_name_col < len(row) else ""
        queue_website = row[queue_website_col].strip() if queue_website_col < len(row) and queue_website_col is not None else ""
        queue_email = row[queue_email_col].strip() if queue_email_col < len(row) and queue_email_col is not None else ""
        
        # Extract domains
        queue_email_domain = extract_domain_from_email(queue_email)
        queue_website_domain = extract_domain_from_website(queue_website)
        
        # Skip if no name
        if not queue_name:
            continue
        
        # Check against exclusion list
        for excl_row in exclusion_data[1:]:
            if len(excl_row) <= excl_name_col:  # Only check if we have enough columns for name
                continue
                
            excl_name = excl_row[excl_name_col].strip() if excl_name_col < len(excl_row) else ""
            excl_email = excl_row[excl_email_col].strip() if excl_email_col < len(excl_row) and excl_email_col is not None else ""
            excl_website = excl_row[excl_website_col].strip() if excl_website_col < len(excl_row) and excl_website_col is not None else ""
            
            # Extract domains from exclusion list
            excl_email_domain = extract_domain_from_email(excl_email)
            excl_website_domain = extract_domain_from_website(excl_website)
            
            # Check for matches
            name_match = is_similar(queue_name, excl_name, NAME_SIMILARITY_THRESHOLD)
            email_match = queue_email and excl_email and is_similar(queue_email, excl_email, EMAIL_SIMILARITY_THRESHOLD)
            website_match = queue_website and excl_website and is_similar(queue_website, excl_website, WEBSITE_SIMILARITY_THRESHOLD)
            
            # Domain matching - only for proprietary domains (not gmail, yahoo, etc.)
            email_domain_match = False
            website_domain_match = False
            cross_domain_match1 = False
            cross_domain_match2 = False
            
            # Only check domains if they are proprietary
            if is_proprietary_domain(queue_email_domain) and is_proprietary_domain(excl_email_domain):
                email_domain_match = (queue_email_domain and excl_email_domain and 
                                      is_similar(queue_email_domain, excl_email_domain, DOMAIN_SIMILARITY_THRESHOLD))
            
            if is_proprietary_domain(queue_website_domain) and is_proprietary_domain(excl_website_domain):
                website_domain_match = (queue_website_domain and excl_website_domain and 
                                       is_similar(queue_website_domain, excl_website_domain, DOMAIN_SIMILARITY_THRESHOLD))
            
            # Cross-domain matching (email domain vs website domain)
            if is_proprietary_domain(queue_email_domain) and is_proprietary_domain(excl_website_domain):
                cross_domain_match1 = (queue_email_domain and excl_website_domain and 
                                      is_similar(queue_email_domain, excl_website_domain, DOMAIN_SIMILARITY_THRESHOLD))
            
            if is_proprietary_domain(queue_website_domain) and is_proprietary_domain(excl_email_domain):
                cross_domain_match2 = (queue_website_domain and excl_email_domain and 
                                      is_similar(queue_website_domain, excl_email_domain, DOMAIN_SIMILARITY_THRESHOLD))
            
            if (name_match or email_match or website_match or 
                email_domain_match or website_domain_match or 
                cross_domain_match1 or cross_domain_match2):
                matches.append({
                    'row': row_idx,
                    'matches': {
                        'name': name_match,
                        'email': email_match,
                        'website': website_match,
                        'email_domain': email_domain_match,
                        'website_domain': website_domain_match,
                        'cross_domain1': cross_domain_match1,
                        'cross_domain2': cross_domain_match2
                    },
                    'queue_data': {
                        'name': queue_name,
                        'email': queue_email,
                        'website': queue_website,
                        'email_domain': queue_email_domain,
                        'website_domain': queue_website_domain
                    },
                    'exclusion_data': {
                        'name': excl_name,
                        'email': excl_email,
                        'website': excl_website,
                        'email_domain': excl_email_domain,
                        'website_domain': excl_website_domain
                    }
                })
                break  # Found a match, no need to check other exclusion list entries
    
    return matches

def get_sheet_id(sheets, spreadsheet_id, sheet_name):
    """Get the numeric sheet ID for a given sheet name."""
    sheet_metadata = sheets.get(spreadsheetId=spreadsheet_id).execute()
    sheets_data = sheet_metadata.get('sheets', [])
    
    for sheet in sheets_data:
        if sheet['properties']['title'] == sheet_name:
            return sheet['properties']['sheetId']
    
    raise ValueError(f"Sheet '{sheet_name}' not found in spreadsheet")

def highlight_matches(sheets, spreadsheet_id, sheet_name, matches):
    """Highlight matching rows in the queue sheet."""
    if not matches:
        print("No matches found to highlight.")
        return
    
    try:
        # Get the numeric sheet ID
        sheet_id = get_sheet_id(sheets, spreadsheet_id, sheet_name)
        
        # Prepare the batch update request
        requests = []
        for match in matches:
            # Create a red background color format
            requests.append({
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': match['row'] - 1,  # Convert to 0-based index
                        'endRowIndex': match['row'],
                        'startColumnIndex': 0,
                        'endColumnIndex': 100  # Cover all columns
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {
                                'red': 1.0,
                                'green': 0.0,
                                'blue': 0.0
                            }
                        }
                    },
                    'fields': 'userEnteredFormat.backgroundColor'
                }
            })
        
        # Execute the batch update
        body = {
            'requests': requests
        }
        
        result = sheets.batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        
        print(f"Highlighted {len(matches)} rows in the queue sheet.")
    except Exception as e:
        print(f"Error highlighting matches: {str(e)}")

def print_queue_data(queue_data):
    """Print queue data in a readable format."""
    if not queue_data:
        print("No queue data to display.")
        return
    
    print("\nRaw queue data:")
    for i, row in enumerate(queue_data):
        print(f"Row {i}: {row}")
    
    # Get column indices
    headers = queue_data[0]
    print("\nLooking for columns in queue sheet:")
    print("Headers:", headers)
    
    name_col, website_col, email_col = get_column_indices(headers, 'queue')
    
    if None in (name_col, website_col, email_col):
        print("Could not find required columns in queue data.")
        return
    
    print(f"\nFound columns at indices: Name={name_col}, Website={website_col}, Email={email_col}")
    
    print("\nQueue Data:")
    print("-" * 140)
    print(f"{'Row':<5} | {'Name':<30} | {'Website':<30} | {'Email':<30} | {'Domain':<20} | {'Proprietary':<10}")
    print("-" * 140)
    
    for row_idx, row in enumerate(queue_data[1:], start=2):  # Start from 2 to account for header row
        if len(row) <= name_col:  # Only check if we have enough columns for name
            print(f"Skipping row {row_idx} - not enough columns for name")
            continue
            
        name = row[name_col].strip() if name_col < len(row) else ""
        website = row[website_col].strip() if website_col < len(row) and website_col is not None else ""
        email = row[email_col].strip() if email_col < len(row) and email_col is not None else ""
        
        # Extract domains
        email_domain = extract_domain_from_email(email)
        website_domain = extract_domain_from_website(website)
        domain = email_domain or website_domain
        is_proprietary = "Yes" if is_proprietary_domain(domain) else "No"
        
        # Skip if no name
        if not name:
            print(f"Skipping row {row_idx} - empty name")
            continue
            
        print(f"{row_idx:<5} | {name[:30]:<30} | {website[:30]:<30} | {email[:30]:<30} | {domain[:20]:<20} | {is_proprietary:<10}")
    
    print("-" * 140)
    print(f"Total rows in queue: {len(queue_data) - 1}")  # Subtract header row

def main():
    """Main function to execute the script."""
    print("Authenticating with Google Sheets API...")
    sheets = authenticate_google_sheets()
    
    # Get data from both sheets
    print("Retrieving queue data...")
    queue_data = get_sheet_data(sheets, QUEUE_SPREADSHEET_ID, QUEUE_SHEET_NAME)
    if not queue_data:
        print(f"No data found in queue sheet '{QUEUE_SHEET_NAME}'.")
        return
    
    # Print queue data for verification
    print_queue_data(queue_data)
    
    print("\nRetrieving exclusion list data...")
    exclusion_data = get_sheet_data(sheets, EXCLUSION_LIST_SPREADSHEET_ID, EXCLUSION_LIST_SHEET_NAME)
    if not exclusion_data:
        print(f"No data found in exclusion list sheet '{EXCLUSION_LIST_SHEET_NAME}'.")
        return
    
    # Find matches
    print("\nFinding potential matches...")
    matches = find_matches(queue_data, exclusion_data)
    
    if matches:
        print(f"\nFound {len(matches)} potential matches:")
        for match in matches:
            print(f"\nRow {match['row']}:")
            print(f"Queue: {match['queue_data']}")
            print(f"Exclusion List: {match['exclusion_data']}")
            print("Matches:")
            for field, is_match in match['matches'].items():
                if is_match:
                    print(f"- {field}")
            
            # Check if the match is based on a proprietary domain
            domains = [
                match['queue_data']['email_domain'],
                match['queue_data']['website_domain'],
                match['exclusion_data']['email_domain'],
                match['exclusion_data']['website_domain']
            ]
            all_common = all(not is_proprietary_domain(d) for d in domains if d)
            if all_common:
                print("⚠️ Warning: This match is based on common email providers only and may be a false positive.")
        
        # Highlight matches in the queue sheet
        print("\nHighlighting matches in the queue sheet...")
        highlight_matches(sheets, QUEUE_SPREADSHEET_ID, QUEUE_SHEET_NAME, matches)
    else:
        print("No potential matches found.")

if __name__ == "__main__":
    main() 