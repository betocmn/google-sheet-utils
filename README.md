# Google Sheet Utils

A collection of utility scripts for working with Google Sheets data.

## Available Utilities

1. **Price Updater** (`scripts/sheet_price_updater.py`) - Updates pricing information by:
   - Converting non-AUD prices to AUD
   - Converting case prices to per-bottle prices

## Prerequisites

- Python 3.6+
- A Google Cloud service account with access to the Google Sheets API
- Google Sheets must be shared with the service account email address with Editor permissions

## Setup

1. Clone this repository:
   ```
   git clone <repository_url>
   cd google-sheet-utils
   ```

2. Create a virtual environment:
   ```bash
   uv venv --python 3.11
   ```

3. Activate the environment:
   ```bash
   # For zsh/bash:
   source .venv/bin/activate

   # For fish shell:
   source .venv/bin/activate.fish
   ```

4. Install the required packages:
   ```bash
   # Using uv (recommended):
   uv pip install -r requirements.txt
   
   # Alternative if uv pip doesn't work:
   .venv/bin/python -m ensurepip --upgrade
   .venv/bin/python -m pip install -r requirements.txt
   ```

5. Place your service account JSON key file in the `keys` directory:
   - This directory is gitignored for security
   - The path to the service account key should be `keys/YOUR_KEY.json`

6. **IMPORTANT:** Share your Google Sheets with the service account email address:
   - Open each Google Sheet you want to manipulate
   - Click the "Share" button in the top right corner
   - Add the service account email as an Editor
   - The email address can be found in your service account JSON file as "client_email"
   - Without this step, the scripts will not be able to access or modify your sheets

## Usage

Each script in the `scripts` directory serves a specific purpose. Run them directly:

```bash
# Process a specific sheet using a full Google Sheet URL (preferred)
python scripts/sheet_price_updater.py --url "https://docs.google.com/spreadsheets/d/1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k/edit?gid=2104612227#gid=2104612227"

# Process a specific sheet by name (requires spreadsheet ID to be set in the script)
python scripts/sheet_price_updater.py --sheet "Sheet Name"

# List all available sheets in the default spreadsheet
python scripts/sheet_price_updater.py --list

# List all available sheets in a specific spreadsheet
python scripts/sheet_price_updater.py --list --url "https://docs.google.com/spreadsheets/d/1JsrKKr_jZJqLmJEKxpyPygZ4lUu19-GW2dnt7p7g-1k/edit"
```

## Adding New Utilities

To add a new utility script:

1. Create a new Python file in the `scripts` directory
2. Add appropriate documentation within the script
3. Update the "Available Utilities" section in this README

## Configuration

Most scripts use environment variables or configuration parameters at the top of the file. See each script's documentation for specific configuration options. 