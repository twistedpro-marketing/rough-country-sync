import os
import json
import pandas as pd
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO


# Load JSON from environment variable
creds_json = os.environ.get("GOOGLE_CREDS_JSON")
if creds_json is None:
    raise Exception("Missing GOOGLE_CREDS_JSON environment variable")

SHEET_NAME = "Rough Country Inventory"

creds_dict = json.loads(creds_json)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)


# --- FTP File (Actually an HTTPS Feed from Rough Country) ---
EXCEL_URL = "https://feeds.roughcountry.com/jobber_pc1.xlsx"

def fetch_excel_from_rough_country():
    response = requests.get(EXCEL_URL)
    response.raise_for_status()
    return BytesIO(response.content)

def upload_to_google_sheet(df):
    scope = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive',
    ]
    
    creds_json = os.environ.get("GOOGLE_CREDS_JSON")
    if creds_json is None:
        raise Exception("Missing GOOGLE_CREDS_JSON environment variable")

    creds_dict = json.loads(creds_json)
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)

    sheet = client.open(SHEET_NAME).sheet1
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())


def main():
    print("📥 Downloading Excel from Rough Country...")
    excel_bytes = fetch_excel_from_rough_country()

    print("📊 Converting to DataFrame...")
    df = pd.read_excel(excel_bytes)

# Remove completely empty rows
df.dropna(how="all", inplace=True)

# Replace NaN with empty strings (safe for Sheets)
df.fillna("", inplace=True)

# Filter: Only include rows where inventory > 0
# (Adjust to the exact column name for inventory)
if "Qty On Hand" in df.columns:
    df = df[df["Qty On Hand"] > 0]

# Optional: Rename columns to match Shopify format
rename_map = {
    "backspacing": "1_backspacing",
    "diameter": "1_wheel_diameter",
    "size_desc": "description_tag",
    # Add more as needed
}

df.rename(columns=rename_map, inplace=True)



    print("📤 Uploading to Google Sheets...")
    upload_to_google_sheet(df)

    print("✅ Sync complete!")

if __name__ == "__main__":
    main()
