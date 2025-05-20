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
    df = df.fillna("")  # Replace NaN with empty string so it's JSON safe


    print("📤 Uploading to Google Sheets...")
    upload_to_google_sheet(df)

    print("✅ Sync complete!")

if __name__ == "__main__":
    main()
