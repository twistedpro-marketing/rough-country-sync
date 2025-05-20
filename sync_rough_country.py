import pandas as pd
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

# --- Google Sheets Config ---
SHEET_NAME = "Rough Country Inventory"  # Change if your sheet name is different
CREDENTIALS_FILE = "credentials.json"   # This will be used in Railway later

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
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)

    sheet = client.open(SHEET_NAME).sheet1
    sheet.clear()  # Clear previous data
    sheet.update([df.columns.values.tolist()] + df.values.tolist())

def main():
    print("📥 Downloading Excel from Rough Country...")
    excel_bytes = fetch_excel_from_rough_country()

    print("📊 Converting to DataFrame...")
    df = pd.read_excel(excel_bytes)

    print("📤 Uploading to Google Sheets...")
    upload_to_google_sheet(df)

    print("✅ Sync complete!")

if __name__ == "__main__":
    main()
