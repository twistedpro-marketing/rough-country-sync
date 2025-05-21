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
    for col in df.columns:
    if df[col].dtype == float:
        df[col] = df[col].fillna(0)  # or NaN-safe value like 0
    else:
        df[col] = df[col].fillna("")


    # Combine NV_Stock and TN_Stock into a new column
    if "NV_Stock" in df.columns and "TN_Stock" in df.columns:
        df["NV_Stock"] = pd.to_numeric(df["NV_Stock"], errors="coerce").fillna(0)
        df["TN_Stock"] = pd.to_numeric(df["TN_Stock"], errors="coerce").fillna(0)
        df["Total_Stock"] = df["NV_Stock"] + df["TN_Stock"]
    else:
        df["Total_Stock"] = 0  # fallback if columns missing

    # Filter: Only include rows
    df = df[df["Inventory"] > 0]
    rename_map = {
    "Inventory": "variant_inventory_qty",
    # ...
    }


    # Optional: Rename columns to match Shopify
    rename_map = {
        "backspacing": "1_backspacing",
        "diameter": "1_wheel_diameter",
        "size_desc": "description_tag",
        # Add more if needed
    }

    df.rename(columns=rename_map, inplace=True)

    print("📤 Uploading to Google Sheets...")
    upload_to_google_sheet(df)

    print("✅ Sync complete!")


if __name__ == "__main__":
    main()
