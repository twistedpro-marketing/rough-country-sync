import os
import json
import pandas as pd
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO

SHEET_NAME = "Rough Country Inventory"
EXCEL_URL = "https://feeds.roughcountry.com/jobber_pc1.xlsx"

# === Google Auth ===
creds_json = os.environ.get("GOOGLE_CREDS_JSON")
if creds_json is None:
    raise Exception("Missing GOOGLE_CREDS_JSON environment variable")

creds_dict = json.loads(creds_json)
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)


# === Fetch Excel File ===
def fetch_excel_from_rough_country():
    response = requests.get(EXCEL_URL)
    response.raise_for_status()
    return BytesIO(response.content)


# === Upload Full Cleaned Data to Main Sheet ===
def upload_to_google_sheet(df):
    sheet = client.open(SHEET_NAME).sheet1
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())


# === Upload Shopify-formatted Data to Separate Tab ===
def upload_shopify_sheet(df, sheet_name="Shopify Export"):
    spreadsheet = client.open(SHEET_NAME)

    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        worksheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")

    print(f"✅ Writing {len(df)} rows to '{sheet_name}' tab")

    worksheet.update([df.columns.values.tolist()] + df.values.tolist())


# === Main ===
def main():
    print("📥 Downloading Excel from Rough Country...")
    excel_bytes = fetch_excel_from_rough_country()

    print("📊 Converting to DataFrame...")
    df = pd.read_excel(excel_bytes)

    print("🛒 Sending Shopify data to export tab...")

    upload_shopify_sheet(shopify_df)

    # Clean and prepare data
    df.dropna(how="all", inplace=True)

    for col in df.columns:
        if df[col].dtype == float:
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("")

    # Combine NV and TN stock
    if "NV_Stock" in df.columns and "TN_Stock" in df.columns:
        df["NV_Stock"] = pd.to_numeric(df["NV_Stock"], errors="coerce").fillna(0)
        df["TN_Stock"] = pd.to_numeric(df["TN_Stock"], errors="coerce").fillna(0)
        df["Inventory"] = df["NV_Stock"] + df["TN_Stock"]
    else:
        df["Inventory"] = 0

    df = df[df["Inventory"] > 0]  # Only keep in-stock items

    print("📤 Uploading full cleaned data...")
    upload_to_google_sheet(df)

    # === Build Shopify-formatted DataFrame ===
    shopify_df = pd.DataFrame()
    shopify_df["Handle"] = df["Part #"].astype(str).str.lower().str.replace(" ", "-").str.replace(r"[^\w\-]", "", regex=True)
    shopify_df["Title"] = df["Description"] if "Description" in df.columns else df["Part #"]
    shopify_df["Vendor"] = "Rough Country"
    shopify_df["Variant SKU"] = df["Part #"]
    shopify_df["Variant Inventory Qty"] = df["Inventory"]
    shopify_df["Variant Price"] = df.get("Jobber", 0)
    shopify_df["Image Src"] = df.get("Image Link", "")

    # === Add Metafields ===
    shopify_df["product.metafields.custom.description_tag"] = df.get("size_desc", "")
    shopify_df["product.metafields.custom.1_backspacing"] = df.get("backspacing", "")
    shopify_df["product.metafields.custom.1_wheel_diameter"] = df.get("diameter", "")

    # === Final column order ===
    shopify_df = shopify_df[
        [
            "Handle",
            "Title",
            "Vendor",
            "Variant SKU",
            "Variant Inventory Qty",
            "Variant Price",
            "Image Src",
            "product.metafields.custom.description_tag",
            "product.metafields.custom.1_backspacing",
            "product.metafields.custom.1_wheel_diameter",
        ]
    ]
