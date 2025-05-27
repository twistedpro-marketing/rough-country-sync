import os
print("🐍 Script entrypoint reached.")

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
    print("🔍 Script has started.")

    try:
        # Fetch and load Excel
        print("📥 Downloading Excel from Rough Country...")
        excel_bytes = fetch_excel_from_rough_country()
        print("📊 Reading into DataFrame...")
        df = pd.read_excel(excel_bytes)
        print(f"✅ DataFrame loaded with {len(df)} rows.")
        print(f"🧠 Columns: {list(df.columns)}")

        # Combine stock columns
        df["NV_Stock"] = pd.to_numeric(df.get("NV_Stock", 0), errors="coerce").fillna(0)
        df["TN_Stock"] = pd.to_numeric(df.get("TN_Stock", 0), errors="coerce").fillna(0)
        df["Inventory"] = df["NV_Stock"] + df["TN_Stock"]

        # Filter and clean
    for col in df.columns:
        if df[col].dtype == float:
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("")


        print(f"🧹 Cleaned DataFrame has {len(df)} in-stock rows.")

        # Upload full inventory
        print("📤 Uploading full cleaned data...")
        upload_to_google_sheet(df)

        # Build Shopify-formatted export with multi-image handling
        shopify_rows = []

        for _, row in df.iterrows():
            handle = row["sku"].lower().replace(" ", "-")
            images = [row.get(f"image_{i}", "") for i in range(1, 7)]
            images = [img for img in images if img]  # Remove blanks

            for i, img in enumerate(images):
                shopify_row = {
                    "Handle": handle,
                    "Image Src": img,
                    "Image Position": i + 1,
                }

                # Add full product details only for the first image
                if i == 0:
                    shopify_row.update({
                        "Title": row.get("description", ""),
                        "Vendor": "Rough Country",
                        "Variant SKU": row["sku"],
                        "Variant Inventory Qty": row["Inventory"],
                        "Variant Price": row.get("price", 0),
                        "product.metafields.custom.description_tag": row.get("size_desc", ""),
                        "product.metafields.custom.1_backspacing": row.get("backspacing", ""),
                        "product.metafields.custom.1_wheel_diameter": row.get("diameter", ""),
                    })

                shopify_rows.append(shopify_row)

        shopify_df = pd.DataFrame(shopify_rows)


        # Reorder columns
        shopify_df = shopify_df[
        [
            "Handle",
            "Title",
            "Vendor",
            "Variant SKU",
            "Variant Inventory Qty",
            "Variant Price",
            "Image Src",
            "Image Position",
            "product.metafields.custom.description_tag",
            "product.metafields.custom.1_backspacing",
            "product.metafields.custom.1_wheel_diameter",
        ]
    ]


        print("🛒 Sending Shopify data to export tab...")
        upload_shopify_sheet(shopify_df)
        print("✅ Sync complete!")

    except Exception as e:
        print(f"❌ Script crashed: {e}")

    # Keep container alive for 30s to view logs
    import time
    time.sleep(30)


if __name__ == "__main__":
    print("🐍 Script entrypoint reached.")
    main()
