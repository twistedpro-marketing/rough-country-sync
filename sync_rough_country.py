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
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="50")

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

        # Separate UTV products
        utv_df = df[df['utv_product'].str.lower() == 'y']
        main_df = df[df['utv_product'].str.lower() != 'y']

        print(f"🥣 Found {len(utv_df)} UTV rows and {len(main_df)} non-UTV rows.")

        # Build Shopify-formatted export with multi-image handling
        shopify_rows = []
        for _, row in main_df.iterrows():
            handle = row["sku"].lower().replace(" ", "-")
            images = [row.get(f"image_{i}", "") for i in range(1, 7)]
            images = [img for img in images if img]
            availability_raw = str(row.get("availability", "")).lower()

            if "backorder" in availability_raw:
                restock_date = row.get("special_from_date", "")
                availability_status = f"Backorder - Restock {restock_date}" if restock_date else "Backorder"
            elif "in stock" in availability_raw or row.get("Inventory", 0) > 0:
                availability_status = "In Stock"
            else:
                availability_status = "Out of Stock"


            for i, img in enumerate(images):
                shopify_row = {
                    "Handle": handle,
                    "Image Src": img,
                    "Image Position": i + 1,
                }

                if i == 0:
                    # Format metafields with simple HTML
                    features = f"<ul><li>{'</li><li>'.join(str(row.get('features', '')).split(';'))}</li></ul>" if row.get("features") else ""
                    notes = f"<ul><li>{'</li><li>'.join(str(row.get('notes', '')).split(';'))}</li></ul>" if row.get("notes") else ""
                    fitment = row.get("fitment", "").replace(":", "|").replace(";", ";")

                    shopify_row.update({
                        "Title": row.get("title", ""),
                        "Body (HTML)": row.get("description", ""),
                        "Vendor": "Rough Country",
                        "Variant SKU": row["sku"],
                        "Variant Inventory Qty": row["Inventory"],
                        "Variant Price": row.get("price", 0),
                        "Variant Cost": row.get("cost", 0),
                        "Cost per item": row.get("cost", ""),
                        "product.metafields.custom.description_tag": row.get("size_desc", ""),
                        "product.metafields.custom.1_backspacing": row.get("backspacing", ""),
                        "product.metafields.custom.1_wheel_diameter": row.get("diameter", ""),
                        "product.metafields.custom.product_features": features,
                        "product.metafields.custom.important_notes": notes,
                        "product.metafields.custom.guides": row.get("instructions", ""),
                        "product.metafields.custom.time": row.get("install_time", ""),
                        "product.metafields.custom.tire_size": row.get("tire_info", ""),
                        "product.metafields.custom.components": f"{row.get('front_components', '')} {row.get('rear_components', '')}".strip(),
                        "product.metafields.custom.video_url": row.get("video", ""),
                        "product.metafields.convermax.fitment": fitment,
                        "Weight": row.get("weight", ""),
                        "Manufacturer": row.get("manufacturer", ""),
                        "UPC": row.get("upc", ""),
                        "Category": row.get("category", ""),
                        "product.metafields.custom.availability_status": availability_status,
                        "Variant Weight": row.get("weight", ""),
                        "Variant Barcode": row.get("upc", ""),
                        "product.metafields.custom.brand": row.get("manufacturer", ""),
                        "product.metafields.custom.category": row.get("category", ""),
                        "product.metafields.custom.variant_height": row.get("height", ""),
                        "product.metafields.custom.variant_width": row.get("width", ""),
                        "product.metafields.custom.variant_length": row.get("length", "")

                    })



                shopify_rows.append(shopify_row)

        shopify_df = pd.DataFrame(shopify_rows)

        # Reorder columns
        shopify_df = shopify_df[
            [
                "Handle",
                "Title",
                "Body (HTML)",
                "Vendor",
                "Variant SKU",
                "Variant Inventory Qty",
                "Variant Price",
                "Cost per item",  # <--- This must match the string key above
                "Image Src",
                "Image Position",
                "product.metafields.custom.availability_status",
                "product.metafields.custom.description_tag",
                "product.metafields.custom.1_backspacing",
                "product.metafields.custom.1_wheel_diameter",
                "product.metafields.custom.product_features",
                "product.metafields.custom.important_notes",
                "product.metafields.custom.guides",
                "product.metafields.custom.time",
                "product.metafields.custom.tire_size",
                "product.metafields.custom.video_url",
                "product.metafields.custom.components",
                "product.metafields.convermax.fitment",
                "Weight",
                "Manufacturer",
                "UPC",
                "Category"
            ]
        ]


        shopify_df = shopify_df.fillna("").astype(str)


        print("🛒 Sending Shopify data to export tab...")
        upload_shopify_sheet(shopify_df)
        print("✅ Shopify Export complete!")

        # Export UTV rows to dedicated sheet
        print("🥣 Sending UTV data to 'rough-country-UTV-only' tab...")
        upload_shopify_sheet(utv_df, sheet_name="rough-country-UTV-only")
        print("✅ UTV sheet export complete.")

    except Exception as e:
        print(f"❌ Script crashed: {e}")

    import time
    time.sleep(30)


if __name__ == "__main__":
    print("🐍 Script entrypoint reached.")
    main()
