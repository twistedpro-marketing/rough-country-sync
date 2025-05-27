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


def fetch_excel_from_rough_country():
    response = requests.get(EXCEL_URL)
    response.raise_for_status()
    return BytesIO(response.content)


def upload_to_google_sheet(df):
    sheet = client.open(SHEET_NAME).sheet1
    sheet.clear()
    sheet.update([df.columns.values.tolist()] + df.values.tolist())


def upload_shopify_sheet(df, sheet_name="Shopify Export"):
    spreadsheet = client.open(SHEET_NAME)
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        worksheet.clear()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")

    print(f"✅ Writing {len(df)} rows to '{sheet_name}' tab")
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())


def format_fitment(raw_fitment):
    try:
        if not raw_fitment or pd.isna(raw_fitment):
            return ""
        groups = raw_fitment.split(";")
        formatted = []
        for group in groups:
            parts = group.split(":")
            if len(parts) >= 4:
                year, drive, make, model = parts[:4]
                formatted.append(f"{year}|{make}|{model}|{drive}")
            elif len(parts) == 3:
                year, make, model = parts
                formatted.append(f"{year}|{make}|{model}")
            else:
                formatted.append(group)
        return ";".join(formatted)
    except:
        return ""


def main():
    print("🔍 Script has started.")

    try:
        print("📅 Downloading Excel from Rough Country...")
        excel_bytes = fetch_excel_from_rough_country()
        print("📊 Reading into DataFrame...")
        df = pd.read_excel(excel_bytes)
        print(f"✅ DataFrame loaded with {len(df)} rows.")
        print(f"🧠 Columns: {list(df.columns)}")

        df["NV_Stock"] = pd.to_numeric(df.get("NV_Stock", 0), errors="coerce").fillna(0)
        df["TN_Stock"] = pd.to_numeric(df.get("TN_Stock", 0), errors="coerce").fillna(0)
        df["Inventory"] = df["NV_Stock"] + df["TN_Stock"]

        for col in df.columns:
            if df[col].dtype == float:
                df[col] = df[col].fillna(0)
            else:
                df[col] = df[col].fillna("")

        df = df[df["Inventory"] > 0]
        print(f"🧹 Cleaned DataFrame has {len(df)} in-stock rows.")
        print("📤 Uploading full cleaned data...")
        upload_to_google_sheet(df)

        utv_df = df[df["utv_product"].str.lower() == "y"]
        non_utv_df = df[df["utv_product"].str.lower() != "y"]
        print(f"🥣 Found {len(utv_df)} UTV rows and {len(non_utv_df)} non-UTV rows.")

        shopify_rows = []
        for _, row in non_utv_df.iterrows():
            handle = row["sku"].lower().replace(" ", "-")
            images = [row.get(f"image_{i}", "") for i in range(1, 7)]
            images = [img for img in images if img]

            for i, img in enumerate(images):
                shopify_row = {
                    "Handle": handle,
                    "Image Src": img,
                    "Image Position": i + 1,
                }
                if i == 0:
                    shopify_row.update({
                        "Title": row.get("title", ""),
                        "Vendor": "Rough Country",
                        "Variant SKU": row["sku"],
                        "Variant Inventory Qty": row["Inventory"],
                        "Variant Price": row.get("price", 0),
                        "Body (HTML)": (
                            f"<p>{row.get('description', '')}</p>"
                            + f"<p><strong>Features:</strong> {row.get('features', '')}</p>"
                            + f"<p><strong>Notes:</strong> {row.get('notes', '')}</p>"
                            + f"<p><strong>Install Time:</strong> {row.get('install_time', '')}</p>"
                            + f"<p><strong>Fitment:</strong> {format_fitment(row.get('fitment', ''))}</p>"
                        ),
                        "Cost": row.get("cost", ""),
                        "Weight": row.get("weight", ""),
                        "Manufacturer": row.get("manufacturer", ""),
                        "UPC": row.get("upc", ""),
                        "Category": row.get("category", ""),
                        "product.metafields.custom.description_tag": row.get("size_desc", ""),
                        "product.metafields.custom.1_backspacing": row.get("backspacing", ""),
                        "product.metafields.custom.1_wheel_diameter": row.get("diameter", ""),
                    })
                shopify_rows.append(shopify_row)

        # === Load existing Shopify handles from separate sheet ===
        def load_existing_handles():
            try:
                handle_sheet = client.open("Rough Country Inventory").worksheet("rc-handles")


                data = handle_sheet.get_all_records()
                return pd.DataFrame(data)
            except Exception as e:
                print(f"⚠️ Could not load Shopify Handles sheet: {e}")
                return pd.DataFrame(columns=["SKU", "Handle"])

        existing_handles_df = load_existing_handles()=

        shopify_df = pd.DataFrame(shopify_rows)
        shopify_df = shopify_df.applymap(lambda x: "" if pd.isna(x) else x).fillna("")

        shopify_df = shopify_df[[
            "Handle",
            "Title",
            "Vendor",
            "Variant SKU",
            "Variant Inventory Qty",
            "Variant Price",
            "Body (HTML)",
            "Cost",
            "Weight",
            "Manufacturer",
            "UPC",
            "Category",
            "Image Src",
            "Image Position",
            "product.metafields.custom.description_tag",
            "product.metafields.custom.1_backspacing",
            "product.metafields.custom.1_wheel_diameter",
        ]]

        print("🛒 Sending Shopify data to export tab...")
        upload_shopify_sheet(shopify_df, sheet_name="Shopify Export")
        print("✅ Shopify Export complete!")

        print("🥣 Sending UTV data to 'rough-country-UTV-only' tab...")
        upload_shopify_sheet(utv_df, sheet_name="rough-country-UTV-only")
        print("✅ UTV sheet export complete.")

        # Merge original handles using SKU
        shopify_df = pd.merge(
            shopify_df,
            existing_handles_df.rename(columns={"SKU": "Variant SKU"}),
            on="Variant SKU",
            how="left"
        )

        # If handle exists in Shopify, use it. Otherwise, generate one
        shopify_df["Handle"] = shopify_df["Handle_y"].combine_first(shopify_df["Handle_x"])
        shopify_df.drop(columns=["Handle_x", "Handle_y"], inplace=True)

    except Exception as e:
        print(f"❌ Script crashed: {e}")

    import time
    time.sleep(30)


if __name__ == "__main__":
    print("🐍 Script entrypoint reached.")
    main()
