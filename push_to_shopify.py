import os
import json
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# === Shopify Credentials from Environment ===
SHOPIFY_API_KEY = os.environ.get("SHOPIFY_API_KEY")
SHOPIFY_API_PASSWORD = os.environ.get("SHOPIFY_API_PASSWORD")
SHOPIFY_STORE_DOMAIN = os.environ.get("SHOPIFY_STORE_DOMAIN")

SHOPIFY_API_BASE = f"https://{SHOPIFY_API_KEY}:{SHOPIFY_API_PASSWORD}@{SHOPIFY_STORE_DOMAIN}/admin/api/2024-04"

# === Google Auth ===
creds_dict = json.loads(os.environ.get("GOOGLE_CREDS_JSON"))
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

# === Load product data from Google Sheet ===
SHEET_NAME = "Rough Country Inventory"
sheet = client.open(SHEET_NAME)
worksheet = sheet.worksheet("Shopify Export")
rows = worksheet.get_all_records()

# === Pick the first product as a test ===
first_product = rows[0]

# === Construct Shopify Product Payload ===
shopify_payload = {
    "product": {
        "title": first_product.get("Title", ""),
        "body_html": first_product.get("Body (HTML)", ""),
        "vendor": first_product.get("Vendor", ""),
        "tags": "",
        "variants": [
            {
                "sku": first_product.get("Variant SKU", ""),
                "price": first_product.get("Variant Price", ""),
                "inventory_quantity": first_product.get("Variant Inventory Qty", ""),
                "weight": first_product.get("Weight", ""),
                "weight_unit": "lb",
                "barcode": first_product.get("UPC", ""),
                "inventory_management": "shopify",
                "cost": first_product.get("Cost per item", "")
            }
        ],
        "images": [
            {"src": first_product.get("Image Src", "")}
        ]
    }
}

# === POST to Shopify ===
response = requests.post(
    f"{SHOPIFY_API_BASE}/products.json",
    headers={"Content-Type": "application/json"},
    data=json.dumps(shopify_payload)
)

# === Output result ===
print("📦 Shopify Response Status:", response.status_code)
try:
    print(json.dumps(response.json(), indent=2))
except:
    print(response.text)
