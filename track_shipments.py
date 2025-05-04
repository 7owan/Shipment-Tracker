import os
import sys
import requests
from xml.etree import ElementTree
from openpyxl import load_workbook
from dotenv import load_dotenv

load_dotenv()  # Load .env values

def print_progress_bar(iteration, total, length=40):
    percent = f"{100 * (iteration / float(total)):.1f}"
    filled_length = int(length * iteration // total)
    bar = '█' * filled_length + '-' * (length - filled_length)
    sys.stdout.write(f'\rProgress: |{bar}| {percent}% ({iteration}/{total})')
    sys.stdout.flush()

def get_auth_token(client_id, client_secret):
    url = "https://apis.fedex.com/oauth/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret
    }
    response = requests.post(url, headers=headers, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

def get_delivery_date_fedex(access_token, tracking_number):
    url = "https://apis.fedex.com/track/v1/trackingnumbers"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {access_token}"
    }
    payload = {
        "trackingInfo": [
            {
                "trackingNumberInfo": {
                    "trackingNumber": tracking_number
                }
            }
        ],
        "includeDetailedScans": False
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    result = response.json()

    try:
        shipment = result["output"]["completeTrackResults"][0]["trackResults"][0]
        dates = shipment.get("dateAndTimes", [])
        for date in dates:
            if date["type"] in ["ACTUAL_DELIVERY", "ESTIMATED_DELIVERY"]:
                return date["dateTime"][:10]
    except Exception as e:
        print(f"\nFedEx error: {e}")
    return None

def get_delivery_date_aduiepyle(user_email, tracking_number):
    url = "https://api.aduiepyle.com/2/shipment/status"
    params = {
        'user': user_email,
        'type': 0,
        'value': tracking_number
    }
    response = requests.get(url, params=params)
    if response.status_code == 200:
        try:
            root = ElementTree.fromstring(response.text)
            for status in root.findall(".//statusDetail"):
                if status.find("description").text == "DELIVERED":
                    return status.find("start").text[:10]
        except ElementTree.ParseError as e:
            print(f"\nXML parse error: {e}")
    return None

def process_tracking_sheet(filename, sheet_name="Sheet1"):
    FEDEX_CLIENT_ID = os.getenv("FEDEX_CLIENT_ID")
    FEDEX_CLIENT_SECRET = os.getenv("FEDEX_CLIENT_SECRET")
    ADUIEPYLE_EMAIL = os.getenv("ADUIEPYLE_EMAIL")

    try:
        fedex_token = get_auth_token(FEDEX_CLIENT_ID, FEDEX_CLIENT_SECRET)
    except Exception as e:
        print("FedEx auth error:", e)
        fedex_token = None

    wb = load_workbook(filename)
    ws = wb.active  # default to first sheet

    # Map headers to their column indices
    header_row = [cell.value.strip() if cell.value else "" for cell in ws[1]]
    col_indices = {
        "Carrier": None,
        "Tracking": None,
        "Delivered": None
    }

    for idx, header in enumerate(header_row):
        header_lower = header.lower()
        if header_lower == "carrier":
            col_indices["Carrier"] = idx
        elif header_lower in ["pro #", "pro number"]:
            col_indices["Tracking"] = idx
        elif header_lower == "delivered date":
            col_indices["Delivered"] = idx

    if None in col_indices.values():
        print("Missing required column(s):", col_indices)
        return

    rows = list(ws.iter_rows(min_row=2))
    total_rows = len(rows)
    updated_rows = 0

    for i, row in enumerate(rows, 1):
        carrier = (row[col_indices["Carrier"]].value or "").strip().upper()
        tracking = (row[col_indices["Tracking"]].value or "").strip()
        delivered_value = row[col_indices["Delivered"]].value

        # Skip if Delivered DATE is already filled
        if delivered_value or not carrier or not tracking:
            print_progress_bar(i, total_rows)
            continue

        try:
            if carrier in ["FEP", "FEE", "FEU", "FEA", "FED", "FEC"] and fedex_token:
                date = get_delivery_date_fedex(fedex_token, tracking)
            elif carrier == "DUE":
                date = get_delivery_date_aduiepyle(ADUIEPYLE_EMAIL, tracking)
            else:
                date = None  # Unknown carrier
        except Exception as e:
            print(f"\nError tracking {tracking}: {e}")
            date = "Error"

        row[col_indices["Delivered"]].value = date
        updated_rows += 1
        print_progress_bar(i, total_rows)

    print()  # Newline after progress bar
    wb.save(filename)
    print(f"Done. Updated '{filename}' with {updated_rows} new delivery dates.")

if __name__ == "__main__":
    for file_name in os.listdir():
        if file_name.endswith(".xlsx"):
            try:
                print(f"\nProcessing: {file_name}")
                process_tracking_sheet(file_name)
            except Exception as e:
                print(f"Error processing {file_name}: {e}")
