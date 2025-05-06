import os
import sys
import requests
import time
from xml.etree import ElementTree
from openpyxl import load_workbook
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

load_dotenv()

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
        "trackingInfo": [{"trackingNumberInfo": {"trackingNumber": tracking_number}}],
        "includeDetailedScans": False
    }
    try:
        response = requests.post(url, headers=headers, json=payload)
        response.raise_for_status()
        result = response.json()
        shipment = result["output"]["completeTrackResults"][0]["trackResults"][0]
        for date in shipment.get("dateAndTimes", []):
            if date["type"] in ["ACTUAL_DELIVERY", "ESTIMATED_DELIVERY"]:
                return date["dateTime"][:10]
    except Exception as e:
        return f"FedEx error: {e}"
    return None

def get_delivery_date_aduiepyle(user_email, tracking_number):
    url = "https://api.aduiepyle.com/2/shipment/status"
    params = {
        'user': user_email,
        'type': 0,
        'value': tracking_number
    }
    try:
        response = requests.get(url, params=params)
        if response.status_code == 200:
            root = ElementTree.fromstring(response.text)
            for status in root.findall(".//statusDetail"):
                if status.find("description").text == "DELIVERED":
                    return status.find("start").text[:10]
    except Exception as e:
        return f"ADP error: {e}"
    return None

def track_package(row_idx, row, col_indices, fedex_token, ad_email):
    carrier = clean_cell(row[col_indices["Carrier"]]).upper()
    tracking = clean_cell(row[col_indices["Tracking"]])
    delivered_value = clean_cell(row[col_indices["Delivered"]])

    if delivered_value or not carrier or not tracking:
        return row_idx, None  # Skip if already filled or missing info

    try:
        if carrier in ["FEP", "FEE", "FEU", "FEA", "FED", "FEC"]:
            date = get_delivery_date_fedex(fedex_token, tracking)
        elif carrier == "DUE":
            date = get_delivery_date_aduiepyle(ad_email, tracking)
        else:
            return row_idx, None  # Skip unknown carriers
    except Exception as e:
        date = f"Tracking error: {e}"

    return row_idx, date

def process_tracking_sheet(filename, fedex_token, ad_email, sheet_name="Sheet1"):
    wb = load_workbook(filename)
    ws = wb.active

    header_row = [cell.value.strip() if cell.value else "" for cell in ws[1]]
    col_indices = {"Carrier": None, "Tracking": None, "Delivered": None}

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

    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [
            executor.submit(track_package, i, row, col_indices, fedex_token, ad_email)
            for i, row in enumerate(rows)
        ]

        progress_lock = Lock()
        completed = 0

        for f in as_completed(futures):
            idx, result = f.result()
            if result:
                rows[idx][col_indices["Delivered"]].value = result
                updated_rows += 1
            with progress_lock:
                completed += 1
                print_progress_bar(completed, total_rows)

    print()
    wb.save(filename)
    print(f"Done. Updated '{filename}' with {updated_rows} new delivery dates.")

def clean_cell(cell):
    return str(cell.value or "").strip()

if __name__ == "__main__":
    start_time = time.time()

    FEDEX_CLIENT_ID = os.getenv("FEDEX_CLIENT_ID")
    FEDEX_CLIENT_SECRET = os.getenv("FEDEX_CLIENT_SECRET")
    ADUIEPYLE_EMAIL = os.getenv("ADUIEPYLE_EMAIL")

    try:
        FEDEX_TOKEN = get_auth_token(FEDEX_CLIENT_ID, FEDEX_CLIENT_SECRET)
    except Exception as e:
        print("FedEx auth error:", e)
        FEDEX_TOKEN = None

    for file_name in os.listdir():
        if file_name.endswith(".xlsx"):
            try:
                print(f"Processing: {file_name}")
                process_tracking_sheet(file_name, fedex_token=FEDEX_TOKEN, ad_email=ADUIEPYLE_EMAIL)
            except Exception as e:
                print(f"Error processing {file_name}: {e}")

    elapsed_time = time.time() - start_time
    minutes, seconds = divmod(elapsed_time, 60)
    print(f"Total time taken: {int(minutes)}m {int(seconds)}s")
