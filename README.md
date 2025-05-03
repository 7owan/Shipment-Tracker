Shipment Tracker is a Python script that reads tracking numbers from Excel files, fetches delivery status from FedEx and A. Duie Pyle APIs, and updates the Excel sheet with the delivery date.

Features:

1. Supports FedEx and A. Duie Pyle tracking.
2. Reads .xlsx files with the following columns: Carrier, Pro #, Delivered DATE.
3. Skips rows that already have a delivery date filled in.
4. Displays a progress bar in the console.

Instructions:

Place hammond_shipment_tracker.exe in a directory along with a .env file and all the spreadsheets you wish to update.

Inside of your .env file include the following:

FEDEX_CLIENT_ID=your_client_id
FEDEX_CLIENT_SECRET=your_secret
ADUIEPYLE_EMAIL=your_email@example.com

Fill in the placeholders with your own credentials.

Your spreadsheets MUST follow this format:

Must have atleast 3 distinct columns, named: Carrier, Pro #, Delivered DATE.
The spreadsheet must have the following 3 headers on the first line, they must be named exactly the same (non-case sensitive), and the order they are in does not matter.

An example of how the project directory should look:

shipment-tracker/
│
├── .env                          # Environment Variables (Not committed to git)
├── hammond_shipment_tracker.exe  # Main Program
├── README.txt                    # Instructions for using the project (your text version)
├── shipping.xlsx                 # Example Excel file to be processed
