# AutoTrader Price Puller

## Description

This script automates the process of fetching average used car prices for specific models and years from AutoTrader's internal API. It retrieves the data daily and logs it into an Excel spreadsheet (`Price_Puller.xlsx`), with each car model having its own sheet.

## Features

* Fetches average price data per model year for a predefined list of cars.
* Uses specific AutoTrader API URLs for targeted searches (model, location, etc.).
* Writes data to an Excel file (`Price_Puller.xlsx`), creating the file or sheets if they don't exist.
* Organizes data with one sheet per car model.
* Adds a new row for the current day's prices, using a fixed anchor date (`22Apr2025`) to determine the correct row.
* Dynamically creates and manages headers (Date, Year columns) in each sheet.
* Applies currency formatting (`$#,##0.00`) to price cells.
* Implements retry logic for robustness against temporary network issues.
* Logs script activity and errors to `price_puller.log`.

## How it Works

1.  **Initialization:** Sets up logging to `price_puller.log`.
2.  **URL Iteration:** The `main()` function iterates through a dictionary of car models (sheet names) and their corresponding AutoTrader API URLs.
3.  **Data Fetching (`get_avg_price`):**
    * For each URL, it sends an HTTP GET request with specific headers, using a session with automatic retries.
    * It parses the JSON response from the API.
    * It extracts the average price (`avgPrice`) for each available model year (`value`).
4.  **Row Calculation:**
    * It determines the target row number in the Excel sheet based on the number of days elapsed since a hardcoded anchor date (`22Apr2025`). Row 2 corresponds to the anchor date. **This assumes the script runs daily without gaps.**
5.  **Excel Writing:**
    * Loads the `Price_Puller.xlsx` workbook or creates a new one if it doesn't exist.
    * Accesses the sheet corresponding to the car model name or creates it.
    * Checks if headers exist in the first row. If not, it writes 'Date' and the sorted model years found in the current data pull as headers (bolded).
    * Writes the current date (formatted as `DDMMMYYYY`) in the 'Date' column for the calculated row.
    * Writes the fetched average prices into the appropriate year columns for the calculated row, applying currency formatting.
    * Saves the workbook.
6.  **Error Handling:** Any exceptions during the request or processing for a specific URL are caught, logged to `price_puller.log`, and the script continues to the next URL.

## Requirements

* Python 3.x
* Required Python libraries:
    * `requests`
    * `openpyxl`

*(Note: `pandas` is imported in the script but currently not used.)*

## Setup

1.  Ensure you have Python 3 installed.
2.  Install the required libraries using pip:
    ```bash
    pip install requests openpyxl
    ```
3.  Place the script (e.g., `price_puller.py`) in your desired directory.

## Usage

Run the script from your terminal:

```bash
python price_puller.py