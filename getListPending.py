import requests
import datetime
import csv
import sys
from bs4 import BeautifulSoup

# Get the page number from command-line arguments
if len(sys.argv) < 2:
    print("Usage: python getListPending.py <page>")
    sys.exit(1)

try:
    page = int(sys.argv[1])
except ValueError:
    print("Error: Page parameter must be an integer.")
    sys.exit(1)

# Pobranie daty dzisiejszej
today_date = (datetime.datetime.today()).strftime('%Y-%m-%d')

# API endpoint
url = f"https://dane.biznes.gov.pl/api/ceidg/v3/firmy?status=OCZEKUJE_NA_ROZPOCZECIE_DZIALANOSCI&page={page}"

headers = {
    "Authorization": "Bearer eyJraWQiOiJjZWlkZyIsImFsZyI6IkhTNTEyIn0.eyJnaXZlbl9uYW1lIjoiUGlvdHIiLCJwZXNlbCI6Ijg0MDMyMDE4MTMyIiwiaWF0IjoxNzQ5ODI2OTE0LCJmYW1pbHlfbmFtZSI6IlN6bGFkZXJiYSIsImNsaWVudF9pZCI6IlVTRVItODQwMzIwMTgxMzItUElPVFItU1pMQURFUkJBIn0.VP_Yp14G36e2QKCnY1C3ZSDMFm5zVmEBfOxWUOvIB4OhRi3YnrEomjC1BMfJ0a5ENZMRTNYopgpjgerzsc__0A", #Sz2
    # "Authorization": "Bearer eyJraWQiOiJjZWlkZyIsImFsZyI6IkhTNTEyIn0.eyJnaXZlbl9uYW1lIjoiUElPVFIiLCJwZXNlbCI6Ijg0MDMyMDE4MTMyIiwiaWF0IjoxNjk3NTUwMDM4LCJmYW1pbHlfbmFtZSI6IlNaTEFERVJCQSIsImNsaWVudF9pZCI6IlVTRVItODQwMzIwMTgxMzItUElPVFItU1pMQURFUkJBIn0.Ycxc5906lA0FAIK0LycrIp-5_BSJQgN2MS-4fs5k2PoinAHnKh-LcDi525NxCE_C49EBSQgma-Ob1uIHvjGQ8Q", #Sz
}

try:
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # Raises an error for HTTP codes 4xx/5xx
    content_type = response.headers.get("Content-Type", "").lower()

    if "application/json" in content_type or "text/json" in content_type:
        responseType = "JSON"
    elif "application/html" in content_type or "text/html" in content_type:
        responseType = "HTML"
    else:
        responseType = "Other (Content-Type: {})".format(content_type)
    print(f"Response type is: {responseType}")

except requests.exceptions.RequestException as e:
    print(f"Failed to fetch data: {e}")
    exit(1)

if responseType == "JSON":

    try:
        data = response.json()  # Parse JSON response
    except requests.exceptions.RequestException as e:
        print(f"Error: Unable to parse JSON response.: {e}")
        sys.exit(1)

    firma = data.get("firma", [])
    count = data.get("count", 0)

    # Define the output CSV file
    output_file = f"data-pending/NEW_OCZEKUJACE_{today_date}.csv"

    # Open CSV file to write data
    with open(output_file, mode="w", newline="", encoding="utf-8") as file:
        csv_writer = csv.writer(file)

        # Check if 'firmy' key exists and is a list
        if "firmy" in data and isinstance(data["firmy"], list):
            # Extract headers dynamically from the first entry
            headers = [
                "id",
                "nazwa",
                "wlasciciel_imie",
                "wlasciciel_nazwisko",
                "wlasciciel_nip",
                "wlasciciel_regon",
                "dataRozpoczecia",
                "status"
            ]
            csv_writer.writerow(headers)  # Write header row

            # Write data rows
            for company in data["firmy"]:
                row = [
                    company.get("id", ""),
                    company.get("nazwa", ""),
                    company.get("wlasciciel", {}).get("imie", ""),
                    company.get("wlasciciel", {}).get("nazwisko", ""),
                    company.get("wlasciciel", {}).get("nip", ""),
                    company.get("wlasciciel", {}).get("regon", ""),
                    company.get("dataRozpoczecia", ""),
                    company.get("status", "")
                ]
                csv_writer.writerow(row)

            print(f"Data saved to {output_file}")

        else:
            print(f"No data found for page {page}.")

    print(f"Count value : {count}")


elif responseType == "HTML":
    try:
        soup = BeautifulSoup(response.text, "html.parser")  # Parse HTML response
        title = soup.title.string if soup.title else "No title found"
    except requests.exceptions.RequestException as e:
        print(f"Error: Unable to parse HTML response.: {e}")
        sys.exit(1)
    print(f"HTML title is: {title}")