import functools
import subprocess
import time
import re
import os
import glob
from datetime import datetime, timedelta
import sys

# # DATA STARTU (ustaw ręcznie datę rozpoczęcia działania programu)
# START_DATE_STR = "2025-04-04"  # format: RRRR-MM-DD
# START_DATE = datetime.strptime(START_DATE_STR, "%Y-%m-%d")

# # Data końcowa: 7 dni od startu
# END_DATE = START_DATE + timedelta(days=7)

# # Dzisiejsza data
# TODAY = datetime.today()

# # Sprawdzenie, czy dzisiaj mieści się w zakresie
# if not (START_DATE <= TODAY < END_DATE):
#     print(f"Program może być uruchamiany tylko od {START_DATE_STR} do {(END_DATE - timedelta(days=1)).strftime('%Y-%m-%d')}.")
#     sys.exit(0)

xlsx_files = glob.glob("data-active/*.xlsx")

for xlsx_file in xlsx_files:
    os.remove(xlsx_file)
    print(f"Deleted: {xlsx_file}")

start_time = time.time()

# Initialize variables
page = 0
base_count = None  # Used to control the loop termination (each page is assumed to return 25 items)
data_found = False  # Will be set to True if getListActive.py returns data at least once
yesterday_date = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')
csv_file = f"data-active/AKTYWNE_{yesterday_date}.csv"  # File to be deleted at the end
xlsx_file = f"data-active/AKTYWNE_{yesterday_date}.xlsx"  # File to be deleted at the end
extracted_title = False

if os.path.exists(xlsx_file):
    print(f"Detected existing XLSX file")
    try:
        os.remove(xlsx_file)
        print(f"XLSX file {xlsx_file} deleted successfully.")
    except Exception as e:
        print(f"Error deleting XLSX file {xlsx_file}: {e}")

while True:
    if base_count is not None and base_count < 1:
        print("No more data available. Stopping process.")
        break

    time.sleep(4)
    current_time = datetime.now().strftime("%H:%M:%S")

    print(f"{current_time} - Running getListActive.py with page={page}...")
    result = subprocess.run(["python3", "getListActive.py", str(page)],
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            universal_newlines=True)

    stdout = result.stdout
    stderr = result.stderr

    print("Full Output from getListActive.py:")

    if result.returncode != 0:
        print(f"Error executing getListActive.py on page {page}: {stderr}")
        print(stdout)
        break

    if base_count is None:
        match = re.search(r"Count value.*?:\s*(\d+)", stdout)
        if match:
            base_count = int(match.group(1))
            print(f"Initial count extracted: {base_count}")
            if base_count > 0:
                data_found = True
        else:
            print("Error: No valid count extracted. Exiting.")
            data_found = False

    if "HTML" in stdout:
        data_found = False
        cleaned_response = "\n".join(line.strip() for line in stdout.splitlines() if line.strip())
        matchHTML = re.search(r"HTML title is:\s*(.+)", cleaned_response)  # Capture everything after "HTML title is:"
        if matchHTML:
            extracted_title = matchHTML.group(1)
            print(cleaned_response)
        else:
            extracted_title = "Nie pobrano listy firm, skontaktuj się z administratorem."
            print(extracted_title)
            print(cleaned_response)
            break

    if "Data saved to" in stdout:
        data_found = True

        print(f"Running convert.py after getListActive.py execution (page {page})...")
        convert_result = subprocess.run(["python3", "convert.py", str(csv_file)],
                                        stdout=subprocess.PIPE,
                                        stderr=subprocess.PIPE,
                                        universal_newlines=True)
        print("Full Output from convert.py:")
        print(convert_result.stdout)
        if convert_result.returncode != 0:
            print(f"Error executing convert.py: {convert_result.stderr}")
            data_found = False
            break
    else:
        print(f"No expected data returned on page {page}. Stopping process.")
        break

    base_count -= 25
    page += 1

print = functools.partial(print, flush=True)

if data_found:
    print("Running collectFirmaData.py after finishing all getListActive.py and convert.py executions in real-time:")
    with subprocess.Popen(["python3", "collectFirmaData.py", str(xlsx_file), str(start_time)],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT,
                        universal_newlines=True,
                        bufsize=1) as proc:
        for line in proc.stdout:
            print(line, end='')  # Already includes newline
        proc.wait()
        if proc.returncode != 0:
            print(f"collectFirmaData.py ended with error code: {proc.returncode}")
else:
    print("No data found from comparePending.py. Skipping collectFirmaData.py.")

if os.path.exists(csv_file):
    try:
        os.remove(csv_file)
        print(f"CSV file {csv_file} deleted successfully.")
    except Exception as e:
        print(f"Error deleting CSV file {csv_file}: {e}")
else:
    print(f"No CSV file found for deletion: {csv_file}")

print("Finished processing all pages.")

# Run sendmail.py as the final step
file_to_send = f"data-active/AKTYWNE_{yesterday_date}.xlsx"
print("Running sendemail.py to send the report...")
command = ["python3", "sendmail.py", str(file_to_send)]
if extracted_title:  # If extracted title is provided from HTML response, add it as an argument
    command.append(str(extracted_title))
sendmail_result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
print("Full Output from sendmail.py:")
print(sendmail_result.stdout)

end_time = time.time()
total_time = end_time - start_time
print(f"Total execution time: {total_time:.2f} seconds")