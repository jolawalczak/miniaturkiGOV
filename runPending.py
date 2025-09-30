import functools
import subprocess
import time
import re
import os
from datetime import datetime, timedelta
import shutil
import glob
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

# Record the start time
start_time = time.time()

today_date = (datetime.today()).strftime('%Y-%m-%d')
yesterday_date = (datetime.today() - timedelta(days=1)).strftime('%Y-%m-%d')

# Delete unecessary file
xlsx_files_to_delete = glob.glob("data-pending/OCZEKUJACE_*.xlsx")
# Delete unecessary files
if xlsx_files_to_delete:  # Check if any files were found
    for file_path in xlsx_files_to_delete:
        if os.path.isfile(file_path):  # Ensure it's a file
            try:
                os.remove(file_path)
                print(f"Deleted: {file_path}")
            except Exception as e:
                print(f"Error deleting XLSX file {file_path}: {e}")
else:
    print("No matching XLSX files found.")

# Initialize variables
page = 0
base_count = None  # Used to control the loop termination (each page is assumed to return 25 items)
data_found = False  # Will be set to True if getList.py returns data at least once
csv_file = f"data-pending/NEW_OCZEKUJACE_{today_date}.csv"  # File to be deleted at the end
xlsx_file = f"data-pending/NEW_OCZEKUJACE_{today_date}.xlsx"
final_file = f"data-pending/OCZEKUJACE_{today_date}.xlsx"
extracted_title = False

while True:
    # If base_count is set and no more data is available, exit early.
    if base_count is not None and base_count < 1:
        print("No more data available. Stopping process.")
        break

    time.sleep(4)  # Delay before each execution
    current_time = datetime.now().strftime("%H:%M:%S")

    print(f"{current_time} - Running getListPending.py with page={page}...")
    result = subprocess.run(["python3", "getListPending.py", str(page)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)

    stdout = result.stdout
    stderr = result.stderr

    print("Full Output from getListPending.py:")

    if result.returncode != 0:
        print(f"Error executing getListPending.py on page {page}: {stderr}")
        print(stdout)
        break

    # Extract count value only the first time
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

    # Check if the expected data was returned by getList.py
    if "Data saved to" in stdout:
        data_found = True  # Mark that data was found in at least one iteration

        print(f"Running convert.py after getListPending.py execution (page {page})...")
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

    # Decrement base_count and move to the next page
    base_count -= 25
    page += 1

if data_found:
    print(f"Running comparePending.py")
    result = subprocess.run(["python3", "comparePending.py"],stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)

    stdout = result.stdout
    stderr = result.stderr

    if result.returncode == 1:
        print("Files are compared successfully")
        print(stdout)

        xlsx_files_to_delete = glob.glob("data-pending/OLD_OCZEKUJACE_*.xlsx")

        # Delete unecessary files
        if xlsx_files_to_delete:  # Check if any files were found
            for file_path in xlsx_files_to_delete:
                if os.path.isfile(file_path):  # Ensure it's a file
                    try:
                        os.remove(file_path)
                        print(f"Deleted: {file_path}")
                    except Exception as e:
                        print(f"Error deleting XLSX file {file_path}: {e}")
        else:
            print("No matching XLSX files found.")

        
        # Change name NEW files to OLD file
        shutil.move(f"data-pending/NEW_OCZEKUJACE_{today_date}.xlsx", f"data-pending/OLD_OCZEKUJACE_{today_date}.xlsx")
        print(f"Changed file name: data-pending/NEW_OCZEKUJACE_{today_date}.xlsx to data-pending/OLD_OCZEKUJACE_{today_date}.xlsx")

    else:
        print("Error executing comparePending.py")
        print(stdout)

print = functools.partial(print, flush=True)

if data_found:
    print("Running collectFirmaData.py after finishing comparePending.py execution in real-time:")
    with subprocess.Popen(["python3", "collectFirmaData.py", str(final_file), str(start_time)],
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

# After processing all pages, delete the CSV file if it exists
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
print("Running sendemail.py to send the report...")
command = ["python3", "sendmail.py", str(final_file)]
if extracted_title:  # If extracted title is provided from HTML response, add it as an argument
    command.append(str(extracted_title))
sendmail_result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
print("Full Output from sendmail.py:")
print(sendmail_result.stdout)

# Calculate and display the total execution time
end_time = time.time()
total_time = end_time - start_time
print(f"Total execution time: {total_time:.2f} seconds")