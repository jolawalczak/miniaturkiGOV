import requests
import json
import sys

# Get the page number from command-line arguments
if len(sys.argv) < 2:
    print("Usage: python getFirmaByID.py <id>")
    sys.exit(1)

try:
    id = sys.argv[1]
except ValueError:
    print("Error: ID parameter must be a string type.")
    sys.exit(1)

# API endpoint
url = f"https://dane.biznes.gov.pl/api/ceidg/v3/firma/{id}"

# Authorization token (ensure this is securely stored in production)
headers = {
    "Authorization": "Bearer eyJraWQiOiJjZWlkZyIsImFsZyI6IkhTNTEyIn0.eyJnaXZlbl9uYW1lIjoiUElPVFIiLCJwZXNlbCI6Ijg0MDMyMDE4MTMyIiwiaWF0IjoxNjk3NTUwMDM4LCJmYW1pbHlfbmFtZSI6IlNaTEFERVJCQSIsImNsaWVudF9pZCI6IlVTRVItODQwMzIwMTgxMzItUElPVFItU1pMQURFUkJBIn0.Ycxc5906lA0FAIK0LycrIp-5_BSJQgN2MS-4fs5k2PoinAHnKh-LcDi525NxCE_C49EBSQgma-Ob1uIHvjGQ8Q", #Sz
}

# Make the request
try:
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # Raises an error for HTTP codes 4xx/5xx
    data = response.json()  # Parse JSON response
except requests.exceptions.RequestException as e:
    print(f"Failed to fetch data: {e}")
    sys.exit(1)
except json.JSONDecodeError:
    print("Error: Unable to parse JSON response.")
    sys.exit(1)

# Print the entire JSON data
print(json.dumps(data, indent=4, ensure_ascii=False))