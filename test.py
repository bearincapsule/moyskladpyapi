import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
from getpass import getpass

# Prompt the user for Moysklad credentials
MOYSKLAD_LOGIN = input("Enter your Moysklad email (login): ")
MOYSKLAD_PASSWORD = getpass("Enter your Moysklad API key (password): ")

# API endpoint for product folders
url = 'https://api.moysklad.ru/api/remap/1.2/entity/productfolder'

# Initialize variables for pagination
limit = 1000  # Maximum number of folders per request
offset = 0
all_folders = []

while True:
    # Add pagination parameters to the request
    params = {
        'limit': limit,
        'offset': offset
    }

    # Make the GET request with Basic Authentication
    response = requests.get(
        url,
        params=params,
        auth=HTTPBasicAuth(MOYSKLAD_LOGIN, MOYSKLAD_PASSWORD)
    )

    # Check if the request was successful
    if response.status_code == 200:
        # Parse the JSON response
        data = response.json()

        # Extract the 'rows' from the response (list of folders)
        folders = data.get('rows', [])
        all_folders.extend(folders)

        # Check if there are more folders to fetch
        if len(folders) < limit:
            break  # Exit the loop if all folders have been fetched
        else:
            offset += limit  # Move to the next page
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}")
        print("Response:", response.text)
        break

# Filter only the required fields: id, name, and pathName
filtered_data = []
for folder in all_folders:
    filtered_item = {
        'id': folder.get('id'),
        'name': folder.get('name'),
        'pathName': folder.get('pathName', '')  # Full path of the folder
    }
    filtered_data.append(filtered_item)

# Convert the filtered data into a pandas DataFrame
df = pd.DataFrame(filtered_data)

# Export the DataFrame to an Excel file
output_file = 'folders_data.xlsx'
df.to_excel(output_file, index=False)

print(f"Data successfully exported to {output_file}")
