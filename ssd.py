import requests
import pandas as pd
from requests.auth import HTTPBasicAuth
from getpass import getpass

# Prompt the user for Moysklad credentials
MOYSKLAD_LOGIN = input("Enter your Moysklad email (login): ")
MOYSKLAD_PASSWORD = getpass("Enter your Moysklad API key (password): ")

# Folder ID (replace with your folder ID)
FOLDER_ID = 'https://api.moysklad.ru/api/remap/1.2/entity/productfolder/9082b9a3-eeca-11eb-0a80-02db000c24ef'

# API endpoint for products
url = 'https://api.moysklad.ru/api/remap/1.2/entity/assortment'

# Initialize variables for pagination
limit = 1000  # Maximum number of products per request
offset = 0
all_products = []

while True:
    # Parameters to filter products by folder ID and handle pagination
    params = {
        'filter': f'productFolder={FOLDER_ID}',
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

        # Extract the 'rows' from the response (list of products)
        products = data.get('rows', [])
        all_products.extend(products)

        # Check if there are more products to fetch
        if len(products) < limit:
            break  # Exit the loop if all products have been fetched
        else:
            offset += limit  # Move to the next page
    else:
        print(f"Failed to fetch data. Status code: {response.status_code}")
        print("Response:", response.text)
        break

# Filter and format the required fields: code, name, and barcodes
filtered_data = []
for product in all_products:
    # Extract barcodes and format them to include only numbers
    barcodes = []
    for barcode in product.get('barcodes', []):
        if isinstance(barcode, dict):
            for key, value in barcode.items():
                if isinstance(value, str) and value.isdigit():
                    barcodes.append(value)
                elif isinstance(value, str):
                    # Extract numbers from strings like "ean13: 2000000079790"
                    numbers = ''.join(filter(str.isdigit, value))
                    if numbers:
                        barcodes.append(numbers)

    # Join barcodes into a single string (comma-separated)
    barcodes_str = ', '.join(barcodes)

    # Add the product data to the filtered list
    filtered_item = {
        'code': product.get('code', ''),
        'name': product.get('name', ''),
        'barcodes': barcodes_str
    }
    filtered_data.append(filtered_item)

# Convert the filtered data into a pandas DataFrame
df = pd.DataFrame(filtered_data)

# Export the DataFrame to an Excel file
output_file = 'products_in_folder.xlsx'
df.to_excel(output_file, index=False)

print(f"Data successfully exported to {output_file}")