import requests
from requests.auth import HTTPBasicAuth

# Replace with your MoySklad credentials
USERNAME = "warehousetseh2@gmail.com"
PASSWORD = "PurestSklad6632!"
BASE_URL = "https://api.moysklad.ru/api/remap/1.2"

# Function to get serial number
def get_serial_number(product_code):
    headers = {
        "Content-Type": "application/json"
    }

    # Authenticate with Basic Auth
    auth = HTTPBasicAuth(USERNAME, PASSWORD)

    # Search for the product by its code
    product_url = f"{BASE_URL}/entity/product?filter=code={product_code}"
    response = requests.get(product_url, headers=headers, auth=auth)

    if response.status_code == 200:
        products = response.json().get("rows", [])
        if not products:
            print("Product not found.")
            return None
        
        product_id = products[0]["id"]  # Get the product ID

        # Get serial numbers for the product
        serial_url = f"{BASE_URL}/entity/serial?filter=assortment.id={product_id}"
        serial_response = requests.get(serial_url, headers=headers, auth=auth)

        if serial_response.status_code == 200:
            serials = serial_response.json().get("rows", [])
            return [s["name"] for s in serials]  # Extract serial numbers

    print(f"Error: {response.status_code}, {response.text}")
    return None


# Example usage
product_code = "123456"  # Replace with your product code
serial_numbers = get_serial_number(product_code)

if serial_numbers:
    print("Serial Numbers:", serial_numbers)
else:
    print("No serial numbers found.")
