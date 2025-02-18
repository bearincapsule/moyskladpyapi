import requests
from requests.auth import HTTPBasicAuth

auth = HTTPBasicAuth('warehousetseh2@gmail.com', 'PurestSklad6632!')
url = "https://api.moysklad.ru/api/remap/1.2/entity/assortment?filter=stockMode=positiveOnly"
response = requests.get(url, auth=auth)
products = response.json().get('rows', [])

# Iterate through the products
for product in products[:5]:  # Show only first 5 for demonstration
    name = product.get('name')
    code = product.get('code')
    path = product.get('pathName')
    stock = product.get('stock', 0)
    
    # Extract folder from product meta href
    product_response = requests.get(product.get('product', {}).get('meta', {}).get('href', {}), auth=auth)
    productid = product_response.json()
    foldername = productid.get('pathName')
    
    # Default category value
    category_value = "Unknown"
    
    # Access characteristics
    characteristics = product.get('characteristics', [])
    for char in characteristics:
        if char.get('name') == "Категория":
            category_value = char.get('value')
            break  # Stop searching once we find the category
    
    # Print the product information
    print(f"Name: {name}, Code: {code}, Path: {path}, Stock: {stock}, Category: {category_value}\nFolder: {foldername}")

