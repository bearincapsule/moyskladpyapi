import aiohttp
import asyncio
import pandas as pd
from datetime import datetime
from requests.auth import HTTPBasicAuth

auth = HTTPBasicAuth('warehousetseh2@gmail.com', 'PurestSklad6632!')
url = "https://api.moysklad.ru/api/remap/1.2/report/stock/all"
included_price_types = ["Цена розница", "Цена маркетплейс", "Цена мелкий опт", "Цена средний опт"]

async def fetch(session, url):
    async with session.get(url, auth=aiohttp.BasicAuth(auth.username, auth.password)) as response:
        return await response.json()

async def fetch_product_details(session, product):
    meta_url = product.get('meta', {}).get('href', {})
    metaid = await fetch(session, meta_url)

    # Default category value
    category_value = "Unknown"
    
    # Access characteristics
    characteristics = metaid.get('characteristics', [])
    for char in characteristics:
        if char.get('name') == "Категория":
            category_value = char.get('value')
            break  # Stop searching once we find the category

    # Get all sales prices and filter based on included price types
    sale_prices = metaid.get('salePrices', [])
    prices = {
        price.get('priceType', {}).get('name', 'Unknown'): price.get('value', 0) / 100  # Convert to proper format if needed
        for price in sale_prices
        if price.get('priceType', {}).get('name', 'Unknown') in included_price_types
    }

    return {
        'name': product.get('name'),
        'code': product.get('code'),
        'path': product.get('folder', {}).get('pathName') + '/' + product.get('folder', {}).get('name'),
        'stock': product.get('stock', 0),
        'days': product.get('stockDays', 0),
        'category': category_value,
        'prices': prices
    }

async def main():
    async with aiohttp.ClientSession() as session:
        response = await fetch(session, url)
        products = response.get('rows', [])

        tasks = [fetch_product_details(session, product) for product in products[:5]]
        results = await asyncio.gather(*tasks)

        # Prepare data for export
        data = []
        for result in results:
            base_data = {
                'Путь': result['path'],
                'Наименование': result['name'],
                'Категория': result['category'],
                'Код товара': result['code'],
                'Порядковый номер': None,
                'Включено в план размещения': ['Да','Нет'],
                'Фото на серевере': ['Да','Нет'],
                'Дней на складе': result['days'],
                'Остаток': result['stock']
            }
            for price_name, price_value in result['prices'].items():
                base_data[price_name] = price_value
            data.append(base_data)

        # Create a DataFrame and export to Excel
        df = pd.DataFrame(data)
        pd.set_option('display.max_colwidth', None)  # Show full content of each cell
        pd.set_option('display.width', 50)         # Increase display width
        # timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        # filename = f"products_{timestamp}.xlsx"
        # df.to_excel(filename, index=False)
        print(df)

# Run the main function
asyncio.run(main())
