import asyncio
import aiohttp
import pandas as pd
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, Menu
from aiohttp import ClientSession
from requests.auth import HTTPBasicAuth
from datetime import datetime

# Base API URL
BASE_URL = "https://api.moysklad.ru/api/remap/1.2/entity"

# API Endpoints
ENDPOINTS = {
    "product": f"{BASE_URL}/product",
    "folders": f"{BASE_URL}/productfolder",
    "assortment": f"{BASE_URL}/assortment",
    "variants": f"{BASE_URL}/variant"
}

EMAILS_FILE = "used_emails.json"

# Global variables
auth = None  # Store authentication credentials
folder_metadata = {}  # Dictionary for folder metadata

def build_folder_tree(folders):
    folder_dict = {None: []}  # Root level folders
    for folder in folders:
        parent_href = folder.get('productFolder', {}).get('meta', {}).get('href')
        if parent_href not in folder_dict:
            folder_dict[parent_href] = []
        folder_dict[parent_href].append(folder)
    return folder_dict

async def fetch_data(session, endpoint, params=None):
    url = ENDPOINTS[endpoint]
    for _ in range(3):  # Retry up to 3 times
        try:
            async with session.get(url, params=params, auth=auth, timeout=10) as response:
                if response.status == 200:
                    return await response.json()
                elif response.status == 429:
                    await asyncio.sleep(1)  # Rate limit handling
        except aiohttp.ClientError:
            await asyncio.sleep(1)
    return None  # Return None if all retries fail

async def fetch_folders():
    async with ClientSession() as session:
        data = await fetch_data(session, "folders")
        return data.get('rows', []) if data else []

async def fetch_products(folder_hrefs):
    async with ClientSession() as session:
        tasks = [fetch_data(session, "assortment", {'filter': f'productFolder={href}', 'limit': 1000}) for href in folder_hrefs]
        responses = await asyncio.gather(*tasks)
        products = [p for res in responses if res for p in res.get('rows', []) if p.get('stock', 0) > 0]
        return products

async def fetch_variants(product_id):
    async with ClientSession() as session:
        data = await fetch_data(session, "variants", {"filter": f"product={product_id}"})
        if not data:
            return "No Category"
        return ", ".join([char.get("value", "") for v in data.get('rows', []) for char in v.get("characteristics", []) if char.get("value")]) or "No Category"

async def process_data(folder_hrefs, folder_name):
    products = await fetch_products(folder_hrefs)
    
    tasks = [fetch_variants(p.get('id', '')) for p in products]
    categories = await asyncio.gather(*tasks)
    
    data = []
    for i, p in enumerate(products):
        data.append([
            p.get('pathName', 'Unknown'),
            p.get('name', 'No Name'),
            categories[i],
            p.get('code', 'No Code'),
            p.get('salePrices', [{}])[0].get('value', 0) / 100,
            p.get('stock', 0)
        ])
    
    df = pd.DataFrame(data, columns=["Category", "Product Name", "Categories", "Code", "Price", "Stock"])
    filename = f"{folder_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    df.to_excel(filename, index=False)
    messagebox.showinfo("Success", f"Data exported to {filename}")

def open_main_menu(root, auth_param):
    global auth
    auth = auth_param
    
    async def init_folders():
        all_folders = await fetch_folders()
        folder_dict = build_folder_tree(all_folders)
        
        for widget in root.winfo_children():
            widget.destroy()
        
        tree = ttk.Treeview(root)
        tree.heading("#0", text="Folders", anchor=tk.W)
        tree.pack(expand=True, fill=tk.BOTH)
        
        def populate_tree(parent, folders):
            for folder in folders:
                folder_id = tree.insert(parent, "end", text=folder['name'], open=False)
                folder_metadata[folder_id] = folder['meta']['href']
                populate_tree(folder_id, folder_dict.get(folder['meta']['href'], []))
        
        populate_tree("", folder_dict.get(None, []))
        
        def on_fetch():
            selected_item = tree.selection()
            if not selected_item:
                messagebox.showerror("Error", "Please select a folder.")
                return
            selected_folder_href = folder_metadata[selected_item[0]]
            asyncio.run(process_data([selected_folder_href], tree.item(selected_item[0], 'text')))
        
        fetch_button = tk.Button(root, text="Fetch Products", command=on_fetch)
        fetch_button.pack(pady=10)
    
    asyncio.run(init_folders())

def create_gui():
    root = tk.Tk()
    root.title("MoySklad Product Fetcher")
    root.geometry("600x400")
    tk.Label(root, text="Enter your MoySklad email:").pack()
    email_entry = tk.Entry(root)
    email_entry.pack()
    tk.Label(root, text="Enter your MoySklad password:").pack()
    password_entry = tk.Entry(root, show="*")
    password_entry.pack()
    
    def authenticate():
        global auth
        auth = HTTPBasicAuth(email_entry.get(), password_entry.get())
        open_main_menu(root, auth)
    
    login_button = tk.Button(root, text="Login", command=authenticate)
    login_button.pack()
    root.mainloop()

if __name__ == "__main__":
    create_gui()
