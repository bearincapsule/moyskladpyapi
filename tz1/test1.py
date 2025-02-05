import requests
import pandas as pd
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, Menu
from requests.auth import HTTPBasicAuth
from datetime import datetime

# MoySklad API credentials
API_URL = "https://api.moysklad.ru/api/remap/1.2/entity/product"
FOLDERS_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/productfolder'
ASSORTMENT_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/assortment'
VARIANTS_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/variant'
EMAILS_FILE = "used_emails.json"

auth = None  # Global variable to store authentication credentials
folder_metadata = {}  # Dictionary to store folder metadata

def fetch_all_folders(auth):
    response = requests.get(FOLDERS_URL, auth=auth)
    if response.status_code != 200:
        messagebox.showerror("Error", "Error fetching folders.")
        return []
    return response.json().get('rows', [])

def build_folder_tree(folders):
    folder_dict = {None: []}  # Root level folders
    for folder in folders:
        parent_href = folder.get('productFolder', {}).get('meta', {}).get('href')
        if parent_href not in folder_dict:
            folder_dict[parent_href] = []
        folder_dict[parent_href].append(folder)
    return folder_dict

def populate_tree(tree, parent, folder_dict, parent_href):
    if parent_href in folder_dict:
        for folder in folder_dict[parent_href]:
            folder_id = tree.insert(parent, "end", text=folder['name'], open=False)
            folder_metadata[folder_id] = folder['meta']['href']  # Store metadata correctly
            populate_tree(tree, folder_id, folder_dict, folder['meta']['href'])

def fetch_products(folder_hrefs, auth):
    products = []
    for folder_href in folder_hrefs:
        params = {'filter': f'productFolder={folder_href}', 'limit': 1000}
        response = requests.get(ASSORTMENT_URL, auth=auth, params=params)
        if response.status_code != 200:
            messagebox.showerror("Error", "Error fetching products.")
            continue
        products.extend([p for p in response.json().get('rows', []) if p.get('stock', 0) > 0])  # Filter only in-stock products
    return products

def fetch_variants(product_id, auth):
    response = requests.get(f"{VARIANTS_URL}?filter=product={product_id}", auth=auth)
    if response.status_code != 200:
        return "No Category"
    variants = response.json().get('rows', [])
    characteristics = []
    for v in variants:
        char_values = [char.get("value", "") for char in v.get("characteristics", []) if char.get("value")]
        characteristics.extend(char_values)
    return ", ".join(characteristics) if characteristics else "No Category"

def export_to_excel(products, folder_name):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{folder_name}_{timestamp}.xlsx"
    
    # Organize products by category
    category_data = {}
    for p in products:
        category = p.get('pathName', 'Unknown')
        product_name = p.get('name', 'No Name')
        product_code = p.get('code', 'No Code')
        product_price = p.get('salePrices', [{}])[0].get('value', 0) / 100
        product_stock = p.get('stock', 0)
        product_categories = fetch_variants(p.get('id', ''), auth) if p.get('id') else "No Category"
        
        if category not in category_data:
            category_data[category] = []
        category_data[category].append([product_name, product_categories, product_code, product_price, product_stock])
    
    df_list = []
    for category, items in category_data.items():
        df = pd.DataFrame(items, columns=["Product Name", "Categories", "Code", "Price", "Stock"])
        df.insert(0, "Category", category)
        df_list.append(df)
    
    final_df = pd.concat(df_list, ignore_index=True)
    final_df.to_excel(filename, index=False)
    messagebox.showinfo("Success", f"Data exported to {filename}")

def open_main_menu(root, auth):
    for widget in root.winfo_children():
        widget.destroy()
    tree = ttk.Treeview(root)
    tree.heading("#0", text="Folders", anchor=tk.W)
    tree.pack(expand=True, fill=tk.BOTH)
    
    all_folders = fetch_all_folders(auth)
    folder_dict = build_folder_tree(all_folders)
    populate_tree(tree, "", folder_dict, None)
    
    def on_fetch():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a folder.")
            return
        selected_folder_href = folder_metadata[selected_item[0]]  # Retrieve folder href from metadata
        products = fetch_products([selected_folder_href], auth)
        if products:
            export_to_excel(products, tree.item(selected_item[0], 'text'))
        else:
            messagebox.showinfo("Info", "No products found.")
    
    fetch_button = tk.Button(root, text="Fetch Products", command=on_fetch)
    fetch_button.pack(pady=10)

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
