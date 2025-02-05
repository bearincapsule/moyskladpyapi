import requests
import pandas as pd
import json
import tkinter as tk
from tkinter import ttk, messagebox
from requests.auth import HTTPBasicAuth

# Load credentials from a file
def load_credentials(filename="credentials.json"):
    try:
        with open(filename, "r") as file:
            credentials = json.load(file)
            return credentials.get("MOYSKLAD_LOGIN"), credentials.get("MOYSKLAD_PASSWORD")
    except FileNotFoundError:
        messagebox.showerror("Error", "Credentials file not found.")
        return None, None

MOYSKLAD_LOGIN, MOYSKLAD_PASSWORD = load_credentials()
if not MOYSKLAD_LOGIN or not MOYSKLAD_PASSWORD:
    messagebox.showerror("Error", "Missing credentials. Please check credentials.json.")
    exit()

# API Endpoints
FOLDERS_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/productfolder'
ASSORTMENT_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/assortment'

# Function to fetch all folders
def fetch_all_folders():
    response = requests.get(FOLDERS_URL, auth=HTTPBasicAuth(MOYSKLAD_LOGIN, MOYSKLAD_PASSWORD))
    if response.status_code != 200:
        messagebox.showerror("Error", "Error fetching folders.")
        return []
    return response.json().get('rows', [])

# Function to filter parent folders manually
def get_parent_folders(folders):
    return [folder for folder in folders if 'productFolder' not in folder]

# Function to get subfolders of a given parent folder
def get_subfolders(folders, parent_folder_id):
    return [folder for folder in folders if folder.get('productFolder', {}).get('meta', {}).get('href') == parent_folder_id]

# Recursive function to fetch all subfolders
def get_all_subfolders(folders, parent_folder_id):
    subfolders = get_subfolders(folders, parent_folder_id)
    all_subfolders = subfolders[:]
    for subfolder in subfolders:
        all_subfolders.extend(get_all_subfolders(folders, subfolder['meta']['href']))
    return all_subfolders

# Function to fetch products from selected folders
def fetch_products(folder_hrefs):
    products = []
    limit = 1000  # API limit
    
    for folder_href in folder_hrefs:
        offset = 0
        while True:
            params = {
                'filter': f'productFolder={folder_href}',
                'limit': limit,
                'offset': offset
            }
            response = requests.get(ASSORTMENT_URL, auth=HTTPBasicAuth(MOYSKLAD_LOGIN, MOYSKLAD_PASSWORD), params=params)
            if response.status_code != 200:
                messagebox.showerror("Error", "Error fetching products.")
                break
            data = response.json().get('rows', [])
            products.extend(data)
            if len(data) < limit:
                break
            offset += limit
    return products

# Function to export product data to Excel
def export_to_excel(products, filename="products.xlsx"):
    df = pd.DataFrame([{ 'Code': p.get('code', ''), 'Name': p.get('name', '') } for p in products])
    df.to_excel(filename, index=False)
    messagebox.showinfo("Success", f"Data exported to {filename}")

# GUI Implementation
def create_gui():
    root = tk.Tk()
    root.title("MoySklad Product Fetcher")
    root.geometry("600x400")
    
    tree = ttk.Treeview(root)
    tree.heading("#0", text="Folders", anchor=tk.W)
    tree.pack(expand=True, fill=tk.BOTH)
    
    all_folders = fetch_all_folders()
    parent_folders = get_parent_folders(all_folders)
    folder_dict = {}
    
    def populate_tree(parent, folders):
        for folder in folders:
            folder_id = tree.insert(parent, "end", text=folder['name'], open=False)
            folder_dict[folder_id] = folder['meta']['href']
            subfolders = get_subfolders(all_folders, folder['meta']['href'])
            if subfolders:
                populate_tree(folder_id, subfolders)
    
    populate_tree("", parent_folders)
    
    def on_fetch():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a folder.")
            return
        selected_folder = folder_dict[selected_item[0]]
        fetch_subfolders = messagebox.askyesno("Fetch", "Do you want to include subfolders?")
        selected_folders = [selected_folder]
        if fetch_subfolders:
            selected_folders.extend([sf['meta']['href'] for sf in get_all_subfolders(all_folders, selected_folder)])
        products = fetch_products(selected_folders)
        if products:
            export_to_excel(products)
        else:
            messagebox.showinfo("Info", "No products found.")
    
    fetch_button = tk.Button(root, text="Fetch Products", command=on_fetch)
    fetch_button.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()
