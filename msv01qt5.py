import requests
import pandas as pd
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, Menu
from requests.auth import HTTPBasicAuth
from datetime import datetime

EMAILS_FILE = "used_emails.json"

# API Endpoints
FOLDERS_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/productfolder'
ASSORTMENT_URL = 'https://api.moysklad.ru/api/remap/1.2/entity/assortment'

auth = None  # Global variable to store authentication credentials

def load_used_emails():
    if os.path.exists(EMAILS_FILE):
        with open(EMAILS_FILE, "r") as file:
            try:
                data = json.load(file)
                return data if isinstance(data, dict) else {}
            except json.JSONDecodeError:
                return {}
    return {}

def save_used_email(email, name):
    emails = load_used_emails()
    emails[email] = name
    with open(EMAILS_FILE, "w") as file:
        json.dump(emails, file)

def delete_email():
    emails = load_used_emails()
    if not emails:
        messagebox.showinfo("Info", "No saved emails to delete.")
        return
    selected_email = simpledialog.askstring("Delete Email", "Select email to delete:", initialvalue=list(emails.keys())[0])
    if selected_email and selected_email in emails:
        del emails[selected_email]
        with open(EMAILS_FILE, "w") as file:
            json.dump(emails, file)
        messagebox.showinfo("Success", f"Deleted {selected_email}")

def fill_email(event):
    selected_name = email_entry.get()
    emails = load_used_emails()
    email = next((e for e, n in emails.items() if n == selected_name), None)
    if email:
        email_entry.set(email)

def fetch_all_folders(auth):
    response = requests.get(FOLDERS_URL, auth=auth)
    if response.status_code != 200:
        messagebox.showerror("Error", "Error fetching folders.")
        return []
    return response.json().get('rows', [])

def get_parent_folders(folders):
    return [folder for folder in folders if 'productFolder' not in folder]

def get_subfolders(folders, parent_folder_id):
    return [folder for folder in folders if folder.get('productFolder', {}).get('meta', {}).get('href') == parent_folder_id]

def get_all_subfolders(folders, parent_folder_id):
    subfolders = get_subfolders(folders, parent_folder_id)
    all_subfolders = subfolders[:]
    for subfolder in subfolders:
        all_subfolders.extend(get_all_subfolders(folders, subfolder['meta']['href']))
    return all_subfolders

def fetch_products(folder_hrefs, auth):
    products = []
    limit = 1000
    
    for folder_href in folder_hrefs:
        offset = 0
        while True:
            params = {
                'filter': f'productFolder={folder_href}',
                'limit': limit,
                'offset': offset
            }
            response = requests.get(ASSORTMENT_URL, auth=auth, params=params)
            if response.status_code != 200:
                messagebox.showerror("Error", "Error fetching products.")
                break
            data = response.json().get('rows', [])
            products.extend(data)
            if len(data) < limit:
                break
            offset += limit
    return products

def export_to_excel(products, folder_name):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{folder_name}_{timestamp}.xlsx"
    for product in products:
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
    df = pd.DataFrame([{ 
        'Code': p.get('code', ''), 
        'Name': p.get('name', ''), 
        'Barcode': barcodes_str
    } for p in products])
    df.to_excel(filename, index=False)
    messagebox.showinfo("Success", f"Data exported to {filename}")

def open_main_menu(root, auth, email, name):
    for widget in root.winfo_children():
        widget.destroy()
    
    tree = ttk.Treeview(root)
    tree.heading("#0", text="Folders", anchor=tk.W)
    tree.pack(expand=True, fill=tk.BOTH)
    
    all_folders = fetch_all_folders(auth)
    parent_folders = get_parent_folders(all_folders)
    folder_dict = {}
    
    def populate_tree(parent, folders):
        for folder in folders:
            folder_id = tree.insert(parent, "end", text=folder['name'], open=False)
            folder_dict[folder_id] = (folder['meta']['href'], folder['name'])
            subfolders = get_subfolders(all_folders, folder['meta']['href'])
            if subfolders:
                populate_tree(folder_id, subfolders)
    
    populate_tree("", parent_folders)
    
    def on_fetch():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a folder.")
            return
        selected_folder, folder_name = folder_dict[selected_item[0]]
        fetch_subfolders = messagebox.askyesno("Fetch", "Do you want to include subfolders?")
        selected_folders = [selected_folder]
        if fetch_subfolders:
            selected_folders.extend([sf['meta']['href'] for sf in get_all_subfolders(all_folders, selected_folder)])
        products = fetch_products(selected_folders, auth)
        if products:
            export_to_excel(products, folder_name)
        else:
            messagebox.showinfo("Info", "No products found.")
    
    fetch_button = tk.Button(root, text="Fetch Products", command=on_fetch)
    fetch_button.pack(pady=10)

def create_gui():
    root = tk.Tk()
    root.title("MoySklad Product Fetcher")
    root.geometry("600x400")
    
    menu_bar = Menu(root)
    root.config(menu=menu_bar)
    
    email_menu = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Saved Emails", menu=email_menu)
    email_menu.add_command(label="Delete Saved Email", command=delete_email)
    
    tk.Label(root, text="Select your account:").pack()
    used_emails = load_used_emails()
    names_list = list(used_emails.values())
    global email_entry
    email_entry = ttk.Combobox(root, values=names_list)
    email_entry.pack(ipadx=50)
    email_entry.bind("<<ComboboxSelected>>", fill_email)
    
    tk.Label(root, text="Enter your MoySklad password:").pack()
    password_entry = tk.Entry(root, show="*")
    password_entry.pack(ipadx=50)
    
    def authenticate(event=None):
        global auth
        emails = load_used_emails()
        email = email_entry.get()
        password = password_entry.get()
        name = next((n for e, n in emails.items() if e == email), None)
        if name is None:
            name = simpledialog.askstring("User Name", "Enter a name for this email:")
        save_used_email(email, name)
        auth = HTTPBasicAuth(email, password)
        open_main_menu(root, auth, email, name)
    
    login_button = tk.Button(root, text="Login", command=authenticate)
    login_button.pack(pady=10)
    root.bind("<Return>", authenticate)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()
