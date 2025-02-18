import requests
import pandas as pd
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, Menu
from requests.auth import HTTPBasicAuth
from datetime import datetime

EMAILS_FILE = "used_emails.json"

API_BASE_URL = 'https://api.moysklad.ru/api/remap/1.2/entity'
FOLDERS_URL = f'{API_BASE_URL}/productfolder'
ASSORTMENT_URL = f'{API_BASE_URL}/assortment?filter=stockMode=positiveOnly'

auth = None

def load_used_emails():
    try:
        with open(EMAILS_FILE, "r") as file:
            data = json.load(file)
            return data if isinstance(data, dict) else {}
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def save_used_email(email, name):
    emails = load_used_emails()
    emails[email] = name
    with open(EMAILS_FILE, "w") as file:
        json.dump(emails, file)

def delete_email():
    emails = load_used_emails()
    if emails:
        selected_email = simpledialog.askstring("Delete Email", "Enter email to delete:", initialvalue=list(emails.keys())[0])
        if selected_email in emails:
            emails.pop(selected_email)
            with open(EMAILS_FILE, "w") as file:
                json.dump(emails, file)
            messagebox.showinfo("Success", f"Deleted {selected_email}")
    else:
        messagebox.showinfo("Info", "No saved emails to delete.")

def fill_email(event):
    selected_name = email_entry.get()
    email = next((e for e, n in load_used_emails().items() if n == selected_name), None)
    if email:
        email_entry.set(email)

def fetch_data(url, params=None):
    response = requests.get(url, auth=auth, params=params)
    if response.status_code == 200:
        return response.json().get('rows', [])
    messagebox.showerror("Error", f"Failed to fetch data from {url}")
    return []

def get_subfolders(folders, parent_folder_id):
    return [folder for folder in folders if folder.get('productFolder', {}).get('meta', {}).get('href') == parent_folder_id]

def get_all_subfolders(folders, parent_folder_id):
    subfolders = get_subfolders(folders, parent_folder_id)
    return subfolders + [sf for folder in subfolders for sf in get_all_subfolders(folders, folder['meta']['href'])]

def fetch_products(folder_hrefs):
    products = []
    for folder_href in folder_hrefs:
        offset = 0
        while True:
            params = {'filter': f'productFolder={folder_href}', 'limit': 1000, 'offset': offset}
            batch = fetch_data(ASSORTMENT_URL, params)
            products.extend(batch)
            if len(batch) < 1000:
                break
            offset += 1000
    return products

def export_to_excel(products, folder_name):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"{folder_name}_{timestamp}.xlsx"
    df = pd.DataFrame([
        {
            'Code': p.get('code', ''),
            'Name': p.get('name', ''),
            'Barcode': ', '.join(
                ''.join(filter(str.isdigit, bc.get('ean13', ''))) for bc in p.get('barcodes', [])
            )
            'Stock': p.get('stock', {}).get('quantity', 0),   
        }
        for p in products
    ])
    df.to_excel(filename, index=False)
    messagebox.showinfo("Success", f"Data exported to {filename}")

def open_main_menu(root, email, name):
    for widget in root.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(root)
    tree.heading("#0", text="Folders", anchor=tk.W)
    tree.pack(expand=True, fill=tk.BOTH)

    all_folders = fetch_data(FOLDERS_URL)
    folder_dict = {}

    def populate_tree(parent, folders):
        for folder in folders:
            folder_id = tree.insert(parent, "end", text=folder['name'], open=False)
            folder_dict[folder_id] = (folder['meta']['href'], folder['name'])
            populate_tree(folder_id, get_subfolders(all_folders, folder['meta']['href']))

    populate_tree("", [f for f in all_folders if 'productFolder' not in f])

    def on_fetch():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "Please select a folder.")
            return
        selected_folder, folder_name = folder_dict[selected_item[0]]
        if messagebox.askyesno("Fetch", "Include subfolders?"):
            selected_folders = [selected_folder] + [sf['meta']['href'] for sf in get_all_subfolders(all_folders, selected_folder)]
        else:
            selected_folders = [selected_folder]
        products = fetch_products(selected_folders)
        if products:
            export_to_excel(products, folder_name)
        else:
            messagebox.showinfo("Info", "No products found.")

    tk.Button(root, text="Fetch Products", command=on_fetch).pack(pady=10)

def create_gui():
    root = tk.Tk()
    root.title("MoySklad Product Fetcher")
    root.geometry("600x400")

    menu_bar = Menu(root)
    menu_bar.add_cascade(label="Saved Emails", menu=Menu(menu_bar, tearoff=0, postcommand=delete_email))
    root.config(menu=menu_bar)

    tk.Label(root, text="Select your account:").pack()
    global email_entry
    email_entry = ttk.Combobox(root, values=list(load_used_emails().values()))
    email_entry.pack(ipadx=50)
    email_entry.bind("<<ComboboxSelected>>", fill_email)

    tk.Label(root, text="Enter your MoySklad password:").pack()
    password_entry = tk.Entry(root, show="*")
    password_entry.pack(ipadx=50)

    def authenticate(event=None):
        global auth
        email = email_entry.get()
        password = password_entry.get()
        name = next((n for e, n in load_used_emails().items() if e == email), None) or simpledialog.askstring("User Name", "Enter a name for this email:")
        save_used_email(email, name)
        auth = HTTPBasicAuth(email, password)
        open_main_menu(root, email, name)

    tk.Button(root, text="Login", command=authenticate).pack(pady=10)
    root.bind("<Return>", authenticate)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
