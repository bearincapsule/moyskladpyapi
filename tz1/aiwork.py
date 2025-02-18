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
    "productfolder": f"{BASE_URL}/productfolder",
    "assortment": f"{BASE_URL}/assortment",
    "variant": f"{BASE_URL}/variant",
    "customerorder": f"{BASE_URL}/customerorder",
    "demand": f"{BASE_URL}/demand",
    "salesreturn": f"{BASE_URL}/salesreturn",
    "invoiceout": f"{BASE_URL}/invoiceout",
    "paymentin": f"{BASE_URL}/paymentin",
    "purchaseorder": f"{BASE_URL}/purchaseorder",
    "supply": f"{BASE_URL}/supply",
    "purchasereturn": f"{BASE_URL}/purchasereturn",
    "invoicein": f"{BASE_URL}/invoicein",
    "paymentout": f"{BASE_URL}/paymentout",
    "organization": f"{BASE_URL}/organization",
    "counterparty": f"{BASE_URL}/counterparty",
    "store": f"{BASE_URL}/store",
    "project": f"{BASE_URL}/project",
    "employee": f"{BASE_URL}/employee"
}

EMAILS_FILE = "used_emails.json"

# Authentication Manager
class AuthManager:
    @staticmethod
    def load_used_emails():
        if os.path.exists(EMAILS_FILE):
            with open(EMAILS_FILE, "r") as file:
                try:
                    data = json.load(file)
                    return data if isinstance(data, dict) else {}
                except json.JSONDecodeError:
                    return {}
        return {}

    @staticmethod
    def save_used_email(email, name):
        emails = AuthManager.load_used_emails()
        emails[email] = name
        with open(EMAILS_FILE, "w") as file:
            json.dump(emails, file)

    @staticmethod
    def delete_email():
        emails = AuthManager.load_used_emails()
        if not emails:
            messagebox.showinfo("Info", "No saved emails to delete.")
            return
        selected_email = simpledialog.askstring("Delete Email", "Select email to delete:", initialvalue=list(emails.keys())[0])
        if selected_email and selected_email in emails:
            del emails[selected_email]
            with open(EMAILS_FILE, "w") as file:
                json.dump(emails, file)
            messagebox.showinfo("Success", f"Deleted {selected_email}")

# API Client
class MoySkladAPI:
    def __init__(self, auth):
        self.auth = auth

    async def fetch_data(self, endpoint, params=None):
        url = ENDPOINTS[endpoint]
        async with ClientSession() as session:
            for _ in range(3):  # Retry up to 3 times
                try:
                    async with session.get(url, params=params, auth=self.auth, timeout=10) as response:
                        if response.status == 200:
                            return await response.json()
                        elif response.status == 429:
                            await asyncio.sleep(1)  # Rate limit handling
                except aiohttp.ClientError:
                    await asyncio.sleep(1)
        return None

    async def fetch_entities(self, entity_type, filters=None):
        data = await self.fetch_data(entity_type, params=filters)
        return data.get('rows', []) if data else []

# Data Exporter
class DataExporter:
    @staticmethod
    def export_to_excel(data, entity_name):
        filename = f"{entity_name}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
        messagebox.showinfo("Success", f"Data exported to {filename}")

# GUI Application
class MoySkladApp:
    def __init__(self, root):
        self.root = root
        self.auth = None
        self.api = None
        self.folder_metadata = {}

        self.setup_gui()

    def setup_gui(self):
        self.root.title("MoySklad API Fetcher")
        self.root.geometry("800x600")

        menu_bar = Menu(self.root)
        self.root.config(menu=menu_bar)

        email_menu = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Saved Emails", menu=email_menu)
        email_menu.add_command(label="Delete Saved Email", command=AuthManager.delete_email)

        tk.Label(self.root, text="Email:").pack()
        self.email_entry = tk.Entry(self.root)
        self.email_entry.pack()

        tk.Label(self.root, text="Password:").pack()
        self.password_entry = tk.Entry(self.root, show="*")
        self.password_entry.pack()

        login_button = tk.Button(self.root, text="Login", command=self.authenticate)
        login_button.pack(pady=10)

    def authenticate(self):
        self.auth = HTTPBasicAuth(self.email_entry.get(), self.password_entry.get())
        self.api = MoySkladAPI(self.auth)
        self.open_main_menu()

    def open_main_menu(self):
        for widget in self.root.winfo_children():
            widget.destroy()

        entity_types = list(ENDPOINTS.keys())
        self.entity_selector = ttk.Combobox(self.root, values=entity_types)
        self.entity_selector.pack()

        fetch_button = tk.Button(self.root, text="Fetch Data", command=self.fetch_data)
        fetch_button.pack(pady=10)

    def fetch_data(self):
        entity_type = self.entity_selector.get()
        if not entity_type:
            messagebox.showerror("Error", "Please select an entity type.")
            return
        asyncio.run(self.process_data(entity_type))

    async def process_data(self, entity_type):
        data = await self.api.fetch_entities(entity_type)
        if data:
            DataExporter.export_to_excel(data, entity_type)
        else:
            messagebox.showinfo("Info", "No data found.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MoySkladApp(root)
    root.mainloop()
