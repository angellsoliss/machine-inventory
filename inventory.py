import os
import csv
from winreg import *
from datetime import datetime

registry = ConnectRegistry(None, HKEY_LOCAL_MACHINE) 
RawKey = OpenKey(registry, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")

if not os.path.exists("inventory_reports"):
    os.makedirs("inventory_reports")

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
file_path = os.path.join("inventory_reports", f'{timestamp}')

with open(file_path, 'w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(['Display Name', 'Subkey Name', '#'])

    i = 1
    print("App Name : Subkey Name : Enum.")
    try:
        while True:
            subkey_name = EnumKey(RawKey, i)
            subkey_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" + '\\' + subkey_name
            try:
                subkey = OpenKey(registry, subkey_path)
                display_name, reg_type = QueryValueEx(subkey, "DisplayName")
                print(f"{display_name} : {subkey_name} : {i}")
                writer.writerow([i, display_name, subkey_name])
                CloseKey(subkey)
            except FileNotFoundError: #for when subkey does not have DisplayName attribute
                pass
            i = i + 1
    except WindowsError as e:
        print(f"Error: {e}" )

CloseKey(RawKey)