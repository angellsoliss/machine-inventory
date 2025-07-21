import os
import csv
from winreg import *
from datetime import datetime
import glob

registry = ConnectRegistry(None, HKEY_LOCAL_MACHINE) 
RawKey = OpenKey(registry, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
scriptPath = str(os.path.dirname(__file__))

def directoryOverhead():
    if not os.path.exists("inventory_reports"):
        os.makedirs("inventory_reports")

def checkLastCSV(directory) -> str:
    path = directory + "\inventory_reports"
    files = list(filter(os.path.isfile, glob.glob(os.path.join(path, "*"))))
    files.sort(key=os.path.getatime, reverse=True)
    if files:
        latestCSV = files[0]
        return(f'{latestCSV}')
    else:
        return False   

def lastCSVArray(file) -> list:
    col_data = []
    with open(file, 'r') as f:
        reader = csv.reader(f)
    
        #skip header row
        next(reader, None)
        for row in reader:
            if row:
                col_data.append(row[1])
    return col_data

def inventory():
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_path = os.path.join("inventory_reports", f'{timestamp}')

    with open(file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['#', 'Display Name', 'Subkey Name'])
        i = 1
        print("Enum. : Display Name : Subkey Name")
        try:
            while True:
                subkey_name = EnumKey(RawKey, i)
                subkey_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" + '\\' + subkey_name
                try:
                    subkey = OpenKey(registry, subkey_path)
                    display_name, reg_type = QueryValueEx(subkey, "DisplayName")
                    print(f"{i} : {display_name} : {subkey_name}")
                    writer.writerow([i, display_name, subkey_name])
                    CloseKey(subkey)
                except FileNotFoundError: #for when subkey does not have DisplayName attribute
                    pass
                i = i + 1
        except WindowsError as e:
            print(f"Error: {e}" )
    CloseKey(RawKey)

directoryOverhead()
lastCSV = checkLastCSV(scriptPath)
print(lastCSV)
if lastCSV:
    print(lastCSVArray(lastCSV))
inventory()