import os
import csv
from winreg import *
from datetime import datetime
import glob

registry = ConnectRegistry(None, HKEY_LOCAL_MACHINE) 
RawKey = OpenKey(registry, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
scriptPath = str(os.path.dirname(__file__))

def directoryOverhead() -> None:
    if not os.path.exists("inventory_reports"):
        os.makedirs("inventory_reports")

def checkLastCSV(directory) -> str | bool:
    path = directory + "\inventory_reports\*.csv"
    files = glob.glob(path)
    if files:
        latest_file = max(files, key=os.path.getctime)
        return(f'{latest_file}')
    else:
        return False   

def lastCSVArray(file) -> list:
    col_data = []
    with open(file, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
    
        #skip header row
        next(reader, None)
        for row in reader:
            if row:
                col_data.append(row[1])
    return col_data

def compareInventories(array1, array2) -> str:
    inventoriesDifference = set(array1) != set(array2)

    if inventoriesDifference:
        new = set(array1) - set(array2)
        removed = set(array2) - set(array1)
        return(f'\nNew Items: {new}\nRemoved Items: {removed}')
    else:
        return('No changes made to system inventory since last report.')

programs = []
def inventory(array) -> str | list:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    file_path = os.path.join("inventory_reports", f'{timestamp}.csv')

    with open(file_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['#', 'Display Name', 'Subkey Name'])
        i = 1
        print("Enum. | Display Name | Subkey Name")
        try:
            while True:
                subkey_name = EnumKey(RawKey, i)
                subkey_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" + '\\' + subkey_name
                try:
                    subkey = OpenKey(registry, subkey_path)
                    display_name, reg_type = QueryValueEx(subkey, "DisplayName")
                    print(f"{i} | {display_name} | {subkey_name}")
                    writer.writerow([i, display_name, subkey_name])
                    array.append(display_name)
                    CloseKey(subkey)
                except FileNotFoundError: #for when subkey does not have DisplayName attribute
                    pass
                i = i + 1
        except WindowsError as e:
            return(f"Error: {e}" )
    CloseKey(RawKey)
    return array

directoryOverhead()
lastCSV = checkLastCSV(scriptPath)
inventory(programs)
if lastCSV:
    lastestInventoryReport = lastCSVArray(lastCSV)
    lastCSVDate = lastCSV[-23:-3]
    print(f'\nLast Inventory Report Taken on {lastCSVDate}')
    print(compareInventories(programs, lastestInventoryReport))