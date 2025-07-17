import os
import csv
from winreg import *
import platform

registry = ConnectRegistry(None, HKEY_LOCAL_MACHINE) 
RawKey = OpenKey(registry, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")

i = 0
print("App Name : Subkey Name : Enum.")
try:
    while True:
        subkey_name = EnumKey(RawKey, i)
        subkey_path = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" + '\\' + subkey_name
        try:
            subkey = OpenKey(registry, subkey_path)
            display_name, reg_type = QueryValueEx(subkey, "DisplayName")
            print(f"{display_name} : {subkey_name} : {i}")
            CloseKey(subkey)
        except FileNotFoundError: #for when subkey does not have DisplayName attribute
            pass
        i = i + 1
except WindowsError as e:
    print(f"Error: {e}" )

CloseKey(RawKey)