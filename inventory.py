import os
import csv
from winreg import *
import platform

registry = ConnectRegistry(None, HKEY_LOCAL_MACHINE) 
RawKey = OpenKey(registry, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")

i = 0
try:
    while True:
        name = EnumKey(RawKey, i)
        print(name, i)
        i = i + 1
except WindowsError as e:
    print(f"Error: {e}" )

print(RawKey)