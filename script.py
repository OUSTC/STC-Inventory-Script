#Much Nicer Inventory Test
import os
import time
import psutil
import platform
from datetime import datetime
import wmi
import subprocess
import winreg
import shutil
import getpass
import sys
from datetime import datetime
import tempfile
uname= platform.uname()

destination = r"\\sharecrypt.oakland.edu\SAT Secure\Inventory\2025 Inventory"

if getattr(sys, 'frozen', False):  # Running as .exe
    os.chdir(os.path.dirname(sys.executable))  # Change to the folder of the .exe

#Pc Name
def pc_name():
    pc_name= uname.node
    return pc_name

#Pc Manufacturer
def system_manufacturer():
    c= wmi.WMI()
    for system in c.Win32_ComputerSystem():
        system_manufacturer= system.Manufacturer
    return system_manufacturer

#System Model
def system_model():
    c= wmi.WMI()
    for system in c.Win32_ComputerSystem():
        system_model= system.Model
    return system_model

#serial number
def get_serial_number():
    c= wmi.WMI()
    bios_info= c.Win32_BIOS()[0]
    serial_number= bios_info.SerialNumber.strip()
    return serial_number

#Connection Specific DNS-Suffix
def fqdn():
    c= wmi.WMI()
    for nic in c.Win32_NetworkAdapterConfiguration():
        if nic.DNSDomain:
            return f"FQDN: {nic.DNSDomain}"
    return "FQDN: N/A"
        
#Win Edition
def win_edition():
    win_edition= uname.release
    return win_edition

#Win Build Version
def win_build():
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            display_version, _ = winreg.QueryValueEx(key, "DisplayVersion")
            return display_version  # e.g., "23H2" or "24H2"
    except FileNotFoundError:
        return "Unknown Windows Version"
    
#Mac Addresses
def get_mac_address():
    c = wmi.WMI()
    mac_addresses = {"Ethernet": "N/A", "Wi-Fi": "N/A"}  # Default values
    
    for nic in c.Win32_NetworkAdapterConfiguration():
        if nic.MACAddress and nic.Description:
            if "Wi-Fi" in nic.Description or "Wireless" in nic.Description:
                mac_addresses["Wi-Fi"] = nic.MACAddress
            elif "Ethernet" in nic.Description:
                mac_addresses["Ethernet"] = nic.MACAddress
    return f"\nEthernet MAC: {mac_addresses['Ethernet']}\nWi-Fi MAC: {mac_addresses['Wi-Fi']}"

#Cpu Infopyth
def get_cpu_name():
    try:
        output = subprocess.check_output(
            ['powershell', '-Command', "(Get-CimInstance Win32_Processor).Name"],
            shell=True
        ).decode().strip()
        return output
    except Exception as e:
        return str(e)
        
#Memory Size
def get_memory_size():
    mem = psutil.virtual_memory().total
    return f"{mem / (1024**3):.2f} GB"

#Memory Speed
def get_memory_speed_windows():
    try:
        output = subprocess.check_output(
            ['powershell', '-Command', "Get-CimInstance Win32_PhysicalMemory | Select-Object -ExpandProperty Speed"],
            shell=True
        ).decode().strip().splitlines()
        speeds = [line.strip() for line in output if line.strip().isdigit()]
        return ", ".join(speeds) if speeds else "N/A"
    except Exception as e:
        return str(e)

#Drive Size and Type
def get_drive_info_windows():
    try:
        output = subprocess.check_output(
            'powershell "Get-PhysicalDisk | Select MediaType, BusType, Size | Format-Table -HideTableHeaders"',
            shell=True
        ).decode().strip().split("\n")

        drives = []
        for line in output:
            parts = line.strip().split()
            if len(parts) >= 3:  # Ensure all fields are present
                media_type = parts[0]
                bus_type = parts[1]
                size_bytes = int(parts[2])

                # Convert size to GB
                size_gb = size_bytes / (1024**3)

                # Classify drive type
                if "SSD" in media_type:
                    drive_type = "SSD (SATA)" if "SATA" in bus_type else "NVMe SSD"
                elif "HDD" in media_type:
                    drive_type = "HDD"
                else:
                    continue  # Skip unknown types

                drives.append(f"{drive_type}: {size_gb:.2f} GB")
        return "\n".join(drives) if drives else "No valid drives detected" 
    except Exception as e:
        return str(e)


def get_info():
    info = "\n".join([

        f"Creation Date: {datetime.now().strftime("%A, %B %d, %Y %I:%M %p")}",
        f"PC Name: {pc_name()}",
        f"System Manufacturer: {system_manufacturer()}",
        f"System Model: {system_model()}",
        f"Serial Number: {get_serial_number()}",
        f"FQDN: {fqdn()}",
        f"Windows Edition: {win_edition()}",
        f"Windows Build: {win_build()}",
        f"Mac Addresses:{get_mac_address()}",
        f"CPU: {get_cpu_name()}",
        f"Memory Size: {get_memory_size()}",
        f"Memory Speed: {get_memory_speed_windows()}",
        f"Drive Info:\n{get_drive_info_windows()}"
    ])
    return info


def disconnect_network_share(destination):
    subprocess.run(f'net use "{destination}" /delete', shell=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def pushFile(destination, source_file):
   
     # Check if the source file exists
    if not os.path.isfile(source_file):
        print(f"ERROR: Source file not found: {source_file}")
        input("Press Enter to continue...")
        sys.exit(1)

    # Force disconnect any existing mapping to the destination
    disconnect_network_share(destination)

    # Prompt for user credentials
    username = input("Enter your admnet username: ")
    password = getpass.getpass("Enter your password: ")

    # Map the network share with supplied credentials
    map_cmd = f'net use "{destination}" /user:admnet\\{username} {password} /persistent:no'
    result = subprocess.run(map_cmd, shell=True, capture_output=True, text=True)

    if result.returncode != 0:
        print("ERROR: Unable to map network share with provided credentials.")
        print("Details:", result.stderr)  # Log the exact error
        input("Press Enter to continue...")
        sys.exit(1)

    # os.chdir(destination)  # Change working directory

    # List available subdirectories in the network share
    try:
        print(f"\nAvailable folders in {destination}:")
        folders = [f for f in os.listdir(destination) if os.path.isdir(os.path.join(destination, f))]
        if not folders:
            print(f"No subdirectories found in {destination}.")
            input("Press Enter to continue...")
            sys.exit(1)

        # Display folders with numbering
        folder_dict = {idx: folder for idx, folder in enumerate(folders, start=1)}
        for idx, folder in folder_dict.items():
            print(f"  {idx}: {folder}")

        while True:
            try:
                folder_num = int(input("Enter the number corresponding to your target folder: "))
                if folder_num in folder_dict:
                    chosen_folder = os.path.join(destination, folder_dict[folder_num])
                    break
                else:
                    print("Invalid selection. Please try again.")
            except ValueError:
                print("Invalid input. Please enter a number.")

        print(f"You have chosen: {chosen_folder}")

        print("Copying source file to chosen folder...")
        shutil.copy(source_file, chosen_folder)
    except Exception as e:
        print("ERROR: Operation failed. Check your permissions.")
        input("Press Enter to continue...")
        sys.exit(1)
    finally:
        disconnect_network_share(destination)
        password = ""  # Clear the password

    print("Files copied successfully.")
    if os.path.exists(source_file):
        os.remove(source_file)
        print(f"Deleted local copy: {source_file}")
    
    input("Press Enter to continue...")
    sys.exit(0)


temp_dir = tempfile.gettempdir()
output_file = os.path.join(tempfile.gettempdir(), get_serial_number() + ".txt")



with open(output_file, "w", encoding="utf-8") as f:
    f.write(get_info())

print(f"Generated file at: {output_file}")

pushFile(destination, output_file)
