import os
import sys
import subprocess
import time
import shutil
import getpass
import socket
import wmi
from datetime import datetime

destination = r"\\sharecrypt.oakland.edu\SAT Secure\Inventory\2025 Inventory"

def disconnect_network_share(destination):
    subprocess.run(f'net use "{destination}" /delete', shell=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def pushFile(destination,source_file):
    
    # Prompt for user credentials
    username = input("Enter your admnet username: ")
    password = getpass.getpass("Enter your password: ")


    # Check if the source file exists
    if not os.path.isfile(source_file):
        print(f"ERROR: Source file not found: {source_file}")
        input("Press Enter to continue...")
        sys.exit(1)

    # Force disconnect any existing mapping to the destination
    disconnect_network_share(destination)

    # Map the network share with supplied credentials
    print("Mapping network share...")
    map_cmd = f'net use "{destination}" /user:admnet\\{username} {password} /persistent:no'
    result = subprocess.run(map_cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if result.returncode != 0:
        print("ERROR: Unable to map network share with provided credentials.")
        input("Press Enter to continue...")
        sys.exit(1)

    # List available subdirectories in the network share
    try:
        print(f"\nAvailable folders in {destination}:")
        folders = [f for f in os.listdir(destination) if os.path.isdir(os.path.join(destination, f))]
        if not folders:
            print(f"No subdirectories found in {destination}.")
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
        sys.exit(1)
    finally:
        disconnect_network_share(destination)
        password = ""  # Clear the password

    print("Files copied successfully.")
    # Countdown before closing
    for i in range(5, 0, -1):
        print(f"Closing in {i}...")
        time.sleep(1)


def generateSysFile():
    # Instantiate a single WMI connection
    c = wmi.WMI()

    def get_system_info():
        system = c.Win32_ComputerSystem()[0]
        os_info = c.Win32_OperatingSystem()[0]
        bios = c.Win32_BIOS()[0]
        cpu = c.Win32_Processor()[0]
        memory = round(int(system.TotalPhysicalMemory) / (1024 ** 3), 2)
        domain = system.Domain
        workgroup = "Workgroup" if not system.PartOfDomain else "Domain"
        return {
            "Computer Name": socket.gethostname(),
            "Serial Number": bios.SerialNumber,
            "Manufacturer": system.Manufacturer,
            "Model": system.Model,
            "Windows Version": os_info.Caption,
            "Windows Build": os_info.BuildNumber,
            "CPU": cpu.Name,
            "Total RAM (GB)": memory,
            "Domain/Workgroup": domain if workgroup == "Domain" else workgroup,
            "System Boot Time": os_info.LastBootUpTime.split('.')[0]
        }

    def get_mac_addresses():
        macs = {}
        for adapter in c.Win32_NetworkAdapterConfiguration():
            if adapter.MACAddress:
                connection_type = "Wireless" if "wireless" in adapter.Description.lower() else "Wired"
                macs[adapter.Description] = {
                    "MAC Address": adapter.MACAddress,
                    "Type": connection_type
                }
        return macs

    def get_disk_info():
        disks = {}
        for disk in c.Win32_DiskDrive():
            interface_type = disk.InterfaceType
            ssd_status = "SSD" if "ssd" in interface_type.lower() else "HDD"
            disks[disk.DeviceID] = {
                "Model": disk.Model,
                "Size (GB)": round(int(disk.Size) / (1024 ** 3), 2),
                "Type": ssd_status,
                "Interface": interface_type
            }
        return disks

    def write_to_file(data, filename):
        with open(filename, "w", encoding="utf-8") as f:
            f.write(f"System Information - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 50 + "\n\n")
            for key, value in data["System Info"].items():
                f.write(f"{key}: {value}\n")
            f.write("\nMAC Addresses:\n" + "-" * 50 + "\n")
            for desc, mac in data["MAC Addresses"].items():
                f.write(f"{desc} ({mac['Type']}): {mac['MAC Address']}\n")
            f.write("\nHard Disks:\n" + "-" * 50 + "\n")
            for dev, disk in data["Disk Info"].items():
                f.write(f"{dev}: {disk['Model']} - {disk['Size (GB)']} GB - {disk['Type']} ({disk['Interface']})\n")
        return filename

    system_info = get_system_info()
    mac_addresses = get_mac_addresses()
    disk_info = get_disk_info()

    data = {
        "System Info": system_info,
        "MAC Addresses": mac_addresses,
        "Disk Info": disk_info
    }
    now = datetime.now().strftime("%Y-%m-%d")
    file_name = f"{system_info['Serial Number']}-{system_info['Computer Name']}-{now}.txt"
    write_to_file(data, file_name)
    print(f"System information has been written to {file_name}")
    return file_name
   

if __name__ == "__main__":
    file_name = generateSysFile()
    pushFile(destination,file_name)

