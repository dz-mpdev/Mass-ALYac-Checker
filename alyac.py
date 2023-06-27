import os
import datetime
import platform
import socket

def get_folder_path(ip):
    return r"\\{}\\c$\ProgramData\ESTsoft\ALYac\log".format(ip)

def get_directory_scan_path(ip):
    architecture = platform.architecture()[0]
    base_directory = "ALYac"
    if architecture == "32bit":
        base_directory = "ALYac"
    elif architecture == "64bit":
        base_directory = "ALYac"
    else:
        raise ValueError("Unsupported architecture.")
    return r"\\{}\\c$\ProgramData\ESTsoft\{}\update".format(ip, base_directory)

def get_latest_modified_date(folder_path, folder_name):
    folder_scan_path = os.path.join(folder_path, folder_name)
    if os.path.exists(folder_scan_path):
        modified_timestamp = os.path.getmtime(folder_scan_path)
        modified_date_scan = datetime.datetime.fromtimestamp(modified_timestamp)
        return modified_date_scan
    else:
        return None

hmi_list = "hmi.txt"
current_time = datetime.datetime.now()

with open(hmi_list, "r") as file:
    ip_list = file.read().strip().split('\n')

# Get the hostname
hostname = socket.gethostname()
counter = 0
separator = "◇─" * 22

for ip in ip_list:
    counter += 1
    folder_path = get_folder_path(ip)
    server_scan_folder_name = "server_scan"
    directory_scan_path = get_directory_scan_path(ip)
    
    try:
        if not os.path.exists(folder_path):
            print("Folder path not found for IP:", ip)
            continue

        modified_date_scan = get_latest_modified_date(folder_path, server_scan_folder_name)
        if modified_date_scan:
            print(separator)
            print(f"{counter}. Hostname: {hostname}")
            print("   IP Address          :", ip)
            print("   ALYac Latest Scan   :", modified_date_scan.strftime("%Y-%m-%d %H:%M:%S"))
        else:
            print("Server_scan folder not detected for IP:", ip)

        if os.path.exists(directory_scan_path):
            timestamp = os.path.getmtime(directory_scan_path)
            latest_update_date = datetime.datetime.fromtimestamp(timestamp)
            print("   ALYac Latest Update :", latest_update_date.strftime("%Y-%m-%d %H:%M:%S"))
        else:
            print("   ALYac Latest Update not detected for IP:", ip)
        
    except Exception as e:
        print("An error occurred for IP:", ip)
        print("Error message :", str(e))

print(separator)
print("Result Checking        :", current_time.strftime("%Y-%m-%d %H:%M:%S"))
