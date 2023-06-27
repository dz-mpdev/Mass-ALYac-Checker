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
    lines = file.read().strip().split('\n')

# Group IPs by room
rooms = {}
current_room = None
for line in lines:
    line = line.strip()
    if line:
        if line.islower():
            current_room = line.capitalize()
            rooms[current_room] = []
        else:
            rooms[current_room].append(line)

# Get the hostname
hostname = socket.gethostname()
separator = "◇─" * 22
print('IP Address       |  HMI Location        | Latest Update     | Latest Scan  ')
for room, ip_list in rooms.items():
    print(f"\n{room}:")
    counter = 0  # Reset the counter for each room
    for ip in ip_list:
        counter += 1
        folder_path = get_folder_path(ip)
        server_scan_folder_name = "server_scan"
        directory_scan_path = get_directory_scan_path(ip)
        
        try:
            if not os.path.exists(folder_path):
                print("HMI Offline:", ip)
                continue

            modified_date_scan = get_latest_modified_date(folder_path, server_scan_folder_name)
            if modified_date_scan:
                print(f"{counter}  IP Address          :", ip, "-",hostname)
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
