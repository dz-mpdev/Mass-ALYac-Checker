import os
import datetime
import platform
import socket
import colorama
from colorama import Fore
import msvcrt

class AlyacChecker:
    def __init__(self, hmi_list):
        self.hmi_list = hmi_list
        self.current_time = datetime.datetime.now()
        self.hostname = socket.gethostname()

    def get_folder_path(self, ip):
        return r"\\{}\\c$\ProgramData\ESTsoft\ALYac\log".format(ip)

    def get_directory_scan_path(self, ip):
        architecture = platform.architecture()[0]
        base_directory = "ALYac"
        if architecture == "32bit":
            base_directory = "ALYac"
        elif architecture == "64bit":
            base_directory = "ALYac"
        else:
            raise ValueError("Unsupported architecture.")
        return r"\\{}\\c$\ProgramData\ESTsoft\{}\update".format(ip, base_directory)

    def get_latest_modified_date(self, folder_path, folder_name):
        folder_scan_path = os.path.join(folder_path, folder_name)
        if os.path.exists(folder_scan_path):
            modified_timestamp = os.path.getmtime(folder_scan_path)
            modified_date_scan = datetime.datetime.fromtimestamp(modified_timestamp)
            return modified_date_scan
        else:
            return None

    def run(self):
        with open(self.hmi_list, "r") as file:
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

        separator = "⊱" * 1
        print(f"{Fore.CYAN}Room - IP - Hostname                       |  Alyac Status")
        print(f"⊱{Fore.RESET}")
        for room, ip_list in rooms.items():
            print(f"{Fore.CYAN}{room}:{Fore.RESET}")
            counter = 0  # Reset the counter for each room
            for ip in ip_list:
                counter += 1
                folder_path = self.get_folder_path(ip)
                server_scan_folder_name = "server_scan"
                directory_scan_path = self.get_directory_scan_path(ip)

                try:
                    if not os.path.exists(folder_path):
                        print(f"{counter}. {Fore.RED}{ip:<17} | HMI Offline         |  -                  {Fore.RESET}")
                        continue

                    modified_date_scan = self.get_latest_modified_date(folder_path, server_scan_folder_name)
                    print(f"{counter}. {ip:<17}- {self.hostname:<20} |  {Fore.GREEN}Last Scan   {Fore.RESET} {modified_date_scan.strftime('%H:%M %Y-%m-%d'):<17}")

                    timestamp = os.path.getmtime(directory_scan_path)
                    latest_update_date = datetime.datetime.fromtimestamp(timestamp)
                    print(f"   {ip:<17}- {self.hostname:<20} |  {Fore.YELLOW}Last Update {Fore.RESET} {latest_update_date.strftime('%H:%M %Y-%m-%d'):<17}")

                except Exception as e:
                    print(f"{counter}. {Fore.RED}{ip:<17} | {self.hostname:<20} | -                  | An error occurred: {str(e)}{Fore.RESET}")

            print(f"{Fore.CYAN}{separator}{Fore.RESET}")

        print(f"Result Checking        : {Fore.YELLOW}{self.current_time.strftime('%Y-%m-%d %H:%M:%S')}{Fore.RESET}")

        self.wait_for_key()

    def wait_for_key(self):
        print("\nPress Enter to re-scan or Esc to exit.")
        while True:
            if msvcrt.kbhit():
                key = msvcrt.getch()
                if key == b'\r':  # Enter key
                    checker.run()
                    break
                elif key == b'\x1b':  # Esc key
                    print("Application stopped.")
                    return

# Usage:
colorama.init()
checker = AlyacChecker("hmi.txt")
checker.run()
