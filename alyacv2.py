import os
import datetime
import platform
import socket
import colorama
import msvcrt
from colorama import Fore

class AlyacChecker:
    def __init__(self, hmi_list):
        self.hmi_list = hmi_list
        self.current_time = datetime.datetime.now()
        self.hostname = socket.gethostname()

    def get_folder_path(self, ip):
        return r"\\{}\\c$\ProgramData\ESTsoft\ALYac\log".format(ip)

    def get_directory_scan_path(self, ip):
        base_directory = "ALYac"
        return r"\\{}\\c$\ProgramData\ESTsoft\{}\update".format(ip, base_directory)

    def get_latest_modified_date(self, folder_path, folder_name):
        folder_scan_path = os.path.join(folder_path, folder_name)
        if os.path.exists(folder_scan_path):
            modified_timestamp = os.path.getmtime(folder_scan_path)
            modified_date_scan = datetime.datetime.fromtimestamp(modified_timestamp)
            return modified_date_scan
        else:
            return None

    def parse_hmi_file(self):
        with open(self.hmi_list, "r") as file:
            lines = [line.strip() for line in file if line.strip()]

        rooms = {}
        current_room = None

        for line in lines:
            if line.islower():
                current_room = line.capitalize()
                rooms[current_room] = []
            else:
                ip, computer_name = line.split("|")
                rooms[current_room].append((ip.strip(), computer_name.strip()))

        return rooms

    def run(self):
        rooms = self.parse_hmi_file()

        separator = "⊱" * 1
        print(f"{Fore.CYAN}Room - IP - Hostname                       |  Alyac Status")
        print(f"⊱{Fore.RESET}")

        for room, ip_list in rooms.items():
            print(f"{Fore.CYAN}{room}:{Fore.RESET}")
            for counter, (ip, computer_name) in enumerate(ip_list, start=1):
                folder_path = self.get_folder_path(ip)
                server_scan_folder_name = "server_scan"
                directory_scan_path = self.get_directory_scan_path(ip)

                try:
                    if not os.path.exists(folder_path):
                        print(f"{counter}. {Fore.RED}{ip:<17} | HMI Offline         |  -                  {Fore.RESET}")
                        continue

                    modified_date_scan = self.get_latest_modified_date(folder_path, server_scan_folder_name)
                    modified_date_scan_str = modified_date_scan.strftime('%H:%M:%S %Y-%m-%d') if modified_date_scan else ""
                    print(f"{counter}. {ip:<17}- {computer_name:<20} |  {Fore.GREEN}Last Scan   {Fore.RESET} {modified_date_scan_str:<17}")

                    timestamp = os.path.getmtime(directory_scan_path)
                    latest_update_date = datetime.datetime.fromtimestamp(timestamp)
                    print(f"   {ip:<17}- {computer_name:<20} |  {Fore.YELLOW}Last Update {Fore.RESET} {latest_update_date.strftime('%H:%M:%S %Y-%m-%d'):<17}")

                except Exception as e:
                    print(f"{counter}. {Fore.RED}{ip:<17} | {computer_name:<20} | -                  | An error occurred: {str(e)}{Fore.RESET}")

            print(f"{Fore.CYAN}{separator}{Fore.RESET}")

        print(f"Checking Time    : {Fore.YELLOW}{self.current_time.strftime('%H:%M:%S %Y-%m-%d')}{Fore.RESET}")

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
