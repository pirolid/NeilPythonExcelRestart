# Import necessary libraries
import os
import subprocess
import time
from colorama import Fore, Style
import xlwings as xw

#------------------------------------------------------------------------------------------------------------------
# Global variable to track if the folder has been opened
#------------------------------------------------------------------------------------------------------------------
folder_opened = False
#------------------------------------------------------------------------------------------------------------------
# Function to open the current script's folder in Windows Explorer
#------------------------------------------------------------------------------------------------------------------
def open_current_folder():
    """
    Opens the folder containing this script in Windows Explorer.
    """
    global folder_opened
    if not folder_opened:
        current_folder = os.path.dirname(os.path.abspath(__file__))
        subprocess.run(f'explorer "{current_folder}"', shell=True)
        print(f"{Fore.GREEN}Opened folder: {current_folder}{Style.RESET_ALL}\n")
        folder_opened = True
    else:
        print(f"{Fore.YELLOW}Folder is already open. Skipping this step.{Style.RESET_ALL}\n")
    return os.path.dirname(os.path.abspath(__file__))
#------------------------------------------------------------------------------------------------------------------
# Function to handle the Excel file
#------------------------------------------------------------------------------------------------------------------
def handle_excel_file(folder_path):
    """
    Continuously handles the Excel file named 'prices' (with .xls or .xlsx extensions).
    Restarts the target workbook only, ensures macros run uninterrupted, and keeps other workbooks unaffected.
    """
    try:
        # Define potential file names
        excel_files = ["prices.xls", "prices.xlsx"]

        # Check for the presence of the file before asking the user for input
        found_file = None
        for file in excel_files:
            if os.path.exists(os.path.join(folder_path, file)):
                found_file = file
                break

        if not found_file:
            print(f"{Fore.RED}No Excel file named 'prices' found. Exiting script.{Style.RESET_ALL}\n")
            return

        print(f"{Fore.YELLOW}File found: {found_file}{Style.RESET_ALL}\n")
        excel_path = os.path.join(folder_path, found_file)

        # Ask user for the timer once
        while True:
            try:
                ResetTimer = int(input(f"{Fore.CYAN}Enter the number of seconds before restarting the file (set once): {Style.RESET_ALL}"))
                if ResetTimer <= 0:
                    print(f"{Fore.RED}Please enter a positive number.{Style.RESET_ALL}")
                else:
                    break
            except ValueError:
                print(f"{Fore.RED}Invalid input. Please enter a valid number.{Style.RESET_ALL}")

        # Open Excel app
        app = xw.App(visible=True, add_book=False)

        # Infinite loop to restart the target workbook
        while True:
            try:
                # Check if the workbook is already open
                workbook = None
                for wb in app.books:
                    if wb.fullname == excel_path:
                        workbook = wb
                        print(f"{Fore.YELLOW}Workbook is already open. Closing and restarting it.{Style.RESET_ALL}\n")
                        workbook.close()
                        break

                # Open the workbook
                print(f"{Fore.CYAN}Opening the workbook: {found_file}{Style.RESET_ALL}")
                workbook = app.books.open(excel_path, update_links=False, read_only=False, ignore_read_only_recommended=True)

                # Wait for the defined time before writing to the workbook
                for i in range(ResetTimer, 0, -1):
                    print(f"{Fore.CYAN}{i} seconds remaining before restarting the workbook...{Style.RESET_ALL}")
                    time.sleep(1)

                # Write to cell A1 and save
                print(f"{Fore.CYAN}Writing 'tos.rtd' to cell A1.{Style.RESET_ALL}")
                sheet = workbook.sheets[0]  # Access the first sheet
                sheet.range("A1").value = "tos.rtd"
                workbook.save()
                print(f"{Fore.GREEN}Updated cell A1 with 'tos.rtd' in file: {found_file}{Style.RESET_ALL}\n")

                # Close the workbook after updating
                workbook.close()
                print(f"{Fore.YELLOW}Workbook closed: {found_file}{Style.RESET_ALL}\n")

            except Exception as e:
                print(f"{Fore.RED}An error occurred while processing the workbook: {e}{Style.RESET_ALL}\n")

    except Exception as e:
        print(f"{Fore.RED}An unexpected error occurred: {e}{Style.RESET_ALL}\n")


#------------------------------------------------------------------------------------------------------------------
# Main function
#------------------------------------------------------------------------------------------------------------------
if __name__ == "__main__":
    # Infinite loop to keep the script running
    while True:
        # Open the current folder in Windows Explorer
        script_folder = open_current_folder()
        
        # Check for the Excel file and handle it
        handle_excel_file(script_folder)
        
        # Prompt the user if they want to exit or continue
        user_input = input(f"{Fore.CYAN}Press 'E' to exit or any other key to restart: {Style.RESET_ALL}")
        if user_input.strip().lower() == 'e':
            print(f"{Fore.GREEN}Exiting the script. Goodbye!{Style.RESET_ALL}")
            break
