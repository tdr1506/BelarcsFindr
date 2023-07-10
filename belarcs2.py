import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import re
import socket
import openpyxl
from bs4 import BeautifulSoup
from collections import OrderedDict
from datetime import datetime
import pandas as pd


class Syst:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.file_dict = OrderedDict()
        self.data = pd.DataFrame(columns=['file', 'contents'])

    def add_file(self, file_name):
        self.file_dict[file_name] = datetime.now()

    def on_file_click(self, file_name):
        file_path = os.path.join(self.folder_path, file_name)
        with open(file_path, 'r', encoding='utf-8') as file_content:
            content = file_content.read()
            self.data = self.data.append({'file': file_name, 'contents': content}, ignore_index=True)
            print("CONTENTS OF PREVIOUS FILE:", content)
        self.file_dict.move_to_end(file_name, last=False)
        print("Order of Files After clicking", list(self.file_dict.keys()))

    def get_next_file(self):
        file_name, time_used = self.file_dict.popitem(last=False)
        self.add_file(file_name)
        return f"{file_name} accessed at {time_used}"


def search_files(file_paths, search_text, output_folder_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = 'Output'

    # Set column headers
    headers = ['SystemName', 'Department', 'Employee Name', 'Branch', 'Floor', 'Port', 'System Model',
               'Processor', 'Main Circuit Board', 'Drives', 'Memory Modules', 'Display']
    worksheet.append(headers)

    for file_path in file_paths:
        with open(file_path, 'r', encoding='utf-8') as file_content:
            print(f"File: {file_path}")

            # Load the HTML into BeautifulSoup
            soup = BeautifulSoup(file_content, 'html.parser')

            # System Model
            desired_caption = 'System Model'
            table = soup.find('caption', string=re.compile(desired_caption)).find_parent(
                'table') if soup.find('caption', string=re.compile(desired_caption)) else None

            system_model = ''
            if table:
                html_content = table.find('td').decode_contents().strip()
                lines = html_content.split('<br>')
                if len(lines) >= 2:
                    system_model = '\n'.join(line.strip() for line in lines[:2])
                else:
                    system_model = BeautifulSoup(html_content, 'html.parser').get_text().strip()

            # Processor
            div2 = soup.find_all("div", {'class': "reportSection rsLeft"})
            processor = div2[1].find('td').get_text(strip=True)

            # Main Circuit Board
            div2 = soup.find_all("div", {'class': "reportSection rsRight"})
            main_circuit_board = div2[1].find('td').get_text(strip=True)

            # Drives
            desired_caption4 = 'Drives'
            table4 = soup.find('caption', string=re.compile(desired_caption4)).find_parent(
                'table') if soup.find('caption', string=re.compile(desired_caption4)) else None
            drives = table4.find('td').contents[0].strip() if table4 else ''

            # Memory Modules
            div2 = soup.find_all("div", {'class': "reportSection rsRight"})
            memory_modules = div2[2].find('td').get_text(strip=True)

            # Display
            desired_caption6 = 'Display'
            table6 = soup.find('caption', string=re.compile(desired_caption6)).find_parent(
                'table') if soup.find('caption', string=re.compile(desired_caption6)) else None
            display = table6.find('td').decode_contents().split('<br>')[0].strip() if table6 else ''

            SystemName, Dept, ename, branch, sym_floor, port_with_extension = os.path.basename(file_path).split('_')[:6]

            # Remove the .html extension from the port
            port = os.path.splitext(port_with_extension)[0]

            # Add data to Excel worksheet
            data = [SystemName, Dept, ename, branch, sym_floor, port, system_model, processor,
                    main_circuit_board, drives, memory_modules, display]
            worksheet.append(data)

    # Save the Excel file
    output_file_path = os.path.join(output_folder_path, "output.xlsx")
    workbook.save(output_file_path)
    print('Excel file saved successfully!')


def confirm_search_text():
    search_text = search_entry.get()
    if search_text.lower() in ["sym"]:
        run_button.config(state=tk.NORMAL)
    else:
        messagebox.showerror("Error", "Please enter 'SYM' as the search text.")
        run_button.config(state=tk.DISABLED)


def select_output_folder():
    # Disable network connection
    socket.setdefaulttimeout(0)

    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_folder_path = filedialog.askdirectory(initialdir=desktop_path)
    output_entry.delete(0, tk.END)
    output_entry.insert(tk.END, output_folder_path)

    # Enable network connection
    socket.setdefaulttimeout(None)



def run_search():
    folder_path = r"C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp"
    search_text = search_entry.get()
    output_folder_path = output_entry.get()

    if not search_text:
        messagebox.showerror("Error", "Please enter search text.")
        return

    if not output_folder_path:
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        output_folder_path = desktop_path

    try:
        file_paths = [os.path.join(folder_path, file_name) for file_name in os.listdir(folder_path) if
                      os.path.isfile(os.path.join(folder_path, file_name))]

        # Sort the files by modification time in descending order
        file_paths.sort(key=lambda x: os.path.getmtime(x), reverse=True)

        search_files(file_paths[:1], search_text, output_folder_path)
        messagebox.showinfo("Success", "Search and file generation completed successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))


# Create the main window
window = tk.Tk()
window.title("BelarcFindr")
window.geometry("550x150")

# Create and position the widgets
search_label = tk.Label(window, text="Enter Search Text:")
search_label.grid(row=0, column=0, sticky=tk.W, padx=10, pady=10)

search_entry = tk.Entry(window, width=50)
search_entry.grid(row=0, column=1, padx=10, pady=10)

confirm_button = tk.Button(window, text="Confirm", command=confirm_search_text)
confirm_button.grid(row=0, column=2, padx=10, pady=10)

output_label = tk.Label(window, text="Output Folder Path:")
output_label.grid(row=1, column=0, sticky=tk.W, padx=10, pady=10)

output_entry = tk.Entry(window, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)

browse_button = tk.Button(window, text="Browse", command=select_output_folder)
browse_button.grid(row=1, column=2, padx=10, pady=10)

run_button = tk.Button(window, text="Run", command=run_search, state=tk.DISABLED)
run_button.grid(row=2, column=1, pady=10)

window.mainloop()