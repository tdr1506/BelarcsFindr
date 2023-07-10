## BelarcsFindr

### Problem Statement:

The code provided is a Python script for a program called "BelarcFindr" that allows users to search for specific information within HTML files and generate an Excel output file containing the extracted data. The script utilizes the Tkinter library to create a graphical user interface (GUI) for the application.

The main functionality of the script includes:
- Prompting the user to enter a search text.
- Verifying the search text and enabling the "Run" button only if the text is "SYM" (case-insensitive).
- Allowing the user to select an output folder where the generated Excel file will be saved.
- Running the search and file generation process when the "Run" button is clicked.

The script uses the `BeautifulSoup` and `openpyxl` libraries to parse the HTML content of the files and extract relevant information. The extracted data is then written to an Excel file.

To use the "BelarcFindr" application:
1. Launch the application by running the script.
2. The GUI window titled "BelarcFindr" will appear.
3. Enter the desired search text in the "Enter Search Text:" entry field. Only the text "SYM" (case-insensitive) is accepted.
4. Click the "Confirm" button to verify the search text. If the text is not "SYM," an error message will be displayed.
5. The "Run" button will be enabled if the search text is "SYM." Click the "Run" button to proceed with the search and file generation.
6. If no output folder path is specified, the default path will be set to the user's desktop.
7. A search will be performed on the HTML files located in the specified folder path.
8. The script will generate an Excel file named "output.xlsx" containing the extracted information.
9. The Excel file will be saved in the selected output folder.
10. A success message will be displayed upon completion of the search and file generation process.
11. The code assumes that there are HTML files located in the folder path specified within the script. The search and file generation process is limited to one file (`file_paths[:1]`) to demonstrate functionality. To process all HTML files within the folder path, modify the code accordingly.

### Solution:
The solution to the problem is to provide detailed documentation for the provided Python code. The documentation will cover the following aspects:

1. Code Structure and Dependencies:
   - The code depends on the following external libraries: `tkinter`, `openpyxl`, `bs4`, `pandas`.
   - The required imports are provided at the beginning of the code.

2. Class `Syst`:
   - The `Syst` class represents a system object and provides methods for managing and accessing files within a specified folder.
   - The class constructor initializes the `Syst` object with a folder path and sets up data structures to track files and their access times.
   - The class provides methods for adding a file, handling file clicks, and retrieving the next file to process.

3. Function `search_files`:
   - The `search_files` function takes a list of file paths, a search text, and an output folder path as input.
   - It processes each file, extracts relevant information using BeautifulSoup, and saves the data to an Excel file.
   - The function uses openpyxl to create and manipulate Excel workbooks.

4. Function `confirm_search_text`:
   - The `confirm_search_text` function is called when the user enters search text in the GUI.
   - It enables or disables the "Run" button based on whether the search text is entered or not.

5. Function `select_output_folder`:
   - The `select_output_folder` function is called when the user clicks the "Browse" button in the GUI.
   - It opens a file dialog to select an output folder and updates the output folder path in the GUI.

6. Function `run_search`:
   - The `run_search` function is called when the user clicks the "Run" button in the GUI.
   - It retrieves the search text and output folder path from the GUI, validates them, and executes the search process using the `search_files` function.
   - If successful, it displays a success message; otherwise, it displays an error message.

7. GUI Initialization:
   - The code initializes a GUI window using the tkinter library.
   - It creates and positions various widgets such as labels, entry fields, and buttons.
   - The GUI components are responsible for taking user inputs, such as search text and output folder path, and triggering the corresponding actions.

### Code:

```
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
```

### Outcome:
```
File: C:\Program Files (x86)\Belarc\BelarcAdvisor\System\tmp\SYMPC017_IT_VINODH_VSPITP_1F_D76.html
Excel file saved successfully!
```

After parsing the HTML file and extracting the relevant information, the script created an Excel file named "output.xlsx". The Excel file was saved successfully, indicating that the search and file generation process completed without errors.
