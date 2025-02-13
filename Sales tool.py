import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime

# Default source folder (Change this path to your actual default folder)
DEFAULT_SOURCE_FOLDER = r"C:\Users\JHalv\OneDrive\Code\Test Sales Tool\Test"

LOG_FILE = "opportunity_log.xlsx"  # Log file name

def log_entry(name, folder_path):
    """Logs the created folder into an Excel file"""
    log_data = {
        "Opportunity Name": [name],
        "Date Created": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        "Folder Path": [folder_path]
    }

    df_new = pd.DataFrame(log_data)

    if os.path.exists(LOG_FILE):
        df_existing = pd.read_excel(LOG_FILE)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new

    df_combined.to_excel(LOG_FILE, index=False)
    print(f"Logged: {name} -> {folder_path}")

def copy_and_rename(names, source_folder, destination_parent):
    if not os.path.exists(source_folder):
        messagebox.showerror("Error", f"Source folder does not exist: {source_folder}")
        return
    
    for name in names:
        name = name.strip()
        if not name:
            continue
        
        new_folder = os.path.join(destination_parent, name)

        try:
            shutil.copytree(source_folder, new_folder)

            # Rename files inside the copied folder
            for file_name in os.listdir(new_folder):
                old_path = os.path.join(new_folder, file_name)
                if os.path.isfile(old_path):
                    base, ext = os.path.splitext(file_name)
                    new_name = f"{name} {base}{ext}"
                    new_path = os.path.join(new_folder, new_name)
                    os.rename(old_path, new_path)

            log_entry(name, new_folder)
            print(f"Successfully created: {new_folder}")
        except FileExistsError:
            messagebox.showwarning("Warning", f"Folder '{name}' already exists! Skipping...")
        except Exception as e:
            print(f"Error processing {name}: {e}")

    messagebox.showinfo("Success", "Folders and files copied, renamed, and logged successfully!")

def select_destination_folder():
    folder_selected = filedialog.askdirectory()
    dest_entry.delete(0, tk.END)
    dest_entry.insert(0, folder_selected)

def start_process():
    names = name_entry.get().split(',')
    source_folder = DEFAULT_SOURCE_FOLDER  # Uses the default folder
    destination_parent = dest_entry.get().strip()

    if not names or not destination_parent:
        messagebox.showerror("Error", "Please enter names and select a destination folder!")
        return

    copy_and_rename(names, source_folder, destination_parent)

# GUI Setup
root = tk.Tk()
root.title("Folder Copy & Rename Tool")

tk.Label(root, text="Enter Names (comma-separated):").grid(row=0, column=0, padx=10, pady=5)
name_entry = tk.Entry(root, width=50)
name_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Source Folder:").grid(row=1, column=0, padx=10, pady=5)
tk.Label(root, text=DEFAULT_SOURCE_FOLDER, fg="blue").grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Select Destination Folder:").grid(row=2, column=0, padx=10, pady=5)
dest_entry = tk.Entry(root, width=40)
dest_entry.grid(row=2, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=select_destination_folder).grid(row=2, column=2, padx=5, pady=5)

tk.Button(root, text="Start", command=start_process, bg="green", fg="white").grid(row=3, column=0, columnspan=3, pady=10)

root.mainloop()
