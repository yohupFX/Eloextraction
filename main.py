import os
import re
import shutil
import pandas as pd
from tkinter import Tk, filedialog, messagebox

def main():
    def get_multiple_folders():
        """Prompt the user to select multiple folders."""
        folders = []
        while True:
            folder = filedialog.askdirectory(title="Select a folder to extract FROM")
            if folder:
                folders.append(folder)
            else:
                break
            if not messagebox.askyesno("Continue?", "Do you want to add another folder?"):
                break
        return folders

    def get_excel_file():
        """Prompt the user to select an Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select an Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        return file_path

    def extract_number_from_folder(folder_name):
        """Extract the 6-digit number from the folder name."""
        match = re.search(r"\b\d{6}\b", folder_name)
        return match.group() if match else None

    def process_folder(folder_path):
        """Process all subfolders and files recursively."""
        print(f"Processing folder: {folder_path}")
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            if os.path.isdir(item_path):  # If it's a folder, process it recursively
                print(f"Found subfolder: {item}")
                folder_number = extract_number_from_folder(item)
                if folder_number and folder_number in name_dict:
                    process_files_in_folder(item_path, name_dict[folder_number])
                else:
                    process_folder(item_path)  # Recurse into the subfolder
            else:
                print(f"Skipping file at root level: {item_path}")

    def process_files_in_folder(folder_path, folder_name):
        """Process all files in a specific folder."""
        print(f"Processing files in folder: {folder_path}")
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                new_name = f"{folder_name}_{file}"
                dst_path = os.path.join(to_folder, new_name)

                # Copy file and rename
                print(f"Copying file from {file_path} to {dst_path}")
                shutil.copy(file_path, dst_path)
                print(f"Copied: {file_path} -> {dst_path}")

    # Main script logic
    root = Tk()
    root.withdraw()

    print("Step 1: Selecting source folders")
    from_folders = get_multiple_folders()
    if not from_folders:
        messagebox.showerror("Error", "You must select at least one source folder.")
        return
    print(f"Selected folders: {from_folders}")

    print("Step 2: Selecting destination folder")
    to_folder = filedialog.askdirectory(title="Select a folder to extract TO")
    if not to_folder:
        messagebox.showerror("Error", "You must select a destination folder.")
        return
    print(f"Selected destination folder: {to_folder}")

    print("Step 3: Selecting Excel file")
    excel_path = get_excel_file()
    if not excel_path:
        print("No Excel file selected. Exiting...")
        return
    print(f"Selected Excel file: {excel_path}")

    print("Step 4: Loading Excel file")
    try:
        df = pd.read_excel(excel_path, usecols=[0, 1], header=0)
        name_dict = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
        print(f"Loaded Excel mapping: {name_dict}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        return

    print("Step 5: Starting file extraction and renaming...")
    for root_folder in from_folders:
        process_folder(root_folder)
    print("Step 6: File extraction and renaming complete!")
    messagebox.showinfo("Done", "Files have been extracted and renamed.")

if __name__ == "__main__":
    main()
