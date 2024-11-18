import os
import re
import shutil
import pandas as pd
import argparse


def extract_number_from_folder(folder_name):
    """Extract the 6-digit number from the folder name."""
    match = re.search(r"\b\d{6}\b", folder_name)
    return match.group() if match else None


def process_folder(folder_path, name_dict, to_folder):
    """Process all subfolders and files recursively."""
    print(f"Processing folder: {folder_path}")
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):  # If it's a folder, process it recursively
            print(f"Found subfolder: {item}")
            folder_number = extract_number_from_folder(item)
            if folder_number and folder_number in name_dict:
                process_files_in_folder(item_path, name_dict[folder_number], to_folder)
            else:
                process_folder(item_path, name_dict, to_folder)  # Recurse into the subfolder
        else:
            print(f"Skipping file at root level: {item_path}")


def process_files_in_folder(folder_path, folder_name, to_folder):
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


def main():
    parser = argparse.ArgumentParser(
        description="Extract and rename files based on an Excel mapping."
    )
    parser.add_argument(
        "--source_folders", nargs="+", required=True, help="Paths to source folders."
    )
    parser.add_argument(
        "--destination_folder", required=True, help="Path to the destination folder."
    )
    parser.add_argument(
        "--excel_file", required=True, help="Path to the Excel file for mapping."
    )

    args = parser.parse_args()

    # Validate input folders and Excel file
    from_folders = args.source_folders
    to_folder = args.destination_folder
    excel_path = args.excel_file

    if not all(os.path.isdir(folder) for folder in from_folders):
        print("Error: One or more source folders are invalid.")
        return

    if not os.path.isdir(to_folder):
        print("Error: Destination folder is invalid.")
        return

    if not os.path.isfile(excel_path):
        print("Error: Excel file is invalid.")
        return

    print(f"Source folders: {from_folders}")
    print(f"Destination folder: {to_folder}")
    print(f"Excel file: {excel_path}")

    # Load Excel file
    try:
        df = pd.read_excel(excel_path, usecols=[0, 1], header=0)
        name_dict = dict(zip(df.iloc[:, 0].astype(str), df.iloc[:, 1]))
        print(f"Loaded Excel mapping: {name_dict}")
    except Exception as e:
        print(f"Error: Failed to read Excel file: {e}")
        return

    print("Starting file extraction and renaming...")
    for root_folder in from_folders:
        process_folder(root_folder, name_dict, to_folder)
    print("File extraction and renaming complete!")


if __name__ == "__main__":
    main()
