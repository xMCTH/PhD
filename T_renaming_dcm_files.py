"""
Script Name: rename_dcm_recursive.py

Description:
This script recursively processes folders containing DICOM (.dcm) files and a
corresponding text (.txt) file with new filenames.

Workflow:
1. The user is prompted to select a root directory via a folder selection dialog.
2. The script scans all subfolders inside the selected directory.
3. In each subfolder, the script checks for:
   - One or more .dcm files
   - Exactly one .txt file containing new names (one per line)
4. If both are present and the number of names in the .txt file matches the
   number of .dcm files:
   - The .dcm files are sorted (e.g. 0001.dcm, 0002.dcm, ...)
   - Each .dcm file is renamed according to the corresponding line in the .txt file
   - Renamed files are copied into a new subfolder named "new" within the same directory
5. If the conditions are not met, the folder is skipped safely.
6. After processing, the script prints a summary to the console listing:
   - Which folders were successfully processed
   - Which folders were skipped and why

Safety:
- Original DICOM files are never modified or overwritten.
- Renamed files are stored separately in a "new" folder.

Requirements:
- Python 3.x
- Standard libraries only (os, shutil, tkinter)

Intended Use:
- Batch renaming of DICOM files based on externally provided name lists
- Suitable for use in PyCharm and other Python IDEs
"""

import os
import shutil
import tkinter as tk
from tkinter import filedialog

# ---------- SELECT ROOT FOLDER ----------
root = tk.Tk()
root.withdraw()

root_folder = filedialog.askdirectory(
    title="Select root folder containing subfolders with DCM and TXT files"
)

if not root_folder:
    raise RuntimeError("No folder selected. Script aborted.")

processed_folders = []
skipped_folders = []

# ---------- WALK THROUGH SUBFOLDERS ----------
for current_path, _, files in os.walk(root_folder):

    dcm_files = sorted([f for f in files if f.lower().endswith(".dcm")])
    txt_files = [f for f in files if f.lower().endswith(".txt")]

    # Require at least one DCM and exactly one TXT
    if not dcm_files or len(txt_files) != 1:
        skipped_folders.append((current_path, "Missing DCM or TXT / multiple TXT"))
        continue

    txt_path = os.path.join(current_path, txt_files[0])

    # Read names
    with open(txt_path, "r", encoding="utf-8") as f:
        new_names = [line.strip() for line in f if line.strip()]

    # Check count match
    if len(dcm_files) != len(new_names):
        skipped_folders.append((current_path, "DCM/TXT count mismatch"))
        continue

    # Create output folder
    output_folder = os.path.join(current_path, "new")
    os.makedirs(output_folder, exist_ok=True)

    # Rename and copy
    for old_file, new_name in zip(dcm_files, new_names):
        old_path = os.path.join(current_path, old_file)
        new_filename = f"{new_name}.dcm"
        new_path = os.path.join(output_folder, new_filename)

        shutil.copy2(old_path, new_path)

    processed_folders.append(current_path)

# ---------- CONSOLE REPORT ----------
print("\n===== SCRIPT SUMMARY =====\n")

print(f"Processed folders ({len(processed_folders)}):")
for folder in processed_folders:
    print(f"  ✔ {folder}")

print(f"\nSkipped folders ({len(skipped_folders)}):")
for folder, reason in skipped_folders:
    print(f"  ✖ {folder} — {reason}")

print("\nProcessing complete.")
