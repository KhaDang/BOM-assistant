import os
import shutil
import pandas as pd
from collections import defaultdict
import openpyxl


# -----------------------------
# Configuration
# -----------------------------

excel_file = r"D:\Users\khda\Desktop\MissingDrw.xlsx"

search_folder = r"D:\AME Vault"

destination_folder = os.path.join(
    os.path.expanduser("~"),
    "Desktop",
    "Copied_Drawings"
)

extension = ".SLDDRW"
drw_col = "Bruker PN"

# -----------------------------
# Create destination
# -----------------------------

os.makedirs(destination_folder, exist_ok=True)

# -----------------------------
# Read Excel
# -----------------------------

df = pd.read_excel(excel_file)

drawing_list = df[drw_col].astype(str).str.strip()

# -----------------------------
# Build file index
# (Much faster than searching repeatedly)
# -----------------------------

print("Indexing files...")

file_index = defaultdict(list)

for root, dirs, files in os.walk(search_folder):
    for file in files:

        if file.lower().endswith(extension.lower()):

            fullpath = os.path.join(root, file)
            file_index[file.lower()].append(fullpath)

print(f"{len(file_index)} unique drawing names.")

# -----------------------------
# Copy files
# -----------------------------

duplicate_files = []
missing_files = []

for drawing in drawing_list:

    filename = drawing + extension
    key = filename.lower()

    if key not in file_index:

        print(f"Missing : {filename}")
        missing_files.append(filename)

        continue

    locations = file_index[key]

    # Only one file found
    if len(locations) == 1:

        shutil.copy2(
            locations[0],
            os.path.join(destination_folder, filename)
        )

        print(f"Copied : {filename}")

    # Duplicate files found
    else:

        duplicate_files.append({
            "Drawing": filename,
            "Count": len(locations),
            "Locations": "\n".join(locations)
        })

        print(f"Duplicate : {filename}")
# -----------------------------
# Save duplicate report
# -----------------------------
if duplicate_files:

    pd.DataFrame(duplicate_files).to_excel(
        os.path.join(destination_folder, "DuplicateFiles.xlsx"),
        index=False
    )

# -----------------------------
# Save missing list
# -----------------------------
if missing_files:

    pd.DataFrame({
        "Missing": missing_files
    }).to_excel(
        os.path.join(destination_folder,
                     "MissingFiles.xlsx"),
        index=False
    )

print("Finished.")