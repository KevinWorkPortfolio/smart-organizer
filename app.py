import os
import shutil
import csv
from datetime import datetime

# This command define the script's root directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# This command acts as a "bridge" between where the script is located and the directories
# Using os.path.join makes the code compatible between different operating systems
CARTELLA_INPUT = os.path.join(BASE_DIR, "FileToSort")
CARTELLA_OUTPUT = os.path.join(BASE_DIR, "OutputData")
REPORT_FILE = os.path.join(BASE_DIR, "report_movements.csv")

# I check for the existence of directories and create them if they are not present.
# The "exist_ok=True" parameter prevents the program from crashing if the folder already exists.
os.makedirs(CARTELLA_OUTPUT, exist_ok=True)
os.makedirs(CARTELLA_INPUT, exist_ok=True)

# This command splits files based on their operational function:
# Reports: Word processing, reports, and presentations.
# Financial Statements: Data Analysis, Spreadsheets, and Reporting.
CATEGORIE = {
    "Reports": ["pdf", "doc", "docx", "txt", "ppt", "pptx"],
    "Financial_Statements": ["xls", "xlsx", "csv"]
}

# With this command I go to do a "reverse engineering" from "category to extension list" to "extension to category"
estensione_to_categoria = {
    ext: cat for cat, exts in CATEGORIE.items() for ext in exts
}

def get_creation_date(file_path):
    stat = os.stat(file_path)
    try:
        return datetime.fromtimestamp(stat.st_ctime)
    except:
        return datetime.fromtimestamp(stat.st_mtime)

# --- PROCESS STARTS ---

# I initialize the activity log in CSV (Sorting Report) format.
# Overwrite any previous reports (mode="w") and define the column header
# to ensure complete traceability of each moved file.
with open(REPORT_FILE, mode="w", newline="", encoding="utf-8") as report:
    writer = csv.writer(report)
    writer.writerow(["File Name", "Original Path", "New Path", "Category", "Year", "Month"])
    
# This command checks if the folder exists
    if not os.path.exists(CARTELLA_INPUT):
          print(f"ERRORE: The directory {CARTELLA_INPUT} doesn't exist!")
    else:
          elementi = os.listdir(CARTELLA_INPUT)
          print(f"DEBUG: Found {len(elementi)} elements in '{CARTELLA_INPUT}'")

          for file_nome in elementi:
            percorso_completo_input = os.path.join(CARTELLA_INPUT, file_nome)

            # Here the "discard" occurs from the search for all files that are folders
            if not os.path.isfile(percorso_completo_input):
                print(f"DEBUG: Discard {file_nome} (It is a folder)")
                continue

            # Here the extension is analyzed
            parts = file_nome.split(".")
            if len(parts) < 2:
                print(f"DEBUG: Discard {file_nome} (extension not found)")
                continue
            
            estensione = parts[-1].lower()
            if estensione not in estensione_to_categoria:
                print(f"DEBUG: Discard {file_nome} (estensione .{estensione} is not supported)")
                continue

            # I extract the parameters needed for cataloging
            # I identify the target category by reverse index.
            # Retrieve the source date to organize the files into history subfolders.
            categoria = estensione_to_categoria[estensione]
            data = get_creation_date(percorso_completo_input)
            anno = str(data.year)
            mese = f"{data.month:02d}"

            destinazione_cartella = os.path.join(CARTELLA_OUTPUT, categoria, anno, mese)
            os.makedirs(destinazione_cartella, exist_ok=True)

            destinazione_finale = os.path.join(destinazione_cartella, file_nome)

            # This command prevents duplicates from being placed between sorted files
            if os.path.exists(destinazione_finale):
                base, ext = os.path.splitext(file_nome)
                count = 1
                while os.path.exists(destinazione_finale):
                    destinazione_finale = os.path.join(destinazione_cartella, f"{base}_{count}{ext}")
                    count += 1

            # This part of the process allows you to move files to the output folder
            try:
                shutil.move(percorso_completo_input, destinazione_finale)
                print(f"OK: Moved {file_nome} -> {categoria}")
                
                # Here the report is written in csv format without being re-sorted infinitely many times in the process
                writer.writerow([
                    file_nome,
                    os.path.abspath(percorso_completo_input),
                    os.path.abspath(destinazione_finale),
                    categoria,
                    anno,
                    mese
                ])
            except Exception as e:
                print(f"ERROR while moving {file_nome}: {e}")

print("\n--- COMPLETED PROCESS ---")
print(f"Check the directory '{CARTELLA_OUTPUT}' for your sorted files.")

# Automatic report opening
try:
    os.startfile(REPORT_FILE)
except:
    print("Note: Could not open your .csv file")