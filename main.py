import os
import shutil
import pandas as pd
from report_generator import create_report
from email_sender import send_email

# ===============================
# Folder Configuration
# ===============================

INPUT_FOLDER = "input_files"
OUTPUT_FOLDER = "reports"
PROCESSED_FOLDER = "processed_files"

# Create folders automatically if not exist
for folder in [INPUT_FOLDER, OUTPUT_FOLDER, PROCESSED_FOLDER]:
    os.makedirs(folder, exist_ok=True)


# ===============================
# Get Files From Input Folder
# ===============================

def get_new_files():
    try:
        files = os.listdir(INPUT_FOLDER)
        # filter only Excel files
        excel_files = [f for f in files if f.endswith(".xlsx")]
        return excel_files
    except Exception as e:
        print("Error reading input folder:", e)
        return []


# ===============================
# Process Each File
# ===============================

def process_file(file_name):

    file_path = os.path.join(INPUT_FOLDER, file_name)

    if not os.path.exists(file_path):
        print("File not found:", file_path)
        return

    try:

        print(f"Processing file: {file_name}")

        # Read Excel file
        df = pd.read_excel(file_path)

        if "Amount" not in df.columns:
            print(f"'Amount' column missing in {file_name}")
            return

        # Calculate summary
        total = df["Amount"].sum()
        average = df["Amount"].mean()

        print("Total:", total)
        print("Average:", average)

        # Generate PDF report
        report_path = create_report(file_name, total, average)

        print("Report generated:", report_path)

        # Send email
        send_email(report_path)

        print("Email sent successfully")

        # Copy file to processed folder (NOT move)
        destination = os.path.join(PROCESSED_FOLDER, file_name)
        shutil.copy(file_path, destination)

        print(f"{file_name} copied to processed_files")

    except Exception as e:
        print(f"Error processing {file_name}:", e)


# ===============================
# Main Runner
# ===============================

def run_system():

    print("Automation started")

    files = get_new_files()

    if not files:
        print("No Excel files found in input folder.")
        return

    for file in files:
        process_file(file)

    print("Automation completed")


# ===============================
# Start Program
# ===============================

if __name__ == "__main__":
    run_system()