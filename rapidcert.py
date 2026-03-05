# Alaways open an excel file before to suppress the activation wizerd of excel when the script is run for the first time.
import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import shutil
import os
import datetime

def get_certificate_files():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(title="Select Certificate Files", filetypes=[("Excel files", "*.xlsx *.xls")])
    return list(file_paths)

# A function to read excel file and extract values of different cells like R72, E72 etc.
def extract_certificate_data(file_path):
    workbook = xw.Book(file_path)
    worksheet = workbook.sheets[0]
    data = {
        "ID": worksheet['M10'].value,
        "CalDate": worksheet['E72'].value,
        "CalDueDate": worksheet['R72'].value,
    }
    workbook.close()
    return data

# A function to update the certificate with proper certificate number
def update_certificate(file_path, cert_no):
    workbook = xw.Book(file_path)
    worksheet = workbook.sheets[0]
    worksheet['L2'].value = cert_no
    workbook.save()
    workbook.close()

# A function to update the tracker file with the extracted data
def update_tracker(workbook, data):
    # active sheet of the workbook
    worksheet = workbook.sheets.active

    last_row = worksheet.cells.last_cell.row
    for index, row in enumerate(worksheet.range(f"G1:G{last_row}").rows):
        if row[0].value == data["ID"]:
            if worksheet[f"M{index+1}"].value == "YES" or worksheet[f"K{index+1}"].value == "YES":
                if worksheet[f"N{index+1}"].value.endswith("2026"):
                    return "Feature of second version is not implemented yet"
                else:
                    cal_date = datetime.datetime.strptime(data["CalDate"], "%d.%m.%Y")
                    worksheet[f"I{index+1}"].value = cal_date
                    depts = {"09-": "VBS", "08-": "BBS", "06-": "AHS", "01-": "PRD"}
                    data = f"CAL/{depts.get(data["ID"][:3])}/IH/{int(worksheet[f"A{index+1}"].value):03d}/01-2026"
                    worksheet[f"N{index+1}"].value = data
                    return data
            else:
                return "Equipment is already calibrated. Please check the tracker file."
    else:
        return "ID Not Found"

if __name__ == "__main__":
    trackerPath = {"09-": "\\\\192.168.5.6\\IPL-VLD- Master Plan\\5. Validation Activities_Vaccine\\3. Viral Bulk Suite\\1. VALIDATION & CALIBRATION TRACKER AT VBS\\1. Calibration\\2026\\Calibration Tracker (VBS).xlsx",
                   "08-": "\\\\192.168.5.6\\IPL-VLD- Master Plan\\5. Validation Activities_Vaccine\\2. Bacterial Bulk Suite\\1. VALIDATION & CALIBRATION TRACKER OF BBS\\1. Calibration\\2026\\Calibration Tracker (BBS).xlsx",
                   "06-": "\\\\192.168.5.6\\IPL-VLD- Master Plan\\5. Validation Activities_Vaccine\\7. Animal house\\1. VALIDATION & CALIBRATION TRACKER\\1. Calibration\\2026\\Calibration Tracker (Animal House).xlsx",
                   "01-": "\\\\192.168.5.6\\IPL-VLD- Master Plan\\5. Validation Activities_Vaccine\\1. Production\\1. VALIDATION & CALIBRATION TRACKERS-VACCINE\\1. Calibration tracker\\2026\\Calibration Tracker (Production).xlsx"}
    workbooks = {}
    cert_files = get_certificate_files()
    log_file = open("rapidcert.log", "a")
    print("ID\t\t\tCalibration Date\tCalibration Due Date\tCertificate No.")
    log_file.write("ID\t\t\tCalibration Date\tCalibration Due Date\tCertificate No.\n")
    for i in cert_files:
        data = extract_certificate_data(i)
        workbook = None
        for key, value in workbooks.items():
            if data["ID"].startswith(key):
                workbook = value
                break
        else:
            for key, value in trackerPath.items():
                if data["ID"].startswith(key):
                    # Create a backup copy of the tracker file with timestamp in a folder named backup in the same directory as the tracker file
                    backup_folder = os.path.join(os.path.dirname(value), "backup")
                    if not os.path.exists(backup_folder):
                        os.makedirs(backup_folder)
                    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                    backup_file = os.path.join(backup_folder, f"Calibration Tracker Backup {timestamp}.xlsx")
                    shutil.copy2(value, backup_file)
                    workbook = xw.Book(value)
                    workbooks[key] = workbook
                    break
            else:
                print(f"No matching workbook found for ID: {data['ID']}")
                continue
        certNo = update_tracker(workbook, data)
        if certNo.startswith("CAL/"):
            update_certificate(i, certNo)
        print(f"{data['ID']}\t{data['CalDate']}\t\t{data['CalDueDate']}\t\t{certNo}")
        # Also log in a log file in current working directory with name rapidcert_log.txt with the same information as printed in console
        log_file.write(f"{data['ID']}\t{data['CalDate']}\t\t{data['CalDueDate']}\t\t{certNo}\n")
    for workbook in workbooks.values():
        workbook.save()
        workbook.close()
    log_file.close()
