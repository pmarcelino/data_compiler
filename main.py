import os
import pandas as pd
from datetime import datetime


def read_excel_files(directory):
    # Initialize an empty DataFrame for the compiled data
    compiled_data = pd.DataFrame()
    log_entries = []

    # List of sheets to check
    sheets_to_check = [
        "castores_3_4",
        "benjamins_5_6",
        "cadetes_7_8",
        "juniores_9_10",
        "seniores_11_12",
    ]

    # Iterate over all Excel files in the specified directory
    for file in os.listdir(directory):
        if (
            file.endswith(".xlsx")
            or file.endswith(".xls")
            or file.endswith(".ods")
            or file.endswith(".csv")
        ):
            file_path = os.path.join(directory, file)
            file_name, _ = os.path.splitext(
                file
            )  # Separate the file name from its extension

            try:
                # Try reading the "info_escola" sheet
                school_info = pd.read_excel(file_path, sheet_name="info_escola")

                # Check if the "info_escola" DataFrame is empty
                if not school_info.empty and pd.notna(school_info.iloc[0, 1]):
                    school_name = school_info.iloc[0, 1]
                else:
                    log_entry = {
                        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "File": file,
                        "School Name": "",
                        "Error Type": "WARN",
                        "Message": "'info_escola' sheet is empty or Cell B2 is empty. Extracting school name from filename.",
                    }
                    log_entries.append(log_entry)
                    if "bebras_2023_papel_respostas" in file_name:
                        school_name = file_name.split("bebras_2023_papel_respostas")[
                            1
                        ].strip(" ()")
                    else:
                        school_name = file_name.strip(" ()")

                # Read and verify data from other sheets
                all_sheets_empty_or_missing = True
                for sheet in sheets_to_check:
                    try:
                        sheet_data = pd.read_excel(file_path, sheet_name=sheet)
                    except ValueError as e:
                        log_entry = {
                            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "File": file,
                            "School Name": school_name,
                            "Error Type": "INFO",
                            "Message": f"{sheet} is missing.",
                        }
                        log_entries.append(log_entry)
                        continue

                    # Check if the sheet is empty
                    if sheet_data.empty:
                        log_entry = {
                            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "File": file,
                            "School Name": school_name,
                            "Error Type": "INFO",
                            "Message": f"{sheet} is empty.",
                        }
                        log_entries.append(log_entry)
                    else:
                        all_sheets_empty_or_missing = False
                        # Check for "AnoEscolar" values
                        if (
                            "AnoEscolar" in sheet_data.columns
                            and not sheet_data["AnoEscolar"].between(3, 12).all()
                        ):
                            log_entry = {
                                "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "File": file,
                                "School Name": school_name,
                                "Error Type": "ERROR",
                                "Message": f"{sheet} contains 'AnoEscolar' value(s) outside the range 3-12.",
                            }
                            log_entries.append(log_entry)

                        sheet_data["Escola"] = school_name
                        cols = sheet_data.columns.tolist()
                        cols.insert(0, cols.pop(cols.index("Escola")))
                        sheet_data = sheet_data.reindex(columns=cols)
                        sheet_data = sheet_data.loc[
                            :, ~sheet_data.columns.str.contains("^Unnamed")
                        ]
                        compiled_data = pd.concat(
                            [compiled_data, sheet_data], ignore_index=True
                        )

                # Log an error if all sheets are empty or missing
                if all_sheets_empty_or_missing:
                    log_entry = {
                        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "File": file,
                        "School Name": school_name,
                        "Error Type": "ERROR",
                        "Message": "All sheets are empty or missing.",
                    }
                    log_entries.append(log_entry)

            except Exception as e:
                log_entry = {
                    "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "File": file,
                    "School Name": "",
                    "Error Type": "ERROR",
                    "Message": str(e),
                }
                log_entries.append(log_entry)

    return compiled_data, pd.DataFrame(log_entries)


# Replace with the actual directory path containing the Excel files
directory = "submissoes"

# Read and compile data from Excel files
compiled_data, log_data = read_excel_files(directory)

# Save the compiled data to a CSV file
compiled_data.to_csv("paper_prepare.csv", index=False)

# Save the log data to a CSV log file
log_data.to_csv("log_file.csv", index=False)

# Output to confirm completion
print("Data compilation and log file generation completed.")
