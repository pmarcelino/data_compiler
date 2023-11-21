### README File for Excel Data Compilation Script

#### Overview
This Python script processes a series of Excel files in a specified directory, extracting and compiling data into a CSV file. It also generates a detailed log file in CSV format, capturing various information and error messages encountered during the data extraction process.

#### Setting Up the Environment
1. **Create a Virtual Environment** (Recommended):
   - Navigate to your project directory in the terminal.
   - Run `python -m venv venv` to create a virtual environment named `venv`.
   - Activate the virtual environment:
     - Windows: `.\venv\Scripts\activate`
     - macOS/Linux: `source venv/bin/activate`

2. **Install Required Libraries**:
   - Run `pip install -r requirements.txt` to install the required libraries.

#### Running the Script
- Place the script in your project directory.
- Ensure the Excel files to be processed are in a directory named `submissoes` in the same location as the script.
- Run the script using `python main.py`.

#### Log File Explanation
- The script generates a log file named `log_file.csv`.
- The log file contains entries for various events and errors encountered during processing, such as:
  - Missing or empty sheets in the Excel files.
  - Invalid data, such as out-of-range values in specific columns.
  - File read errors or other exceptions.
- Each entry in the log file includes a timestamp, file name, school name (if applicable), error type (INFO/WARN/ERROR), and a descriptive message.

#### Compiled Data File
- The compiled data is saved in a file named `paper_prepare.csv`.
- The data from different sheets of each Excel file is combined into this single CSV file, with appropriate headers and formatting.