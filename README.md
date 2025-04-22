# Excel Supplier Summary Extractor

This is a Python desktop application that reads an Excel workbook with multiple sheets and extracts supplier names and total amounts from each sheet. The extracted data is compiled into a summary Excel file.

## Features

- Automatically detects and extracts supplier names from each sheet.
- Extracts total values from the last row of each sheet.
- Ignores empty rows and columns.
- User-friendly file selection dialogs using a simple GUI.
- Saves the summary to a new Excel file.

## How It Works

1. The user selects an input Excel file (`.xlsx` format).
2. The script processes each sheet in the workbook:
   - Finds the row containing the "Supplier Code".
   - Extracts the supplier name.
   - Gets the total from the last row with data.
3. The user is prompted to save the summarized output.
4. The output is saved as a new Excel file.

## Requirements

- Python 3.7+
- `tkinter` (usually included with Python)
- The following Python packages (see `requirements.txt`):
  - `pandas`
  - `openpyxl`

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/excel-supplier-summary.git
   cd excel-supplier-summary
2. Install dependencies:
   ```bash
   pip install -r requirements.txt

## Usage

Simply run the script:
```bash
python main.py
```
- A file dialog will prompt you to select the source Excel file.
- After processing, another dialog will prompt you to choose a save location for the summary.

## Example

### Sample Input Sheet (Excel)

|   | A               | B        |
|---|-----------------|----------|
| 1 | Supplier Code:  | SUP123   |
| 2 | Item            | Amount   |
| 3 | Widget A        | 5000     |
| 4 | Widget B        | 7345.67  |
| 5 | Grand Total     | 12345.67 |

### Output Summary

| Supplier Name | Total    |
|---------------|----------|
| SUP123        | 12345.67 |


   
