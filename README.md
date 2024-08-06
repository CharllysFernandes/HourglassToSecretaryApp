# Hourglass Data Export and Processing

This project includes Python scripts to export data from the Hourglass app, process it, and create Excel files for importing into the Secretary app. Follow the steps below to execute the scripts and generate the required Excel files.

## Prerequisites

1. **Python**: Ensure Python 3.x is installed on your system.
2. **Python Packages**: You need `pandas` and `openpyxl` libraries. Install them using pip:

    ```bash
    pip install pandas openpyxl
    ```

3. **Data Files**: Make sure you have the following JSON files in the same directory:
    - `dados.json` (Contains data for both publishers and reports)

## Scripts Overview

There are two main scripts provided:
1. **`export_publishers.py`**: Extracts publisher data from `dados.json` and generates an Excel file.
2. **`export_reports.py`**: Extracts report data from `dados.json` and generates an Excel file.

## Steps to Run the Scripts

### 1. Prepare Your Directory

Ensure the following files are in the same directory:
- `dados.json` (Your source data file)
- `export_publishers.py` (Script for processing publisher data)
- `export_reports.py` (Script for processing report data)

### 2. Run `export_publishers.py`

This script will process the publisher data and create an Excel file named `publishers_data.xlsx`.

```bash
python export_publishers.py
```

### 3. Run `export_publishers.py`

This script will process the publisher data and create an Excel file named `publishers_data.xlsx`.

```bash
python export_publishers.py
```

### 4. Run `export_reports.py`

This script will process the report data and create an Excel file named `reports_data.xlsx`.

```bash
python export_reports.py
```


### 5. Import into Secretary

After running both scripts, you will have two Excel files:
- `publishers_data.xlsx`
- `reports_data.xlsx`

These files can now be imported into the Secretary app.

## Script Details

### `export_publishers.py`

This script processes publisher data from `dados.json` and creates an Excel file with the following columns:
- A: Empty
- B: `firstname`
- C: `middlename`
- D: `lastname`
- E: Empty
- F: Status values (2 for "Regular Pioneer", 1 for "Continuous Auxiliary Pioneer")
- G: Appointment type (`1` for "MS", `0` otherwise)
- H: Appointment type (`1` for "Elder", `0` otherwise)
- I: Empty
- J: Empty
- K: Empty
- L: Birthdate in "DD/MM/YYYY" format
- M: Baptism date in "DD/MM/YYYY" format
- N: Gender (`1` for "Male", `0` otherwise)
- O: Cellphone numbers extracted
- P-Q-S-U-V-W: Empty
- R: Email

### `export_reports.py`

This script processes report data from `dados.json` and creates an Excel file with the following columns:
- A: Full name (`firstname middlename lastname`)
- B: Year
- C: Month
- D: Placements (0 if null)
- E: Videoshowings (0 if null)
- F: Minutes divided by 60
- G: Empty
- H: Credithours (0 if null)
- I: Return visits (0 if null)
- J: Studies (0 if null)
- K: Pioneer status (2 for "Regular", 1 for "Auxiliary", 0 if null)
- M: Remarks (empty if null)

## Troubleshooting

- Ensure all JSON data is correctly formatted.
- Verify that the file paths and names are correct.
- If you encounter any issues, check the error messages for clues.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
