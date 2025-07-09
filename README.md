# Duplicate Eliminator for CRM

A simple desktop application with a graphical user interface to help you identify duplicate contacts before importing them into your CRM.
This tool compares a main CRM contact list (as a CSV file) against a file containing new records (either Excel or CSV) and separates the new contacts from those that are already in the CRM.

![screenshot](/screenshot.png)


## Key Features

- **Compares Files**: Compares a "master" CRM export against a "new records" file.
- **Flexible Input**: Accepts CSV for the CRM export and either **Excel (.xlsx/.xls) or CSV** for the new records.
- **Robust Matching**: Uses advanced name normalization to improve matching accuracy (converts to lowercase, removes accents and most punctuation).
- **Clean Outputs**: Generates two separate, ready-to-use CSV files:
    1.  `contacts_to_import.csv` (unique records not found in the CRM).
    2.  `duplicates_to_review.csv` (records that were already in the CRM).

## How to Use

1.  Run the `duplicate_eliminator_2.py` script.
2.  Click **Browse...** to select your master **CRM Export (CSV)** file.
3.  Click **Browse...** to select your **New Records (Excel/CSV)** file.
4.  (Optional) Click the **Options...** button to change the column names or delimiters to match your files. The current settings are always displayed on the main window.
5.  Click the **Process Files** button.
6.  The application will compare the files and log its progress. Once complete, the "Save" buttons will be enabled.
7.  Click **Save Unique Contacts (CSV)** to save a file with only the new contacts.
8.  Click **Save Duplicates for Review (CSV)** to save a file listing the contacts that were already found in your CRM for your reference.

## Installation
Use the compiled version in Relases. 
Otherwise, use python and the following dependencies: 

### Dependencies
This script requires a few Python libraries. You can install them all with a single command using pip:

```bash
pip install pandas openpyxl unidecode xlrd
```


