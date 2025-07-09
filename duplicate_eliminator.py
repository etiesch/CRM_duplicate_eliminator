import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Toplevel, Frame, Label, Entry, Button, StringVar
from tkinter import font as tkFont 
import csv
import os
import pandas as pd
import re
from unidecode import unidecode 


# --- Configuration Variables ---
CONFIG = {
    'CRM_DELIMITER': ';',
    'CRM_LAST_NAME_COL': 'Nom',
    'CRM_FIRST_NAME_COL': 'Prénom',
    'NEW_RECORDS_NAME_COL': 'Nom',
    'NEW_RECORDS_FORENAME_COL': 'Prénom',
    'NEW_RECORDS_CSV_DELIMITER': ','
}


# --- Other variables ---

APP_TITLE = "CRM Duplicate Eliminator"

processed_uniques = []
processed_duplicates = []
new_records_header_global = [] # Store the header of the new records file

crm_delimiter_sv = None
crm_last_name_sv = None
crm_first_name_sv = None
new_records_name_sv = None
new_records_forename_sv = None
new_records_csv_delimiter_sv = None


# --- Core Logic ---

def normalize_name_part(name_part):
    # Normalize a name to allow robust comparison
    if pd.isna(name_part) or not str(name_part).strip():
        return ""
    
    s = str(name_part).strip()
    
    # Transliterate ("René" -> "Rene", "Müller" -> "Muller")
    s = unidecode(s)
    
    # Convert to lowercase for consistent comparison
    s = s.lower()
    
    # Remove all periods, commas, apostrophes, etc
    s = re.sub(r'[^a-z0-9\s-]', '', s)
    
    # 4. Normalize minus uses ("Smith - Jones" -> "smith-jones")
    s = re.sub(r'\s*-\s*', '-', s)
    
    # 5. Normalize spaces
    s = re.sub(r'\s+', ' ', s)
    
    return s.strip() # Final strip

def create_unique_id(last_name, first_name):
    # Creates a standardized unique ID from name and forename
    ln_normalized = normalize_name_part(last_name)
    fn_normalized = normalize_name_part(first_name)

    if ln_normalized and fn_normalized:
        return f"{ln_normalized} {fn_normalized}"
    elif ln_normalized: # Only last name is valid/present
        return ln_normalized
    elif fn_normalized: # Only first name is valid/present
        return fn_normalized
    else:
        return ""

def read_crm_csv_file(filepath, delimiter, expected_lastname_col, expected_firstname_col, status_widget):
    records = []
    unique_ids = set() # Allows to separate entries 
    header = []
    status_widget.insert(tk.END, f"Attempting to read CRM file: {os.path.basename(filepath)}\n")
    status_widget.insert(tk.END, f"  Using delimiter: '{delimiter}'\n")
    status_widget.insert(tk.END, f"  Using Last Name column: '{expected_lastname_col}'\n")
    status_widget.insert(tk.END, f"  Using First Name column: '{expected_firstname_col}'\n")

    try:
        with open(filepath, mode='r', encoding='utf-8-sig') as file:
            reader = csv.reader(file, delimiter=delimiter)
            header = next(reader, None)
            if not header:
                return None, None, None, f"Error: CRM file '{os.path.basename(filepath)}' is empty or has no header."

            normalized_header = [h.strip() for h in header]
            try:
                lastname_col_idx = normalized_header.index(expected_lastname_col)
                firstname_col_idx = normalized_header.index(expected_firstname_col)
            except ValueError:
                err_msg = (f"Error: Required columns ('{expected_lastname_col}', '{expected_firstname_col}') not found in CRM file '{os.path.basename(filepath)}'.\n"
                           f"Found headers: {', '.join(header)}")
                return None, None, None, err_msg

            for i, row in enumerate(reader):
                if len(row) != len(header):
                    status_widget.insert(tk.END, f"Warning: Row {i+2} in CRM file has incorrect columns ({len(row)} vs {len(header)} expected). Skipping.\n")
                    continue
                try:
                    record_dict = dict(zip(header, row))
                    last_name = row[lastname_col_idx]
                    first_name = row[firstname_col_idx]
                    
                    unique_id = create_unique_id(last_name, first_name)
                    if unique_id: # Only add if a valid ID could be generated
                        records.append(record_dict)
                        unique_ids.add(unique_id)
                    else:
                        status_widget.insert(tk.END, f"Warning: Row {i+2} in CRM file has empty/invalid name/forename after normalization. Skipping.\n")
                        
                except Exception as e:
                    status_widget.insert(tk.END, f"Warning: Error processing row {i+2} in CRM file: {e}. Skipping.\n")
        
        status_widget.insert(tk.END, f"Successfully processed {len(records)} records from CRM file ({len(unique_ids)} unique IDs generated).\n")
        return records, unique_ids, header, None
    except FileNotFoundError:
        return None, None, None, f"Error: CRM File not found at '{filepath}'."
    except Exception as e:
        return None, None, None, f"An unexpected error occurred while reading CRM file: {e}"

def read_new_records_file(filepath, expected_name_col, expected_forename_col, csv_delimiter, status_widget):
    all_records = []
    final_header = None
    filename = os.path.basename(filepath)
    file_ext = os.path.splitext(filename)[1].lower()

    status_widget.insert(tk.END, f"Attempting to read New Records file: {filename}\n")
    status_widget.insert(tk.END, f"  Using Name column: '{expected_name_col}'\n")
    status_widget.insert(tk.END, f"  Using Forename column: '{expected_forename_col}'\n")

    try:
        if file_ext in ['.xlsx', '.xls']:
            status_widget.insert(tk.END, f"  Reading as Excel file (will process all sheets)...\n")
            try:
                excel_data = pd.read_excel(filepath, sheet_name=None, engine=None, dtype=str) # Read all as string at firts
            except Exception as e_pandas_read:
                status_widget.insert(tk.END, f"Pandas read_excel (all sheets) failed: {e_pandas_read}.\n")
                engine_to_try = 'openpyxl' if file_ext == '.xlsx' else 'xlrd' if file_ext == '.xls' else None
                if engine_to_try:
                    try:
                        status_widget.insert(tk.END, f"  Retrying with engine: {engine_to_try}...\n")
                        excel_data = pd.read_excel(filepath, sheet_name=None, engine=engine_to_try, dtype=str)
                    except Exception as e_engine_retry:
                        return None, None, f"Error reading Excel '{filename}' with {engine_to_try}: {e_engine_retry}."
                else:
                    return None, None, f"Error reading Excel file '{filename}': {e_pandas_read}."
            
            if not excel_data:
                return None, None, f"Error: Excel file '{filename}' is empty or no sheets could be read."

            for sheet_name, df in excel_data.items():
                status_widget.insert(tk.END, f"  Processing sheet: '{sheet_name}'...\n")
                if df.empty:
                    status_widget.insert(tk.END, f"    Sheet '{sheet_name}' is empty. Skipping.\n")
                    continue
                
                df = df.fillna('') # Deals with annoying NaNs (--> empty strings)
                current_sheet_header_orig = list(df.columns) 
                current_sheet_header_norm = [str(h).strip() for h in current_sheet_header_orig]
                
                actual_sheet_name_col = expected_name_col
                actual_sheet_forename_col = expected_forename_col

                if expected_name_col not in current_sheet_header_norm or expected_forename_col not in current_sheet_header_norm:
                    norm_expected_name = expected_name_col.lower()
                    norm_expected_forename = expected_forename_col.lower()
                    found_name = False
                    found_forename = False
                    
                    temp_actual_name_col = None
                    temp_actual_forename_col = None

                    for idx, h_norm in enumerate(current_sheet_header_norm):
                        if not found_name and h_norm.lower() == norm_expected_name:
                            temp_actual_name_col = current_sheet_header_orig[idx]
                            found_name = True
                        if not found_forename and h_norm.lower() == norm_expected_forename:
                            temp_actual_forename_col = current_sheet_header_orig[idx]
                            found_forename = True
                        if found_name and found_forename:
                            break
                    
                    if found_name and found_forename:
                        actual_sheet_name_col = temp_actual_name_col
                        actual_sheet_forename_col = temp_actual_forename_col
                        status_widget.insert(tk.END, f"    Note: Found '{expected_name_col}' (as '{actual_sheet_name_col}') and '{expected_forename_col}' (as '{actual_sheet_forename_col}') case-insensitively in sheet '{sheet_name}'.\n")
                    else:
                        status_widget.insert(tk.END, f"    Warning: Sheet '{sheet_name}' missing required columns. Skipping.\n")
                        status_widget.insert(tk.END, f"    Expected: '{expected_name_col}', '{expected_forename_col}'. Found: {', '.join(current_sheet_header_orig)}\n")
                        continue
                
                if final_header is None: # Use header from the first sheet that has the required columns
                    final_header = [str(h) for h in current_sheet_header_orig]
                    status_widget.insert(tk.END, f"    Using header from sheet '{sheet_name}' for output: {', '.join(final_header)}\n")

                for i, row_series in df.iterrows():
                    record_dict = {str(k): str(v) for k, v in row_series.to_dict().items()} # Ensure all values are strings
                    name = record_dict.get(str(actual_sheet_name_col), "")
                    forename = record_dict.get(str(actual_sheet_forename_col), "")
                    
                    # Check for empty name/forename before creating ID (create_unique_id handles internal emptiness)
                    if not (str(name).strip() or str(forename).strip()):
                        status_widget.insert(tk.END, f"    Warning: Row {i+2} in sheet '{sheet_name}' has empty name/forename fields. Skipping.\n")
                        continue
                    all_records.append(record_dict)
            
            if not all_records and final_header is None: 
                 return None, None, f"Error: No sheet in '{filename}' contained the required columns ('{expected_name_col}', '{expected_forename_col}')."

        elif file_ext == '.csv':
            status_widget.insert(tk.END, f"  Reading as CSV file with delimiter '{csv_delimiter}'...\n")
            with open(filepath, mode='r', encoding='utf-8-sig') as file:
                reader = csv.reader(file, delimiter=csv_delimiter)
                header_csv_original = next(reader, None)
                if not header_csv_original:
                    return None, None, f"Error: New Records CSV file '{filename}' is empty or has no header."
                
                final_header = [str(h) for h in header_csv_original]
                normalized_csv_header = [h.strip() for h in final_header]
                
                name_col_idx, forename_col_idx = -1, -1
                try: # Exact match first
                    name_col_idx = normalized_csv_header.index(expected_name_col)
                    forename_col_idx = normalized_csv_header.index(expected_forename_col)
                except ValueError: # Case-insensitive fallback
                    norm_expected_name = expected_name_col.lower()
                    norm_expected_forename = expected_forename_col.lower()
                    for idx, h_norm in enumerate(normalized_csv_header):
                        if name_col_idx == -1 and h_norm.lower() == norm_expected_name:
                            name_col_idx = idx
                        if forename_col_idx == -1 and h_norm.lower() == norm_expected_forename:
                            forename_col_idx = idx
                        if name_col_idx != -1 and forename_col_idx != -1:
                            status_widget.insert(tk.END, f"    Note: Used case-insensitive matching for Name/Forename columns in '{filename}'.\n")
                            break
                    if name_col_idx == -1 or forename_col_idx == -1:
                        err_msg = (f"Error: Required columns not found in New Records CSV '{filename}'.\n"
                                   f"Expected: '{expected_name_col}', '{expected_forename_col}'. Found: {', '.join(final_header)}")
                        return None, None, err_msg

                for i, row_list in enumerate(reader):
                    if len(row_list) != len(final_header):
                        status_widget.insert(tk.END, f"    Warning: Row {i+2} in CSV has incorrect columns ({len(row_list)} vs {len(final_header)} expected). Skipping.\n")
                        continue
                    record_dict = dict(zip(final_header, row_list))
                    name = row_list[name_col_idx]
                    forename = row_list[forename_col_idx]
                    if not (str(name).strip() or str(forename).strip()):
                        status_widget.insert(tk.END, f"    Warning: Row {i+2} in CSV has empty name/forename fields. Skipping.\n")
                        continue
                    all_records.append(record_dict)
        else:
            return None, None, f"Error: Unsupported file type for New Records: '{file_ext}'."
        
        if not all_records and final_header:
             status_widget.insert(tk.END, f"  Processed file '{filename}', found header but no valid data rows.\n")
        elif all_records:
            status_widget.insert(tk.END, f"  Successfully processed {len(all_records)} records from New Records file '{filename}'.\n")
        
        if not final_header and not all_records: # This case means no sheets had headers/data or CSV was empty
            return None, None, f"Error: Could not determine a header or find any records in '{filename}' that meet criteria."
        
        return all_records, final_header, None
        
    except FileNotFoundError:
        return None, None, f"Error: New Records file not found at '{filepath}'."
    except Exception as e:
        return None, None, f"An unexpected error occurred while reading New Records file '{filename}': {e}"

def process_files():
    global processed_uniques, processed_duplicates, new_records_header_global
    processed_uniques = []
    processed_duplicates = []
    new_records_header_global = []

    crm_filepath = crm_file_entry.get()
    new_records_filepath = new_records_file_entry.get()

    status_text.config(state=tk.NORMAL)
    status_text.delete('1.0', tk.END)

    if not crm_filepath or not new_records_filepath:
        status_text.insert(tk.END, "Error: Both CRM export file and New Records file must be selected.\n")
        status_text.config(state=tk.DISABLED)
        enable_save_buttons(False)
        return

    # --- Use current CONFIG values ---
    current_crm_delimiter = CONFIG['CRM_DELIMITER']
    current_crm_last_name = CONFIG['CRM_LAST_NAME_COL']
    current_crm_first_name = CONFIG['CRM_FIRST_NAME_COL']
    current_new_records_name = CONFIG['NEW_RECORDS_NAME_COL']
    current_new_records_forename = CONFIG['NEW_RECORDS_FORENAME_COL']
    current_new_records_csv_delimiter = CONFIG['NEW_RECORDS_CSV_DELIMITER']

    # --- Read CRM File ---
    status_text.insert(tk.END, f"Processing CRM file: {os.path.basename(crm_filepath)}...\n")
    _, crm_unique_ids, _, crm_error = read_crm_csv_file(
        crm_filepath, current_crm_delimiter, current_crm_last_name, current_crm_first_name, status_text
    )
    if crm_error:
        status_text.insert(tk.END, crm_error + "\n")
        status_text.config(state=tk.DISABLED); enable_save_buttons(False); return
    if crm_unique_ids is None: 
        status_text.insert(tk.END, "Critical error reading CRM file (crm_unique_ids is None).\n")
        status_text.config(state=tk.DISABLED); enable_save_buttons(False); return
    status_text.insert(tk.END, f"Found {len(crm_unique_ids)} unique IDs in CRM file.\n" if crm_unique_ids else f"Warning: CRM file yielded no unique IDs to compare against.\n")

    # --- Read New Records File ---
    status_text.insert(tk.END, f"\nProcessing New Records file: {os.path.basename(new_records_filepath)}...\n")
    new_records_list, header_from_new_file, new_records_error = read_new_records_file(
        new_records_filepath, current_new_records_name, current_new_records_forename, current_new_records_csv_delimiter, status_text
    )
    if new_records_error:
        status_text.insert(tk.END, new_records_error + "\n")
        status_text.config(state=tk.DISABLED); enable_save_buttons(False); return
    
    new_records_header_global = header_from_new_file 
    if not new_records_header_global:
        status_text.insert(tk.END, f"Critical Error: No valid header could be determined from New Records file. Cannot compare or save.\n")
        status_text.config(state=tk.DISABLED); enable_save_buttons(False); return
    
    status_text.insert(tk.END, f"Using header for New Records: {', '.join(map(str,new_records_header_global))}\n")
    
    if not new_records_list: # Check if list is empty after successful header read
        status_text.insert(tk.END, "Warning: New Records file has a valid header but contains no data rows to process.\n")
    else:
        status_text.insert(tk.END, f"Total records to process from New Records file: {len(new_records_list)}\n")


    # --- Comparison Logic ---
    status_text.insert(tk.END, "\nComparing records...\n")
    
    normalized_global_header_map = {str(h).strip().lower(): str(h) for h in new_records_header_global}
    actual_name_col_in_header = normalized_global_header_map.get(current_new_records_name.lower())
    actual_forename_col_in_header = normalized_global_header_map.get(current_new_records_forename.lower())

    if not actual_name_col_in_header or not actual_forename_col_in_header:
        status_text.insert(tk.END, f"Critical Error: Configured Name/Forename columns ('{current_new_records_name}', '{current_new_records_forename}') not found in determined New Records header after normalization. Cannot compare.\n")
        status_text.insert(tk.END, f"Determined header: {', '.join(new_records_header_global)}\n")
        status_text.insert(tk.END, f"Normalized map: {normalized_global_header_map}\n")
        status_text.config(state=tk.DISABLED); enable_save_buttons(False); return

    for idx, record_dict in enumerate(new_records_list):
        name_val = record_dict.get(actual_name_col_in_header, "")
        forename_val = record_dict.get(actual_forename_col_in_header, "")
        
        current_id = create_unique_id(name_val, forename_val)
        if not current_id: 
            status_text.insert(tk.END, f"Warning: Skipping record (row {idx+2} approx, ID empty after normalization): { {actual_name_col_in_header: name_val, actual_forename_col_in_header: forename_val} }\n")
            continue

        if crm_unique_ids is not None and current_id in crm_unique_ids:
            processed_duplicates.append(record_dict)
        else:
            processed_uniques.append(record_dict)

    status_text.insert(tk.END, f"\n--- Processing Complete ---\n")
    status_text.insert(tk.END, f"Unique new contacts to import: {len(processed_uniques)}\n")
    status_text.insert(tk.END, f"Duplicate contacts found (for review): {len(processed_duplicates)}\n")
    
    if new_records_header_global and (processed_uniques or processed_duplicates):
        enable_save_buttons(True)
    else:
        enable_save_buttons(False)
        if not new_records_header_global:
             status_text.insert(tk.END, "Cannot enable save: New records header missing.\n")
        if not processed_uniques and not processed_duplicates:
            status_text.insert(tk.END, "No unique or duplicate records to save (possibly due to empty input or all records skipped).\n")

    status_text.config(state=tk.DISABLED); status_text.see(tk.END)

# --- UI Functions ---
def open_options_window():
    options_win = Toplevel(root)
    options_win.title("Configuration Options")
    options_win.geometry("450x300")
    options_win.transient(root) 
    options_win.grab_set() 

    main_frame = Frame(options_win, padx=10, pady=10)
    main_frame.pack(fill=tk.BOTH, expand=True)

    temp_sv = {key: StringVar(value=CONFIG[key]) for key, value in CONFIG.items()}

    fields = [
        ("CRM Delimiter:", 'CRM_DELIMITER'),
        ("CRM Last Name Column:", 'CRM_LAST_NAME_COL'),
        ("CRM First Name Column:", 'CRM_FIRST_NAME_COL'),
        ("New Records Name Column:", 'NEW_RECORDS_NAME_COL'),
        ("New Records Forename Column:", 'NEW_RECORDS_FORENAME_COL'),
        ("New Records CSV Delimiter:", 'NEW_RECORDS_CSV_DELIMITER')
    ]

    for i, (label_text, key) in enumerate(fields):
        Label(main_frame, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=2)
        Entry(main_frame, textvariable=temp_sv[key], width=30).grid(row=i, column=1, sticky=tk.EW, padx=5, pady=2)
    
    main_frame.columnconfigure(1, weight=1)

    def save_config():
        global CONFIG 
        for key in CONFIG:
            CONFIG[key] = temp_sv[key].get()
        
        update_main_window_info_labels()
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, "Configuration updated.\n")
        status_text.config(state=tk.DISABLED)
        status_text.see(tk.END)
        options_win.destroy()

    def cancel_config():
        options_win.destroy()

    button_frame = Frame(main_frame)
    button_frame.grid(row=len(fields), column=0, columnspan=2, pady=10, sticky=tk.E)
    Button(button_frame, text="Save", command=save_config, width=10, bg="#ABEBC6").pack(side=tk.LEFT, padx=5)
    Button(button_frame, text="Cancel", command=cancel_config, width=10).pack(side=tk.LEFT)

def update_main_window_info_labels():
    #Updates the StringVars for the infos on the main window.
    if crm_delimiter_sv: crm_delimiter_sv.set(CONFIG['CRM_DELIMITER'])
    if crm_last_name_sv: crm_last_name_sv.set(CONFIG['CRM_LAST_NAME_COL'])
    if crm_first_name_sv: crm_first_name_sv.set(CONFIG['CRM_FIRST_NAME_COL'])
    if new_records_name_sv: new_records_name_sv.set(CONFIG['NEW_RECORDS_NAME_COL'])
    if new_records_forename_sv: new_records_forename_sv.set(CONFIG['NEW_RECORDS_FORENAME_COL'])
    if new_records_csv_delimiter_sv: new_records_csv_delimiter_sv.set(CONFIG['NEW_RECORDS_CSV_DELIMITER'])


def browse_file(entry_widget, title="Select File", filetypes=(("All files", "*.*"),)):
    filepath = filedialog.askopenfilename(title=title, filetypes=filetypes)
    if filepath:
        entry_widget.config(state=tk.NORMAL)
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, filepath)
        entry_widget.config(state=tk.DISABLED) 
        status_text.config(state=tk.NORMAL)
        status_text.delete('1.0', tk.END)
        status_text.insert(tk.END, "Files selected.\n")
        status_text.config(state=tk.DISABLED)
        enable_save_buttons(False)
        global processed_uniques, processed_duplicates, new_records_header_global
        processed_uniques, processed_duplicates, new_records_header_global = [], [], []

def enable_save_buttons(enable=True):
    state = tk.NORMAL if enable else tk.DISABLED
    save_uniques_button.config(state=state)
    save_duplicates_button.config(state=state)

def save_output_file(data_to_save, header_row, default_filename, title="Save CSV File"):
    if not header_row:
        messagebox.showerror("Error", "Cannot save: Output header missing.")
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, f"Error saving {default_filename}: Output header missing.\n")
        status_text.config(state=tk.DISABLED); return
    if not data_to_save:
        messagebox.showinfo("No Data", f"No data to save for '{default_filename}'."); return

    filepath = filedialog.asksaveasfilename(
        defaultextension=".csv", initialfile=default_filename,
        filetypes=(("CSV files", "*.csv"), ("All files", "*.*")), title=title
    )
    if not filepath: return

    try:
        with open(filepath, mode='w', newline='', encoding='utf-8') as file:
            # Ensure all header elements are strings
            str_header_row = [str(h) for h in header_row]
            writer = csv.DictWriter(file, fieldnames=str_header_row, delimiter=',', extrasaction='ignore')
            writer.writeheader()
            for row_dict in data_to_save:
                # Ensure all keys and values in the row are strings for DictWriter
                stringified_row = {str(k): str(v) if pd.notna(v) else "" for k, v in row_dict.items()}
                # Filter row to only include keys present in the header to avoid DictWriter errors
                filtered_row = {h_key: stringified_row.get(h_key, "") for h_key in str_header_row}
                writer.writerow(filtered_row)
        messagebox.showinfo("Success", f"File saved: {os.path.basename(filepath)}")
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, f"Saved: {os.path.basename(filepath)}\n")
        status_text.config(state=tk.DISABLED); status_text.see(tk.END)
    except Exception as e:
        messagebox.showerror("Error Saving File", f"Error saving '{os.path.basename(filepath)}': {e}")
        status_text.config(state=tk.NORMAL)
        status_text.insert(tk.END, f"Error saving {os.path.basename(filepath)}: {e}\n")
        status_text.config(state=tk.DISABLED); status_text.see(tk.END)

# --- TKINTER base Setup ---
root = tk.Tk()
root.title(APP_TITLE)
root.geometry("700x800") 

crm_delimiter_sv = StringVar(value=CONFIG['CRM_DELIMITER'])
crm_last_name_sv = StringVar(value=CONFIG['CRM_LAST_NAME_COL'])
crm_first_name_sv = StringVar(value=CONFIG['CRM_FIRST_NAME_COL'])
new_records_name_sv = StringVar(value=CONFIG['NEW_RECORDS_NAME_COL'])
new_records_forename_sv = StringVar(value=CONFIG['NEW_RECORDS_FORENAME_COL'])
new_records_csv_delimiter_sv = StringVar(value=CONFIG['NEW_RECORDS_CSV_DELIMITER'])

file_frame = tk.LabelFrame(root, text="Select Input Files", padx=10, pady=10)
file_frame.pack(fill=tk.X, padx=10, pady=5)
file_frame.columnconfigure(1, weight=1) 
tk.Label(file_frame, text="CRM Export (CSV):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=3)
crm_file_entry = tk.Entry(file_frame, width=55, state=tk.DISABLED) 
crm_file_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=3) 
tk.Button(file_frame, text="Select database",  width=15, command=lambda: browse_file(crm_file_entry, "Select CRM CSV", (("CSV files", "*.csv"), ("All files", "*.*")))).grid(row=0, column=2, padx=5, pady=3)
tk.Label(file_frame, text="New Records (Excel/CSV):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=3)
new_records_file_entry = tk.Entry(file_frame, width=55, state=tk.DISABLED) 
new_records_file_entry.grid(row=1, column=1, sticky=tk.EW, padx=5, pady=3) 
tk.Button(file_frame, text="Select new records", width=15, command=lambda: browse_file(new_records_file_entry, "Select New Records File", (("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")))).grid(row=1, column=2, padx=5, pady=3)

info_frame = tk.LabelFrame(root, text="Current Input Files Configuration", padx=10, pady=10)
info_frame.pack(fill=tk.X, padx=10, pady=5)

monospace_font = tkFont.Font(family="Courier New", size=11)


new_records_info_subframe = Frame(info_frame)
new_records_info_subframe.pack(fill=tk.X, anchor=tk.W)
Label(info_frame, text="CRM File     --->      LAST NAME: '", font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(info_frame, textvariable=crm_last_name_sv, font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(info_frame, text="'     FIRST NAME: '",font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(info_frame, textvariable=crm_first_name_sv, font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(info_frame, text="'     DELIMITER: '",font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(info_frame, textvariable=crm_delimiter_sv,font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)

Label(info_frame, text="'").pack(side=tk.LEFT, anchor=tk.W)

Label(new_records_info_subframe, text="New Records  --->      LAST NAME: '",font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(new_records_info_subframe, textvariable=new_records_name_sv,font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(new_records_info_subframe, text="'     FIRST NAME: '",font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(new_records_info_subframe, textvariable=new_records_forename_sv,font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(new_records_info_subframe, text="'     DELIMITER: '",font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(new_records_info_subframe, textvariable=new_records_csv_delimiter_sv,font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)
Label(new_records_info_subframe, text="'",font=monospace_font).pack(side=tk.LEFT, anchor=tk.W)


action_buttons_frame = Frame(root)
action_buttons_frame.pack(pady=10)
process_button = tk.Button(action_buttons_frame, text="Process Files", command=process_files, font=('Arial', 12, 'bold'), bg="#AED6F1", width=20, pady=5)
process_button.pack(side=tk.LEFT, padx=5)
options_button = tk.Button(action_buttons_frame, text="Options", command=open_options_window, width=15, pady=5)
options_button.pack(side=tk.LEFT, padx=5)

status_frame = tk.LabelFrame(root, text="Status & Messages", padx=10, pady=10)
status_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
status_text = scrolledtext.ScrolledText(status_frame, height=12, width=80, state=tk.DISABLED, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1)
status_text.pack(pady=5, padx=5, fill=tk.BOTH, expand=True)

save_frame = tk.LabelFrame(root, text="Save Output Files", padx=10, pady=10)
save_frame.pack(fill=tk.X, padx=10, pady=(5,10))
save_uniques_button = tk.Button(save_frame, text="Save Unique Contacts (CSV)", command=lambda: save_output_file(processed_uniques, new_records_header_global, "contacts_to_import.csv", "Save Unique Contacts"), state=tk.DISABLED, width=30, bg="#ABEBC6")
save_uniques_button.pack(side=tk.LEFT, padx=10, pady=5, expand=True)
save_duplicates_button = tk.Button(save_frame, text="Save Duplicates for Review (CSV)", command=lambda: save_output_file(processed_duplicates, new_records_header_global, "duplicates_to_review.csv", "Save Duplicates for Review"), state=tk.DISABLED, width=30, bg="#FAD7A0")
save_duplicates_button.pack(side=tk.RIGHT, padx=10, pady=5, expand=True)

if __name__ == "__main__":
    status_text.config(state=tk.NORMAL)
    status_text.insert(tk.END, f"Welcome to {APP_TITLE}!\n\n")
    status_text.insert(tk.END, "1. Select CRM export & New Records files.\n")
    status_text.insert(tk.END, "2. (Optional) Click 'Options...' to change delimiters/column names.\n")
    status_text.insert(tk.END, "3. Click 'Process Files'.\n\n")
    status_text.insert(tk.END, "Names are normalized (lowercase, no accents/punctuation, etc.) for better matching.\n\n")
    status_text.insert(tk.END, "Enjoy :)\n")

    status_text.config(state=tk.DISABLED)
    root.mainloop()
