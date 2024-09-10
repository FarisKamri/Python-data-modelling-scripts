##Insert below to Python

import pandas as pd
from datetime import datetime, timedelta
import warnings
import os
import glob
from pathlib import Path


# Get the time_target
current_time = datetime.now()
time_15min = current_time - timedelta(minutes=15)
time_target = time_15min.strftime('%H:%M')
print(time_target)

#Get the files and categorise
def get_files_in_directory(directory):
    return glob.glob(os.path.join(directory, '*'))

def get_most_recent_files(directory, num_files=3):
    files = get_files_in_directory(directory)
    files_with_times = [(f, os.path.getmtime(f)) for f in files]
    files_with_times.sort(key=lambda x: x[1], reverse=True)
    return files_with_times[:num_files]

def assign_files_to_categories(files_with_times, categories):
    filepaths = {category: None for category in categories}
    for file, _ in files_with_times:
        for category in categories:
            if category.lower() in os.path.basename(file).lower():
                if filepaths[category] is None:
                    filepaths[category] = file
                    break 

    return filepaths

downloads_folder = str(Path.home() / 'Downloads')
most_recent_files = get_most_recent_files(downloads_folder, num_files=3)
categories = ['GMV', 'Orders', 'Cost']
filepaths = assign_files_to_categories(most_recent_files, categories)

print("Updated filepaths dictionary:")
for category, path in filepaths.items():
    print(f"{category}: {path}")
# Suppress specific warnings from openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Column mappings
columns = {
    'GMV': {'time': 'Time', 'cumulative': 'Cumulative GMV'},
    'Orders': {'time': 'Time', 'cumulative': 'Cumulative Gross Orders'},
    'Cost': {'time': 'Time', 'cumulative': 'Cumulative Total Gross Cost'}
}

# Target local time and timezone offset
time_zone_offset = 8

def convert_to_utc(local_time_str, offset):
    local_time = datetime.strptime(local_time_str, '%H:%M')
    utc_time = local_time - timedelta(hours=offset)
    return utc_time.strftime('%H:%M')

def process_file(filepath, time_column, cumulative_column):
    try:
        excel_file = pd.ExcelFile(filepath)
        utc_time_str = convert_to_utc(time_target, time_zone_offset)
        utc_time = datetime.strptime(utc_time_str, '%H:%M').time()
        
        for sheet_index in range(min(4, len(excel_file.sheet_names))):
            sheet_name = excel_file.sheet_names[sheet_index]
            df = pd.read_excel(filepath, sheet_name=sheet_index)
            
            # Convert time column to datetime and extract time
            df[time_column] = pd.to_datetime(df[time_column], errors='coerce').dt.time
            
            # Find the row matching the UTC time
            matching_row = df[df[time_column] == utc_time]
            final_data_value = df[cumulative_column].iloc[-1] if not df[cumulative_column].empty else 'No data'
            
            # Print results
            if not matching_row.empty:
                data_value = matching_row[cumulative_column].values[0]
                print(data_value)
            else:
                print('No matching time found')
            
            print(final_data_value)
    
    except Exception as e:
        print(f"Error processing file {filepath}: {e}")

# Print the sheet name for tab 3 of the first file
first_file = list(filepaths.values())[0]  # Get the first file's path

try:
    first_excel_file = pd.ExcelFile(first_file)
    if len(first_excel_file.sheet_names) > 3:
        sheet_name_tab_3 = first_excel_file.sheet_names[3]  # Sheet tab 3 is index 3
        
        # Assuming the date is always the last part of the sheet name after the last underscore
        date_str = sheet_name_tab_3.split('_')[-1]
        
        # Print the extracted date in the desired format
        print(date_str)
    else:
        print("The first file does not have a tab 3.")
except Exception as e:
    print(f"Error accessing the first file: {e}")

# Process each file
for key in filepaths:
    process_file(filepaths[key], columns[key]['time'], columns[key]['cumulative'])


