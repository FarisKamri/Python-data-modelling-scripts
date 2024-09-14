import warnings
import pandas as pd
from pathlib import Path
import glob
import os

# Suppress specific warnings
warnings.filterwarnings("ignore", category=UserWarning, module='openpyxl')

def get_files_in_directory(directory):
    """Return a list of files in the specified directory."""
    return glob.glob(os.path.join(directory, '*'))

def get_most_recent_files(directory, num_files=3):
    """Return the most recent files in the given directory."""
    files = get_files_in_directory(directory)
    files_with_times = [(f, os.path.getmtime(f)) for f in files]
    files_with_times.sort(key=lambda x: x[1], reverse=True)
    return files_with_times[:num_files]

def assign_files_to_categories(files_with_times, categories):
    """Assign files to categories based on their names."""
    filepaths = {category: None for category in categories}
    for file, _ in files_with_times:
        for category in categories:
            if category.lower() in os.path.basename(file).lower():
                if filepaths[category] is None:
                    filepaths[category] = file
                    break 
    return filepaths

def get_most_recent_file(directory, keyword):
    """Return the path to the most recent file in the given directory containing the keyword."""
    files = [f for f in Path(directory).iterdir() if f.is_file() and keyword in f.name]
    if not files:
        raise FileNotFoundError(f"No files found containing '{keyword}' in the directory.")
    most_recent_file = max(files, key=lambda f: f.stat().st_mtime)
    return most_recent_file

def extract_third_column(file_path):
    """Extract the third column from an Excel file."""
    try:
        df = pd.read_excel(file_path, sheet_name=0)  # Read the first sheet
    except Exception as e:
        raise Exception(f"An error occurred while reading the Excel file: {e}")
    
    if df.shape[1] < 3:
        raise ValueError("The file does not have a third column.")
    
    # Extract the third column
    third_column_data = df.iloc[:, 2]  # Third column (0-based index)
    return third_column_data

def main():
    """Main function to get the most recent files and extract data."""
    # Replace this with the path to your downloads directory
    downloads_directory = str(Path.home() / 'Downloads')
    
    try:
        # Get the most recent GMV and Cost files
        gmv_file = get_most_recent_file(downloads_directory, 'GMV')
        cost_file = get_most_recent_file(downloads_directory, 'COST')
        
        # Extract third column from both files
        gmv_data = extract_third_column(gmv_file)
        cost_data = extract_third_column(cost_file)
        
        # Print the data in a format suitable for Excel
        for gmv_value, cost_value in zip(gmv_data, cost_data):
            print(f"{gmv_value:.2f}\t{cost_value:.2f}")
        
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
