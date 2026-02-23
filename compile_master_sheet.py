import os
import pandas as pd

# Path to downloads directory
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")

# List all subfolders (criteria)
criteria_folders = [f for f in os.listdir(DOWNLOAD_DIR) if os.path.isdir(os.path.join(DOWNLOAD_DIR, f))]

all_dfs = []
for folder in criteria_folders:
    folder_path = os.path.join(DOWNLOAD_DIR, folder)
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx'):
            file_path = os.path.join(folder_path, file)
            try:
                df = pd.read_excel(file_path)
                df['Criteria'] = folder.replace('_', ' ')
                df['Year'] = file.split('_')[-1].replace('.xlsx', '')
                all_dfs.append(df)
            except Exception as e:
                print(f"Failed to read {file_path}: {e}")

if all_dfs:
    master_df = pd.concat(all_dfs, ignore_index=True)
    master_df.to_excel(os.path.join(DOWNLOAD_DIR, 'master_sheet.xlsx'), index=False)
    print('Master sheet created at:', os.path.join(DOWNLOAD_DIR, 'master_sheet.xlsx'))
else:
    print('No Excel files found to compile.')
