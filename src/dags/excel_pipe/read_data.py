import pandas as pd
import os

def read_data(file_paths, output_file):
    print('FILE PATHS', file_paths)
    print('OUTPUT FILE', output_file)
    with pd.ExcelWriter(output_file, mode='w') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='Test', index=False)

    def load_all_worksheets(file_paths):
        all_dataframes = {}
        missed_files = []

        for file_path in file_paths:
            try:
                xls = pd.ExcelFile(file_path)
                base_name = os.path.basename(file_path).split('.')[0]
                for sheet_name in xls.sheet_names:
                    key = f"{base_name} - {sheet_name}"
                    all_dataframes[key] = xls.parse(sheet_name)
            except Exception as e:
                missed_files.append(file_path)
                print(f"Error reading the file {file_path}: {e}")

        print(f'Input: {len(file_paths)} files')
        print(f'No errors: {len(file_paths) - len(missed_files)} files')
        print(f'Files with errors: {missed_files}')

        return all_dataframes

    all_sheets_data = load_all_worksheets(file_paths)
    return all_sheets_data
