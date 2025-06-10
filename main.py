import os
import sys
import traceback
import configparser
import pandas as pd
from openpyxl.utils import cell as cell_utils


EXIT_PROMPT = "[Press Enter or close the window to exit]"

def load_config(config_path="config.ini"):
    if not os.path.exists(config_path):
        print("Error: config.ini not found in the current directory.")
        input(EXIT_PROMPT)
        sys.exit(1)

    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8_sig")

    try:
        folder_path = config.get('Settings', 'WorkingFolderPath')
        sheet_name = config.get('Settings', 'SheetName', fallback=None)
        if sheet_name is None:
            sheet_number = config.getint('Settings', 'SheetNumber', fallback=1)
            sheet_name = sheet_number - 1
        output_file = config.get('Settings', 'OutputFileName')
        cell_dict = {v: k for k, v in dict(config.items('Cells')).items()}
    except (configparser.NoSectionError, configparser.NoOptionError, ValueError) as e:
        print(f"Error parsing config.ini: {e}")
        input(EXIT_PROMPT)
        sys.exit(1)

    return folder_path, sheet_name, output_file, cell_dict


def get_excel_files(folder_path):
    excel_files = []
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith(('.xlsx', '.xls', '.xlsm')):
                excel_files.append(os.path.join(root, file))
    return sorted(excel_files)


def extract_cells_from_excel(file_path, cell_dict, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
    except Exception as e:
        print(f"Failed to read {file_path}: {e}")
        return None

    selected_cells = []
    for cell in cell_dict.keys():
        try:
            row, col = cell_utils.coordinate_to_tuple(cell)
            selected_cells.append(df.iloc[row - 1, col - 1])
        except Exception:
            print(f"Cell {cell} out of range or invalid in file: {file_path}")
            selected_cells.append(None)

    return pd.DataFrame([selected_cells], columns=cell_dict.values())


def write_output_csv(df, output_path):
    try:
        df.to_csv(output_path, index=False)
        print(f"Conversion complete: All extracted data written to {output_path}")
    except Exception as e:
        print(f"Failed to write CSV: {e}")
        input(EXIT_PROMPT)
        sys.exit(1)


def main():
    try:
        folder_path, sheet_name, output_file, cell_dict = load_config()
        excel_files = get_excel_files(folder_path)

        combined_df = pd.DataFrame(columns=cell_dict.values())

        for file_path in excel_files:
            result_df = extract_cells_from_excel(file_path, cell_dict, sheet_name)
            if result_df is not None:
                combined_df = pd.concat([combined_df, result_df], ignore_index=True)
                print(f"Extraction complete: {os.path.basename(file_path)}")

        output_path = os.path.join(folder_path, output_file)
        write_output_csv(combined_df, output_path)

        input(EXIT_PROMPT)

    except Exception as e:
        print(f"An unexpected error occurred: {str(e)}")
        traceback.print_exc()
        input(EXIT_PROMPT)
        sys.exit(1)


if __name__ == "__main__":
    main()
