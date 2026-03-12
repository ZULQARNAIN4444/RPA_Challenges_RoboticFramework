# excel_writer.py
import pandas as pd


def save_excel_data(file_path, data, sheet_name="Sheet1"):
    """
    Saves a list of dictionaries to an Excel file using pandas.
    Overwrites the existing sheet completely.

    :param file_path: Path to Excel file
    :param data: List of dictionaries [{col1: val1, col2: val2}, ...]
    :param sheet_name: Excel sheet name
    """
    # Read existing Excel
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError:
        # If file doesn't exist, create empty df
        df = pd.DataFrame()

    # Convert new data to DataFrame
    new_df = pd.DataFrame(data)

    # If original file exists, remove rows with same index as new data
    # (optional, depends on your workflow)

    # Save new data to Excel
    new_df.to_excel(file_path, sheet_name=sheet_name, index=False)