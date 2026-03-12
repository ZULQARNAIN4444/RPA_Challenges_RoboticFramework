import pandas as pd        # pandas is used to read and write Excel files
import os                  # os is used to check if files/folders exist
import requests            # requests is used to download files from the internet


# URL of the Excel file that contains the notary data
URL = "https://botsdna.com/notaries/AP-ADVOCATES.xlsx"

# Global dataframe variable
# It will store the Excel data so multiple functions can access and modify it
df_global = None


# This function cleans district names
# Example: "HYDERABAD DIST" -> "HYDERABAD"
def clean_district(text):

    return str(text).upper().replace("DIST", "").strip()



# This function ensures the Excel file exists
# If not, it downloads it from the website
def ensure_excel_downloaded(path):

    # Check if file already exists
    if os.path.exists(path):
        return   # If it exists, do nothing

    print("Downloading Excel file...")

    # Download Excel file from website
    r = requests.get(URL)

    # Save downloaded file to given path
    with open(path, "wb") as f:
        f.write(r.content)



# This function loads the Excel rows and prepares them for Robot Framework
def load_notary_rows(path):

    global df_global

    # Read Excel file
    # dtype=str ensures all values are read as text
    df_global = pd.read_excel(path, dtype=str).fillna("")

    # If "Transaction Number" column doesn't exist, create it
    if "Transaction Number" not in df_global.columns:
        df_global["Transaction Number"] = ""

    rows = []                # List to store rows that need processing
    current_district = None  # Variable to track which district we are currently in

    # Loop through each row in Excel
    for idx, row in df_global.iterrows():

        # Read SL.NO column
        sl = row["SL.NO."].strip()

        # If this row contains district information
        if "DIST" in sl.upper():

            # Extract district name and clean it
            current_district = clean_district(sl)

            # Skip this row because it's just a district heading
            continue

        # If district not yet found, skip row
        if not current_district:
            continue

        # Skip rows that already have transaction numbers
        # (means they were already processed earlier)
        if row["Transaction Number"].strip():
            continue

        # Get notary name
        notary = row["NOTARY ADVOCATE NAME"]

        # Get practice area
        area = row["AREA OF PRACTICE"]

        # If notary name is empty skip the row
        if not notary:
            continue

        # Add row data to list for Robot Framework processing
        rows.append({
            "index": idx,                # Save Excel row index
            "district": current_district,
            "notary": notary,
            "area": area
        })

    # Return rows list to Robot Framework
    return rows



# This function updates transaction number in dataframe
def set_transaction_number(row, txn):

    global df_global

    # Get row index saved earlier
    idx = row["index"]

    # Update transaction number in dataframe
    df_global.at[idx, "Transaction Number"] = txn



# This function saves updated dataframe back to Excel
def save_excel_file(path):

    global df_global

    # Write dataframe to Excel file
    df_global.to_excel(path, index=False)

    print("Excel saved successfully")