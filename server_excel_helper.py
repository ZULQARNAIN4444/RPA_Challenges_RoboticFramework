import pandas as pd      # pandas library used to read and write Excel files
import os                # os library used for file operations

# Global dataframe variable
# It stores the Excel data so multiple functions can update it
df_global = None


# Function to read server requests from Excel
def load_server_requests(path):

    global df_global

    # Read Excel file into pandas dataframe
    # dtype=str ensures every column is treated as text
    df_global = pd.read_excel(path, dtype=str).fillna("")

    # Check if "status" column exists
    # If not, create it
    if "status" not in df_global.columns:
        df_global["status"] = ""

    rows = []   # List that will store rows for Robot Framework

    # Loop through each row in Excel
    for idx, row in df_global.iterrows():

        # If status already filled, skip that request
        if row["status"].strip():
            continue

        # Split Applications column by comma
        # Example: "Docker,Jenkins,IIS" → ["Docker","Jenkins","IIS"]
        apps = [a.strip() for a in row["Applications"].split(",")]

        # Store row data as dictionary
        rows.append({
            "index": idx,                     # Excel row index (needed for updating later)
            "RequestID": row["RequestID"],    # Request ID from Excel
            "OS": row["OS"],                  # Operating system
            "RAM": row["RAM"],                # RAM value
            "HDD": row["HDD"],                # HDD size
            "Applications": apps              # List of applications
        })

    # Return list of dictionaries to Robot Framework
    return rows


# Function to update server status after server creation
def set_server_status(index, status):

    global df_global

    # Update status column for the specific row
    df_global.at[index, "status"] = status


# Function to save final Excel output
def save_server_excel(path):

    global df_global

    # Write dataframe back to Excel file
    df_global.to_excel(path, index=False)

    print("Server output saved successfully")