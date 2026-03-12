import pandas as pd        # pandas library used for reading, processing, and writing Excel data
import os                  # os library used for file and folder operations
import requests            # requests library used for downloading files from the internet
import zipfile             # zipfile library used to extract ZIP files
import glob                # glob library used to search files using patterns (*.zip, *.txt)
import re                  # re library used for regular expressions (pattern matching)


# Function to read Excel file and prepare dataframe structure
def prepare_dataframe(file_path):

    # Read Excel file into pandas dataframe and convert all values to string
    df = pd.read_excel(file_path,dtype=str)

    # List of columns that must exist in Excel
    needed = ["Status","PAN NUMBER","Bank","Branch","Loan Taken On","Amount","EMI(month)"]

    # Check if each column exists, if not create it
    for col in needed:
        if col not in df.columns:
            df[col] = ""

    # Extract last 4 digits from AccountNumber column
    df["Last4"] = df["AccountNumber"].astype(str).str[-4:]

    # Convert dataframe into list of dictionaries (Robot Framework works better with this format)
    return df.to_dict(orient="records")


# Function to extract last 4 digits from loan text
def get_last4_digits(text):

    # Regular expression to match exactly 4 digits at the end of a string
    m = re.search(r'\d{4}$',str(text))

    # If match found
    if m:
        return m.group()   # return the matched digits

    # If no digits found
    return ""


# Function to update dataframe using last 4 digits
def update_dataframe_from_last4(df,last4,status,pan):

    # Loop through each row in dataframe (which is actually a list of dictionaries)
    for row in df:

        # If the last 4 digits match
        if row.get("Last4") == last4:

            # Update status column
            row["Status"] = status

            # Update PAN number column
            row["PAN NUMBER"] = pan


# Function to download files using requests session
def download_with_session(url,download_path):

    # Extract filename from URL
    filename = url.split("/")[-1]

    # Create full file path where file will be saved
    filepath = os.path.join(download_path,filename)

    # Create a requests session
    session = requests.Session()

    # Fake browser header so server allows download
    headers = {
        "User-Agent":"Mozilla/5.0"
    }

    # Send GET request to download file
    r = session.get(url,headers=headers)

    # Save file in binary mode
    with open(filepath,"wb") as f:
        f.write(r.content)


# Function to extract all ZIP files from download folder
def extract_all_zips(download_path):

    # Find all zip files inside the folder
    zips = glob.glob(os.path.join(download_path,"*.zip"))

    # Loop through each zip file
    for z in zips:

        try:
            # Open the zip file
            with zipfile.ZipFile(z,"r") as zip_ref:

                # Extract contents to the same folder
                zip_ref.extractall(download_path)

            # Delete the zip file after extraction
            os.remove(z)

        # If zip is corrupted or invalid
        except:
            print("Skipping invalid zip:",z)


# Function to read TXT files and update dataframe
def update_from_txt(df,download_path):

    # Find all txt files in folder
    txt_files = glob.glob(os.path.join(download_path,"*.txt"))

    # Loop through each txt file
    for txt in txt_files:

        data = {}   # dictionary to store parsed values from txt

        # Open text file
        with open(txt,"r",encoding="utf-8") as f:

            # Read file line by line
            for line in f:

                # If line contains key:value format
                if ":" in line:

                    # Split key and value
                    k,v = line.split(":",1)

                    # Remove extra spaces and store in dictionary
                    data[k.strip()] = v.strip()

        # Get account number from txt file
        acc = data.get("Account Number","")

        # Find matching row in dataframe
        for row in df:

            if str(row.get("AccountNumber","")) == acc:

                # Update loan information from txt file
                row["Bank"] = data.get("Bank","")
                row["Branch"] = data.get("Branch","")
                row["Loan Taken On"] = data.get("Loan Taken On","")
                row["Amount"] = data.get("Amount","")
                row["EMI(month)"] = data.get("EMI(month)","")

        # Delete txt file after processing
        os.remove(txt)


# Function to save updated dataframe back to Excel
def save_dataframe(df,file_path):

    # Remove helper column Last4 before saving
    for row in df:
        row.pop("Last4",None)

    # Convert list of dictionaries back to pandas dataframe
    pd.DataFrame(df).to_excel(file_path,index=False)