*** Settings ***
Library    SeleniumLibrary        # Library for browser automation (open browser, click elements, get text, etc.)
Library    OperatingSystem        # Library for file and folder operations
Library    Collections            # Library for list and dictionary operations
Library    excel_helper.py        # Custom Python library that contains helper functions for Excel and file processing


*** Variables ***
${URL}                https://botsdna.com/ActiveLoans/     # Website URL containing the Active Loans table
${DOWNLOAD_PATH}      B:/BotsDNA/xlsx_activeloan           # Folder where downloaded files will be stored
${INPUT_FILE}         ${DOWNLOAD_PATH}/input.xlsx          # Excel file that contains loan account data
${TABLE_ROWS}         //table//tr[position()>1]             # XPath to select all table rows except the header row


*** Keywords ***
Setup Browser
    Create Directory    ${DOWNLOAD_PATH}        # Create download folder if it does not exist
    Open Browser    ${URL}    chrome             # Open Chrome browser and navigate to the Active Loans webpage
    Maximize Browser Window                     # Make browser window full screen
    Set Selenium Implicit Wait    5              # Wait up to 5 seconds for elements to appear automatically


*** Test Cases ***
Active Loan Automation

    Setup Browser        # Call custom keyword to open browser and prepare environment


    # Check if the Excel input file already exists
    ${exists}=    Evaluate    os.path.exists(r'''${INPUT_FILE}''')    modules=os

    # If Excel file does not exist, download it from the website
    IF    not ${exists}
        Download With Session    ${URL}input.xlsx    ${DOWNLOAD_PATH}
    END


    # Read the Excel file and convert it into a dataframe (list of dictionaries)
    ${df}=    Prepare Dataframe    ${INPUT_FILE}


    # Get all rows from the website table
    ${rows}=    Get WebElements    ${TABLE_ROWS}

    # Count how many rows exist
    ${row_count}=    Get Length    ${rows}

    # Increase range by 1 because Robot Framework range excludes last value
    ${end}=    Evaluate    ${row_count}+1


    # Loop through each row of the table
    FOR    ${i}    IN RANGE    1    ${end}

        # Get loan status from column 1
        ${status}=    Get Text    xpath=(${TABLE_ROWS})[${i}]/td[1]

        # Get loan link element from column 2
        ${loan_elem}=    Get WebElement    xpath=(${TABLE_ROWS})[${i}]/td[2]/a

        # Extract loan text (e.g. Loan1234)
        ${loan_text}=    Get Text    ${loan_elem}

        # Get PAN number from column 3
        ${pan}=    Get Text    xpath=(${TABLE_ROWS})[${i}]/td[3]

        # Extract the download link from the loan element
        ${link}=    Get Element Attribute    ${loan_elem}    href

        # Extract last 4 digits from the loan text (used for matching with Excel)
        ${last4}=    Get Last4 Digits    ${loan_text}

        # Update dataframe row where last 4 digits match the loan account
        Update Dataframe From Last4    ${df}    ${last4}    ${status}    ${pan}

        # Download the ZIP file containing loan details
        Download With Session    ${link}    ${DOWNLOAD_PATH}

    END


    # Extract all downloaded ZIP files
    Extract All Zips    ${DOWNLOAD_PATH}


    # Read TXT files extracted from ZIPs and update dataframe with loan information
    Update From Txt    ${df}    ${DOWNLOAD_PATH}


    # Save the updated dataframe back into the Excel file
    Save Dataframe    ${df}    ${INPUT_FILE}


    # Print completion message in Robot logs
    Log    AUTOMATION COMPLETED