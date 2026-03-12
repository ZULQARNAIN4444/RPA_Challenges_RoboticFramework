*** Settings ***
Library    SeleniumLibrary        # Used for browser automation (open browser, click, input text, etc.)
Library    Collections            # Used for dictionary and list operations
Library    OperatingSystem        # Used for file and folder operations
Library    notary_excel_helper.py     # Custom Python library that handles Excel reading/writing


*** Variables ***
${URL}                   https://botsdna.com/notaries/    # Website where notary form exists
${DOWNLOAD_PATH}         B:/BotsDNA                       # Folder where Excel file will be stored
${EXCEL_FILE}            ${DOWNLOAD_PATH}/AP-ADVOCATES.xlsx  # Full path of Excel file containing notary data

${DISTRICT_XPATH}        //*[@id="DIST"]                  # XPath for District dropdown
${NOTARY_NAME_XPATH}     //*[@id="notary"]                # XPath for Notary name input field
${AREA_XPATH}            //*[@id="area"]                  # XPath for Area input field
${SUBMIT_XPATH}          //input[@value='Submit Notary']  # XPath for Submit button
${TRANS_XPATH}           //*[@id="TransNo"]               # XPath where transaction number appears after submission


*** Keywords ***
Setup Browser
    Create Directory    ${DOWNLOAD_PATH}     # Create download folder if it doesn't exist
    Open Browser    ${URL}    chrome         # Open Chrome browser and navigate to the notary website
    Maximize Browser Window                  # Maximize the browser window
    Set Selenium Implicit Wait    5          # Selenium waits up to 5 seconds for elements to appear


Submit Notary
    [Arguments]    ${district}    ${notary}    ${area}   # Keyword accepts district, notary name, and area

    # Select district from dropdown using visible label
    Select From List By Label    ${DISTRICT_XPATH}    ${district}

    # Enter notary name
    Input Text    ${NOTARY_NAME_XPATH}    ${notary}

    # Enter area name
    Input Text    ${AREA_XPATH}    ${area}

    # Click submit button
    Click Element    ${SUBMIT_XPATH}

    # After submission, get the transaction number displayed on page
    ${txn}=    Get Text    ${TRANS_XPATH}

    [Return]    ${txn}    # Return transaction number back to test case


*** Test Cases ***
Fast Notary Automation

    Setup Browser     # Open browser and prepare environment

    # Ensure Excel file exists (Python helper will download it if missing)
    Ensure Excel Downloaded    ${EXCEL_FILE}

    # Load all rows from Excel file
    ${rows}=    Load Notary Rows    ${EXCEL_FILE}

    # Loop through each row from Excel
    FOR    ${row}    IN    @{rows}

        # Extract district value from row dictionary
        ${district}=    Set Variable    ${row}[district]

        # Extract notary name from row dictionary
        ${notary}=      Set Variable    ${row}[notary]

        # Extract area from row dictionary
        ${area}=        Set Variable    ${row}[area]

        # Log which record is currently being processed
        Log    Processing ${district} | ${notary}

        # Call keyword to submit the form and capture transaction number
        ${txn}=    Submit Notary    ${district}    ${notary}    ${area}

        # Save transaction number back into row data
        Set Transaction Number    ${row}    ${txn}

        # Go back to the form page for next submission
        Go To    ${URL}

    END

    # Save updated Excel file once after processing all rows
    Save Excel File    ${EXCEL_FILE}

    Close Browser    # Close browser after automation finishes