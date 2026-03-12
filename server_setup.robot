*** Settings ***
Library    SeleniumLibrary          # Used for browser automation (open browser, click, input text)
Library    Collections              # Used for list and dictionary operations
Library    OperatingSystem          # Used for file and directory operations
Library    server_excel_helper.py   # Custom Python helper for Excel reading/writing


*** Variables ***
${URL}                     https://botsdna.com/server/      # Website where server request form exists
${DOWNLOAD_PATH}           B:/BotsDNA                       # Folder where Excel file will be stored
${EXCEL_FILE}              ${DOWNLOAD_PATH}/input.xlsx      # Path of Excel input file

${EXCEL_DOWNLOAD_XPATH}    //a[contains(text(),'input.xlsx')]  # XPath for downloading Excel file
${OS_XPATH}                //*[@id='os']                       # XPath for OS dropdown
${RAM_XPATH}               //*[@id='Ram']                      # XPath for RAM dropdown
${CREATE_SERVER_XPATH}     //*[@id='CreateServer']             # XPath for Create Server button


*** Keywords ***

# Opens browser and prepares environment
Setup Browser
    Create Directory    ${DOWNLOAD_PATH}      # Create download folder if not present
    Open Browser    ${URL}    chrome          # Open Chrome and navigate to server page
    Maximize Browser Window                  # Maximize browser window
    Set Selenium Implicit Wait    5           # Wait up to 5 seconds for elements


# Downloads Excel file only if it doesn't exist
Download Excel If Missing
    ${exists}=    Evaluate    os.path.exists(r'''${EXCEL_FILE}''')    modules=os
    IF    not ${exists}
        Click Element    ${EXCEL_DOWNLOAD_XPATH}   # Click download link
        Sleep    5                                  # Wait for download to complete
    END


# Select HDD radio button based on value from Excel
Select HDD
    [Arguments]    ${hdd_value}

    # Loop through all HDD options on webpage
    FOR    ${i}    IN RANGE    1    6

        # Get label text of each HDD option
        ${label}=    Get Text    xpath=//tr[3]/td[2]/label[${i}]

        # Clean webpage text (remove spaces and lowercase)
        ${clean_label}=    Evaluate    "${label}".replace(" ","").strip().lower()

        # Clean Excel value the same way
        ${clean_excel}=    Evaluate    "${hdd_value}".replace(" ","").strip().lower()

        # Compare Excel value with webpage option
        IF    '${clean_label}' == '${clean_excel}'

            # Click corresponding HDD radio button
            Click Element    xpath=//tr[3]/td[2]/input[${i}]

            Exit For Loop
        END

    END


# Select multiple application checkboxes
Select Applications
    [Arguments]    @{apps}

    # Loop through applications coming from Excel
    FOR    ${app}    IN    @{apps}

        ${clean_excel}=    Evaluate    "${app}".strip().lower()

        # Loop through application options on webpage
        FOR    ${i}    IN RANGE    1    10

            # Get label text from webpage
            ${label}=    Get Text    xpath=//tr[4]/td[2]/label[${i}]

            ${clean_label}=    Evaluate    "${label}".strip().lower()

            # Compare Excel value with webpage option
            IF    '${clean_label}' == '${clean_excel}'

                # Check if checkbox already selected
                ${checked}=    Run Keyword And Return Status
                ...    Checkbox Should Be Selected    xpath=//tr[4]/td[2]/input[${i}]

                # If not selected, click checkbox
                IF    not ${checked}
                    Click Element    xpath=//tr[4]/td[2]/input[${i}]
                END

                Exit For Loop
            END

        END

    END


# Complete server creation form
Create Server
    [Arguments]    ${os}    ${ram}    ${hdd}    @{apps}

    # Clean OS text
    ${os_clean}=    Evaluate    "${os}".strip()

    # Select OS dropdown value
    Select From List By Label    ${OS_XPATH}    ${os_clean}

    # Clean RAM value
    ${ram_clean}=    Evaluate    "${ram}".strip()

    # Select RAM dropdown value
    Select From List By Label    ${RAM_XPATH}   ${ram_clean}

    # Select HDD radio option
    Select HDD    ${hdd}

    # Select required applications
    Select Applications    @{apps}

    # Click create server button
    Click Element    ${CREATE_SERVER_XPATH}

    # Wait until result table appears
    Wait Until Page Contains Element    //table    timeout=10


*** Test Cases ***
Server Automation

    Setup Browser
    Download Excel If Missing

    # Load server requests from Excel using Python helper
    ${rows}=    Load Server Requests    ${EXCEL_FILE}

    # Loop through each request
    FOR    ${row}    IN    @{rows}

        ${index}=    Set Variable    ${row}[index]
        ${request}=  Set Variable    ${row}[RequestID]
        ${os}=       Set Variable    ${row}[OS]
        ${ram}=      Set Variable    ${row}[RAM]
        ${hdd}=      Set Variable    ${row}[HDD]
        ${apps}=     Set Variable    ${row}[Applications]

        Log    Processing ${request}

        # Reload page for new request
        Go To    ${URL}

        # Fill form and create server
        Create Server    ${os}    ${ram}    ${hdd}    @{apps}

        # Update Excel status
        Set Server Status    ${index}    server confirmed

    END

    # Save final Excel output
    Save Server Excel    ${DOWNLOAD_PATH}/Server_Output.xlsx

    Close Browser