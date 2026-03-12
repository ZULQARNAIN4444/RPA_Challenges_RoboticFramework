*** Settings ***
Library    SeleniumLibrary          # Used for browser automation
Library    Collections              # Used for dictionaries and lists
Library    OperatingSystem          # Used for file/folder operations
Library    String                   # Used for string manipulation
Library    BuiltIn                  # Default Robot keywords
Library    excel_writer.py          # Custom Python helper to save Excel data

*** Variables ***
${BASE_URL}            https://botsdna.com/school/
${DOWNLOAD_PATH}       B:/BotsDNA
${EXCEL_NAME}          Master Template.xlsx
${EXCEL_PATH}          ${DOWNLOAD_PATH}/${EXCEL_NAME}
${EXCEL_DOWNLOAD_XPATH}    //html/body/center/font/a[2]

# Dictionary storing column names and their XPaths on the webpage
&{XPATHS}
...    School Name=//html/body/center/h1[1]
...    School Address=//html/body/center/table[1]/tbody/tr[1]/td[2]
...    School Phonenumber=//html/body/center/table[1]/tbody/tr[2]/td[2]
...    Number of Student =//html/body/center/table[1]/tbody/tr[3]/td[2]
...    Prncipal Name=//html/body/center/table[1]/tbody/tr[4]/td[2]
...    Number of TeachingStaff=//html/body/center/table[1]/tbody/tr[5]/td[2]
...    Number of Non-TeachingStaff =//html/body/center/table[1]/tbody/tr[6]/td[2]
...    Number of School buses=//html/body/center/table[1]/tbody/tr[7]/td[2]
...    School Playground=//html/body/center/table[1]/tbody/tr[8]/td[2]
...    Facilities=//html/body/center/table[1]/tbody/tr[9]/td[2]
...    School Accrediation=//html/body/center/table[1]/tbody/tr[10]/td[2]
...    School Hostel=//html/body/center/table[1]/tbody/tr[11]/td[2]
...    School Canteen=//html/body/center/table[1]/tbody/tr[12]/td[2]
...    School Stationary=//html/body/center/table[1]/tbody/tr[13]/td[2]
...    School Teaching method's=//html/body/center/table[1]/tbody/tr[14]/td[2]
...    School Timing=//html/body/center/table[1]/tbody/tr[15]/td[2]
...    School Achivements=//html/body/center/table[1]/tbody/tr[16]/td[2]
...    School Awards=//html/body/center/table[1]/tbody/tr[17]/td[2]
...    School Uniform=//html/body/center/table[1]/tbody/tr[18]/td[2]
...    School type=//html/body/center/table[1]/tbody/tr[19]/td[2]

*** Keywords ***
Setup Browser
    Create Directory    ${DOWNLOAD_PATH}
    ${options}=    Evaluate    sys.modules['selenium.webdriver'].ChromeOptions()    sys, selenium.webdriver
    Call Method    ${options}    add_argument    --start-maximized
    Call Method    ${options}    add_argument    --disable-notifications
    &{prefs}=    Create Dictionary
    ...    download.default_directory=${DOWNLOAD_PATH}
    ...    download.prompt_for_download=False
    ...    safebrowsing.enabled=True
    Call Method    ${options}    add_experimental_option    prefs    ${prefs}
    Open Browser    ${BASE_URL}    chrome    options=${options}

Safe Get Text
    [Arguments]    ${xpath}
    ${text}=    Run Keyword And Ignore Error    Get Text    ${xpath}
    ${result}=    Run Keyword If    '${text[0]}'=='PASS'    Set Variable    ${text[1]}    ELSE    Set Variable    ${EMPTY}
    RETURN    ${result}

*** Test Cases ***
School Automation
    Setup Browser

    # Download Excel if missing
    ${exists}=    Evaluate    os.path.exists(r'''${EXCEL_PATH}''')    modules=os
    IF    not ${exists}
        Click Element    ${EXCEL_DOWNLOAD_XPATH}
        Sleep    5s
    END

    # Read Excel using pandas
    ${rows}=    Evaluate    pandas.read_excel(r'''${EXCEL_PATH}''').to_dict(orient='records')    modules=pandas

    # Loop over school codes and scrape
    FOR    ${row}    IN    @{rows}
        ${school_code}=    Strip String    ${row['School Code']}
        IF    '${school_code}'==''
            CONTINUE
        END
        Log    Processing School Code: ${school_code}
        Go To    ${BASE_URL}${school_code}.html

        # Extract data for each column
        FOR    ${column}    ${xpath}    IN    &{XPATHS}
            ${value}=    Safe Get Text    ${xpath}
            Set To Dictionary    ${row}    ${column}=${value}
        END
    END

    # Save back using python function
    Save Excel Data    ${EXCEL_PATH}    ${rows}    Sheet1

    Close Browser
    Log    SCHOOL AUTOMATION COMPLETED SUCCESSFULLY