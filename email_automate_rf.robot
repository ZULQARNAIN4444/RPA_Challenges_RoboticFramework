*** Settings ***
Library    SeleniumLibrary
Library    Collections
Library    OperatingSystem
Library    RPA.Email.ImapSmtp
Library    email_automate.py


*** Variables ***
${URL}                     https://botsdna.com/server/
${DOWNLOAD_PATH}           B:/BotsDNA
${EXCEL_FILE}              ${DOWNLOAD_PATH}/input.xlsx

${EMAIL_SENDER}            zulqarnainzia8@gmail.com
${EMAIL_PASSWORD}          rmhkpvdggxlwimjk  # App password (same as before)

${EXCEL_DOWNLOAD_XPATH}    //a[contains(text(),'input.xlsx')]
${OS_XPATH}                //*[@id='os']
${RAM_XPATH}               //*[@id='Ram']
${CREATE_SERVER_XPATH}     //*[@id='CreateServer']


*** Keywords ***

Setup Browser
    Create Directory    ${DOWNLOAD_PATH}
    Open Browser    ${URL}    chrome
    Maximize Browser Window
    Set Selenium Implicit Wait    5


Download Excel If Missing
    ${exists}=    Evaluate    os.path.exists(r'''${EXCEL_FILE}''')    modules=os
    IF    not ${exists}
        Click Element    ${EXCEL_DOWNLOAD_XPATH}
        Sleep    5
    END


Send Email Report
    [Arguments]    ${to_email}    ${request_id}    ${result}

    ${subject}=    Set Variable    Server Created - ${request_id}

    ${body}=    Set Variable
    ...    Hello,\n\nYour server request (${request_id}) has been processed.\n\nResult:\n${result}\n\nRegards,\nRPA Bot

    Authorize SMTP
...    smtp_server=smtp.gmail.com
...    smtp_port=587
...    account=${EMAIL_SENDER}
...    password=${EMAIL_PASSWORD}

    Send Message
    ...    sender=${EMAIL_SENDER}
    ...    recipients=${to_email}
    ...    subject=${subject}
    ...    body=${body}


Select HDD
    [Arguments]    ${hdd_value}
    FOR    ${i}    IN RANGE    1    6
        ${label}=    Get Text    xpath=//tr[3]/td[2]/label[${i}]
        ${clean_label}=    Evaluate    "${label}".replace(" ","").strip().lower()
        ${clean_excel}=    Evaluate    "${hdd_value}".replace(" ","").strip().lower()

        IF    '${clean_label}' == '${clean_excel}'
            Click Element    xpath=//tr[3]/td[2]/input[${i}]
            Exit For Loop
        END
    END


Select Applications
    [Arguments]    @{apps}
    FOR    ${app}    IN    @{apps}
        ${clean_excel}=    Evaluate    "${app}".strip().lower()

        FOR    ${i}    IN RANGE    1    10
            ${label}=    Get Text    xpath=//tr[4]/td[2]/label[${i}]
            ${clean_label}=    Evaluate    "${label}".strip().lower()

            IF    '${clean_label}' == '${clean_excel}'
                ${checked}=    Run Keyword And Return Status
                ...    Checkbox Should Be Selected    xpath=//tr[4]/td[2]/input[${i}]

                IF    not ${checked}
                    Click Element    xpath=//tr[4]/td[2]/input[${i}]
                END

                Exit For Loop
            END
        END
    END


Create Server
    [Arguments]    ${os}    ${ram}    ${hdd}    @{apps}

    ${os_clean}=    Evaluate    "${os}".strip()
    Select From List By Label    ${OS_XPATH}    ${os_clean}

    ${ram_clean}=    Evaluate    "${ram}".strip()
    Select From List By Label    ${RAM_XPATH}   ${ram_clean}

    Select HDD    ${hdd}
    Select Applications    @{apps}

    Click Element    ${CREATE_SERVER_XPATH}

    Wait Until Page Contains Element    //table    timeout=10
    Sleep    5s

    ${table_cells}=    Get Table Cell    //table    1    2
    RETURN    ${table_cells}


*** Test Cases ***
Server Automation

    Setup Browser
    Download Excel If Missing

    ${rows}=    Load Server Requests    ${EXCEL_FILE}

    FOR    ${row}    IN    @{rows}

        ${request}=  Set Variable    ${row}[RequestID]
        ${os}=       Set Variable    ${row}[OS]
        ${ram}=      Set Variable    ${row}[RAM]
        ${hdd}=      Set Variable    ${row}[HDD]
        ${apps}=     Set Variable    ${row}[Applications]
        ${email}=    Set Variable    ${row}[Email]

        Log    Processing ${request}

        Go To    ${URL}

        ${result}=    Create Server    ${os}    ${ram}    ${hdd}    @{apps}

        # PURE RF EMAIL (NO PYTHON)
        Send Email Report    ${email}    ${request}    ${result}

    END

    Close Browser