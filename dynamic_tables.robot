*** Settings ***
Library    SeleniumLibrary        # Library used to automate browser actions (open browser, get text, click, etc.)
Library    Collections            # Library used for list and dictionary operations
Library    RPA.Excel.Files        # Library used to create and write Excel files

*** Variables ***
${URL}    https://botsdna.com/locator/                 # Website that contains the dynamic customer table
${OUTPUT}    B:/BotsDNA/customers_by_country.xlsx      # Output Excel file path

*** Test Cases ***
Solve Dynamic Table Challenge

    Open Browser    ${URL}    chrome        # Opens Chrome browser and navigates to the given URL
    Maximize Browser Window                 # Makes browser window full screen

    Wait Until Page Contains Element    //table    30s    # Wait until a table element appears (max 30 seconds)
    Sleep    5s                                           # Extra wait to ensure table loads completely

    ${table_count}=    Get Element Count    //table        # Count how many tables exist on the webpage
    ${limit}=    Evaluate    ${table_count} + 1            # Add 1 because Robot RANGE loop excludes the last number

    ${target_index}=    Set Variable    -1                 # Initialize variable to store the correct table index
    ${headers}=    Create List                             # Create an empty list to store table headers

    # Loop through each table to find the one that contains "Customer" header
    FOR    ${i}    IN RANGE    1    ${limit}

        ${ths}=    Get WebElements    xpath=(//table)[${i}]//th    # Get all header elements (th) from table i
        ${header_texts}=    Create List                            # Create temporary list to store header text

        # Loop through each header element
        FOR    ${th}    IN    @{ths}
            ${text}=    Get Text    ${th}               # Extract text from the header element
            Append To List    ${header_texts}    ${text}   # Add the header text to the list
        END

        # Check if any header contains the word "Customer"
        ${found}=    Evaluate    any("Customer" in h for h in $header_texts)

        # If the correct table is found
        IF    ${found}
            ${target_index}=    Set Variable    ${i}        # Store the index of that table
            ${headers}=    Set Variable    ${header_texts}  # Save the header list
            Exit For Loop                                   # Stop searching further tables
        END

    END

    # If no table was found containing the Customer header, stop the test
    Run Keyword If    ${target_index} == -1    Fail    Customer table not found

    # Replace the first header name with "customer name"
    Set List Value    ${headers}    0    customer name

    ${country_data}=    Create Dictionary        # Create a dictionary to store data grouped by country

    ${header_len}=    Get Length    ${headers}   # Get number of headers in the table

    # Loop through headers except the first column (customer name)
    FOR    ${i}    IN RANGE    1    ${header_len}
        ${country}=    Get From List    ${headers}    ${i}   # Get the country name from headers
        ${empty_list}=    Create List                       # Create empty list for that country
        Set To Dictionary    ${country_data}    ${country}    ${empty_list}   # Add country key with empty list
    END

    # Get all rows from the identified table
    ${rows}=    Get WebElements    xpath=(//table)[${target_index}]//tr
    ${row_len}=    Get Length    ${rows}          # Count total rows

    # Loop through table rows (starting from row 2 to skip headers)
    FOR    ${r}    IN RANGE    2    ${row_len}

        ${cells}=    Get WebElements    xpath=(//table)[${target_index}]//tr[${r}]//td   # Get all cells in the row
        ${customer}=    Get Text    ${cells}[0]    # First column contains customer name

        ${cell_len}=    Get Length    ${cells}     # Count number of columns

        # Loop through country columns
        FOR    ${c}    IN RANGE    1    ${cell_len}

            ${value}=    Get Text    ${cells}[${c}]    # Get value for that country
            ${valid}=    Evaluate    $value not in ["","0","0.0","-","N/A"]   # Check if value is meaningful

            # If the value is valid
            IF    ${valid}

                ${country}=    Get From List    ${headers}    ${c}   # Get country name from header

                ${record}=    Create Dictionary                  # Create dictionary record
                ...    customer name=${customer}                 # Store customer name
                ...    location value=${value}                   # Store location value

                ${list}=    Get From Dictionary    ${country_data}    ${country}  # Get list for that country
                Append To List    ${list}    ${record}             # Add the record to the country list
            END

        END

    END

    Create Excel    ${country_data}     # Call custom keyword to create Excel output

    Close Browser                       # Close the browser after processing

#
*** Keywords ***
Create Excel
    [Arguments]    ${data}             # Accept dictionary data as input

    Create Workbook    ${OUTPUT}       # Create new Excel workbook

    ${countries}=    Get Dictionary Keys    ${data}   # Get list of all country names

    # Loop through each country
    FOR    ${country}    IN    @{countries}

        ${records}=    Get From Dictionary    ${data}    ${country}   # Get records for that country

        ${count}=    Get Length    ${records}        # Count number of records

        Run Keyword If    ${count} == 0    Continue For Loop   # Skip country if it has no records

        Create Worksheet    ${country}        # Create worksheet named after the country

        Append Rows To Worksheet    ${records}   # Write records to Excel sheet

    END

    Remove Worksheet      Sheet     # Remove default sheet created automatically

    Save Workbook                     # Save the Excel file