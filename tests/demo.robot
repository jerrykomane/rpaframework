*** Settings ***
Library    RPA.Excel.Files
Library    Collections
Library     SeleniumLibrary

*** Variables ***
${EXCEL_FILE_PATH}     C:/practice/Excel_Demo/tests/test_data.xlsx
${LastName}             xpath=//input[@ng-reflect-name="labelLastName"]
${phone}                xpath=//input[@ng-reflect-name="labelPhone"]
${Email}                xpath=//input[@ng-reflect-name="labelEmail"]
${Firstname}            xpath=//input[@ng-reflect-name="labelFirstName"]
${CompanyName}          xpath=//input[@ng-reflect-name="labelCompanyName"]
${Address}              xpath=//input[@ng-reflect-name="labelAddress"]
${CompanyRole}          xpath=//input[@ng-reflect-name="labelRole"]
${Submit}               xpath=//input[@type="submit"]

*** Test Cases ***
Fill Form From Excel Data
    [Tags]  POC
    Open The Excel File
    Get Excel Data
    Log To Console    ${row}  # This is for debugging; remove once confirmed.
    Fill Form Using Extracted Data
    Close The Excel File


*** Keywords ***
Open The Excel File
    Open Workbook    ${EXCEL_FILE_PATH}

Get Excel Data
    ${data}=    Read Worksheet As Table    header=True
    ${row}=     Get From List    ${data}    0
    Set Suite Variable    ${row}

Fill Form Using Extracted Data
    Open Browser    ${row[0]}    chrome  # ${row[0]} corresponds to "URL"
    Maximize Browser Window
    Wait Until Element Is Visible    ${LastName}
    Input Text    ${lastName}        ${row[1]}  # ${row[1]} corresponds to "Company"
    Input Text    ${phone}           ${row[2]}  # ${row[2]} corresponds to "Phone"
    Input Text    ${Email}           ${row[3]}  # ${row[3]} corresponds to "Email"
    Input Text    ${Firstname}       ${row[4]}  # ${row[4]} corresponds to "Firstname"
    Input Text    ${CompanyName}     ${row[5]}  # ${row[5]} corresponds to "CompanyName"
    Input Text    ${Address}         ${row[6]}  # ${row[5]} corresponds to "CompanyName"
    Input Text    ${CompanyRole}     ${row[7]}  # ${row[5]} corresponds to "CompanyName"
    Click Element   ${Submit}


    Sleep    5

Close The Excel File
    Close Workbook
