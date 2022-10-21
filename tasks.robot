*** Settings ***
Documentation       Template robot main suite.
Library             Collections
Library             MyLibrary
Library             RPA.Browser
Library             RPA.Excel.Files
Resource            keywords.robot
Variables           MyVariables.py

*** keywords ***
Create Excel Report 
     Create workbook   D:\\robotResults.xlsx
     Save workbook

Read Excel 
    open workbook    D:\\robot.xlsx
    ${list}      read worksheet      header=true 
    log to console   ${list}
    close workbook 
    FOR         ${index}    IN    @{list}
       Search Cars     ${index}
    END   
    
Search Cars 
    [Arguments]   ${index} 
    GO to  %{C_URL}
    Maximize Browser Window
    Wait Until Element Is Visible    xpath:/html/body/div[1]/div/main/div[1]/div[1]/div[1]/div/div[1]/form/div[1]/div[1]/div/div/div/div/div[1]/div[2]
    Click Element    xpath:/html/body/div[1]/div/main/div[1]/div[1]/div[1]/div/div[1]/form/div[1]/div[1]/div/div/div/div/div[1]/div[2]
    Press Keys  NONE   ${index}[make] 
    Sleep  333ms
    Press Keys  NONE   TAB  
    Press Keys  NONE   TAB  
    Sleep  555ms 
    Press Keys  NONE   ${index}[Model] 
    Sleep  333ms
    Press Keys  NONE   TAB  
    Press Keys  NONE   TAB  
    Sleep  555ms 
    Press Keys  NONE   ${index}[max_km]
    Sleep  500ms 
    Press Keys  NONE   ENTER
    Sleep  5000ms 
    # Click Sort by button  
    Click Element    xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[1]/div[2]/div[1]/div[2]/div/div
    Sleep  500ms 
    # Choose Cheapest first 
    Click Element    xpath:/html/body/div[8]/div/div/div/div[2]/div/div[6]/p
    Sleep  3000ms 
    ${name}  Get Text   xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/h6
    Sleep   1s
    ${Total_km}  Get Text  xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[1]/div[1]/span[2]
    Sleep   1s
    ${seller}   Get Text   xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[3]/div[1]/div/div[1]/div/div/span[2]
    Sleep   1s
    ${country}  Get Text  xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[3]/div[1]/div/div[2]/span
    Sleep   1s
    ${fuel}   Get Text   xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[1]/div[5]
    Sleep   1s 
    ${transmition}  Get Text  xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[1]/div[4]/span[2]
    Sleep   1s
    ${price}  Get Text  xpath:/html/body/div[1]/div/main/div[2]/div[2]/section/div/div[2]/div[1]/div/a/div/div[2]/div[3]/div[2]/div/div[2]/div[2]/div/div[1]
    Sleep   1s
    ${car_dict}    Create Dictionary  
    ...        name=${name}
    ...        Total_km=${Total_km} 
    ...        seller=${seller}
    ...        country=${country}
    ...        feul=${fuel}
    ...        transmition=${transmition}
    ...        price=${price}
    log to console  ${car_dict}

    Append To Excel   ${car_dict}

Append To Excel
    [Arguments]      ${car_dict}
    open workbook    D:\\robotResults.xlsx
    Append Rows To worksheet   ${car_dict}    header=true
    Save workbook



*** Tasks ***
main 
    Create Excel Report
    Open Available Browser 
    Read Excel
    Close All Browsers 

