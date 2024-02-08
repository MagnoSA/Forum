*** Settings ***
Test Teardown       Close All Excel Documents
Library 			ExcelLibrary
Library 			Collections

*** Variables ***
${filepath}    ${CURDIR}/../resources/Excel/ExcelRobotTest.xlsx

*** Test Cases ***
Verify Excel Cell
	Read Cell    ${filepath}    TestSheet1    2    1    

Verify Excell Row
    Read Row
    
Create Excell File
    Write Cell

*** Keywords ***
Read Cell
    [Arguments]    ${filepath}    ${sheetname}    ${rownum}    ${colnum}
    Open Excel Document    ${filepath}    1
    Get Sheet    ${sheetname}
    ${data}    Read Excel Cell    ${rownum}    ${colnum}

Read Row
    ${doc1}=	Create Excel Document	doc_id=docname1		
    ${row_data}=	Create List	t1	t2	t3
    Write Excel Row	row_num=5	row_data=${row_data}	
    ${rd}=	Read Excel Row	row_num=5	max_num=3
    Lists Should Be Equal	${row_data}	${rd}		
    	
Write Cell
    Create Excel Document	doc_id=docname1		
    Write Excel Cell	1	1	Hello
    Save Excel Document     ${CURDIR}/../resources/Excel/file.xlsx