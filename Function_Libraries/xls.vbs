Option explicit
Dim xlApp,myXLS,sheet,sheetName,xlPath,fso
Set xlApp= CreateObject("Excel.Application")
xlApp.Visible=false
xlApp.DisplayAlerts=false
Set sheet= nothing
Set myXLS= nothing



'saves the xl file and destroys all objectects related to it
Function destroyFile
' destroys sheet
	If NOT sheet is nothing Then
		Set sheet=nothing
	End If
' destroy xl file
	If NOT  myXLS is nothing Then
		myXLS.save
		myxls.close
		Set myxls=nothing
	End If
End Function

Function destroyXLSApp
	destroyFile
	If  NOT xlApp is nothing Then
		xlApp.Application.Quit
		Set xlApp=nothing
	End If
End Function

'  Reads the data from XLS File
Function readData(xlFilePath,sName,row,col)
	
	If  xlPath <> xlFilePath Then
' destroy previous xls opened - if
		destroyFile
'  check if the file is existing
		If NOT isFileExisting(xlFilePath) Then
			msgbox "File not found " & xlFilePath
			exitTest
		End If
' open the xl file
		Set myXls = xlApp.Workbooks.Open(xlFilePath)
		xlPath=xlFilePath
		
' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
			msgbox  xlFilePath & " has not got sheet -  " & sName
			exitTest
		End If
' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName
		
'  file is same but sheet is diff
	ElseIf sheetName <>  sName Then
' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
			msgbox  xlFilePath & " has not got sheet -  " & sName
			exitTest
		End If
		
' destroys sheet
		If NOT sheet is nothing Then
			Set sheet=nothing
		End If
		
' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName
		
	End If
	
' read the data from the sheet
	
	readData = sheet.cells(row,col)
End Function

' write data in xls file
Function writeData(xlFilePath,sName,row,col,data)
	If  xlPath <> xlFilePath Then
' destroy previous xls opened - if
		destroyFile
'  check if the file is existing
		If NOT isFileExisting(xlFilePath) Then
			msgbox "File not found " & xlFilePath
			exitTest
		End If
' open the xl file
		Set myXls = xlApp.Workbooks.Open(xlFilePath)
		xlPath=xlFilePath
		
' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
			msgbox  xlFilePath & " has not got sheet -  " & sName
			exitTest
		End If
' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName
		
'  file is same but sheet is diff
	ElseIf sheetName <>  sName Then
' check if sheet is present
		If NOT isSheetExisting(xlFilePath,sName) Then
			msgbox  xlFilePath & " has not got sheet -  " & sName
			exitTest
		End If
		
' destroys sheet
		If NOT sheet is nothing Then
			Set sheet=nothing
		End If
		
' open the sheet of xl file
		Set sheet=myXls.sheets(sName)
		sheetName=sName
		
	End If
' write data
	'If data <> empty OR data = 0 OR data <> null Then
		If IsNumeric(data) Then
			sheet.cells(row,col).value = CInt(data)
		Else
			sheet.cells(row,col).value = data
		End If
		destroyFile
	'Else
		'destroyFile
	'End If
End Function

'returns total cols in xls file
Function getColumnCount(xlFilePath,sheetName)
	Dim totalCols
	totalCols=0
	While readData(xlFilePath,sheetName,1,(totalCols+1)) <> ""
		totalCols=totalCols+1
	Wend
	getColumnCount=totalCols
	
End Function

'Returns total Rows
Function getRowCount(xlFilePath,sheetName)
	Dim totalRows,cId
	totalRows=0
	For cId=1 to getColumnCount(xlFilePath,sheetName)
		
		While readData(xlFilePath,sheetName,(totalRows+1),cId) <> ""
			totalRows=totalRows+1
		Wend
		
	Next
	getRowCount= totalRows
	
End Function

' checks if a file exists
Function isFileExisting(filePath)
	Set fso = createObject("Scripting.FileSystemObject")
	
	If  fso.FileExists(filePath)  Then
		isFileExisting=true
	else
		isFileExisting=false
	End If
	
End Function

'  Checks if sheet is existing
Function isSheetExisting(filePath,sName)
	Dim totalsheets,sNum
	totalsheets = myXLS.Worksheets.count
	For sNum=1 to totalsheets
		If  myXLS.Worksheets(sNum).name = sName Then
			isSheetExisting=true
			Exit Function
		End If
	Next
	isSheetExisting=false
End Function


Function readEnvironmentVariables

	Environment.Value("estrTestResultHomeFolder") = readData(xlFilePath_env,"Sheet1",2,2)
	Environment.Value("TestSuite") = readData(xlFilePath_env,"Sheet1",3,2)
	Environment.Value("Browser") = readData(xlFilePath_env,"Sheet1",4,2)
	Environment.Value("estrEnvPath") = readData(xlFilePath_env,"Sheet1",6,2)
	Environment.Value("estrTestCaseStatus") = readData(xlFilePath_env,"Sheet1",7,2)
	Environment.Value("eintGlobalCheckPointCount") = readData(xlFilePath_env,"Sheet1",8,2)
	Environment.Value("STATUS_PASS") = readData(xlFilePath_env,"Sheet1",9,2)
	Environment.Value("STATUS_FAIL") = readData(xlFilePath_env,"Sheet1",10,2)

End Function
Function writeEnvironmentVariables	
	writeData xlFilePath_env,"Sheet1",8,2,Environment.Value("eintGlobalCheckPointCount")
	destroyFile
End Function  

'Method to fetch Test Data
'returns data as dictionary object
Function getData(xlFilePath,sheetName)
	Dim columnCount, data, rowNumber ,i

	columnCount = getColumnCount(xlFilePath,sheetName)
	rowNumber = getRowNumber(xlFilePath,sheetName,2)

	Set data=createObject("Scripting.Dictionary")
	For i=2 to columnCount
	data.Add readData(xlFilePath,sheetName,1,i),readData(xlFilePath,sheetName,rowNumber,i)
	Next

	Set getData = data
	destroyFile
End Function







