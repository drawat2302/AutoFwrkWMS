'******variiables***********
'Public strStepNum
Public intPassTestCaseCounter
Public intFailTestCaseCounter
Public strResultsFolder


fninitialize
executeScript
deinitialize


Function executeScript
	Set xlsApp = CreateObject("Excel.Application")
	xlsApp.visible=false
	Set myFile =  xlsApp.Workbooks.Open(Environment.Value("TestSuite"))
	Set suiteSheet = myFile.sheets("Test Suite 1")
	rNum=2
	Dim testCaseName, testCaseResult
	
	
	While suiteSheet.cells(rNum,1) <> "" 
			testCaseName = suiteSheet.cells(rNum,2)
			testdataSheet = suiteSheet.cells(rNum,4)
			If suiteSheet.cells(rNum,3).value = "Y" Then
				Environment("TestCaseName") = testCaseName
				Environment("estrTestDataSheet") = testdataSheet
				'add component,columns functions
				AddComponentNameInExcelReport testCaseName
				AddColumnNamesInReport
				'fnInitializeVariablesforTestCase
				Environment("eintGlobalCheckPointCount") = 0
				Environment("estrTestCaseStatus") = ""
				'writeEnvironmentVariables
				'print testCaseName
				Execute testCaseName
				'set pass/fail status to env
				If InStr(Environment("estrTestCaseStatus"),"PASS") > 0 Then
					intPassTestCaseCounter = intPassTestCaseCounter+1
					Environment.Value("intPassTestCaseCounter") = intPassTestCaseCounter
				ElseIf InStr(Environment("estrTestCaseStatus"),"FAIL") > 0 Then
					intFailTestCaseCounter = intFailTestCaseCounter+1
					Environment.Value("intFailTestCaseCounter") = intFailTestCaseCounter
				End If
				'write testcasestatus to env sheet
				'writeData xlFilePath_env,"Sheet1",7,2,Environment.Value("estrTestCaseStatus")
				'writeData xlFilePath_env,"Sheet1",11,2,Environment.Value("intPassTestCaseCounter")
				'writeData xlFilePath_env,"Sheet1",12,2,Environment.Value("intFailTestCaseCounter")
				
				print Environment("estrTestCaseStatus")
				print Environment.Value("intPassTestCaseCounter")
				print Environment.Value("intFailTestCaseCounter")
				
				'writeEnvironmentVariables
				'ReportResult(testCaseStatus)
				'increment the counter for pass /fail
				'fndeinitializevariablesforTestCase
			End if
			rNum = rNum + 1
	Wend 
	
	myFile.save
	myFile.close
	xlsApp.quit
	
	Set xlsApp = nothing
	Set myFile = Nothing
	Set suiteSheet = Nothing
End Function
Function fninitialize()
	'variables to increment the step , passtc and failed Tc.
	readEnvironmentVariables
	Environment("eExecutionStartTime") = Now()
	intPassTestCaseCounter = 0
	intFailTestCaseCounter = 0
	
	Environment.Value("intPassTestCaseCounter") = intPassTestCaseCounter
	Environment.Value("intFailTestCaseCounter") = intFailTestCaseCounter
	
	strResultsFolder= CreateTestResultsFolder()
	'print strResultsFolder
	Environment.Value("estrResultsFolder") = strResultsFolder
	strExcelReportFileName = CreateExcelReportTemplate(strResultsFolder)
	Environment.Value("estrExcelReportFileName") = strExcelReportFileName
	writeData xlFilePath_env,"Sheet1",5,2,Environment.Value("estrExcelReportFileName")
End Function
Function deinitialize()
	Environment("eExecutionEndTime") = Now()
	ReportExecutionSummary
End Function
