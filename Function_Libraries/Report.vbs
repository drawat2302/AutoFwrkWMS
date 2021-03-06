	'********************************************************************************************************
	'Method Name: CreateTestResultsFolder
	'Description: Creates Test Results Folder
	'Arguments: NULL
	'Returns Values: NULL
	'Date Of Creation: 28/12/2016
	'Last Modified: 28/12/2016
	'Author : Pritam Kadam
	'********************************************************************************************************
Function CreateTestResultsFolder ()
		Dim strFolderName, strScriptId, strYear, strMonth, strDay, strHour, strMinute, strSecond, strDate, fso, objFolder

		Set fso = CreateObject("Scripting.FileSystemObject")
		
		Wait(2)
        strDate = Now
		strDay = Day(strDate)
		strMonth = Month(strDate)
		strYear = Year(strDate)
		strHour = Hour(strDate)
		strMinute = Minute(strDate)
		strSecond = Second(strDate)
		strScriptId = Environment.Value("estrScriptId")
		strResultsHomeFolder = Environment.Value("estrResultsHomeFolder")

        strFolderName = strScriptId & "_"
        'strScriptId is being in env sheet
		strFolderName = strFolderName & strDay & "_" & strMonth & "_" & strYear & "_" & strHour & "_" & strMinute & "_" & strSecond

		Set objFolder = fso.CreateFolder(strResultsHomeFolder & "\" & strFolderName)
		'create environmnet variables for 
		strReportFolder = strResultsHomeFolder&"\"&strFolderName
		Environment("estrReportFolder") = strReportFolder
		CreateTestResultsFolder = strReportFolder
End Function
 '********************************************************************************************************
	'Module Name: AddComponentNameInExcelReport
	'Description: AddComponentInExcelReport
	'Arguments: 
	'Returns Values: NULL
	'Date Of Creation: /06/2016
	'Last Modified: 28/12/2016
	'Author: Rohan Karkanis
	'********************************************************************************************************
Sub AddComponentNameInExcelReport (strTCode)
	
		Dim strActionName, strReportExcelFileName,strComponentName
		'strActionName = Environment.Value("ActionName")
		'strComponentName  = Environment.Value("TestName")
		'strIterationID = Environment.Value("estrIterationId")  'Added this additionally
		Dim fso, objFile, strFileName,strStepNum
		Dim objExcel, objWorkBook, objWorkSheet2, strExcelFileName, intLastRow, intRowToWrite
		Const COLOR_WHITE = 2
		Const COLOR_BLUE = 5
		Const COLOR_LIGHT_GREY = 15
		Const ACTION_NAME_ROW_HEIGHT = 27
		Const COLOR_BLACK = 1
		Const XL_CENTER = -4108
		
		
				'strReportExcelFileName = Environment("estrExcelReportFileName")
				strReportExcelFileName = Environment("estrExcelReportFileName")
				
				Set objExcel = CreateObject("Excel.Application")	'Open Excel application
				Set objWorkBook = objExcel.WorkBooks.Open(strReportExcelFileName)		' Open Excel file
				Set objWorkSheet2 = objWorkBook.WorkSheets(2)
				'IsNewIteration = Environment("eIsNewIteration")
				strStepNum = Environment("estrStepNum") 
				strStepNum = strStepNum + 1
				Environment("estrStepNum") = strStepNum
				objExcel.DisplayAlerts = False
				
				'If isNewIteration = "True" Then
				intLastRow = objWorkSheet2.UsedRange.Rows.Count		'Last Row used
				intRowToWrite = intLastRow + 1
				objWorkSheet2.Cells (intRowToWrite,1).Interior.ColorIndex = COLOR_WHITE
				objWorkSheet2.Cells (intRowToWrite,1).Font.ColorIndex = COLOR_WHITE
				objWorkSheet2.Range("A" & intRowToWrite & ":E" & intRowToWrite).MergeCells = True
				objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).BorderAround(7)
				
				intLastRow = objWorkSheet2.UsedRange.Rows.Count		'Last Row used
				intRowToWrite = intLastRow + 1
				'objWorkSheet2.Cells (intRowToWrite,1).Value = "IterationID:"&strIterationID

				objWorkSheet2.Cells (intRowToWrite,1).HorizontalAlignment = XL_CENTER
				objWorkSheet2.Cells (intRowToWrite,1).Font.Bold = TRUE
				objWorkSheet2.Cells (intRowToWrite,1).Font.Size = 15
				objWorkSheet2.Cells (intRowToWrite,1).Interior.ColorIndex = COLOR_BLACK
				objWorkSheet2.Cells (intRowToWrite,1).Font.ColorIndex = COLOR_WHITE
				objWorkSheet2.Range("A" & intRowToWrite & ":E" & intRowToWrite).MergeCells = True
				objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).BorderAround(7)

				'isNewIteration = "False"
				'Environment("eIsNewIteration") = isNewIteration
				'End If
				
				intLastRow = objWorkSheet2.UsedRange.Rows.Count		'Last Row used
				intRowToWrite = intLastRow + 1
	
				objWorkSheet2.Rows (intRowToWrite).RowHeight = ACTION_NAME_ROW_HEIGHT
				'objWorkSheet2.Cells (intRowToWrite,1).Value = "IterationID:"&strIterationID&" "&strActionName
				objWorkSheet2.Cells (intRowToWrite,1).Value = "Test Case."&strStepNum&" "&strTCode
				'objWorkSheet2.Cells (intRowToWrite,1).Value = "Step Number "& strStepNum&" : "&strComponentName
				
				objWorkSheet2.Cells (intRowToWrite,1).HorizontalAlignment = XL_CENTER
				objWorkSheet2.Cells (intRowToWrite,1).Font.Bold = TRUE
				objWorkSheet2.Cells (intRowToWrite,1).Font.Size = 15
				objWorkSheet2.Cells (intRowToWrite,1).Interior.ColorIndex = COLOR_LIGHT_GREY
				objWorkSheet2.Cells (intRowToWrite,1).Font.ColorIndex = COLOR_BLUE
				objWorkSheet2.Range("A" & intRowToWrite & ":E" & intRowToWrite).MergeCells = True
				objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).BorderAround(7)
				
				'Save Workbook
				objWorkbook.SaveAs(strReportExcelFileName)
				objWorkBook.Close
				objExcel.Quit
				Set objExcel = Nothing		

'				If strStepNum > 0 then
'				 AddTestStepDescInExcelReport
'				End if 
	End Sub
	'********************************************************************************************************
	'Module Name: CreateExcelReportTemplate
	'Description: CreateExcelReportTemplate
	'Arguments: NULL
	'Returns Values: NULL
	'Date Of Creation: 28/12/2016
	'Last Modified: 28/12/2016
	'Author: Rohan Karkhanis
	'********************************************************************************************************
	
	Function CreateExcelReportTemplate(strResultsFolder)

		Dim strFileName, strScenarioId, strYear, strMonth, strDay, strHour, strMinute, strSecond, strDate, objExcel, objWorkBook, objWorkSheet1, intCtr, objRange, objWorkSheet2, objWorkSheet3, strScriptId

		Const COLOR_DARK_BLUE = 25
		Const COLOR_WHITE = 2
		Const COLOR_BLUE = 5
		Const TEST_RESULTS = "Test Execution Details"
		Const COLUMN_ONE_WIDTH = 16
		Const COLUMN_TWO_WIDTH = 42
		Const COLUMN_THREE_WIDTH = 56
		Const COLUMN_FOUR_WIDTH = 13
		Const COLUMN_FIVE_WIDTH = 61
		Const EXECUTION_SUMMARY = "Execution Summary"
		Const XL_CENTER = -4108
			
		
        strDate = Now
		strDay = Day(strDate)
		strMonth = Month(strDate)
		strYear = Year(strDate)
		strHour = Hour(strDate)
		strMinute = Minute(strDate)
		strSecond = Second(strDate)
		
		'Get the Report Folder name from the Environment
		
		strScriptId = Environment.Value("estrScriptId")
		
		
		'strReportFolder = Environment("estrTestResultHomeFolder")
		
'		strFileName = QTP_HOME_FOLDER & "\"
'		strFileName = strFileName & TEST_RESULTS_FOLDER & "\"
		strFileName = strResultsFolder & "\"
		strFileName = strFileName & strScriptId & "_Test Results_"
		strFileName = strFileName & strDay & "_" & strMonth & "_" & strYear & "_" & strHour & "_" & strMinute & "_" & strSecond & ".xlsx"
'		strFileName = strFileName & strDay & "_" & strMonth & "_" & strYear & "_" & strHour & "_" & strMinute & "_" & strSecond & ".xls"

		'Create excel object
		Set objExcel = CreateObject("Excel.Application")

		'Add workbook
        Set objWorkBook = objExcel.WorkBooks.Add()

		objExcel.DisplayAlerts = False

		'Delete two empty worksheets
		'Set objWorkSheet3 = objWorkBook.WorkSheets(3)
		'objWorkSheet3.Delete
		objWorkBook.Worksheets.Add
		Set objWorkSheet1 = objWorkBook.WorkSheets(1)
		objWorkSheet1.Name = EXECUTION_SUMMARY

		'Rename WorkSheet
		Set objWorkSheet2 = objWorkBook.WorkSheets(2)
		objWorkSheet2.Name = "RESULTS"

		'Write Header
		objWorkSheet2.Columns(1).ColumnWidth = COLUMN_ONE_WIDTH
		objWorkSheet2.Columns(2).ColumnWidth = COLUMN_TWO_WIDTH
		objWorkSheet2.Columns(3).ColumnWidth = COLUMN_THREE_WIDTH
		objWorkSheet2.Columns(4).ColumnWidth = COLUMN_FOUR_WIDTH
		objWorkSheet2.Columns(5).ColumnWidth = COLUMN_FIVE_WIDTH
		
		objWorkSheet2.Pictures.Insert ("C:\WMS\Environment Variables\DeloitteLogo.gif")
		objWorkSheet2.Cells (1,1).Value = REPORT_HEADER
		objWorkSheet2.Cells (1,1).Font.Bold = TRUE
		objWorkSheet2.Cells (1,1).Font.Size = 26
		objWorkSheet2.Cells (1,1).Font.ColorIndex = COLOR_DARK_BLUE
		objWorkSheet2.Cells (1,1).HorizontalAlignment = XL_CENTER
		objWorkSheet2.Range("A1:E1").MergeCells = True
		
		objWorkSheet2.Range("A2:E4").MergeCells = True

		objWorkSheet2.Cells (5,1).Value = TEST_RESULTS
		objWorkSheet2.Cells (5,1).Font.Bold = TRUE
		objWorkSheet2.Cells (5,1).HorizontalAlignment = XL_CENTER	
		objWorkSheet2.Cells (5,1).Font.Size = 20
		objWorkSheet2.Cells (5,1).Interior.ColorIndex = COLOR_DARK_BLUE
		objWorkSheet2.Cells (5,1).Font.ColorIndex = COLOR_WHITE
		objWorkSheet2.Range ("A5:E5").MergeCells = True

		objWorkSheet2.Range("A6:E6").MergeCells = True
		
		objWorkSheet2.Cells (7,1).Value = "Execution summary."
		objWorkSheet2.Hyperlinks.Add objWorkSheet2.Cells (7, 1), "", "'" & EXECUTION_SUMMARY & "'!A1"
		objWorkSheet2.Cells (7,1).Font.Bold = TRUE
		objWorkSheet2.Cells (7,1).Font.Underline = TRUE
		objWorkSheet2.Cells (7,1).Font.Size = 12
		objWorkSheet2.Cells (7,1).Font.ColorIndex = COLOR_BLUE
		objWorkSheet2.Range("A7:E7").MergeCells = True
		objWorkSheet2.Range("A7:E7").Interior.ColorIndex = COLOR_WHITE

        objWorkSheet2.Range("A8:E8").MergeCells = True
		
        'Save Workbook
		objWorkbook.SaveAs(strFileName)
		objWorkBook.Close
		objExcel.Quit
        Set objExcel = Nothing

' 		Set objShell = CreateObject("WScript.Shell")
'		Set objEnv = objShell.Environment("USER")
'		
'		objEnv("estrExcelReportFileName") = strFileName
		
		Environment.Value("estrExcelReportFileName") = strFileName
		
		CreateExcelReportTemplate = strFileName
End Function
Sub AddColumnNamesInReport ()
	
		Dim fso, objFile, strHTMLFileName, strReportExcelFileName
		Dim objExcel, objWorkBook, objWorkSheet2, strExcelFileName, intLastRow, intRowToWrite

		Const COLOR_LIGHT_GREY = 15
		Const FIELD_NAME_ROW_HEIGHT = 25
		Const XL_CENTER = -4108
		
			strReportExcelFileName = Environment("estrExcelReportFileName")
			Set objExcel = CreateObject("Excel.Application")	'Open Excel application
			Set objWorkBook = objExcel.WorkBooks.Open(strReportExcelFileName)		' Open Excel file
			Set objWorkSheet2 = objWorkBook.WorkSheets(2)

			objExcel.DisplayAlerts = False
			
			intLastRow = objWorkSheet2.UsedRange.Rows.Count		'Last Row used
			intRowToWrite = intLastRow + 1
			
			objWorkSheet2.Rows (intRowToWrite).RowHeight = FIELD_NAME_ROW_HEIGHT
			objWorkSheet2.Cells (intRowToWrite,1).Value = "Validations"
			objWorkSheet2.Cells (intRowToWrite,2).Value = "Validation Step"
			objWorkSheet2.Cells (intRowToWrite,3).Value = "Actual Result"
			objWorkSheet2.Cells (intRowToWrite,4).Value = "Status"
			objWorkSheet2.Cells (intRowToWrite,5).Value = "Screenshot Link"
			objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).Font.Bold = TRUE
			objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).Font.Size = 13
			objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).Interior.ColorIndex = COLOR_LIGHT_GREY
			objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).HorizontalAlignment = XL_CENTER	
			objWorkSheet2.Range ("A" & intRowToWrite & ":A" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("B" & intRowToWrite & ":B" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("C" & intRowToWrite & ":C" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("D" & intRowToWrite & ":D" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("E" & intRowToWrite & ":E" & intRowToWrite).BorderAround(7)
         
			'Save Workbook
			objWorkbook.SaveAs(strReportExcelFileName)
			objWorkBook.Close
			objExcel.Quit
			Set objExcel = Nothing			

	End Sub
Sub ReportResult (strStatus , strStepName, strStatMessage)
	   'Dim strStepName, strFileName, strFullFileName, strReportFolder, strNewStatus
	   strFileName = ""
	   strResultsHomeFolder = Environment("estrTestResultHomeFolder")
	   'objMessagesRecordSet.MoveFirst
'	   Do
'				   If NOT(objMessagesRecordSet.EOF) Then
'					   If ((objMessagesRecordSet(0) = intMajorCode)AND(objMessagesRecordSet(1) = intMinorCode)) Then
'						   strStepName = objMessagesRecordSet(2)
'						   Exit Do
'					   Else
'						   objMessagesRecordSet.MoveNext
'					   End If
'				   Else
'					   Exit Do
'				   End If
'	   Loop
		
		'strStepName = ""'argumnet received from the fn
		'Excel Report
		intGlobalCheckPointCount = Environment("eintGlobalCheckPointCount")
		intGlobalCheckPointCount = intGlobalCheckPointCount + 1
		Environment("eintGlobalCheckPointCount") = intGlobalCheckPointCount
			
			Dim objExcel, objWorkBook, objWorkSheet2, strReportExcelFileName, intLastRow, intRowToWrite, strCellColor

			Const COLOR_GREEN = 10
			Const COLOR_YELLOW = 6
			Const COLOR_RED = 3
			Const COLOR_WHITE = 2
			Const STEP_ROW_HEIGHT = 15
			CAPTURE_SCREENSHOT_LEVEL = 3
			Const XL_CENTER = -4108
			CAPTURE_SCREENSHOT = "Y"
'			Const STATUS_PASS = 0
'			Const STATUS_FAIL = 1
			
            Select Case strStatus
				Case 0
					strNewStatus = "PASS"
					strCellColor = COLOR_GREEN
					Environment("estrTestCaseStatus") = "PASS"
				Case 1
					strNewStatus = "FAIL"
					strCellColor = COLOR_RED
					Environment("estrTestCaseStatus") = "FAIL"
				Case 2
					strNewStatus = "DONE"
					strCellColor = COLOR_WHITE
				Case 3
					strNewStatus = "WARNING"
					strCellColor = COLOR_YELLOW
					Environment("estrTestCaseStatus") = "FAIL"
			End Select				
			
		   	strReportExcelFileName = Environment("estrExcelReportFileName")
			Set objExcel = CreateObject("Excel.Application")
			'objExcel.Visible = True	'Open Excel application
			Set objWorkBook = objExcel.WorkBooks.Open(strReportExcelFileName)		' Open Excel file
			Set objWorkSheet2 = objWorkBook.WorkSheets(2)

			objExcel.DisplayAlerts = False
			
			intLastRow = objWorkSheet2.UsedRange.Rows.Count		'Last Row used
			intRowToWrite = intLastRow + 1
			objWorkSheet2.Rows (intRowToWrite).RowHeight = STEP_ROW_HEIGHT
			
			objWorkSheet2.Range ("A" & intRowToWrite & ":E" & intRowToWrite).HorizontalAlignment = XL_CENTER
			
			'If (strStepName = FALSE) Then
'				objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
'				objWorkSheet2.Cells (intRowToWrite,2).Value = "Code Error"
'				objWorkSheet2.Cells (intRowToWrite,3).Value = "Please verify Major Code and Minor Code."
'				objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
'				objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
			'Else
			   If (CAPTURE_SCREENSHOT = "Y") Then
				   If (CAPTURE_SCREENSHOT_LEVEL = "3") Then
 						If (strFileName = "") Then
							strFileName = CaptureScreenshot()
						End If
                        			objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
						objWorkSheet2.Cells (intRowToWrite,2).Value = strStepName
						objWorkSheet2.Cells (intRowToWrite,3).Value = strStatMessage
						objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
						objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
                        			objWorkSheet2.Cells (intRowToWrite,5).Value = strFileName
					ElseIf (CAPTURE_SCREENSHOT_LEVEL = "2") Then
						If (strStatus <> STATUS_PASS) Then
							If (strFileName = "") Then
								strFileName = CaptureScreenshot()
							End If
							objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
							objWorkSheet2.Cells (intRowToWrite,2).Value = strStepName
							objWorkSheet2.Cells (intRowToWrite,3).Value = strStatMessage
							objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
							objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
							objWorkSheet2.Cells (intRowToWrite,5).Value = strFileName
						Else
							objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
							objWorkSheet2.Cells (intRowToWrite,2).Value = strStepName
							objWorkSheet2.Cells (intRowToWrite,3).Value = strStatMessage
							objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
							objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
						End If
					ElseIf (CAPTURE_SCREENSHOT_LEVEL = "1") Then
						If (strStatus = STATUS_FAIL) Then
							If (strFileName = "") Then
								strFileName = CaptureScreenshot()
							End If
							objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
							objWorkSheet2.Cells (intRowToWrite,2).Value = strStepName
							objWorkSheet2.Cells (intRowToWrite,3).Value = strStatMessage
							objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
							objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
							objWorkSheet2.Cells (intRowToWrite,5).Value = strFileName
						Else
							objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
							objWorkSheet2.Cells (intRowToWrite,2).Value = strStepName
							objWorkSheet2.Cells (intRowToWrite,3).Value = strStatMessage
							objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
							objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
						End If
				   End If
			   Else
					objWorkSheet2.Cells (intRowToWrite,1).Value = intGlobalCheckPointCount
					objWorkSheet2.Cells (intRowToWrite,2).Value = strStepName
					objWorkSheet2.Cells (intRowToWrite,3).Value = strStatMessage
					objWorkSheet2.Cells (intRowToWrite,4).Value = strNewStatus
					objWorkSheet2.Cells (intRowToWrite,4).Font.ColorIndex = COLOR_WHITE
			   End If
			'End If

			If (NOT((IsNull(strFileName)) OR (strFileName = ""))) Then
				objWorkSheet2.Hyperlinks.Add objWorkSheet2.Cells (intRowToWrite, 5), strFileName
			End If

			objWorkSheet2.Cells (intRowToWrite,4).Interior.ColorIndex = strCellColor
			objWorkSheet2.Range ("A" & intRowToWrite & ":A" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("B" & intRowToWrite & ":B" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("C" & intRowToWrite & ":C" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("D" & intRowToWrite & ":D" & intRowToWrite).BorderAround(7)
			objWorkSheet2.Range ("E" & intRowToWrite & ":E" & intRowToWrite).BorderAround(7)
			
			'Save Workbook
			objWorkbook.SaveAs(strReportExcelFileName)
			objWorkBook.Close
			objExcel.Quit
			Set objExcel = Nothing
End Sub
	'********************************************************************************************************
	'Method Name: CaptureScreenshot
	'Description: Saves screnshot of SAP windown for all validations. This is called in ReportResult function
	'Arguments: NULL
	'Returns Values: Screenshot file name
	'Date Of Creation: 01/01/2017
	'Last Modified: 01/01/2017
	'Author: Ashish Vartak
	'********************************************************************************************************
Function CaptureScreenshot ()
		Dim strFileName, strScriptId, strYear, strMonth, strDay, strHour, strMinute, strSecond, strDate, strReportFolder, strFullFileName

		Wait(2)
        strDate = Now
		strDay = Day(strDate)
		strMonth = Month(strDate)
		strYear = Year(strDate)
		strHour = Hour(strDate)
		strMinute = Minute(strDate)
		strSecond = Second(strDate)
		
		'strScriptId = Environment.Value("estrScriptId")
		'strReportFolder = Environment.Value("estrResultsHomeFolder")

		strFileName = Environment.Value("TestName") &"_StepNumber " & intGlobalCheckPointCount & "_" & strDay & "_" & strMonth & "_" & strYear & "_" & strHour & "_" & strMinute & "_" & strSecond & ".png"

		strFullFileName = Environment.Value("estrResultsFolder") & "\" & strFileName
		
		Desktop.CaptureBitmap strFullFileName
		CaptureScreenshot = strFileName 
End Function
'********************************************************************************************************
	'Module Name: ReportExecutionSummary
	'Description: Reports execution summary at the end of the execution
	'Arguments: Scenario
	'Returns Values: NULL
	'Date Of Creation: 23/04/2017
	'Last Modified: 05/08/2017
	'Author: Rohan Karkhanis
	
	'Changes
	'Author: Pritam Kadam
	'Description: intPassTcode,intFailTcode will be used to generate count of pass/fail count
	'********************************************************************************************************
Sub ReportExecutionSummary ()
	
		Dim fso, objFile, strHTMLFileName, intPassedTestCases, intFailedTestCases, intTotalTestCases, strExecutionStartTime, strExecutionEndTime, strRunName, strTimeElapsed, strResultName, strRunTimeTestDataMatrixFile
		Dim strReportExcelFileName, strScriptName, strRunTimeTestDataFileName,strScenarioId,intTotalNumberOfIterations,strTestCaseStatus
		Dim intPassTcode, intFailTcode
		
		strScriptName = "WMS Automation"
		Const TEST_EXECUTION_SUMMARY = "Test Execution Summary"
		Const SCRIPT_NAME = "Scenario Name"
		Const START_TIME = "Start Time"
		Const END_TIME = "End Time"
		Const TIME_ELAPSED = "Time Elapsed"
		Const TOTAL_EXECUTED = "Total No of Scenarios Tested"
		Const TOTAL_PASSED = "No of Passed Test Cases"
		Const TOTAL_FAILED = "No of Failed Test Cases"
		Const RUN_TIME_TEST_DATA_FILE= "Run Time Test Data File"
		Const OVERALL_STATUS = "Overall Status of Scenario: " 
		Const COLOR_LIGHT_GREY = 15
		Const COLOR_GREEN = 10
		Const COLOR_RED = 3
		Const TOTAL_NUMBER_OF_ITERATIONS = "Total Number of Iterations" 
		
		'intPassTcode = Environment.Value("intPassTCode")
		'intFailTcode = Environment.Value("intFailTCode")
        intPassedTestCases = Environment.Value("intPassTestCaseCounter")
		intFailedTestCases = Environment.Value("intFailTestCaseCounter")
		strExecutionStartTime = Environment("eExecutionStartTime")
		strExecutionEndTime = Environment("eExecutionEndTime")
		strResultName = Environment("estrReportFolder")
		strTimeElapsed = getElapsedTime(strExecutionStartTime, strExecutionEndTime)
		strRunTimeTestDataFileName = Environment("TestSuite")
		strScenarioId = Environment.Value("estrScriptId")'get test suite name
		print strScenarioId
		'intTotalNumberOfIterations = Environment("estrNumOfItr")
		 strTestCaseStatus = Environment.value("estrTestCaseStatus")
		intTotalTestCases = intPassedTestCases + intFailedTestCases
			Dim objExcel, objWorkBook, objWorkSheet1, strExcelFileName

			Const COLUMN_TWO_WIDTH = 44
			Const COLUMN_THREE_WIDTH = 44
			Const COLOR_DARK_BLUE = 25
			Const COLOR_WHITE = 2
			Const COLOR_YELLOW = 6

			strReportExcelFileName = Environment("estrExcelReportFileName")
			Set objExcel = CreateObject("Excel.Application")
			Set objWorkBook = objExcel.WorkBooks.Open(strReportExcelFileName)		' Open Excel file
			Set objWorkSheet1 = objWorkBook.WorkSheets(1)
			objExcel.DisplayAlerts = False
			
			objWorkSheet1.Columns(2).ColumnWidth = COLUMN_TWO_WIDTH
			objWorkSheet1.Columns(3).ColumnWidth = COLUMN_THREE_WIDTH

			objWorkSheet1.Cells (2,2).Value = TEST_EXECUTION_SUMMARY
			objWorkSheet1.Cells (2,2).Font.Bold = TRUE
			objWorkSheet1.Cells (2,2).Font.Size = 14
			objWorkSheet1.Cells (2,2).Interior.ColorIndex = COLOR_LIGHT_GREY
			objWorkSheet1.Cells (2,2).Font.ColorIndex = COLOR_DARK_BLUE
			objWorkSheet1.Cells (2,2).HorizontalAlignment = XL_CENTER
			
			objWorkSheet1.Range("B2:C2").MergeCells = True
			objWorkSheet1.Range ("B2:C2").BorderAround(7)

			objWorkSheet1.Cells (3,2).Value = SCRIPT_NAME
			objWorkSheet1.Cells (3,2).Font.Bold = TRUE
			objWorkSheet1.Cells (3,2).Font.Size = 12				
			objWorkSheet1.Cells (3,3).Value = strScenarioId
			objWorkSheet1.Cells (3,3).Font.Bold = TRUE
            objWorkSheet1.Cells (3,3).Font.Size = 12
			objWorkSheet1.Cells (3,3).HorizontalAlignment = XL_RIGHT

			objWorkSheet1.Cells (4,2).Value = START_TIME
			objWorkSheet1.Cells (4,2).Font.Bold = TRUE
			objWorkSheet1.Cells (4,2).Font.Size = 10				
			objWorkSheet1.Cells (4,3).Value = strExecutionStartTime
            objWorkSheet1.Cells (4,3).Font.Size = 10
			objWorkSheet1.Cells (4,3).HorizontalAlignment = XL_RIGHT

			objWorkSheet1.Cells (5,2).Value = END_TIME
			objWorkSheet1.Cells (5,2).Font.Bold = TRUE
			objWorkSheet1.Cells (5,2).Font.Size = 10				
			objWorkSheet1.Cells (5,3).Value = strExecutionEndTime
            objWorkSheet1.Cells (5,3).Font.Size = 10			
			objWorkSheet1.Cells (5,3).HorizontalAlignment = XL_RIGHT

			objWorkSheet1.Cells (6,2).Value = TIME_ELAPSED
			objWorkSheet1.Cells (6,2).Font.Bold = TRUE
			objWorkSheet1.Cells (6,2).Font.Size = 10				
			objWorkSheet1.Cells (6,3).Value = strTimeElapsed
            objWorkSheet1.Cells (6,3).Font.Size = 10			
			objWorkSheet1.Cells (6,3).HorizontalAlignment = XL_RIGHT

			'objWorkSheet1.Cells (8,2).Value = TOTAL_NUMBER_OF_ITERATIONS
			'objWorkSheet1.Cells (8,2).Font.Bold = TRUE
			'objWorkSheet1.Cells (8,2).Font.Size = 10				
			'objWorkSheet1.Cells (8,3).Value = intTotalNumberOfIterations
            'objWorkSheet1.Cells (8,3).Font.Size = 10			
			'objWorkSheet1.Cells (8,3).HorizontalAlignment = XL_RIGHT
			
		
			objWorkSheet1.Cells (8,2).Value = TOTAL_PASSED
			objWorkSheet1.Cells (8,2).Font.Bold = TRUE
			objWorkSheet1.Cells (8,2).Font.Size = 10				
			objWorkSheet1.Cells (8,3).Value = intPassedTestCases
            objWorkSheet1.Cells (8,3).Font.Size = 10
			objWorkSheet1.Cells (8,3).HorizontalAlignment = XL_RIGHT

			objWorkSheet1.Cells (9,2).Value = TOTAL_FAILED
			objWorkSheet1.Cells (9,2).Font.Bold = TRUE
			objWorkSheet1.Cells (9,2).Font.Size = 10				
			objWorkSheet1.Cells (9,3).Value = intFailedTestCases
			'objWorkSheet1.Cells (9,3).Value = intFailTcode
            objWorkSheet1.Cells (9,3).Font.Size = 10			
			objWorkSheet1.Cells (9,3).HorizontalAlignment = XL_RIGHT		

			objWorkSheet1.Cells (10,2).Value = RUN_TIME_TEST_DATA_FILE
			objWorkSheet1.Cells (10,2).Font.Bold = TRUE
			objWorkSheet1.Cells (10,2).Font.Size = 10				
			objWorkSheet1.Cells (10,3).Value = strRunTimeTestDataFileName
			objWorkSheet1.Hyperlinks.Add objWorkSheet1.Cells (10,3), strRunTimeTestDataFileName
            objWorkSheet1.Cells (10,3).Font.Size = 10			
			objWorkSheet1.Cells (10,3).HorizontalAlignment = XL_RIGHT			
			
			objWorkSheet1.Cells (11,2).Value = OVERALL_STATUS &strScenarioId
			objWorkSheet1.Cells (11,2).Font.Bold = TRUE
			objWorkSheet1.Cells (11,2).Font.Size = 10				
		If strTestCaseStatus = "PASS" THEN
			objWorkSheet1.Cells (11,3).Value = strTestCaseStatus
			objWorkSheet1.Cells (11,3).Interior.ColorIndex = COLOR_GREEN
		Else	
			objWorkSheet1.Cells (11,3).Value = strTestCaseStatus
			objWorkSheet1.Cells (11,3).Interior.ColorIndex = COLOR_RED
		End If
			'objWorkSheet1.Hyperlinks.Add objWorkSheet1.Cells (10,3), strRunTimeTestDataFileName
            objWorkSheet1.Cells (11,3).Font.Size = 10			
			objWorkSheet1.Cells (11,3).HorizontalAlignment = XL_RIGHT

			objWorkSheet1.Range ("B3:C12").BorderAround(7)
			
			'Save Workbook
			objWorkbook.SaveAs(strReportExcelFileName)
			objWorkBook.Close
			objExcel.Quit
			Set objExcel = Nothing
End Sub
Function getElapsedTime (dtStart, dtEnd)
	
		Dim dtDiff, intSec, intMin, intHour
		intSec = DateDiff("s", dtStart, dtEnd)
		If (intSec > 3600) Then
			intHour = Fix(intSec/3600)
			intSec = intSec Mod 3600
			If (intSec > 60) Then
				intMin = Fix(intSec/60)
				intSec = intSec Mod 60
			Else
				intMin = 0
			End If
		Else
			intHour = 0
			If (intSec > 60) Then
				intMin = Fix(intSec/60)
				intSec = intSec Mod 60
			Else
				intMin = 0
			End If			
		End If

		dtDiff = intHour & " Hours, " & intMin & " Minutes, " & intSec & " Seconds"
		getElapsedTime = dtDiff
		
	End Function
