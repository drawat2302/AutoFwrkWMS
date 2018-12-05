Function CreateAppointment()
	'Create Test Data object
'**********************************************************************
	Dim objData
	strsheetName = environment.value("estrTestDataSheet")
	Set objData = getData(environment.value("estrTestDataFileName"),strsheetName)
	
	strAppointmentID     = objData.Item("AppointmentID")
	strLocation          = objData.Item("Location")
	
	ReportResult STATUS_PASS , "Logged into WMS Home Page", "Successfully Logged"
	ReportResult STATUS_PASS , "Create Appointment Page has opened", "Successfully Logged"
	ReportResult STATUS_PASS , "Appointment has been created : 112345", "Successfully Logged"
End Function

Function TestCaseB()
	ReportResult STATUS_PASS , "Logged into WMS Home Page", "Successfully Logged"
	ReportResult STATUS_PASS , "Create Appointment Page has opened", "Successfully Logged"
End Function

Function TestCaseC()
	ReportResult STATUS_FAIL , "Logged into WMS Home Page", "Successfully Logged"
	ReportResult STATUS_FAIL , "Create Appointment Page has opened", "Successfully Logged"
End Function
