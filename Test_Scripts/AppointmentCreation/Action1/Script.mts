systemutil.Run "https://wmstokedev101.tjx.com:11011"

'Sign in to the WMS TJX Application
findObject("j_username").Set "lak00099"
findObject("j_password").Click
Set wshell = CreateObject("WScript.Shell")
wshell.SendKeys "Welcome123"
wait 1
findObject("SignIn").click

findObject("HambergerIcon").click
'Creating an appointment
'findObject("MenuSearch").Set "Appointment"

'findObject("Appointments_Add").click

'findObject("Appointment_DateTime").set "11/30/18 08:43"
'findObject("Appointment_Type").select "Live Unload"
'findObject("Appointment_Save").click
' over ride code is required message is seen should check with team

'WAIT 2
'Set AppointNumVal=findObject("AppointmentNum_Rvalue").GetROProperty("innertext")
'msgbox AppointNumVal

'Set AppointNumVal=Browser("Sign In | Manhattan Associates,").Page("Manhattan Associates").Frame("Appointments | Manhattan").WebElement("AppointmentNum_RValue").GetROProperty("innertext")
'msgbox AppointNumVal






'Set obj=Browser("Sign In | Manhattan Associates,").Page("Sign In | Manhattan Associates,")
'Set obj=Browser("creationtime:=0").Page("title:=.*")
'
'Set objdesc = Description.Create
'objdesc("micclass").value = "WebEdit"
'objdesc("html tag").value = "INPUT"
'objdesc("name").value = "j_username"
'
'Set objchild = obj.ChildObjects(objdesc)
'msgbox objchild.count

' --- from test case ----

' show all for search
findObject("MenuSearch_ShowAllBtn").click
findObject("Appointments_Distribution").click
findObject("ShowAll_Appointments").click

findObject("AppointmentNum_Filter").set "111604"
findObject("AppointmentNum_SearchIcon").click
findObject("AppointmentNum_selectId").Select "111604"
findObject("Appointment_SelectBtn").click
findObject("Appointments_ApplyBtn").click
wait 8
findObject("Appointment_ChckBoxAppNum").set "ON"
findObject("Appointments_CheckinBtn").click
'findObject("Appointments_CarrierSearch").click
findObject("Appointment_carrierTxt").set "MISC"


findObject("Appointment_TrailerName").set "ttc1001"
findObject("Appointment_trailerTypeTxt").set "EQ1"
findObject("HambergerIcon").click
findObject("MenuSearch_ShowAllBtn").click
findObject("Appointments_Distribution").click
findObject("ShowAll_YardSlots").click
'Fetching the Yard ZOne Slot which is in open status
Set yardTable = findObject("Yard")

rCount =  yardTable.GetROProperty("rows")
cCount = yardTable.GetROProperty("cols")
For colIterator = 1 To cCount Step 1
    colName = yardTable.GetCellData(1,colIterator)
    If InStr(colName,"Status") > 0 Then
        statusColNum = colIterator
    End If
    If InStr(colName,"Yard Zone Slot") > 0 Then
        strYardZoneColNum = colIterator
    End If
Next

Set StokeYardTable = findObject("StokeYard")
rCount =  StokeYardTable.GetROProperty("rows")
cCount = StokeYardTable.GetROProperty("cols")
For rowIterator = 1 To rCount Step 1
    strStatusValue = StokeYardTable.GetCellData(rowIterator,statusColNum)
    If InStr(strStatusValue,"Open") > 0 Then
        strYardZoneSlot = StokeYardTable.GetCellData(rowIterator,strYardZoneColNum)
        print strYardZoneSlot
        Exit For
    End If
    'print strStatusValue
Next

print "Yard Zone Slot : "&strYardZoneSlot&" whose Status is "&strStatusValue



findObject("YardSlots_CloseIcon").click
findObject("Appointments_currentLocation").set strYardZoneSlot
findObject("CheckIn_CurrentLocationsearch").click
findObject("YardSlot_Selection").click
findObject("YardSlot_SelectBtn").click
findObject("YardSlot_DoneBtn").click
findObject("CheckIn_ConfirmBtn").click
If findObject("AppointmentCreated_SuccessTxt").Exist Then
print pass
else
print fail
End if

findObject("HambergerIcon").click

findObject("MenuSearch_ShowAllBtn").click
findObject("Appointments_Distribution").click
findObject("ShowAll_YardSlots").click
' Checking in yardslots wether the status of the yard zone has been changed from open to inuse
Set StokeYardTable = findObject("StokeYard")
rCount =  StokeYardTable.GetROProperty("rows")
cCount = StokeYardTable.GetROProperty("cols")
For rowIterator = 1 To rCount Step 1
    strexpectedYardZoneSlot = StokeYardTable.GetCellData(rowIterator,strYardZoneColNum)
    If InStr(strYardZoneSlot,strexpectedYardZoneSlot) > 0 Then
        strInUseStatus = StokeYardTable.GetCellData(rowIterator,statusColNum)
        print strInUseStatus
        If InStr(strInUseStatus,"") > 0 Then
        	Print "Yard Zone Slot has changed to "&strInUseStatus
        	Exit for 
        End If
        
    End If
    'print strStatusValue
Next

