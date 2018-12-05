Set App = CreateObject("QuickTest.Application")
App.Launch
App.Visible = True
App.WindowState = "Maximized"' Maximize the QuickTest window
App.ActivateView "ExpertView"' Display the Expert View
App.open "..\..\Driver_Script\DriverScript", False
'Opens the test in editable mode
' Get the test object.
 Set qtTest = App.Test
 ' Run the script using the default test results options.
 qtTest.Run
 ' Close the test.
 qtTest.Close 'Close the test
 ' Close QuickTest Professional.
 App.Quit
 ' Release the created objects.
 set qtTest = nothing
 set App = nothing