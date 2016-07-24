
'Add QTP native reporter to the party. This way we wont need 2 reporting commands - one to ReporterManager, and one to the QTP result log
Call oReporterManager.StartEngine("QTP", "QTP", "")

'Set up a complete backup log - the basics
Call oReporterManager.StartEngine("Text", "Backup", "File>C:\Logs\QTPLog.txt")
Call oReporterManager.StartEngine("XML", "XML Backup", "File>C:\Logs\QTPLog.XML")

'Make sure to copy the demo MDB database to c:\QTP Log.mdb
Call oReporterManager.StartEngine("DB", "DB Backup", "CreateNew>True")

'Set up statistics log - with it we can later produce the step pass/fail ratio with a click of a button
Call oReporterManager.StartEngine("Text", "Stats", "File>C:\Logs\Statistics.CSV|BodyTemplate>""%Status%, %StepName%, %Expected%, %Actual%, %Details%"" & vbcrlf|NewTestCaseTemplate>""General, Switched TestCase, , , %TestCaseName%"" & vbcrlf")

'Set up dedicated performence log - we can load it to excel and imiddiatly extract the time-step differences between relevant steps
Call oReporterManager.StartEngine("Text", "Performance", "File>C:\Logs\Performance.CSV|BodyTemplate>""%Time%, %StepName%"" & vbcrlf|NewTestCaseTemplate>"" "" ")

'Set up a top view excel log, so we'll have something pretty to send to management
Call oReporterManager.StartEngine("Excel", "TopView", "File>C:\Logs\TopView.xls")

'Set up a dedicated error & warnning logs
 Call oReporterManager.StartEngine("Text", "Errors", "File>C:\Logs\Errors.txt|BodyTemplate>""%Time%, %Status%, %StepName%, %Expected%, %Actual%"" & vbcrlf")
 Call oReporterManager.StartEngine("Winlog", "Win Errors", "ShowAs>QTP Errors")
 Call oReporterManager.StartEngine("ScreenCapture", "Error Captures", "Path>C:\Logs\|Prefix>Error -")

'Set up the filter for errors
	Call oReporterManager.AddFilter("RegEx", "Errors>Win Errors>Error Captures", "Pattern>Fail|WhatToSearch>Status")

'Set up a user popup that alerts whoever's watching that an error has occured  - enable major time saves as a script can be stopped imidiatly
 Call oReporterManager.StartEngine("User", "Errors Pop Up", "Timer>2")
	Call oReporterManager.AddFilter("RegEx", "Errors Pop Up", "Pattern>Fail|WhatToSearch>Status")

 'Test the logs :
Call oReporterManager.Report ("Pass", "Step 1" ,"Window Should Open" ,"All OK", "No Details")
Call oReporterManager.Report ("Warning", "Step 2" ,"Window Should Open" ,"It Didn't", "Some Details")
Call oReporterManager.Report ("Fail", "Step 3" ,"Window Should Open" ,"App. Crashed", "Bla Bla Bla")

