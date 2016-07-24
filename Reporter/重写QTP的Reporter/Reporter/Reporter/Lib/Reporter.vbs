'Override it with our custom class
Dim Reporter
Set Reporter = New clsReporter

'Define a funnel function to be called from the test actions
Public Function GetReporter
    Set GetReporter = Reporter
End Function

'Class definition
'In the example, our class just reporter to a text file
Class clsReporter
    Dim oFileReporter
    Public Sub ReportEvent(iStatus, sStepName, sDetails)
        oFileReporter.AppendAllText "c:\log.txt", iStatus & " : " & sStepName & " - " & sDetails & vbcrlf
    End Sub

    Private Sub Class_Initialize
       Set oFIleReporter = DotNetFactory("System.IO.File")
    End Sub
End Class
