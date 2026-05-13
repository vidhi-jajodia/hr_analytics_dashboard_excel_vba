Attribute VB_Name = "modNavigation"
Sub GoToEmployeeSearch()

    On Error GoTo ErrHandler

    With ThisWorkbook
        .Sheets("Employee_Search").Activate
        .Sheets("Employee_Search").Range("C4").Select
    End With

    Exit Sub

ErrHandler:
    MsgBox "Employee_Search sheet not found!", vbCritical

End Sub



Sub GoToDashboard()

    On Error GoTo ErrHandler

    With ThisWorkbook
        .Sheets("Dashboard").Activate
        .Sheets("Dashboard").Range("C1").Select
    End With

    Exit Sub

ErrHandler:
    MsgBox "Dashboard sheet not found!", vbCritical

End Sub



Sub GoToInsights()

    On Error GoTo ErrHandler

    With ThisWorkbook
        .Sheets("Insights").Activate
        .Sheets("Insights").Range("A2").Select
    End With

    Exit Sub

ErrHandler:
    MsgBox "Insights sheet not found!", vbCritical

End Sub



Sub GoToSettings()

    On Error GoTo ErrHandler

    With ThisWorkbook
        .Sheets("Settings").Activate
        .Sheets("Settings").Range("A2").Select
    End With

    Exit Sub

ErrHandler:
    MsgBox "Settings sheet not found!", vbCritical

End Sub


Sub OpenEmployeeForm()
    frmEmployeeMgmt.Show
End Sub

