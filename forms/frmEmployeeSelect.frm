Private Sub cmdSelect_Click()

    If lstEmployees.ListIndex = -1 Then
        MsgBox "Please select an employee", vbExclamation
        Exit Sub
    End If
    
    selectedRow = lstEmployees.List(lstEmployees.ListIndex, 3)
    
    Unload Me
    
    ' DIRECTLY LOAD SELECTED EMPLOYEE
    Call LoadSelectedEmployee

End Sub

Private Sub UserForm_Click()

End Sub
