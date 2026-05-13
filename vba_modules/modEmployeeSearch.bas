Attribute VB_Name = "modEmployeeSearch"
Option Explicit

Public wsData As Worksheet
Public wsSearch As Worksheet

Public matchRows() As Long
Public selectedRow As Long

Sub InitSheets()

    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsSearch = ThisWorkbook.Sheets("Employee_Search")

End Sub




Sub SearchEmployee()

    Call InitSheets
    
    Dim searchValue As String
    Dim i As Long
    Dim lastRow As Long
    Dim matchCount As Long
    
    Dim empID As String, empName As String
    
    Dim colEmployeeID As Long, colEmployeeName As Long, colDepartment As Long
    Dim colJobRole As Long, colSalary As Long, colAttrition As Long
    Dim colSatisfaction As Long, colOvertime As Long, colPromotionGap As Long
    
    searchValue = LCase(Trim(wsSearch.Range("C4").Value))
    
    If searchValue = "" Then
        MsgBox "Please enter Employee ID or Name", vbExclamation
        Exit Sub
    End If
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' map columns safely
    Call MapColumns(colEmployeeID, colEmployeeName, colDepartment, colJobRole, _
                    colSalary, colAttrition, colSatisfaction, colOvertime, colPromotionGap)
    
    matchCount = 0
    ReDim matchRows(1 To 1)
    
    For i = 2 To lastRow
        
        empID = LCase(wsData.Cells(i, colEmployeeID).Value)
        empName = LCase(wsData.Cells(i, colEmployeeName).Value)
        
        If empID = searchValue Or InStr(empName, searchValue) > 0 Then
            
            matchCount = matchCount + 1
            
            If matchCount = 1 Then
                ReDim matchRows(1 To 1)
            Else
                ReDim Preserve matchRows(1 To matchCount)
            End If
            
            matchRows(matchCount) = i
            
        End If
        
    Next i
    
    If matchCount = 0 Then
    
        MsgBox "Employee not found", vbExclamation
        Exit Sub
        
    ElseIf matchCount = 1 Then
    
        Call LoadEmployee(matchRows(1), colEmployeeName, colDepartment, colJobRole, _
                          colSalary, colAttrition, colSatisfaction, colOvertime, colPromotionGap)
    
    Else
    
        Call ShowSelector(colEmployeeID, colEmployeeName, colDepartment)
    
    End If

End Sub



Sub MapColumns(ByRef colEmployeeID As Long, ByRef colEmployeeName As Long, ByRef colDepartment As Long, _
               ByRef colJobRole As Long, ByRef colSalary As Long, ByRef colAttrition As Long, _
               ByRef colSatisfaction As Long, ByRef colOvertime As Long, ByRef colPromotionGap As Long)

    Dim c As Range
    
    For Each c In wsData.Rows(1).Cells
        
        Select Case LCase(Trim(c.Value))
            
            Case "employeeid": colEmployeeID = c.Column
            Case "employeename": colEmployeeName = c.Column
            Case "department": colDepartment = c.Column
            Case "jobrole": colJobRole = c.Column
            Case "salary": colSalary = c.Column
            Case "attrition": colAttrition = c.Column
            Case "jobsatisfaction": colSatisfaction = c.Column
            Case "overtime": colOvertime = c.Column
            Case "yearssincelastpromotion": colPromotionGap = c.Column
            
        End Select
        
    Next c

End Sub







Sub LoadEmployee(ByVal i As Long, _
                 colEmployeeName As Long, colDepartment As Long, colJobRole As Long, _
                 colSalary As Long, colAttrition As Long, colSatisfaction As Long, _
                 colOvertime As Long, colPromotionGap As Long)

    Dim riskScore As Integer
    
    riskScore = 0
    
    If LCase(wsData.Cells(i, colOvertime).Value) = "yes" Then riskScore = riskScore + 2
    If wsData.Cells(i, colSatisfaction).Value <= 2 Then riskScore = riskScore + 2
    If wsData.Cells(i, colPromotionGap).Value >= 3 Then riskScore = riskScore + 2
    If wsData.Cells(i, colSalary).Value < 40000 Then riskScore = riskScore + 1
    
    wsSearch.Range("C7").Value = wsData.Cells(i, colEmployeeName).Value
    wsSearch.Range("C8").Value = wsData.Cells(i, colDepartment).Value
    wsSearch.Range("C9").Value = wsData.Cells(i, colJobRole).Value
    wsSearch.Range("C10").Value = wsData.Cells(i, colSalary).Value
    wsSearch.Range("C11").Value = wsData.Cells(i, colAttrition).Value
    wsSearch.Range("C12").Value = wsData.Cells(i, colSatisfaction).Value
    wsSearch.Range("C13").Value = wsData.Cells(i, colOvertime).Value
    wsSearch.Range("C14").Value = wsData.Cells(i, colPromotionGap).Value
    wsSearch.Range("C15").Value = riskScore & "/7"

End Sub








Sub ShowSelector(colEmployeeID As Long, colEmployeeName As Long, colDepartment As Long)

    Dim i As Long
    
    frmEmployeeSelect.lstEmployees.Clear
    
    For i = 1 To UBound(matchRows)
        
        frmEmployeeSelect.lstEmployees.AddItem wsData.Cells(matchRows(i), colEmployeeID).Value
        
        frmEmployeeSelect.lstEmployees.List(frmEmployeeSelect.lstEmployees.ListCount - 1, 1) = _
            wsData.Cells(matchRows(i), colEmployeeName).Value
        
        frmEmployeeSelect.lstEmployees.List(frmEmployeeSelect.lstEmployees.ListCount - 1, 2) = _
            wsData.Cells(matchRows(i), colDepartment).Value
        
        frmEmployeeSelect.lstEmployees.List(frmEmployeeSelect.lstEmployees.ListCount - 1, 3) = matchRows(i)
        
    Next i
    
    frmEmployeeSelect.Show

End Sub






Sub LoadSelectedEmployee()

    Call InitSheets
    
    Dim colEmployeeID As Long, colEmployeeName As Long, colDepartment As Long
    Dim colJobRole As Long, colSalary As Long, colAttrition As Long
    Dim colSatisfaction As Long, colOvertime As Long, colPromotionGap As Long
    
    Call MapColumns(colEmployeeID, colEmployeeName, colDepartment, colJobRole, _
                    colSalary, colAttrition, colSatisfaction, colOvertime, colPromotionGap)
    
    Call LoadEmployee(selectedRow, colEmployeeName, colDepartment, colJobRole, _
                      colSalary, colAttrition, colSatisfaction, colOvertime, colPromotionGap)

End Sub





Sub ClearEmployeeSearch()

    Call InitSheets
    
    ' Clear input
    wsSearch.Range("C4").ClearContents
    
    ' Clear output
    wsSearch.Range("C7:C15").ClearContents
    
    ' Reset variables
    Erase matchRows
    selectedRow = 0
    
    MsgBox "Search cleared!", vbInformation

End Sub



