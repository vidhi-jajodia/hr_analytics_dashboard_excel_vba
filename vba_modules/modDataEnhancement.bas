Attribute VB_Name = "modDataEnhancement"
Sub AddSalaryBandColumn()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim incomeCol As Variant
    Dim newCol As Long
    Dim i As Long
    Dim salaryVal As Double
    
    Set ws = ThisWorkbook.Sheets("Cleaned_Data")
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Locate MonthlyIncome column
    incomeCol = Application.Match("MonthlyIncome", ws.Rows(1), 0)
    
    If IsError(incomeCol) Then
        MsgBox "MonthlyIncome column not found!", vbCritical
        Exit Sub
    End If
    
    ' Check if SalaryBand already exists
    On Error Resume Next
    newCol = Application.Match("SalaryBand", ws.Rows(1), 0)
    On Error GoTo 0
    
    If IsError(newCol) Or newCol = 0 Then
        newCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, newCol).Value = "SalaryBand"
    End If
    
    ' Populate SalaryBand
    For i = 2 To lastRow
        
        salaryVal = ws.Cells(i, incomeCol).Value
        
        If salaryVal < 50000 Then
            ws.Cells(i, newCol).Value = "Low"
        ElseIf salaryVal < 75000 Then
            ws.Cells(i, newCol).Value = "Medium"
        Else
            ws.Cells(i, newCol).Value = "High"
        End If
        
    Next i
    
    MsgBox "SalaryBand column created successfully!", vbInformation

End Sub





Sub AddYearsSinceLastPromotion()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim promoCol As Variant
    Dim newCol As Long
    Dim i As Long
    Dim promoDate As Variant
    
    Set ws = ThisWorkbook.Sheets("Cleaned_Data")
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Find LastPromotionDate column
    promoCol = Application.Match("LastPromotionDate", ws.Rows(1), 0)
    
    If IsError(promoCol) Then
        MsgBox "LastPromotionDate column not found!", vbCritical
        Exit Sub
    End If
    
    ' Check if column already exists
    On Error Resume Next
    newCol = Application.Match("YearsSinceLastPromotion", ws.Rows(1), 0)
    On Error GoTo 0
    
    If IsError(newCol) Or newCol = 0 Then
        newCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(1, newCol).Value = "YearsSinceLastPromotion"
    End If
    
    ' Populate values
    For i = 2 To lastRow
        
        promoDate = ws.Cells(i, promoCol).Value
        
        If IsDate(promoDate) Then
            ws.Cells(i, newCol).Value = Int((Date - promoDate) / 365)
        Else
            ws.Cells(i, newCol).Value = ""
        End If
        
    Next i
    
    MsgBox "YearsSinceLastPromotion column is added!", vbInformation

End Sub
