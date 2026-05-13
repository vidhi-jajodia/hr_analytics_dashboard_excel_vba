Attribute VB_Name = "modInsights"
Function GetColumnIndex(ws As Worksheet, colName As String) As Long

    Dim lastCol As Long, i As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        If Trim(UCase(ws.Cells(1, i).Value)) = Trim(UCase(colName)) Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    
    GetColumnIndex = -1

End Function



Function GetSetting(paramName As String) As Double

    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Settings")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Settings sheet not found!", vbCritical
        Exit Function
    End If
    
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        If Trim(UCase(ws.Cells(i, 1).Value)) = Trim(UCase(paramName)) Then
            GetSetting = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    
    MsgBox "Setting not found: " & paramName, vbCritical
    GetSetting = 0

End Function



Sub GenerateInsights_DataDriven()

    Dim wsData As Worksheet, wsIns As Worksheet
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsIns = ThisWorkbook.Sheets("Insights")
    
    wsIns.Cells.Clear
    wsIns.Range("A1:E1").Value = Array("Priority", "Area", "Insight", "Impact", "Recommendation")
    
    Dim rowOut As Long: rowOut = 2
    
    Dim colAttr As Long, colDept As Long, colSal As Long
    Dim colExp As Long, colPromo As Long, colOT As Long
    
    colAttr = GetColumnIndex(wsData, "Attrition")
    colDept = GetColumnIndex(wsData, "Department")
    colSal = GetColumnIndex(wsData, "Salary")
    colExp = GetColumnIndex(wsData, "TotalExperience")
    colPromo = GetColumnIndex(wsData, "LastPromotionDate")
    colOT = GetColumnIndex(wsData, "Overtime")
    
    If colAttr = -1 Or colDept = -1 Or colSal = -1 Or colExp = -1 Or colPromo = -1 Or colOT = -1 Then
        MsgBox "Column missing!", vbCritical
        Exit Sub
    End If
    
    ' ?? SETTINGS
    Dim salaryThreshold As Double
    Dim expThreshold As Double
    Dim promoGap As Double
    
    salaryThreshold = GetSetting("Salary_Threshold")
    expThreshold = GetSetting("Early_Experience_Years")
    promoGap = GetSetting("Promotion_Gap_Years")
    
    Dim lastRow As Long, i As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    
    ' =========================
    ' OVERALL
    ' =========================
    Dim total As Long, attrCount As Long
    
    For i = 2 To lastRow
        total = total + 1
        If UCase(Trim(wsData.Cells(i, colAttr).Value)) = "YES" Then attrCount = attrCount + 1
    Next i
    
    If total = 0 Then Exit Sub
    
    Dim overallRate As Double
    overallRate = attrCount / total
    
    ' =========================
    ' 1. DEPARTMENT ANALYSIS
    ' =========================
    Dim dictT As Object, dictA As Object
    Set dictT = CreateObject("Scripting.Dictionary")
    Set dictA = CreateObject("Scripting.Dictionary")
    
    Dim dept As String
    
    For i = 2 To lastRow
        dept = wsData.Cells(i, colDept).Value
        
        If dept <> "" Then
            dictT(dept) = dictT(dept) + 1
            
            If UCase(wsData.Cells(i, colAttr).Value) = "YES" Then
                dictA(dept) = dictA(dept) + 1
            End If
        End If
    Next i
    
    Dim key As Variant, maxRate As Double, maxDept As String
    
    For Each key In dictT.Keys
        Dim r As Double
        r = dictA(key) / dictT(key)
        
        If r > maxRate Then
            maxRate = r
            maxDept = key
        End If
    Next key
    
    If maxDept <> "" Then
        
        wsIns.Cells(rowOut, 1).Value = IIf(maxRate > overallRate * 1.2, "HIGH", "MEDIUM")
        wsIns.Cells(rowOut, 2).Value = "Department"
        wsIns.Cells(rowOut, 3).Value = maxDept & " has highest attrition at " & Format(maxRate, "0.0%")
        wsIns.Cells(rowOut, 4).Value = Format((maxRate - overallRate) / overallRate, "0%")
        wsIns.Cells(rowOut, 5).Value = "Investigate leadership & workload"
        
        rowOut = rowOut + 1
        
    End If
    
    ' =========================
    ' 2. SALARY
    ' =========================
    Dim lowAttr As Long, lowTotal As Long
    
    For i = 2 To lastRow
        If Val(wsData.Cells(i, colSal).Value) < salaryThreshold Then
            
            lowTotal = lowTotal + 1
            
            If UCase(wsData.Cells(i, colAttr).Value) = "YES" Then
                lowAttr = lowAttr + 1
            End If
        End If
    Next i
    
    If lowTotal > 0 Then
        
        Dim lowRate As Double
        lowRate = lowAttr / lowTotal
        
        wsIns.Cells(rowOut, 1).Value = IIf(lowRate > overallRate, "MEDIUM", "LOW")
        wsIns.Cells(rowOut, 2).Value = "Salary"
        wsIns.Cells(rowOut, 3).Value = "Low salary attrition = " & Format(lowRate, "0.0%")
        wsIns.Cells(rowOut, 4).Value = Format((lowRate - overallRate) / overallRate, "0%")
        wsIns.Cells(rowOut, 5).Value = "Review pay structure"
        
        rowOut = rowOut + 1
        
    End If
    
    ' =========================
    ' 3. EXPERIENCE
    ' =========================
    Dim earlyAttr As Long, earlyTotal As Long
    
    For i = 2 To lastRow
        If Val(wsData.Cells(i, colExp).Value) <= expThreshold Then
            
            earlyTotal = earlyTotal + 1
            
            If UCase(wsData.Cells(i, colAttr).Value) = "YES" Then
                earlyAttr = earlyAttr + 1
            End If
        End If
    Next i
    
    If earlyTotal > 0 Then
        
        Dim earlyRate As Double
        earlyRate = earlyAttr / earlyTotal
        
        wsIns.Cells(rowOut, 1).Value = IIf(earlyRate > overallRate, "MEDIUM", "LOW")
        wsIns.Cells(rowOut, 2).Value = "Experience"
        wsIns.Cells(rowOut, 3).Value = "Early career attrition = " & Format(earlyRate, "0.0%")
        wsIns.Cells(rowOut, 4).Value = Format((earlyRate - overallRate) / overallRate, "0%")
        wsIns.Cells(rowOut, 5).Value = "Improve onboarding & mentoring"
        
        rowOut = rowOut + 1
        
    End If
    
    ' =========================
    ' 4. PROMOTION
    ' =========================
    Dim promoAttr As Long, promoTotal As Long
    
    For i = 2 To lastRow
        If IsDate(wsData.Cells(i, colPromo).Value) Then
            
            If DateDiff("yyyy", wsData.Cells(i, colPromo).Value, Date) >= promoGap Then
                
                promoTotal = promoTotal + 1
                
                If UCase(wsData.Cells(i, colAttr).Value) = "YES" Then
                    promoAttr = promoAttr + 1
                End If
                
            End If
            
        End If
    Next i
    
    If promoTotal > 0 Then
        
        Dim promoRate As Double
        promoRate = promoAttr / promoTotal
        
        wsIns.Cells(rowOut, 1).Value = IIf(promoRate > overallRate * 1.2, "HIGH", "MEDIUM")
        wsIns.Cells(rowOut, 2).Value = "Promotion"
        wsIns.Cells(rowOut, 3).Value = "Delayed promotion attrition = " & Format(promoRate, "0.0%")
        wsIns.Cells(rowOut, 4).Value = Format((promoRate - overallRate) / overallRate, "0%")
        wsIns.Cells(rowOut, 5).Value = "Improve promotion cycle"
        
        rowOut = rowOut + 1
        
    End If
    
    ' =========================
    ' 5. OVERTIME
    ' =========================
    Dim otAttr As Long, otTotal As Long
    
    For i = 2 To lastRow
        If UCase(wsData.Cells(i, colOT).Value) = "YES" Then
            
            otTotal = otTotal + 1
            
            If UCase(wsData.Cells(i, colAttr).Value) = "YES" Then
                otAttr = otAttr + 1
            End If
            
        End If
    Next i
    
    If otTotal > 0 Then
        
        Dim otRate As Double
        otRate = otAttr / otTotal
        
        wsIns.Cells(rowOut, 1).Value = IIf(otRate > overallRate, "HIGH", "LOW")
        wsIns.Cells(rowOut, 2).Value = "Overtime"
        wsIns.Cells(rowOut, 3).Value = "Overtime attrition = " & Format(otRate, "0.0%")
        wsIns.Cells(rowOut, 4).Value = Format((otRate - overallRate) / overallRate, "0%")
        wsIns.Cells(rowOut, 5).Value = "Reduce workload"
        
        rowOut = rowOut + 1
        
    End If
    
    ' =========================
    ' SORT
    ' =========================
    wsIns.Range("A1").CurrentRegion.Sort _
        Key1:=wsIns.Range("A2"), Order1:=xlAscending, Header:=xlYes

End Sub





Sub StyleInsights()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Insights")
    
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Header styling
    With ws.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Clear old formatting ONLY in key columns
    ws.Range("A2:A" & lastRow).Interior.Pattern = xlNone
    ws.Range("C2:C" & lastRow).Interior.Pattern = xlNone
    
    ' Apply styling
    For i = 2 To lastRow
        
        Select Case UCase(ws.Cells(i, 1).Value)
            
            Case "HIGH"
                ws.Cells(i, 1).Interior.Color = RGB(255, 199, 206) ' Priority
                ws.Cells(i, 3).Interior.Color = RGB(255, 235, 238) ' Light highlight
                
            Case "MEDIUM"
                ws.Cells(i, 1).Interior.Color = RGB(255, 235, 156)
                ws.Cells(i, 3).Interior.Color = RGB(255, 242, 204)
                
            Case "LOW"
                ws.Cells(i, 1).Interior.Color = RGB(198, 239, 206)
                ws.Cells(i, 3).Interior.Color = RGB(226, 239, 218)
                
        End Select
        
    Next i
    
    ' Alignments (nice polish)
    ws.Columns("A").HorizontalAlignment = xlCenter
    ws.Columns("D").HorizontalAlignment = xlCenter
    
    ' Autofit
    ws.Columns("A:E").AutoFit

End Sub
