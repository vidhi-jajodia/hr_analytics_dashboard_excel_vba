Attribute VB_Name = "modPivotTables"
Public pcHRDashboard As PivotCache

Sub InitializePivotCache()

    Dim wsData As Worksheet
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    Set pcHRDashboard = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
        
End Sub

Sub CreateDepartmentAttritionPivot()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    
    ' Clear old pivot area
    wsPivot.Cells.Clear
    
    ' Define source range
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    ' Create pivot cache
    Set pc = pcHRDashboard
    
    ' Delete old pivot if exists
    On Error Resume Next
    wsPivot.PivotTables("ptDepartmentAttrition").TableRange2.Clear
    On Error GoTo 0
    
    ' Create pivot table
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="ptDepartmentAttrition")
    
    With pt
        
        ' Department as rows
        With .PivotFields("Department")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Attrition filter = Yes
        With .PivotFields("Attrition")
            .Orientation = xlPageField
            .Position = 1
            .CurrentPage = "Yes"
        End With
        
        ' Employee count
        .AddDataField .PivotFields("EmployeeID"), _
            "Attrition Count", xlCount
        
        ' Sort descending
        .PivotFields("Department").AutoSort _
            xlDescending, "Attrition Count"
        
        ' Apply style
        .TableStyle2 = "PivotStyleMedium9"
        
    End With
    
    MsgBox "Department Attrition Pivot Table Created Successfully!", vbInformation

End Sub



' Create Salary Attrition Pivot Table via VBA
Sub CreateSalaryAttritionPivot()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    
    ' Clear old area
    wsPivot.Range("J3:N20").Clear
    
    ' Define source
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    ' Create pivot cache
    Set pc = pcHRDashboard
    
    ' Delete existing pivot if exists
    On Error Resume Next
    wsPivot.PivotTables("ptSalaryAttrition").TableRange2.Clear
    On Error GoTo 0
    
    ' Create pivot
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("J3"), _
        TableName:="ptSalaryAttrition")
    
    With pt
        
        ' SalaryBand rows
        With .PivotFields("SalaryBand")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Attrition filter
        With .PivotFields("Attrition")
            .Orientation = xlPageField
            .CurrentPage = "Yes"
        End With
        
        ' Employee count
        .AddDataField .PivotFields("EmployeeID"), _
            "Attrition Count", xlCount
        
        ' Sort descending
        .PivotFields("SalaryBand").AutoSort _
            xlDescending, "Attrition Count"
        
        ' Style
        .TableStyle2 = "PivotStyleMedium9"
        
    End With
    
    MsgBox "Salary Attrition Pivot Table Created Successfully!", vbInformation

End Sub





Sub CreateGenderDistributionPivot()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    
    ' Clear old area
    wsPivot.Range("Q3:U20").Clear
    
    ' Define source
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    ' Create cache
    Set pc = pcHRDashboard
    
    ' Delete old pivot
    On Error Resume Next
    wsPivot.PivotTables("ptGenderDistribution").TableRange2.Clear
    On Error GoTo 0
    
    ' Create pivot
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("Q3"), _
        TableName:="ptGenderDistribution")
    
    With pt
        
        ' Gender rows
        With .PivotFields("Gender")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Count EmployeeID
        .AddDataField .PivotFields("EmployeeID"), _
            "Employee Count", xlCount
        
        ' Style
        .TableStyle2 = "PivotStyleMedium9"
        
    End With
    
    MsgBox "Gender Distribution Pivot Created Successfully!", vbInformation

End Sub





Sub CreateOvertimeAttritionPivot()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    
    ' Clear old area
    wsPivot.Range("A25:F40").Clear
    
    ' Source data
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    ' Create cache
    Set pc = pcHRDashboard
    
    ' Delete existing pivot
    On Error Resume Next
    wsPivot.PivotTables("ptOvertimeAttrition").TableRange2.Clear
    On Error GoTo 0
    
    ' Create pivot
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A25"), _
        TableName:="ptOvertimeAttrition")
    
    With pt
        
        ' Overtime rows
        With .PivotFields("Overtime")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Attrition filter
        With .PivotFields("Attrition")
            .Orientation = xlPageField
            .CurrentPage = "Yes"
        End With
        
        ' Count employees
        .AddDataField .PivotFields("EmployeeID"), _
            "Attrition Count", xlCount
        
        ' Style
        .TableStyle2 = "PivotStyleMedium9"
        
    End With
    
    MsgBox "Overtime Attrition Pivot Created Successfully!", vbInformation

End Sub






Sub CreatePromotionGapPivot()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    
    ' Delete old pivot safely
    On Error Resume Next
    If Not wsPivot.PivotTables("ptPromotionGap") Is Nothing Then
        wsPivot.PivotTables("ptPromotionGap").TableRange2.Clear
    End If
    On Error GoTo 0
    
    ' Clear destination area
    wsPivot.Range("J25:O60").Clear
    
    ' Define source
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    ' Create cache
    Set pc = pcHRDashboard
    
    ' Create pivot
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("J25"), _
        TableName:="ptPromotionGap")
    
    With pt
        
        ' Rows
        With .PivotFields("YearsSinceLastPromotion")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Filter
        With .PivotFields("Attrition")
            .Orientation = xlPageField
            .Position = 1
            .CurrentPage = "Yes"
        End With
        
        ' Values
        .AddDataField .PivotFields("EmployeeID"), _
            "Attrition Count", xlCount
        
        ' Sort ascending
        .PivotFields("YearsSinceLastPromotion").AutoSort _
            xlAscending, "YearsSinceLastPromotion"
        
        ' Layout
        .RowAxisLayout xlTabularRow
        
        ' Style
        .TableStyle2 = "PivotStyleMedium9"
        
    End With
    
    MsgBox "Promotion Gap Pivot Updated Successfully!", vbInformation

End Sub








Sub CreateEducationAttritionPivot()

    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    
    Set wsData = ThisWorkbook.Sheets("Cleaned_Data")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    
    ' Delete old pivot
    On Error Resume Next
    If Not wsPivot.PivotTables("ptEducationAttrition") Is Nothing Then
        wsPivot.PivotTables("ptEducationAttrition").TableRange2.Clear
    End If
    On Error GoTo 0
    
    ' Clear area
    wsPivot.Range("Q25:V50").Clear
    
    ' Source data
    Set dataRange = wsData.Range("A1").CurrentRegion
    
    ' Cache
    Set pc = pcHRDashboard
    
    ' Create pivot
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("Q25"), _
        TableName:="ptEducationAttrition")
    
    With pt
        
        ' Rows
        With .PivotFields("EducationLevel")
            .Orientation = xlRowField
            .Position = 1
        End With
        
        ' Filter
        With .PivotFields("Attrition")
            .Orientation = xlPageField
            .Position = 1
            .CurrentPage = "Yes"
        End With
        
        ' Values
        .AddDataField .PivotFields("EmployeeID"), _
            "Attrition Count", xlCount
        
        ' Style
        .TableStyle2 = "PivotStyleMedium9"
        
        ' Layout
        .RowAxisLayout xlTabularRow
        
    End With
    
    MsgBox "Education Attrition Pivot Created Successfully!", vbInformation

End Sub

