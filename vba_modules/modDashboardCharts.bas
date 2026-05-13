Attribute VB_Name = "modDashboardCharts"
Sub CreateDepartmentAttritionChart()

    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    Set pt = wsPivot.PivotTables("ptDepartmentAttrition")
    
    ' Delete old chart if it exists
    On Error Resume Next
    wsDash.ChartObjects("DeptAttritionChart").Delete
    On Error GoTo 0
    
    ' Create new chart
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("C12").Left, _
        Top:=wsDash.Range("C12").Top, _
        Width:=wsDash.Range("C12:H30").Width, _
        Height:=wsDash.Range("C12:H30").Height)
    
    chtObj.Name = "DeptAttritionChart"
    
    With chtObj.Chart
        
        ' Set source
        .SetSourceData Source:=pt.TableRange1
        
        ' Chart type
        .ChartType = xlBarClustered
        
        ' Title
        .HasTitle = True
        .ChartTitle.Text = "Department Attrition Analysis"
        
        ' Remove legend
        .HasLegend = False
        
        ' Data labels
        .SeriesCollection(1).ApplyDataLabels
        
        ' Bar color
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(52, 152, 219)
        
        ' Chart background
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' Gridlines off
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        On Error GoTo 0
        
        ' Axis labels
        .Axes(xlCategory).TickLabelPosition = xlLow
        
        ' Title formatting
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14
            .Bold = True
        End With
        
    End With
    
    MsgBox "Department Attrition Chart Created Successfully!", vbInformation

End Sub





Sub CreateSalaryAttritionChart()

    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    Set pt = wsPivot.PivotTables("ptSalaryAttrition")
    
    ' Delete existing chart
    On Error Resume Next
    wsDash.ChartObjects("SalaryAttritionChart").Delete
    On Error GoTo 0
    
    ' Create chart
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("I12").Left, _
        Top:=wsDash.Range("I12").Top, _
        Width:=wsDash.Range("I12:N30").Width, _
        Height:=wsDash.Range("I12:N30").Height)
    
    chtObj.Name = "SalaryAttritionChart"
    
    With chtObj.Chart
        
        ' Source
        .SetSourceData Source:=pt.TableRange1
        
        ' Type
        .ChartType = xlColumnClustered
        
        ' Title
        .HasTitle = True
        .ChartTitle.Text = "Salary Impact on Attrition"
        
        ' Remove legend
        .HasLegend = False
        
        ' Apply labels
        .SeriesCollection(1).ApplyDataLabels
        
        ' Column color
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(46, 204, 113)
        
        ' Clean background
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' Remove major gridlines
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        On Error GoTo 0
        
        ' Title formatting
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14
            .Bold = True
        End With
        
    End With
    
    MsgBox "Salary Attrition Chart Created Successfully!", vbInformation

End Sub




Sub CreateGenderDistributionChart()

    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    Set pt = wsPivot.PivotTables("ptGenderDistribution")
    
    ' Delete old chart
    On Error Resume Next
    wsDash.ChartObjects("GenderDistributionChart").Delete
    On Error GoTo 0
    
    ' Create chart
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("O12").Left, _
        Top:=wsDash.Range("O12").Top, _
        Width:=wsDash.Range("O12:T30").Width, _
        Height:=wsDash.Range("O12:T30").Height)
    
    chtObj.Name = "GenderDistributionChart"
    
With chtObj.Chart
    
    ' Set source
    .SetSourceData Source:=pt.TableRange1
    
    ' Donut chart
    .ChartType = xlDoughnut
    
    ' Title
    .HasTitle = True
    .ChartTitle.Text = "Gender Workforce Distribution"
    
    ' Legend
    .HasLegend = True
    .Legend.Position = xlRight
    
    ' Data labels
    .SeriesCollection(1).ApplyDataLabels
    
    ' Correct donut hole size
    .ChartGroups(1).DoughnutHoleSize = 55
    
    ' Slice colors
    If .SeriesCollection(1).Points.Count >= 2 Then
        .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(52, 152, 219)
        .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(155, 89, 182)
    End If
    
    ' Background
    .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    
    ' Title formatting
    With .ChartTitle.Format.TextFrame2.TextRange.Font
        .Size = 14
        .Bold = True
    End With
    
End With
    
    MsgBox "Gender Distribution Chart Created Successfully!", vbInformation

End Sub







Sub CreateOvertimeRiskChart()

    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    Set pt = wsPivot.PivotTables("ptOvertimeAttrition")
    
    ' Delete old chart
    On Error Resume Next
    wsDash.ChartObjects("OvertimeRiskChart").Delete
    On Error GoTo 0
    
    ' Create chart
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("C32").Left, _
        Top:=wsDash.Range("C32").Top, _
        Width:=wsDash.Range("C32:H49").Width, _
        Height:=wsDash.Range("C32:H49").Height)
    
    chtObj.Name = "OvertimeRiskChart"
    
    With chtObj.Chart
        
        ' Set source
        .SetSourceData Source:=pt.TableRange1
        
        ' Chart type
        .ChartType = xlBarClustered
        
        ' Title
        .HasTitle = True
        .ChartTitle.Text = "Overtime Risk Analysis"
        
        ' Remove legend
        .HasLegend = False
        
        ' Data labels
        .SeriesCollection(1).ApplyDataLabels
        
        ' Bar color
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(231, 76, 60)
        
        ' Clean background
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' Remove gridlines
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        On Error GoTo 0
        
        ' Axis formatting
        .Axes(xlCategory).TickLabelPosition = xlLow
        
        ' Title styling
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14
            .Bold = True
        End With
        
    End With
    
    MsgBox "Overtime Risk Chart Created Successfully!", vbInformation

End Sub







Sub CreatePromotionGapChart()

    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    Set pt = wsPivot.PivotTables("ptPromotionGap")
    
    ' Delete old chart
    On Error Resume Next
    wsDash.ChartObjects("PromotionGapChart").Delete
    On Error GoTo 0
    
    ' Create chart
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("I32").Left, _
        Top:=wsDash.Range("I32").Top, _
        Width:=wsDash.Range("I32:N49").Width, _
        Height:=wsDash.Range("I32:N49").Height)
    
    chtObj.Name = "PromotionGapChart"
    
    With chtObj.Chart
        
        ' Set source
        .SetSourceData Source:=pt.TableRange1
        
        ' Chart type
        .ChartType = xlColumnClustered
        
        ' Title
        .HasTitle = True
        .ChartTitle.Text = "Promotion Delay Risk Analysis"
        
        ' Remove legend
        .HasLegend = False
        
        ' Apply data labels
        .SeriesCollection(1).ApplyDataLabels
        
        ' Bar color
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(241, 196, 15)
        
        ' Background styling
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' Remove gridlines
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        On Error GoTo 0
        
        ' Axis title
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Years Since Last Promotion"
        
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Attrition Count"
        
        ' Chart title formatting
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14
            .Bold = True
        End With
        
        ' Axis formatting
        With .Axes(xlCategory)
            .TickLabelSpacing = 1
        End With
        
    End With
    
    MsgBox "Promotion Gap Chart Updated Successfully!", vbInformation

End Sub








Sub CreateEducationAttritionChart()

    Dim wsDash As Worksheet
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsPivot = ThisWorkbook.Sheets("Pivot_Backend")
    Set pt = wsPivot.PivotTables("ptEducationAttrition")
    
    ' Delete old chart
    On Error Resume Next
    wsDash.ChartObjects("EducationAttritionChart").Delete
    On Error GoTo 0
    
    ' Create chart
    Set chtObj = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("O32").Left, _
        Top:=wsDash.Range("O32").Top, _
        Width:=wsDash.Range("O32:T49").Width, _
        Height:=wsDash.Range("O32:T49").Height)
    
    chtObj.Name = "EducationAttritionChart"
    
    With chtObj.Chart
        
        ' Source
        .SetSourceData Source:=pt.TableRange1
        
        ' Chart type
        .ChartType = xlBarClustered
        
        ' Title
        .HasTitle = True
        .ChartTitle.Text = "Education-Based Attrition Analysis"
        
        ' Remove legend
        .HasLegend = False
        
        ' Data labels
        .SeriesCollection(1).ApplyDataLabels
        
        ' Bar color
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(155, 89, 182)
        
        ' Background
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        ' Remove gridlines
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Delete
        On Error GoTo 0
        
        ' Axis labels
        .Axes(xlCategory).TickLabelPosition = xlLow
        
        ' Title formatting
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Size = 14
            .Bold = True
        End With
        
    End With
    
    MsgBox "Education Attrition Chart Created Successfully!", vbInformation

End Sub

