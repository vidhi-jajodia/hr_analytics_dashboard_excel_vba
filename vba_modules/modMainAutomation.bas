Attribute VB_Name = "modMainAutomation"
Sub RunFullDashboardAutomation()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' ==========================
    ' DATA ENHANCEMENT
    ' ==========================
    Call AddSalaryBandColumn
    Call AddYearsSinceLastPromotion
    
    ' ==========================
    ' PIVOT TABLES
    ' ==========================
    Call InitializePivotCache
    Call CreateDepartmentAttritionPivot
    Call CreateSalaryAttritionPivot
    Call CreateGenderDistributionPivot
    Call CreateOvertimeAttritionPivot
    Call CreatePromotionGapPivot
    Call CreateEducationAttritionPivot
    
    ' ==========================
    ' DASHBOARD CHARTS
    ' ==========================
    Call CreateDepartmentAttritionChart
    Call CreateSalaryAttritionChart
    Call CreateGenderDistributionChart
    Call CreateOvertimeRiskChart
    Call CreatePromotionGapChart
    Call CreateEducationAttritionChart
    
    ' ==========================
    ' REFRESH DASHBOARD
    ' ==========================
    ThisWorkbook.RefreshAll
    
    
    ' ==========================
    ' GENERATE INSIGHTS
    ' ==========================
    Call GenerateInsights_DataDriven
    Call StyleInsights

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "HR Analytics Dashboard Refreshed Successfully!", vbInformation, "Automation Complete"
    
    Exit Sub

ErrorHandler:
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Error occurred: " & Err.Description, vbCritical, "Automation Failed"

End Sub
