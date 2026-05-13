Attribute VB_Name = "modExportToPDF"
Sub ExportDashboardWithInsights_PDF()

    Dim wsDash As Worksheet, wsIns As Worksheet
    Dim pdfPath As String, fileName As String

    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsIns = ThisWorkbook.Sheets("Insights")

    ' =========================
    ' STEP 1: SET PRINT AREAS
    ' =========================
    wsDash.PageSetup.PrintArea = "A1:T50"
    wsIns.PageSetup.PrintArea = wsIns.UsedRange.Address

    ' =========================
    ' STEP 2: PAGE SETUP (DASHBOARD)
    ' =========================
    With wsDash.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
    End With

    ' =========================
    ' STEP 3: PAGE SETUP (INSIGHTS)
    ' =========================
    With wsIns.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False   ' allow multiple rows if needed
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
    End With

    ' =========================
    ' STEP 4: FILE NAME
    ' =========================
    fileName = "HR_DASHBOARD_" & Format(Now, "yyyy-mm-dd_hhmmss") & ".pdf"
    pdfPath = ThisWorkbook.Path & "\export\" & fileName

    ' =========================
    ' STEP 5: EXPORT BOTH SHEETS
    ' =========================
    Sheets(Array("Dashboard", "Insights")).Select

    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    ' Return focus to Dashboard
    wsDash.Select

    MsgBox "PDF exported (Dashboard + Insights):" & vbCrLf & pdfPath, vbInformation

End Sub

