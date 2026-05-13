Private Sub UserForm_Initialize()
    Me.BackColor = RGB(245, 247, 250)

    Dim ctrl As Control

    For Each ctrl In Me.Controls
        
        If TypeName(ctrl) = "CommandButton" Then
            ctrl.BackColor = RGB(0, 120, 215)
            ctrl.ForeColor = RGB(255, 255, 255)
            ctrl.Font.Bold = True
        End If
        
        If TypeName(ctrl) = "Frame" Then
            ctrl.BackColor = RGB(255, 255, 255)
        End If

    Next ctrl
    
    ' Disable Update button initially
    btnUpdate.Enabled = False

    ' Dropdowns
    cmbGender.List = Array("Male", "Female", "Other")
    cmbDepartment.List = Array("HR", "IT", "FINANCE", "MARKETING", "OPERATIONS", "SALES")
    cmbJobRole.List = Array("Analyst", "Associate", "Engineer", "Manager", "Director", "Executive")
    cmbMaritalStatus.List = Array("Single", "Married")
    cmbOvertime.List = Array("Yes", "No")
    cmbAttrition.List = Array("Yes", "No")
    cmbJobSatisfaction.List = Array("1", "2", "3", "4", "5")
    cmbPerformanceRating.List = Array("1", "2", "3", "4", "5")
    cmbWorkLifeBalance.List = Array("1", "2", "3", "4")

End Sub




Private Sub btnSave_Click()

    Dim ws As Worksheet
    Dim r As Long
    Dim f As Range

    Set ws = ThisWorkbook.Sheets("Cleaned_Data")

    ' 1. Check Employee ID is not empty
    If Trim(txtEmployeeID.Value) = "" Then
        MsgBox "Employee ID cannot be empty!", vbExclamation
        txtEmployeeID.SetFocus
        Exit Sub
    End If

    ' 2. Check duplicate Employee ID
    Set f = ws.Columns(1).Find(txtEmployeeID.Value, LookIn:=xlValues, LookAt:=xlWhole)

    If Not f Is Nothing Then
        MsgBox "Employee ID already exists! Please use a unique ID.", vbExclamation
        txtEmployeeID.SetFocus
        Exit Sub
    End If

    ' 3. Find next empty row
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' 4. Save data
    ws.Cells(r, 1) = txtEmployeeID
    ws.Cells(r, 2) = txtFirstName
    ws.Cells(r, 3) = txtLastName
    ws.Cells(r, 4) = cmbGender
    ws.Cells(r, 5) = txtAge
    ws.Cells(r, 6) = cmbDepartment
    ws.Cells(r, 7) = cmbJobRole
    ws.Cells(r, 8) = txtSalary
    ws.Cells(r, 9) = txtCTCPerAnnum
    ws.Cells(r, 10) = txtYearsAtCompany
    ws.Cells(r, 11) = txtTotalExperience
    ws.Cells(r, 12) = txtLastPromotionDate
    ws.Cells(r, 13) = cmbEducationLevel
    ws.Cells(r, 14) = cmbMaritalStatus
    ws.Cells(r, 15) = cmbJobSatisfaction
    ws.Cells(r, 16) = cmbPerformanceRating
    ws.Cells(r, 17) = cmbOvertime
    ws.Cells(r, 18) = cmbWorkLifeBalance
    ws.Cells(r, 19) = txtDistanceFromHome
    ws.Cells(r, 20) = txtMonthlyIncome
    ws.Cells(r, 21) = cmbAttrition

    MsgBox "Employee Details Saved Successfully!", vbInformation

End Sub




Private Sub btnSearch_Click()

    Dim f As Range
    Dim empID As String

    ' Trim spaces and normalize input
    empID = Trim(txtEmployeeID.Value)

    ' Validate empty / space-only input
    If empID = "" Then
        MsgBox "Please enter a valid Employee ID.", vbExclamation
        txtEmployeeID.SetFocus
        btnUpdate.Enabled = False
        Exit Sub
    End If

    ' Search only if valid input exists
    Set f = Sheets("Cleaned_Data").Columns(1).Find(empID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not f Is Nothing Then
        
        txtEmployeeID = f.Offset(0, 0)
        txtFirstName = f.Offset(0, 1)
        txtLastName = f.Offset(0, 2)
        cmbGender = f.Offset(0, 3)
        txtAge = f.Offset(0, 4)
        cmbDepartment = f.Offset(0, 5)
        cmbJobRole = f.Offset(0, 6)
        txtSalary = f.Offset(0, 7)
        txtCTCPerAnnum = f.Offset(0, 8)
        txtYearsAtCompany = f.Offset(0, 9)
        txtTotalExperience = f.Offset(0, 10)
        txtLastPromotionDate = f.Offset(0, 11)
        cmbEducationLevel = f.Offset(0, 12)
        cmbMaritalStatus = f.Offset(0, 13)
        cmbJobSatisfaction = f.Offset(0, 14)
        cmbPerformanceRating = f.Offset(0, 15)
        cmbOvertime = f.Offset(0, 16)
        cmbWorkLifeBalance = f.Offset(0, 17)
        txtDistanceFromHome = f.Offset(0, 18)
        txtMonthlyIncome = f.Offset(0, 19)
        cmbAttrition = f.Offset(0, 20)

        ' Enable update only after valid record is found
        btnUpdate.Enabled = True

    Else
        MsgBox "Employee not found.", vbInformation
        btnUpdate.Enabled = False
    End If

End Sub






Private Sub btnUpdate_Click()

    Dim f As Range
    
    Set f = Sheets("Cleaned_Data").Columns(1).Find(txtEmployeeID, , , xlWhole)

    If Not f Is Nothing Then
        
        f.Offset(0, 1) = txtFirstName
        f.Offset(0, 2) = txtLastName
        
        MsgBox "Employee Details Updated!"
        
    Else
        MsgBox "Not found"
    End If

End Sub





Private Sub btnClear_Click()

    Dim c As Control
    
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Or TypeName(c) = "ComboBox" Then
            c.Value = ""
        End If
    Next
    
    ' Disable Update again
    btnUpdate.Enabled = False

End Sub



Private Sub txtSalary_Change()

    If IsNumeric(txtSalary.Value) And Val(txtSalary.Value) > 0 Then
        txtCTCPerAnnum.Value = Val(txtSalary.Value) * 12
    Else
        txtCTCPerAnnum.Value = ""
    End If

End Sub
