Imports System.Data.SqlClient
Module Module1
     Public OleCn As New SqlConnection()
    Public Dst, Dst1 As New DataSet()
    Public myDA, DA As New SqlDataAdapter()
    Public FrmState As Integer = 0

    Public Function StrConnection() As String
        StrConnection = "Data Source=HASSAN-PC\SQLEXPRESS;Initial Catalog=Collage;User ID=sa;Password=ali123;"
        Return StrConnection
    End Function


    Dim EmpID As New DataGridViewTextBoxColumn()
    Dim EmpName As New DataGridViewTextBoxColumn()
    Dim Desig As New DataGridViewTextBoxColumn()
    Dim BSalary As New DataGridViewTextBoxColumn()
    Dim AGP As New DataGridViewTextBoxColumn()
    Dim Add As New DataGridViewTextBoxColumn()
    Dim city As New DataGridViewTextBoxColumn()
    Dim State As New DataGridViewTextBoxColumn()
    Dim Zipcode As New DataGridViewTextBoxColumn()
    Dim MobileNo As New DataGridViewTextBoxColumn()
    Dim emailid As New DataGridViewTextBoxColumn()


    Public Sub GridViewEmpDetails1(ByVal dgvEmpDetails1 As DataGridView)
        dgvEmpDetails1.AutoGenerateColumns = False
        With EmpID
            .DataPropertyName = "EmpID"
            .HeaderText = "Employee ID"
            .Width = 85
        End With
        With EmpName
            .DataPropertyName = "EmpName"
            .HeaderText = "Employee Name"
            .Width = 130
        End With
        With Desig
            .DataPropertyName = "Designation"
            .HeaderText = "Designation"
            .Width = 115
        End With
        With BSalary
            .DataPropertyName = "BSalary"
            .HeaderText = "Basic Salary"
            .Width = 70
        End With
        With AGP
            .DataPropertyName = "AGP"
            .HeaderText = "AGP"
            .Width = 70
        End With
        With Add
            .DataPropertyName = "Address"
            .HeaderText = "Address"
            .Width = 165
        End With
        With city
            .DataPropertyName = "City"
            .HeaderText = "City"
            .Width = 90
        End With
        With State
            .DataPropertyName = "State"
            .HeaderText = "State"
            .Width = 100
        End With
        With Zipcode
            .DataPropertyName = "zipcode"
            .HeaderText = "zip code"
            .Width = 60
        End With
        With MobileNo
            .DataPropertyName = "MobileNo"
            .HeaderText = "Mobile No"
            .Width = 90
        End With
        With emailid
            .DataPropertyName = "EmailID"
            .HeaderText = "EmailID"
            .Width = 135
        End With


        With dgvEmpDetails1
            .DataSource = Dst
            .DataMember = "PGStaffEmpDetails"
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ShowRowErrors = False
            .ShowCellErrors = False
            .AllowUserToAddRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .RowHeadersVisible = True
            .RowHeadersWidth = 32
            .DefaultCellStyle.SelectionBackColor = Color.SteelBlue
            .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGoldenrodYellow
            .Columns.AddRange(New DataGridViewColumn() {EmpID, EmpName, Desig, BSalary, AGP, Add, city, State, Zipcode, MobileNo, emailid})
        End With
    End Sub



    Dim EmployeeID As New DataGridViewTextBoxColumn()
    Dim EmpNme As New DataGridViewTextBoxColumn()
    Dim Designation As New DataGridViewTextBoxColumn()
    Dim BasicSalary As New DataGridViewTextBoxColumn()
    Dim Addre As New DataGridViewTextBoxColumn()
    Dim city1 As New DataGridViewTextBoxColumn()
    Dim State1 As New DataGridViewTextBoxColumn()
    Dim Zipcode1 As New DataGridViewTextBoxColumn()
    Dim MobNo As New DataGridViewTextBoxColumn()
    Dim emailid1 As New DataGridViewTextBoxColumn()


    Public Sub GridViewEmpDetails2(ByVal dgvEmpDetails2 As DataGridView)
        dgvEmpDetails2.AutoGenerateColumns = False
        With EmployeeID
            .DataPropertyName = "EmpID"
            .HeaderText = "Employee ID"
            .Width = 85
        End With
        With EmpNme
            .DataPropertyName = "EmpName"
            .HeaderText = "Employee Name"
            .Width = 130
        End With
        With Designation
            .DataPropertyName = "Designation"
            .HeaderText = "Designation"
            .Width = 115
        End With
        With BasicSalary
            .DataPropertyName = "BSalary"
            .HeaderText = "Basic Salary"
            .Width = 70
        End With
        With Addre
            .DataPropertyName = "Address"
            .HeaderText = "Address"
            .Width = 165
        End With
        With city1
            .DataPropertyName = "City"
            .HeaderText = "City"
            .Width = 90
        End With
        With State1
            .DataPropertyName = "State"
            .HeaderText = "State"
            .Width = 100
        End With
        With Zipcode1
            .DataPropertyName = "zipcode"
            .HeaderText = "zip code"
            .Width = 60
        End With
        With MobNo
            .DataPropertyName = "MobileNo"
            .HeaderText = "Mobile No"
            .Width = 90
        End With
        With emailid1
            .DataPropertyName = "EmailID"
            .HeaderText = "EmailID"
            .Width = 135
        End With


        With dgvEmpDetails2
            .DataSource = Dst
            .DataMember = "NonTechStaffEmpDetails"
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ShowRowErrors = False
            .ShowCellErrors = False
            .AllowUserToAddRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .RowHeadersVisible = True
            .RowHeadersWidth = 32
            .DefaultCellStyle.SelectionBackColor = Color.SteelBlue
            .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGoldenrodYellow
            .Columns.AddRange(New DataGridViewColumn() {EmployeeID, EmpNme, Designation, BasicSalary, Addre, city1, State1, Zipcode1, MobNo, emailid1})
        End With
    End Sub



    Dim PGStaffSalaryCode As New DataGridViewTextBoxColumn()
    Dim SNO As New DataGridViewTextBoxColumn()
    Dim EmplyID As New DataGridViewTextBoxColumn()
    Dim EmplyName As New DataGridViewTextBoxColumn()
    Dim Desigtion As New DataGridViewTextBoxColumn()
    Dim EmpBSalary As New DataGridViewTextBoxColumn()
    Dim EmpAGP As New DataGridViewTextBoxColumn()
    Dim Total As New DataGridViewTextBoxColumn()
    Dim EmpDA As New DataGridViewTextBoxColumn()
    Dim HRA As New DataGridViewTextBoxColumn()
    Dim GSalary As New DataGridViewTextBoxColumn()
    Dim IT As New DataGridViewTextBoxColumn()
    Dim ProfTax As New DataGridViewTextBoxColumn()
    Dim TotDeduction As New DataGridViewTextBoxColumn()
    Dim NetSalary As New DataGridViewTextBoxColumn()
   
    Public Sub GridViewPGStaffSalaryDetails(ByVal dgvPGStaffSalaryDetails As DataGridView)
        dgvPGStaffSalaryDetails.AutoGenerateColumns = False

        With PGStaffSalaryCode
            .DataPropertyName = "PGStaffSalaryCode"
            .HeaderText = "Salary Code"
            .Width = 80
        End With
        With SNO
            .DataPropertyName = "SlNo"
            .HeaderText = "Sl. No"
            .Width = 45
        End With
        With EmplyID
            .DataPropertyName = "EmpID"
            .HeaderText = "Employee ID"
            .Width = 85
        End With
        With EmplyName
            .DataPropertyName = "EmpName"
            .HeaderText = "Employee Name"
            .Width = 130
        End With
        With Desigtion
            .DataPropertyName = "Designation"
            .HeaderText = "Designation"
            .Width = 110
        End With
        With EmpBSalary
            .DataPropertyName = "BSalary"
            .HeaderText = "Basic Salary"
            .Width = 70
        End With
        With EmpAGP
            .DataPropertyName = "AGP"
            .HeaderText = "AGP"
            .Width = 70
        End With
        With Total
            .DataPropertyName = "Total"
            .HeaderText = "Total"
            .Width = 75
        End With
        With EmpDA
            .DataPropertyName = "DA"
            .HeaderText = "D.A"
            .Width = 70
        End With
        With HRA
            .DataPropertyName = "HRA"
            .HeaderText = "HRA"
            .Width = 70
        End With
        With GSalary
            .DataPropertyName = "GSalary"
            .HeaderText = "Gross Salary"
            .Width = 75
        End With
        With IT
            .DataPropertyName = "IT"
            .HeaderText = "IT"
            .Width = 70
        End With
        With ProfTax
            .DataPropertyName = "ProfTax"
            .HeaderText = "Prof Tax"
            .Width = 70
        End With
        With TotDeduction
            .DataPropertyName = "TotDeduction"
            .HeaderText = "Total Deduction"
            .Width = 75
        End With
        With NetSalary
            .DataPropertyName = "NetSalary"
            .HeaderText = "Net Salary"
            .Width = 75
        End With


        With dgvPGStaffSalaryDetails
            .DataSource = Dst
            .DataMember = "PGStaffSalaryDetails"
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ShowRowErrors = False
            .ShowCellErrors = False
            .AllowUserToAddRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .RowHeadersVisible = True
            .RowHeadersWidth = 32
            .DefaultCellStyle.SelectionBackColor = Color.SteelBlue
            .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGoldenrodYellow
            .Columns.AddRange(New DataGridViewColumn() {PGStaffSalaryCode, SNO, EmplyID, EmplyName, Desigtion, EmpBSalary, EmpAGP, Total, EmpDA, HRA, GSalary, IT, ProfTax, TotDeduction, NetSalary})
        End With
    End Sub




    Dim NonTeachSalaryCode As New DataGridViewTextBoxColumn()
    Dim SlNO As New DataGridViewTextBoxColumn()
    Dim EmployID As New DataGridViewTextBoxColumn()
    Dim EmployName As New DataGridViewTextBoxColumn()
    Dim Design As New DataGridViewTextBoxColumn()
    Dim EmpGSalary As New DataGridViewTextBoxColumn()
    Dim PF As New DataGridViewTextBoxColumn()
    Dim EmpProfTax As New DataGridViewTextBoxColumn()
    Dim TotDed As New DataGridViewTextBoxColumn()
    Dim NetSalry As New DataGridViewTextBoxColumn()
    Dim MGMPF As New DataGridViewTextBoxColumn()
    
    Public Sub GridViewNonTeachStaffSalaryDetails(ByVal dgvNonTeachStaffSalaryDetails As DataGridView)
        dgvNonTeachStaffSalaryDetails.AutoGenerateColumns = False

        With NonTeachSalaryCode
            .DataPropertyName = "NonTeachSalaryCode"
            .HeaderText = "Salary Code"
            .Width = 80
        End With
        With SlNO
            .DataPropertyName = "SlNo"
            .HeaderText = "Sl. No"
            .Width = 45
        End With
        With EmployID
            .DataPropertyName = "EmpID"
            .HeaderText = "Employee ID"
            .Width = 85
        End With
        With EmployName
            .DataPropertyName = "EmpName"
            .HeaderText = "Employee Name"
            .Width = 130
        End With
        With Design
            .DataPropertyName = "Designation"
            .HeaderText = "Designation"
            .Width = 110
        End With
        With EmpGSalary
            .DataPropertyName = "GrossSalary"
            .HeaderText = "Gross Salary"
            .Width = 70
        End With
        With PF
            .DataPropertyName = "PF"
            .HeaderText = "PF"
            .Width = 70
        End With
        With EmpProfTax
            .DataPropertyName = "ProfTax"
            .HeaderText = "Prof Tax"
            .Width = 70
        End With
        With TotDed
            .DataPropertyName = "TotDeduction"
            .HeaderText = "Total Deduction"
            .Width = 75
        End With
        With NetSalry
            .DataPropertyName = "NetSalary"
            .HeaderText = "Net Salary"
            .Width = 75
        End With
        With MGMPF
            .DataPropertyName = "MGMPF"
            .HeaderText = "MGMPF"
            .Width = 70
        End With


        With dgvNonTeachStaffSalaryDetails
            .DataSource = Dst
            .DataMember = "DetailNonTechingStaff"
            .ReadOnly = True
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ShowRowErrors = False
            .ShowCellErrors = False
            .AllowUserToAddRows = False
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .RowHeadersVisible = True
            .RowHeadersWidth = 32
            .DefaultCellStyle.SelectionBackColor = Color.SteelBlue
            .AlternatingRowsDefaultCellStyle.BackColor = Color.LightGoldenrodYellow
            .Columns.AddRange(New DataGridViewColumn() {NonTeachSalaryCode, SlNO, EmployID, EmployName, Design, EmpGSalary, PF, EmpProfTax, TotDed, NetSalry, MGMPF})
        End With
    End Sub
End Module
