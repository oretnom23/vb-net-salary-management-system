Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text

Public Class FrmNonTeachStaffSalaryDetails

    Dim rdr As SqlDataReader = Nothing
    Dim cmd As SqlCommand = Nothing
    Dim sSql As String
    Dim con As SqlConnection = Nothing
    Dim cs As String = "Data Source=HASSAN-PC\SQLEXPRESS;Initial Catalog=Collage;User ID=sa;Password=ali123;"

    Private Sub auto()
        txtSalaryCode.Text = "SCd-" & GetUniqueKey(4)
    End Sub
    Public Shared Function GetUniqueKey(ByVal maxSize As Integer) As String
        Dim chars As Char() = New Char(61) {}
        chars = "123456789".ToCharArray()
        Dim data As Byte() = New Byte(0) {}
        Dim crypto As New RNGCryptoServiceProvider()
        crypto.GetNonZeroBytes(data)
        data = New Byte(maxSize - 1) {}
        crypto.GetNonZeroBytes(data)
        Dim result As New StringBuilder(maxSize)
        For Each b As Byte In data
            result.Append(chars(b Mod (chars.Length)))
        Next
        Return result.ToString()
    End Function

    Private Sub FrmNonTeachStaffSalaryDetails_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            Me.Dispose()
            FrmMain.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Hide()
            FrmNonTeachEmpDetails2.Show()
            FrmNonTeachEmpDetails2.Button3.PerformClick()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            Me.Close()
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        If Len(Trim(txtEmpID.Text)) = 0 Then
            MessageBox.Show("Please select employee id!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtEmpID.Focus()
            Exit Sub
        End If
        If Len(Trim(txtEmpName.Text)) = 0 Then
            MessageBox.Show("Please enter employee name!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtEmpName.Focus()
            Exit Sub
        End If
        If Len(Trim(txtdesignation.Text)) = 0 Then
            MessageBox.Show("Please enter designation!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtdesignation.Focus()
            Exit Sub
        End If

        If Len(Trim(txtGrossSalary.Text)) = 0 Then
            MessageBox.Show("Please enter gross salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtGrossSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(txtPF.Text)) = 0 Then
            MessageBox.Show("Please enter PF!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtPF.Focus()
            Exit Sub
        End If
        If Len(Trim(txtProfTax.Text)) = 0 Then
            MessageBox.Show("Please enter prof tax!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtProfTax.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTotDeduction.Text)) = 0 Then
            MessageBox.Show("Please enter total deduction!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotDeduction.Focus()
            Exit Sub
        End If
        If Len(Trim(txtNetSalary.Text)) = 0 Then
            MessageBox.Show("Please enter net salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtNetSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(txtMGMPF.Text)) = 0 Then
            MessageBox.Show("Please enter MGMPF!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtMGMPF.Focus()
            Exit Sub
        End If

        Try
            Dim temp As Integer
            temp = ListView1.Items.Count()
            If temp = 0 Then
                Dim i As Integer
                Dim lst As New ListViewItem(i)
                lst.SubItems.Add(temp + 1)
                lst.SubItems.Add(txtEmpID.Text)
                lst.SubItems.Add(txtEmpName.Text)
                lst.SubItems.Add(txtdesignation.Text)
                lst.SubItems.Add(txtGrossSalary.Text)
                lst.SubItems.Add(txtPF.Text)
                lst.SubItems.Add(txtProfTax.Text)
                lst.SubItems.Add(txtTotDeduction.Text)
                lst.SubItems.Add(txtNetSalary.Text)
                lst.SubItems.Add(txtMGMPF.Text)

                ListView1.Items.Add(lst)
                i = i + 1

                totGSalary.Text = Grossalary()
                totPF.Text = PF()
                totProfTax.Text = ProfTax()
                totDeduction.Text = TotalDeduction()
                totNetSalary.Text = NetSalary()
                totMGMPF.Text = MGMPF()

                txtEmpID.Text = ""
                txtEmpName.Text = ""
                txtdesignation.Text = ""
                txtGrossSalary.Text = ""
                txtPF.Text = ""
                txtProfTax.Text = ""
                txtTotDeduction.Text = ""
                txtNetSalary.Text = ""
                txtMGMPF.Text = ""

                Button1.Focus()
                Exit Sub
            End If

            For j = 0 To temp - 1
                If (ListView1.Items(j).SubItems(2).Text = txtEmpID.Text) Then
                    MessageBox.Show("This employee record already exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtEmpID.Text = ""
                    txtEmpName.Text = ""
                    txtdesignation.Text = ""
                    txtGrossSalary.Text = ""
                    txtPF.Text = ""
                    txtProfTax.Text = ""
                    txtTotDeduction.Text = ""
                    txtNetSalary.Text = ""
                    txtMGMPF.Text = ""

                    Button1.Focus()
                    Exit Sub

                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)

            lst1.SubItems.Add(temp + 1)
            lst1.SubItems.Add(txtEmpID.Text)
            lst1.SubItems.Add(txtEmpName.Text)
            lst1.SubItems.Add(txtdesignation.Text)
            lst1.SubItems.Add(txtGrossSalary.Text)
            lst1.SubItems.Add(txtPF.Text)
            lst1.SubItems.Add(txtProfTax.Text)
            lst1.SubItems.Add(txtTotDeduction.Text)
            lst1.SubItems.Add(txtNetSalary.Text)
            lst1.SubItems.Add(txtMGMPF.Text)

            ListView1.Items.Add(lst1)
            k = k + 1


            totGSalary.Text = Grossalary()
            totPF.Text = PF()
            totProfTax.Text = ProfTax()
            totDeduction.Text = TotalDeduction()
            totNetSalary.Text = NetSalary()
            totMGMPF.Text = MGMPF()

            txtEmpID.Text = ""
            txtEmpName.Text = ""
            txtdesignation.Text = ""
            txtGrossSalary.Text = ""
            txtPF.Text = ""
            txtProfTax.Text = ""
            txtTotDeduction.Text = ""
            txtNetSalary.Text = ""
            txtMGMPF.Text = ""

            Button1.Focus()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function Grossalary() As Double

        Dim i, j As Integer
        Dim Gsalary As Double
        i = 0
        j = 0
        Gsalary = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                Gsalary = Gsalary + CInt(ListView1.Items(i).SubItems(5).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Gsalary

    End Function

    Public Function PF() As Double

        Dim i, j As Integer
        Dim ValPF As Double
        i = 0
        j = 0
        ValPF = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                ValPF = ValPF + CInt(ListView1.Items(i).SubItems(6).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return ValPF

    End Function

    Public Function ProfTax() As Double

        Dim i, j As Integer
        Dim ValProfTax As Double
        i = 0
        j = 0
        ValProfTax = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                ValProfTax = ValProfTax + CInt(ListView1.Items(i).SubItems(7).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return ValProfTax

    End Function

    Public Function TotalDeduction() As Double

        Dim i, j As Integer
        Dim ValTotDeduction As Double
        i = 0
        j = 0
        ValTotDeduction = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                ValTotDeduction = ValTotDeduction + CInt(ListView1.Items(i).SubItems(8).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return ValTotDeduction

    End Function

    Public Function NetSalary() As Double

        Dim i, j As Integer
        Dim ValNetSalary As Double
        i = 0
        j = 0
        ValNetSalary = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                ValNetSalary = ValNetSalary + CInt(ListView1.Items(i).SubItems(9).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return ValNetSalary

    End Function

    Public Function MGMPF() As Double

        Dim i, j As Integer
        Dim ValMgmpf As Double
        i = 0
        j = 0
        ValMgmpf = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                ValMgmpf = ValMgmpf + CInt(ListView1.Items(i).SubItems(10).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return ValMgmpf

    End Function

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        btnRemove.Enabled = True
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try

            If ListView1.Items.Count = 0 Then
                MsgBox("No items to remove", MsgBoxStyle.Critical, "Error")
            Else

                ListView1.FocusedItem.Remove()

                totGSalary.Text = Grossalary()
                totPF.Text = PF()
                totProfTax.Text = ProfTax()
                totDeduction.Text = TotalDeduction()
                totNetSalary.Text = NetSalary()
                totMGMPF.Text = MGMPF()

                Dim listCount, p As Integer
                listCount = ListView1.Items.Count
                For p = 0 To listCount - 1
                    ListView1.Items(p).SubItems(1).Text = p + 1
                Next

            End If

            btnRemove.Enabled = False
            If ListView1.Items.Count = 0 Then

                totGSalary.Text = ""
                totPF.Text = ""
                totProfTax.Text = ""
                totDeduction.Text = ""
                totNetSalary.Text = ""
                totMGMPF.Text = ""

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click

        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry! no employee record is added..", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If


        If Len(Trim(totGSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of gross salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totGSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(totPF.Text)) = 0 Then
            MessageBox.Show("Please enter total of PF!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totPF.Focus()
            Exit Sub
        End If
        If Len(Trim(totProfTax.Text)) = 0 Then
            MessageBox.Show("Please enter total of prof tax!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totProfTax.Focus()
            Exit Sub
        End If
        If Len(Trim(totDeduction.Text)) = 0 Then
            MessageBox.Show("Please enter sub total of total deduction!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totDeduction.Focus()
            Exit Sub
        End If
        If Len(Trim(totNetSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of net salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totNetSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(totMGMPF.Text)) = 0 Then
            MessageBox.Show("Please enter total of MGMPF!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totMGMPF.Focus()
            Exit Sub
        End If

        Try
            auto()
            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Open()

            Dim ct As String = "select NonTeachSalaryCode from MasterNonTechingStaff where NonTeachSalaryCode=@find"

            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.VarChar, 15, "NonTeachSalaryCode"))
            cmd.Parameters("@find").Value = txtSalaryCode.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Salary Code.  Already Exists!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If Not rdr Is Nothing Then
                    rdr.Close()
                End If

            Else



                con = New SqlConnection(cs)
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Open()

                Dim cb As String = "insert Into MasterNonTechingStaff(NonTeachSalaryCode,SalaryDT1,SalaryDT2,TotGrossSalary,TotPF,TotProfTax,TotalDeduction,TotNetSalary,TotMGMPF) VALUES ('" & txtSalaryCode.Text & "','" & dtpSalaryDate1.Text & "','" & dtpSalaryDate2.Text & "','" & CInt(totGSalary.Text) & "','" & CInt(totPF.Text) & "','" & CDbl(totProfTax.Text) & "','" & CInt(totDeduction.Text) & "','" & CInt(totNetSalary.Text) & "','" & CInt(totMGMPF.Text) & "')"
                cmd = New SqlCommand(cb)
                cmd.Connection = con
                cmd.ExecuteReader()
                con.Close()


                For i = 0 To ListView1.Items.Count - 1
                    con = New SqlConnection(cs)
                    If con.State = ConnectionState.Open Then
                        con.Close()
                    End If
                    con.Open()
                    Dim cd As String = "insert Into DetailNonTechingStaff(NonTeachSalaryCode,SalaryDT1,SalaryDT2,SlNo,EmpID,EmpName,Designation,GrossSalary,PF,ProfTax,TotDeduction,NetSalary,MGMPF) VALUES (@NonTeachSalaryCode,@SalaryDT1,@SalaryDT2,@SlNo,@EmpID,@EmpName,@Designation,@GrossSalary,@PF,@ProfTax,@TotDeduction,@NetSalary,@MGMPF)"
                    cmd = New SqlCommand(cd)

                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("NonTeachSalaryCode", txtSalaryCode.Text)
                    cmd.Parameters.AddWithValue("SalaryDT1", dtpSalaryDate1.Text)
                    cmd.Parameters.AddWithValue("SalaryDT2", dtpSalaryDate2.Text)
                    cmd.Parameters.AddWithValue("SlNo", ListView1.Items(i).SubItems(1).Text)
                    cmd.Parameters.AddWithValue("EmpID", ListView1.Items(i).SubItems(2).Text)
                    cmd.Parameters.AddWithValue("EmpName", ListView1.Items(i).SubItems(3).Text)
                    cmd.Parameters.AddWithValue("Designation", ListView1.Items(i).SubItems(4).Text)
                    cmd.Parameters.AddWithValue("GrossSalary", ListView1.Items(i).SubItems(5).Text)
                    cmd.Parameters.AddWithValue("PF", ListView1.Items(i).SubItems(6).Text)
                    cmd.Parameters.AddWithValue("ProfTax", ListView1.Items(i).SubItems(7).Text)
                    cmd.Parameters.AddWithValue("TotDeduction", ListView1.Items(i).SubItems(8).Text)
                    cmd.Parameters.AddWithValue("NetSalary", ListView1.Items(i).SubItems(9).Text)
                    cmd.Parameters.AddWithValue("MGMPF", ListView1.Items(i).SubItems(10).Text)
                    
                    cmd.ExecuteNonQuery()
                    con.Close()
                Next
                Save.Enabled = False
                Save.Visible = False
                NewRecord.Enabled = True
                NewRecord.Visible = True
                cmdCancel.Enabled = False
                cmdCancel.Visible = False
                Update_record.Enabled = True
                Delete.Enabled = True
                cmdClose.Enabled = True
                cmdClose.Visible = True
                Button2.Enabled = True
                BSubmit.Enabled = True
                btnRemove.Enabled = False
                Button2.Focus()
                MessageBox.Show("Successfully placed", "Order", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub NewRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        Try

            txtSalaryCode.Text = ""
            dtpSalaryDate1.Text = Today
            dtpSalaryDate2.Text = Today
            txtEmpID.Text = ""
            txtEmpName.Text = ""
            txtdesignation.Text = ""
            txtGrossSalary.Text = ""
            txtPF.Text = ""
            txtProfTax.Text = ""
            txtTotDeduction.Text = ""
            txtNetSalary.Text = ""
            txtMGMPF.Text = ""
            

            ListView1.Items.Clear()
            Save.Enabled = True
            Save.Visible = True
            NewRecord.Enabled = False
            NewRecord.Visible = False
            cmdCancel.Enabled = True
            cmdCancel.Visible = True
            Update_record.Enabled = False
            Delete.Enabled = False
            cmdClose.Enabled = False
            cmdClose.Visible = False
            Button1.Focus()
            btnRemove.Enabled = False

            Button1.Enabled = True
            Button7.Enabled = True
            Button2.Enabled = False
            BSubmit.Enabled = False
            Button1.Focus()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Dispose()
            NonTeachStaffSalaryDetails.Show()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Save.Enabled = False
        Save.Visible = False
        NewRecord.Enabled = True
        NewRecord.Visible = True
        cmdCancel.Enabled = False
        cmdCancel.Visible = False
        Update_record.Enabled = False
        Delete.Enabled = False
        cmdClose.Enabled = True
        cmdClose.Visible = True
        Button2.Enabled = True
        BSubmit.Enabled = True
        Button1.Enabled = False
        Button7.Enabled = False
        Call clear()
        btnRemove.Enabled = False
        ListView1.Items.Clear()
        Button2.Focus()
    End Sub

    Sub clear()
        txtSalaryCode.Text = ""
        dtpSalaryDate1.Text = Today
        dtpSalaryDate2.Text = Today
        txtEmpID.Text = ""
        txtEmpName.Text = ""
        txtdesignation.Text = ""
        txtGrossSalary.Text = ""
        txtPF.Text = ""
        txtProfTax.Text = ""
        txtTotDeduction.Text = ""
        txtNetSalary.Text = ""
        txtMGMPF.Text = ""
        totGSalary.Text = ""
        totPF.Text = ""
        totDeduction.Text = ""
        totProfTax.Text = ""
        totNetSalary.Text = ""
        totMGMPF.Text = ""
    End Sub

    Public Sub FillList()
        With ListView1
            .Clear()
            .Columns.Add("ColumnHeader", 0)
            .Columns.Add("Sl No", 50, HorizontalAlignment.Left)
            .Columns.Add("Emp ID", 70, HorizontalAlignment.Center)
            .Columns.Add("Employee Name", 125, HorizontalAlignment.Center)
            .Columns.Add("Designation", 110, HorizontalAlignment.Center)
            .Columns.Add("G Salary", 70, HorizontalAlignment.Center)
            .Columns.Add("PF", 65, HorizontalAlignment.Center)
            .Columns.Add("Prof Tax", 70, HorizontalAlignment.Center)
            .Columns.Add("Tot Deduction", 75, HorizontalAlignment.Center)
            .Columns.Add("Net Salary", 75, HorizontalAlignment.Center)
            .Columns.Add("MGMPF", 75, HorizontalAlignment.Center)
            FillListView(ListView1, GetData(sSql))
        End With
    End Sub

    Public Function GetData(ByVal sSQL As String)

        Dim sqlCmd As SqlCommand = New SqlCommand(sSQL)
        Dim myData As SqlDataReader
        con = New SqlConnection(cs)
        Try
            con.Open()
            sqlCmd.Connection = con
            myData = sqlCmd.ExecuteReader
            Return myData
        Catch ex As Exception
            Return ex
        End Try
    End Function
    Private Sub BSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSubmit.Click
        Try
            If (txtSalaryCode.Text = "") Then
                MessageBox.Show("Retrieve salary code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Button2.Focus()
                Exit Sub
            End If
            sSql = "SELECT SlNo,SlNo,EmpID,EmpName,Designation,GrossSalary,PF,ProfTax,TotDeduction,NetSalary,MGMPF from DetailNonTechingStaff  where NonTeachSalaryCode = '" & txtSalaryCode.Text & "'"
            Call FillList()

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim ct As String = "select SalaryDT1,SalaryDT2,TotGrossSalary,TotPF,TotProfTax,TotalDeduction,TotNetSalary,TotMGMPF from MasterNonTechingStaff where NonTeachSalaryCode=@find"

            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.VarChar, 15, "NonTeachSalaryCode"))
            cmd.Parameters("@find").Value = Trim(txtSalaryCode.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then

                dtpSalaryDate1.Text = Trim(rdr("SalaryDT1")).ToString()
                dtpSalaryDate2.Text = Trim(rdr("SalaryDT2")).ToString()
                totGSalary.Text = Trim(rdr("TotGrossSalary")).ToString()
                totPF.Text = Trim(rdr("TotPF")).ToString()
                totProfTax.Text = Trim(rdr("TotProfTax")).ToString()
                totDeduction.Text = Trim(rdr("TotalDeduction")).ToString()
                totNetSalary.Text = Trim(rdr("TotNetSalary")).ToString()
                totMGMPF.Text = Trim(rdr("TotMGMPF")).ToString()
                
            End If
            Button1.Enabled = True
            Button7.Enabled = True
            Me.Delete.Enabled = True
            Me.Update_record.Enabled = True
            Me.Save.Enabled = False
            Me.Save.Visible = False
            Me.NewRecord.Enabled = True
            Me.NewRecord.Visible = True
            Me.cmdCancel.Enabled = False
            Me.cmdCancel.Visible = False
            Me.cmdClose.Enabled = True
            Me.cmdClose.Visible = True
            Me.Update_record.Enabled = True
            Me.Delete.Enabled = True
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FrmNonTeachStaffSalaryDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Button1.Enabled = False
            Button7.Enabled = False
            NewRecord.Focus()
            dtpSalaryDate1.Text = Today()
            dtpSalaryDate2.Text = Today()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Update_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Update_record.Click
        If Save.Visible = True Or Save.Enabled = True Then
            MsgBox("Rcorcd modification failed ", MsgBoxStyle.Critical, "Updation Failed")
            Exit Sub
        End If

        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry! no employee record is added..", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If


        If Len(Trim(totGSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of gross salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totGSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(totPF.Text)) = 0 Then
            MessageBox.Show("Please enter total of PF!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totPF.Focus()
            Exit Sub
        End If
        If Len(Trim(totProfTax.Text)) = 0 Then
            MessageBox.Show("Please enter total of prof tax!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totProfTax.Focus()
            Exit Sub
        End If
        If Len(Trim(totDeduction.Text)) = 0 Then
            MessageBox.Show("Please enter sub total of total deduction!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totDeduction.Focus()
            Exit Sub
        End If
        If Len(Trim(totNetSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of net salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totNetSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(totMGMPF.Text)) = 0 Then
            MessageBox.Show("Please enter total of MGMPF!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totMGMPF.Focus()
            Exit Sub
        End If
      

        Try

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cb As String = "update MasterNonTechingStaff set SalaryDT1='" & dtpSalaryDate1.Text & "',SalaryDT2='" & dtpSalaryDate2.Text & "',TotGrossSalary='" & totGSalary.Text & "',TotPF='" & totPF.Text & "',TotProfTax= '" & totProfTax.Text & "',TotalDeduction='" & totDeduction.Text & "',TotNetSalary='" & totNetSalary.Text & "',TotMGMPF='" & totMGMPF.Text & "' where NonTeachSalaryCode='" & txtSalaryCode.Text & "'"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cq1 As String = "delete from DetailNonTechingStaff where NonTeachSalaryCode=@DELETE1;"

            cmd = New SqlCommand(cq1)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@DELETE1", System.Data.SqlDbType.VarChar, 15, "NonTeachSalaryCode"))
            cmd.Parameters("@DELETE1").Value = Trim(txtSalaryCode.Text)
            cmd.ExecuteNonQuery()
            con.Close()

            For i = 0 To ListView1.Items.Count - 1
                con = New SqlConnection(cs)
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Open()
                Dim cd As String = "insert Into DetailNonTechingStaff(NonTeachSalaryCode,SlNo,SalaryDT1,SalaryDT2,EmpID,EmpName,Designation,GrossSalary,PF,ProfTax,TotDeduction,NetSalary,MGMPF) VALUES (@NonTeachSalaryCode,@SlNo,@SalaryDT1,@SalaryDT2,@EmpID,@EmpName,@Designation,@GrossSalary,@PF,@ProfTax,@TotDeduction,@NetSalary,@MGMPF)"
                cmd = New SqlCommand(cd)

                cmd.Connection = con
                cmd.Parameters.AddWithValue("NonTeachSalaryCode", txtSalaryCode.Text)
                cmd.Parameters.AddWithValue("SlNo", ListView1.Items(i).SubItems(1).Text)
                cmd.Parameters.AddWithValue("SalaryDT1", dtpSalaryDate1.Text)
                cmd.Parameters.AddWithValue("SalaryDT2", dtpSalaryDate2.Text)
                cmd.Parameters.AddWithValue("EmpID", ListView1.Items(i).SubItems(2).Text)
                cmd.Parameters.AddWithValue("EmpName", ListView1.Items(i).SubItems(3).Text)
                cmd.Parameters.AddWithValue("Designation", ListView1.Items(i).SubItems(4).Text)
                cmd.Parameters.AddWithValue("GrossSalary", ListView1.Items(i).SubItems(5).Text)
                cmd.Parameters.AddWithValue("PF", ListView1.Items(i).SubItems(6).Text)
                cmd.Parameters.AddWithValue("ProfTax", ListView1.Items(i).SubItems(7).Text)
                cmd.Parameters.AddWithValue("TotDeduction", ListView1.Items(i).SubItems(8).Text)
                cmd.Parameters.AddWithValue("NetSalary", ListView1.Items(i).SubItems(9).Text)
                cmd.Parameters.AddWithValue("MGMPF", ListView1.Items(i).SubItems(10).Text)
                cmd.ExecuteNonQuery()
                con.Close()
            Next

            NewRecord.Focus()
            Update_record.Enabled = False
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub delete_records()
        Try

            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cq1 As String = "delete from DetailNonTechingStaff where NonTeachSalaryCode=@DELETE1;"

            cmd = New SqlCommand(cq1)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@DELETE1", System.Data.SqlDbType.VarChar, 15, "NonTeachSalaryCode"))
            cmd.Parameters("@DELETE1").Value = Trim(TxtSalaryCode.Text)
            cmd.ExecuteNonQuery()
            con.Close()

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cq As String = "delete from MasterNonTechingStaff where NonTeachSalaryCode=@DELETE2;"

            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@DELETE2", System.Data.SqlDbType.VarChar, 15, "NonTeachSalaryCode"))

            cmd.Parameters("@DELETE2").Value = Trim(TxtSalaryCode.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Call clear()
                Delete.Enabled = False
                Update_record.Enabled = False
                Button1.Enabled = False
                Button7.Enabled = False
                ListView1.Items.Clear()
                con.Close()

            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Call clear()
                Delete.Enabled = False
                Update_record.Enabled = False
                Button1.Enabled = False
                Button7.Enabled = False
                con.Close()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Delete.Click
        If Save.Visible = True Or Save.Enabled = True Then
            MsgBox("Record deletion failed ", MsgBoxStyle.Critical, "Deletion Failed")
            Exit Sub
        End If
        Try

            If MessageBox.Show("Do you really want to delete this record?", "Order Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub totProfTax_TextChanged(sender As Object, e As EventArgs) Handles totProfTax.TextChanged

    End Sub
End Class