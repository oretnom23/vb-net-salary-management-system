Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text


Public Class EmployeeSalaryDetails

    Dim rdr As SqlDataReader = Nothing
    Dim cmd As SqlCommand = Nothing
    Dim sSql As String
    Dim con As SqlConnection = Nothing
    Dim cs As String = "Data Source=HASSAN-PC\SQLEXPRESS;Initial Catalog=Collage;User ID=sa;Password=ali123;"

    Private Sub auto()
        TxtSalaryCode.Text = "SCd-" & GetUniqueKey(4)
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Hide()
            FrmPGStaffSalaryDetails.Show()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EmployeeSalaryDetails_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Dispose()
            FrmMain.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Me.Hide()
            FrmEmpDetails2.Show()
            FrmEmpDetails2.Button3.PerformClick()
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
            FrmMain.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        If Len(Trim(txtDesignation.Text)) = 0 Then
            MessageBox.Show("Please enter designation!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtDesignation.Focus()
            Exit Sub
        End If
        If Len(Trim(txtBSalary.Text)) = 0 Then
            MessageBox.Show("Please enter basic salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtBSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(txtAGP.Text)) = 0 Then
            MessageBox.Show("Please enter AGP!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtAGP.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTotal.Text)) = 0 Then
            MessageBox.Show("Please enter total!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotal.Focus()
            Exit Sub
        End If
        If Len(Trim(txtDA.Text)) = 0 Then
            MessageBox.Show("Please enter DA!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtDA.Focus()
            Exit Sub
        End If
        If Len(Trim(txtHRA.Text)) = 0 Then
            MessageBox.Show("Please enter HRA!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtHRA.Focus()
            Exit Sub
        End If
        If Len(Trim(txtGrossSalary.Text)) = 0 Then
            MessageBox.Show("Please enter gross salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtGrossSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(txtIT.Text)) = 0 Then
            MessageBox.Show("Please enter IT!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtIT.Focus()
            Exit Sub
        End If
        If Len(Trim(txtProfTax.Text)) = 0 Then
            MessageBox.Show("Please enter Prof Tax!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtProfTax.Focus()
            Exit Sub
        End If
        If Len(Trim(txtTotalDeduction.Text)) = 0 Then
            MessageBox.Show("Please enter total deduction!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalDeduction.Focus()
            Exit Sub
        End If
        If Len(Trim(txtNetSalary.Text)) = 0 Then
            MessageBox.Show("Please enter net salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtNetSalary.Focus()
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
                lst.SubItems.Add(txtDesignation.Text)
                lst.SubItems.Add(txtBSalary.Text)
                lst.SubItems.Add(txtAGP.Text)
                lst.SubItems.Add(txtTotal.Text)
                lst.SubItems.Add(txtDA.Text)
                lst.SubItems.Add(txtHRA.Text)
                lst.SubItems.Add(txtGrossSalary.Text)
                lst.SubItems.Add(txtIT.Text)
                lst.SubItems.Add(txtProfTax.Text)
                lst.SubItems.Add(txtTotalDeduction.Text)
                lst.SubItems.Add(txtNetSalary.Text)

                ListView1.Items.Add(lst)
                i = i + 1

                TotBSalary.Text = Basicsalary()
                TotalAGP.Text = TotAGP()
                TotTotal.Text = SubTot()
                TotalDa.Text = TotDA()
                TotalHRA.Text = TotHRA()
                TotalGSalary.Text = TotGSalary()
                totalIT.Text = TotIT()
                TotalProfTax.Text = TotPofTax()
                TotalDeduction.Text = DeductionSubTot()
                TotalNetSalary.Text = TotNetSalary()

                txtEmpID.Text = ""
                txtEmpName.Text = ""
                txtDesignation.Text = ""
                txtBSalary.Text = ""
                txtAGP.Text = ""
                txtTotal.Text = ""
                txtDA.Text = ""
                txtHRA.Text = ""
                txtGrossSalary.Text = ""
                txtIT.Text = ""
                txtProfTax.Text = " "
                txtTotalDeduction.Text = ""
                txtNetSalary.Text = ""

                Button1.Focus()
                Exit Sub
            End If

            For j = 0 To temp - 1
                If (ListView1.Items(j).SubItems(2).Text = txtEmpID.Text) Then
                    MessageBox.Show("This employee record already exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtEmpID.Text = ""
                    txtEmpName.Text = ""
                    txtDesignation.Text = ""
                    txtBSalary.Text = ""
                    txtAGP.Text = ""
                    txtTotal.Text = ""
                    txtDA.Text = ""
                    txtHRA.Text = ""
                    txtGrossSalary.Text = ""
                    txtIT.Text = ""
                    txtProfTax.Text = " "
                    txtTotalDeduction.Text = ""
                    txtNetSalary.Text = ""
                    Button1.Focus()
                    Exit Sub

                End If
            Next j
            Dim k As Integer
            Dim lst1 As New ListViewItem(k)

            lst1.SubItems.Add(temp + 1)
            lst1.SubItems.Add(txtEmpID.Text)
            lst1.SubItems.Add(txtEmpName.Text)
            lst1.SubItems.Add(txtDesignation.Text)
            lst1.SubItems.Add(txtBSalary.Text)
            lst1.SubItems.Add(txtAGP.Text)
            lst1.SubItems.Add(txtTotal.Text)
            lst1.SubItems.Add(txtDA.Text)
            lst1.SubItems.Add(txtHRA.Text)
            lst1.SubItems.Add(txtGrossSalary.Text)
            lst1.SubItems.Add(txtIT.Text)
            lst1.SubItems.Add(txtProfTax.Text)
            lst1.SubItems.Add(txtTotalDeduction.Text)
            lst1.SubItems.Add(txtNetSalary.Text)
            ListView1.Items.Add(lst1)
            k = k + 1

            TotBSalary.Text = Basicsalary()
            TotalAGP.Text = TotAGP()
            TotTotal.Text = SubTot()
            TotalDa.Text = TotDA()
            TotalHRA.Text = TotHRA()
            TotalGSalary.Text = TotGSalary()
            totalIT.Text = TotIT()
            TotalProfTax.Text = TotPofTax()
            TotalDeduction.Text = DeductionSubTot()
            TotalNetSalary.Text = TotNetSalary()

            txtEmpID.Text = ""
            txtEmpName.Text = ""
            txtDesignation.Text = ""
            txtBSalary.Text = ""
            txtAGP.Text = ""
            txtTotal.Text = ""
            txtDA.Text = ""
            txtHRA.Text = ""
            txtGrossSalary.Text = ""
            txtIT.Text = ""
            txtProfTax.Text = " "
            txtTotalDeduction.Text = ""
            txtNetSalary.Text = ""

            Button1.Focus()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function Basicsalary() As Double

        Dim i, j As Integer
        Dim Bsalary As Double
        i = 0
        j = 0
        Bsalary = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                Bsalary = Bsalary + CInt(ListView1.Items(i).SubItems(5).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Bsalary

    End Function

    Public Function TotAGP() As Double

        Dim i, j As Integer
        Dim AGP As Double
        i = 0
        j = 0
        AGP = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                AGP = AGP + CInt(ListView1.Items(i).SubItems(6).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return AGP

    End Function

    Public Function SubTot() As Double

        Dim i, j As Integer
        Dim SubTotal As Double
        i = 0
        j = 0
        SubTotal = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                SubTotal = SubTotal + CInt(ListView1.Items(i).SubItems(7).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return SubTotal

    End Function

    Public Function TotDA() As Double

        Dim i, j As Integer
        Dim DA As Double
        i = 0
        j = 0
        DA = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                DA = DA + CInt(ListView1.Items(i).SubItems(8).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return DA

    End Function

    Public Function TotHRA() As Double

        Dim i, j As Integer
        Dim HRA As Double
        i = 0
        j = 0
        HRA = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                HRA = HRA + CInt(ListView1.Items(i).SubItems(9).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return HRA

    End Function

    Public Function TotGSalary() As Double

        Dim i, j As Integer
        Dim GSalary As Double
        i = 0
        j = 0
        GSalary = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                GSalary = GSalary + CInt(ListView1.Items(i).SubItems(10).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return GSalary

    End Function

    Public Function TotIT() As Double

        Dim i, j As Integer
        Dim IT As Double
        i = 0
        j = 0
        IT = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                IT = IT + CInt(ListView1.Items(i).SubItems(11).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return IT

    End Function
    Public Function TotPofTax() As Double

        Dim i, j As Integer
        Dim PofTax As Double
        i = 0
        j = 0
        PofTax = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                PofTax = PofTax + CInt(ListView1.Items(i).SubItems(12).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return PofTax

    End Function

    Public Function DeductionSubTot() As Double

        Dim i, j As Integer
        Dim Subtot As Double
        i = 0
        j = 0
        Subtot = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                Subtot = Subtot + CInt(ListView1.Items(i).SubItems(13).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Subtot

    End Function

    Public Function TotNetSalary() As Double

        Dim i, j As Integer
        Dim NetSalary As Double
        i = 0
        j = 0
        NetSalary = 0

        Try
            j = ListView1.Items.Count
            For i = 0 To j - 1
                NetSalary = NetSalary + CInt(ListView1.Items(i).SubItems(14).Text)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return NetSalary

    End Function

    Public Sub FillList()
        With ListView1
            .Clear()
            .Columns.Add("ColumnHeader", 0)
            .Columns.Add("Sl No", 50, HorizontalAlignment.Left)
            .Columns.Add("Emp ID", 60, HorizontalAlignment.Center)
            .Columns.Add("Employee Name", 125, HorizontalAlignment.Center)
            .Columns.Add("Designation", 110, HorizontalAlignment.Center)
            .Columns.Add("B Salary", 65, HorizontalAlignment.Center)
            .Columns.Add("AGP", 65, HorizontalAlignment.Center)
            .Columns.Add("Total", 65, HorizontalAlignment.Center)
            .Columns.Add("D.A", 65, HorizontalAlignment.Center)
            .Columns.Add("HRA", 65, HorizontalAlignment.Center)
            .Columns.Add("G Salary", 75, HorizontalAlignment.Center)
            .Columns.Add("IT", 65, HorizontalAlignment.Center)
            .Columns.Add("Prof Tax", 65, HorizontalAlignment.Center)
            .Columns.Add("Tot Deduction", 75, HorizontalAlignment.Center)
            .Columns.Add("Net Salary", 75, HorizontalAlignment.Center)
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

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Save.Click


        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry! no employee record is added..", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If


        If Len(Trim(TotBSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of basic salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotBSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalAGP.Text)) = 0 Then
            MessageBox.Show("Please enter total of AGP!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalAGP.Focus()
            Exit Sub
        End If
        If Len(Trim(TotTotal.Text)) = 0 Then
            MessageBox.Show("Please enter sub total of total!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotTotal.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalDa.Text)) = 0 Then
            MessageBox.Show("Please enter total of DA!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalDa.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalHRA.Text)) = 0 Then
            MessageBox.Show("Please enter total of HRA!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalHRA.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalGSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of gross salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalGSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(totalIT.Text)) = 0 Then
            MessageBox.Show("Please enter total of IT!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totalIT.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalProfTax.Text)) = 0 Then
            MessageBox.Show("Please enter total of Prof Tax!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalProfTax.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalDeduction.Text)) = 0 Then
            MessageBox.Show("Please enter sub total of total deduction!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalDeduction.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalNetSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of net salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalNetSalary.Focus()
            Exit Sub
        End If

        Try
            auto()
            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Open()

            Dim ct As String = "select PGStaffSalaryCode from PGStaffSalary where PGStaffSalaryCode=@find"

            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.VarChar, 15, "PGStaffSalaryCode"))
            cmd.Parameters("@find").Value = TxtSalaryCode.Text
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

                Dim cb As String = "insert Into PGStaffSalary(PGStaffSalaryCode,SalaryDT1,SalaryDT2,TotBSalary,TotAGP,GrandTotal,TotDA,TotHRA,TotGSalary,TotIT,TotProfTax,GrandTotDeduction,NetSalary) VALUES ('" & TxtSalaryCode.Text & "','" & dtpSalaryDate1.Text & "','" & dtpSalaryDate2.Text & "','" & CInt(TotBSalary.Text) & "','" & CInt(TotalAGP.Text) & "','" & CDbl(TotTotal.Text) & "','" & CInt(TotalDa.Text) & "','" & CInt(TotalHRA.Text) & "','" & CInt(TotalGSalary.Text) & "','" & CInt(totalIT.Text) & "','" & CInt(TotalProfTax.Text) & "','" & CInt(TotalDeduction.Text) & "','" & CInt(TotalNetSalary.Text) & "')"
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
                    Dim cd As String = "insert Into PGStaffSalaryDetails(PGStaffSalaryCode,SlNo,SalaryDT1,SalaryDT2,EmpID,EmpName,Designation,BSalary,AGP,Total,DA,HRA,GSalary,IT,ProfTax,TotDeduction,NetSalary) VALUES (@PGStaffSalaryCode,@SlNo,@SalaryDT1,@SalaryDT2,@EmpID,@EmpName,@Designation,@BSalary,@AGP,@Total,@DA,@HRA,@GSalary,@IT,@ProfTax,@TotDeduction,@NetSalary)"
                    cmd = New SqlCommand(cd)

                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("PGStaffSalaryCode", TxtSalaryCode.Text)
                    cmd.Parameters.AddWithValue("SlNo", ListView1.Items(i).SubItems(1).Text)
                    cmd.Parameters.AddWithValue("SalaryDT1", dtpSalaryDate1.Text)
                    cmd.Parameters.AddWithValue("SalaryDT2", dtpSalaryDate2.Text)
                    cmd.Parameters.AddWithValue("EmpID", ListView1.Items(i).SubItems(2).Text)
                    cmd.Parameters.AddWithValue("EmpName", ListView1.Items(i).SubItems(3).Text)
                    cmd.Parameters.AddWithValue("Designation", ListView1.Items(i).SubItems(4).Text)
                    cmd.Parameters.AddWithValue("BSalary", ListView1.Items(i).SubItems(5).Text)
                    cmd.Parameters.AddWithValue("AGP", ListView1.Items(i).SubItems(6).Text)
                    cmd.Parameters.AddWithValue("Total", ListView1.Items(i).SubItems(7).Text)
                    cmd.Parameters.AddWithValue("DA", ListView1.Items(i).SubItems(8).Text)
                    cmd.Parameters.AddWithValue("HRA", ListView1.Items(i).SubItems(9).Text)
                    cmd.Parameters.AddWithValue("GSalary", ListView1.Items(i).SubItems(10).Text)
                    cmd.Parameters.AddWithValue("IT", ListView1.Items(i).SubItems(11).Text)
                    cmd.Parameters.AddWithValue("ProfTax", ListView1.Items(i).SubItems(12).Text)
                    cmd.Parameters.AddWithValue("TotDeduction", ListView1.Items(i).SubItems(13).Text)
                    cmd.Parameters.AddWithValue("NetSalary", ListView1.Items(i).SubItems(14).Text)

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

            TxtSalaryCode.Text = ""
            dtpSalaryDate1.Text = Today
            dtpSalaryDate2.Text = Today
            txtEmpID.Text = ""
            txtEmpName.Text = ""
            txtDesignation.Text = ""
            txtBSalary.Text = ""
            txtAGP.Text = ""
            txtDA.Text = ""
            txtHRA.Text = ""
            txtGrossSalary.Text = ""
            txtIT.Text = ""
            txtNetSalary.Text = ""
            txtTotal.Text = ""
            txtTotalDeduction.Text = ""
            txtProfTax.Text = ""
            TotalAGP.Text = ""
            TotalDa.Text = ""
            TotalGSalary.Text = ""
            TotalHRA.Text = ""
            totalIT.Text = ""
            TotalNetSalary.Text = ""
            TotalProfTax.Text = ""
            TotTotal.Text = ""
            TotBSalary.Text = ""
            TotalDeduction.Text = ""


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


    Sub clear()
        TxtSalaryCode.Text = ""
        dtpSalaryDate1.Text = Today
        dtpSalaryDate2.Text = Today
        txtEmpID.Text = ""
        txtEmpName.Text = ""
        txtDesignation.Text = ""
        txtBSalary.Text = ""
        txtAGP.Text = ""
        txtDA.Text = ""
        txtHRA.Text = ""
        txtGrossSalary.Text = ""
        txtIT.Text = ""
        txtNetSalary.Text = ""
        txtTotal.Text = ""
        txtTotalDeduction.Text = ""
        txtProfTax.Text = ""
        TotalAGP.Text = ""
        TotalDa.Text = ""
        TotalGSalary.Text = ""
        TotalHRA.Text = ""
        totalIT.Text = ""
        TotalNetSalary.Text = ""
        TotalProfTax.Text = ""
        TotTotal.Text = ""
        TotBSalary.Text = ""
        TotalDeduction.Text = ""
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        btnRemove.Enabled = True
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Try

            If ListView1.Items.Count = 0 Then
                MsgBox("No items to remove", MsgBoxStyle.Critical, "Error")
            Else

                ListView1.FocusedItem.Remove()

                TotBSalary.Text = Basicsalary()
                TotalAGP.Text = TotAGP()
                TotTotal.Text = SubTot()
                TotalDa.Text = TotDA()
                TotalHRA.Text = TotHRA()
                TotalGSalary.Text = TotGSalary()
                totalIT.Text = TotIT()
                TotalProfTax.Text = TotPofTax()
                TotalDeduction.Text = DeductionSubTot()
                TotalNetSalary.Text = TotNetSalary()

                Dim listCount, p As Integer
                listCount = ListView1.Items.Count
                For p = 0 To listCount - 1
                    ListView1.Items(p).SubItems(1).Text = p + 1
                Next

            End If

            btnRemove.Enabled = False
            If ListView1.Items.Count = 0 Then

                TotBSalary.Text = ""
                TotalAGP.Text = ""
                TotTotal.Text = ""
                TotalDa.Text = ""
                TotalHRA.Text = ""
                TotalGSalary.Text = ""
                totalIT.Text = ""
                TotalProfTax.Text = ""
                TotalDeduction.Text = ""
                TotalNetSalary.Text = ""

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Try
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
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EmployeeSalaryDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Button1.Enabled = False
            Button7.Enabled = False
            NewRecord.Focus()
            dtpSalaryDate1.Text = Today()
            dtpSalaryDate2.Text = Today()

        Catch ex As Exception
        End Try

    End Sub

    Private Sub BSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BSubmit.Click
        Try
            If (TxtSalaryCode.Text = "") Then
                MessageBox.Show("Retrieve salary code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Button2.Focus()
                Exit Sub
            End If
            sSql = "SELECT SlNo,SlNo,EmpID,EmpName,Designation,BSalary,AGP,Total,DA,HRA,GSalary,IT,ProfTax,TotDeduction,NetSalary from PGStaffSalaryDetails  where PGStaffSalaryCode = '" & TxtSalaryCode.Text & "'"
            Call FillList()

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim ct As String = "select SalaryDT1,SalaryDT2,TotBSalary,TotAGP,GrandTotal,TotDA,TotHRA,TotGSalary,TotIT,TotProfTax,GrandTotDeduction,NetSalary from PGStaffSalary where PGStaffSalaryCode=@find"

            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.VarChar, 15, "PGStaffSalaryCode"))
            cmd.Parameters("@find").Value = Trim(TxtSalaryCode.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then

                dtpSalaryDate1.Text = Trim(rdr("SalaryDT1")).ToString()
                dtpSalaryDate2.Text = Trim(rdr("SalaryDT2")).ToString()
                TotBSalary.Text = Trim(rdr("TotBSalary")).ToString()
                TotalAGP.Text = Trim(rdr("TotAGP")).ToString()
                TotTotal.Text = Trim(rdr("GrandTotal")).ToString()
                TotalDa.Text = Trim(rdr("TotDA")).ToString()
                TotalHRA.Text = Trim(rdr("TotHRA")).ToString()
                TotalGSalary.Text = Trim(rdr("TotGSalary")).ToString()
                totalIT.Text = Trim(rdr("TotIT")).ToString()
                TotalProfTax.Text = Trim(rdr("TotProfTax")).ToString()
                TotalDeduction.Text = Trim(rdr("GrandTotDeduction")).ToString()
                TotalNetSalary.Text = Trim(rdr("NetSalary")).ToString()

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

    Private Sub Update_record_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Update_record.Click
        If Save.Visible = True Or Save.Enabled = True Then
            MsgBox("Rcorcd modification failed ", MsgBoxStyle.Critical, "Updation Failed")
            Exit Sub
        End If
        If ListView1.Items.Count = 0 Then
            MessageBox.Show("sorry! no employee record is added..", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If


        If Len(Trim(TotBSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of basic salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotBSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalAGP.Text)) = 0 Then
            MessageBox.Show("Please enter total of AGP!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalAGP.Focus()
            Exit Sub
        End If
        If Len(Trim(TotTotal.Text)) = 0 Then
            MessageBox.Show("Please enter sub total of total!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotTotal.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalDa.Text)) = 0 Then
            MessageBox.Show("Please enter total of DA!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalDa.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalHRA.Text)) = 0 Then
            MessageBox.Show("Please enter total of HRA!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalHRA.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalGSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of gross salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalGSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(totalIT.Text)) = 0 Then
            MessageBox.Show("Please enter total of IT!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            totalIT.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalProfTax.Text)) = 0 Then
            MessageBox.Show("Please enter total of Prof Tax!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalProfTax.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalDeduction.Text)) = 0 Then
            MessageBox.Show("Please enter sub total of total deduction!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalDeduction.Focus()
            Exit Sub
        End If
        If Len(Trim(TotalNetSalary.Text)) = 0 Then
            MessageBox.Show("Please enter total of net salary!!!", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TotalNetSalary.Focus()
            Exit Sub
        End If


        Try

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cb As String = "update PGStaffSalary set SalaryDT1='" & dtpSalaryDate1.Text & "',SalaryDT2='" & dtpSalaryDate2.Text & "',TotBSalary='" & TotBSalary.Text & "',TotAGP='" & TotalAGP.Text & "',GrandTotal= '" & TotTotal.Text & "',TotDA='" & TotalDa.Text & "',TotHRA='" & TotalHRA.Text & "',TotGSalary='" & TotalGSalary.Text & "',TotIT='" & totalIT.Text & "',TotProfTax='" & TotalProfTax.Text & "',GrandTotDeduction='" & TotalDeduction.Text & "',NetSalary='" & TotalNetSalary.Text & "' where PGStaffSalaryCode='" & TxtSalaryCode.Text & "'"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cq1 As String = "delete from PGStaffSalaryDetails where PGStaffSalaryCode=@DELETE1;"

            cmd = New SqlCommand(cq1)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@DELETE1", System.Data.SqlDbType.VarChar, 15, "PGStaffSalaryCode"))
            cmd.Parameters("@DELETE1").Value = Trim(TxtSalaryCode.Text)
            cmd.ExecuteNonQuery()
            con.Close()

            For i = 0 To ListView1.Items.Count - 1
                con = New SqlConnection(cs)
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Open()
                Dim cd As String = "insert Into PGStaffSalaryDetails(PGStaffSalaryCode,SlNo,SalaryDT1,SalaryDT2,EmpID,EmpName,Designation,BSalary,AGP,Total,DA,HRA,GSalary,IT,ProfTax,TotDeduction,NetSalary) VALUES (@PGStaffSalaryCode,@SlNo,@SalaryDT1,@SalaryDT2,@EmpID,@EmpName,@Designation,@BSalary,@AGP,@Total,@DA,@HRA,@GSalary,@IT,@ProfTax,@TotDeduction,@NetSalary)"
                cmd = New SqlCommand(cd)

                cmd.Connection = con
                cmd.Parameters.AddWithValue("PGStaffSalaryCode", TxtSalaryCode.Text)
                cmd.Parameters.AddWithValue("SlNo", ListView1.Items(i).SubItems(1).Text)
                cmd.Parameters.AddWithValue("SalaryDT1", dtpSalaryDate1.Text)
                cmd.Parameters.AddWithValue("SalaryDT2", dtpSalaryDate2.Text)
                cmd.Parameters.AddWithValue("EmpID", ListView1.Items(i).SubItems(2).Text)
                cmd.Parameters.AddWithValue("EmpName", ListView1.Items(i).SubItems(3).Text)
                cmd.Parameters.AddWithValue("Designation", ListView1.Items(i).SubItems(4).Text)
                cmd.Parameters.AddWithValue("BSalary", ListView1.Items(i).SubItems(5).Text)
                cmd.Parameters.AddWithValue("AGP", ListView1.Items(i).SubItems(6).Text)
                cmd.Parameters.AddWithValue("Total", ListView1.Items(i).SubItems(7).Text)
                cmd.Parameters.AddWithValue("DA", ListView1.Items(i).SubItems(8).Text)
                cmd.Parameters.AddWithValue("HRA", ListView1.Items(i).SubItems(9).Text)
                cmd.Parameters.AddWithValue("GSalary", ListView1.Items(i).SubItems(10).Text)
                cmd.Parameters.AddWithValue("IT", ListView1.Items(i).SubItems(11).Text)
                cmd.Parameters.AddWithValue("ProfTax", ListView1.Items(i).SubItems(12).Text)
                cmd.Parameters.AddWithValue("TotDeduction", ListView1.Items(i).SubItems(13).Text)
                cmd.Parameters.AddWithValue("NetSalary", ListView1.Items(i).SubItems(14).Text)

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

            Dim cq1 As String = "delete from PGStaffSalaryDetails where PGStaffSalaryCode=@DELETE1;"

            cmd = New SqlCommand(cq1)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@DELETE1", System.Data.SqlDbType.VarChar, 15, "PGStaffSalaryCode"))
            cmd.Parameters("@DELETE1").Value = Trim(TxtSalaryCode.Text)
            cmd.ExecuteNonQuery()
            con.Close()

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()

            Dim cq As String = "delete from PGStaffSalary where PGStaffSalaryCode=@DELETE2;"

            cmd = New SqlCommand(cq)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@DELETE2", System.Data.SqlDbType.VarChar, 15, "PGStaffSalaryCode"))

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
End Class
