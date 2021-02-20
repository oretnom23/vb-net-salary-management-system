Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class FrmPGStaffSalaryDetails2

    Dim rdr As SqlDataReader = Nothing
    Dim dtable As DataTable
    Dim con As SqlConnection = Nothing
    Dim adp As SqlDataAdapter
    Dim ds As DataSet

    Dim cs As String = "Data Source=HASSAN-PC\SQLEXPRESS;Initial Catalog=Collage;User ID=sa;Password=ali123;"

    Sub fillEmpName()
        Try
            Dim CN As New SqlConnection(cs)
            If CN.State = ConnectionState.Open Then
                CN.Close()
            End If
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct  (EmpID) FROM PGStaffSalaryDetails", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmdEmpName.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmdEmpName.Items.Add(drow(0).ToString())

            Next
            CN.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Sub fillSalaryCode()
        Try
            Dim CN As New SqlConnection(cs)
            If CN.State = ConnectionState.Open Then
                CN.Close()
            End If
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct  (PGStaffSalaryCode) FROM PGStaffSalaryDetails", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSalaryCode.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbSalaryCode.Items.Add(drow(0).ToString())

            Next
            CN.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FrmPGStaffSalaryDetails2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Dispose()
            OleCn.Close()
            FrmMain.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FrmPGStaffSalaryDetails2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Call fillSalaryCode()
            Call fillEmpName()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdEmpName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmpName.SelectedIndexChanged
        DataGridView3.DataSource = Nothing
        DataGridView3.Columns.Clear()

        If (DataGridView3.Rows.Count > 0) Then
            DataGridView3.Rows(0).Selected = True
        End If


        Try
            With OleCn
                If .State <> ConnectionState.Open Then
                    .ConnectionString = StrConnection()
                    .Open()
                    With myDA

                        .SelectCommand = New SqlCommand()
                        .SelectCommand.CommandText = "SELECT PGStaffSalaryCode,SlNo,EmpID,EmpName,Designation,BSalary,AGP,Total,DA,HRA,GSalary,IT,ProfTax,TotDeduction,NetSalary FROM PGStaffSalaryDetails where EmpID = '" & cmdEmpName.Text & "' order by EmpID"
                        .SelectCommand.Connection = OleCn

                        Dim Ole As New SqlCommandBuilder(myDA)

                        Dst.Clear()
                        .Fill(Dst, "PGStaffSalaryDetails")
                        Call GridViewPGStaffSalaryDetails(Me.DataGridView3)
                        Me.LblRecord2.Text = Me.BindingContext(Dst, "PGStaffSalaryDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffSalaryDetails").Count
                    End With
                End If
            End With




            Dim BSalary As Int64 = 0
            Dim AGP As Int64 = 0
            Dim Total As Int64 = 0
            Dim DA As Int64 = 0
            Dim HRA As Int64 = 0
            Dim GSalary As Int64 = 0
            Dim IT As Int64 = 0
            Dim ProfTax As Int64 = 0
            Dim TotDeduction As Int64 = 0
            Dim NetSalary As Int64 = 0


            For Each r As DataGridViewRow In Me.DataGridView3.Rows
                BSalary = BSalary + r.Cells(5).Value
                AGP = AGP + r.Cells(6).Value
                Total = Total + r.Cells(7).Value
                DA = DA + r.Cells(8).Value
                HRA = HRA + r.Cells(9).Value
                GSalary = GSalary + r.Cells(10).Value
                IT = IT + r.Cells(11).Value
                ProfTax = ProfTax + r.Cells(12).Value
                TotDeduction = TotDeduction + r.Cells(13).Value
                NetSalary = NetSalary + r.Cells(14).Value
            Next

            GroupBox3.Visible = True
            txtBSalary.Text = BSalary
            txtAGP.Text = AGP
            txtTotal.Text = Total
            txtDA.Text = DA
            txtHRA.Text = HRA
            txtGSalary.Text = GSalary
            txtIT.Text = IT
            txtProfTax.Text = ProfTax
            txtTotalDeduction.Text = TotDeduction
            txtNetSalary.Text = NetSalary

        Catch ex As Exception
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Error")
        Finally
            OleCn.Close()
        End Try
    End Sub

    Private Sub cmbSalaryCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSalaryCode.SelectedIndexChanged
        DataGridView2.DataSource = Nothing
        DataGridView2.Columns.Clear()

        If (DataGridView2.Rows.Count > 0) Then
            DataGridView2.Rows(0).Selected = True
        End If

        Try
            With OleCn
                If .State <> ConnectionState.Open Then
                    .ConnectionString = StrConnection()
                    .Open()
                    With myDA

                        .SelectCommand = New SqlCommand()
                        .SelectCommand.CommandText = "SELECT PGStaffSalaryCode,SlNo,EmpID,EmpName,Designation,BSalary,AGP,Total,DA,HRA,GSalary,IT,ProfTax,TotDeduction,NetSalary FROM PGStaffSalaryDetails where PGStaffSalaryCode = '" & cmbSalaryCode.Text & "' order by PGStaffSalaryCode"
                        .SelectCommand.Connection = OleCn

                        Dim Ole As New SqlCommandBuilder(myDA)

                        Dst.Clear()
                        .Fill(Dst, "PGStaffSalaryDetails")
                        Call GridViewPGStaffSalaryDetails(Me.DataGridView2)
                        Me.LblRecord1.Text = Me.BindingContext(Dst, "PGStaffSalaryDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffSalaryDetails").Count
                    End With
                End If
            End With

            Dim BSalary As Int64 = 0
            Dim AGP As Int64 = 0
            Dim Total As Int64 = 0
            Dim DA As Int64 = 0
            Dim HRA As Int64 = 0
            Dim GSalary As Int64 = 0
            Dim IT As Int64 = 0
            Dim ProfTax As Int64 = 0
            Dim TotDeduction As Int64 = 0
            Dim NetSalary As Int64 = 0


            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                BSalary = BSalary + r.Cells(5).Value
                AGP = AGP + r.Cells(6).Value
                Total = Total + r.Cells(7).Value
                DA = DA + r.Cells(8).Value
                HRA = HRA + r.Cells(9).Value
                GSalary = GSalary + r.Cells(10).Value
                IT = IT + r.Cells(11).Value
                ProfTax = ProfTax + r.Cells(12).Value
                TotDeduction = TotDeduction + r.Cells(13).Value
                NetSalary = NetSalary + r.Cells(14).Value
            Next

            GroupBox7.Visible = True
            txtBSalary1.Text = BSalary
            txtAGP1.Text = AGP
            txtTotal1.Text = Total
            txtDA1.Text = DA
            txtHRA1.Text = HRA
            txtGSalary1.Text = GSalary
            txtIT1.Text = IT
            txtProfTax1.Text = ProfTax
            txtTotalDeduction1.Text = TotDeduction
            txtNetSalary1.Text = NetSalary

        Catch ex As Exception
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Error")
        Finally
            OleCn.Close()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.Columns.Clear()

        If (DataGridView1.Rows.Count > 0) Then
            DataGridView1.Rows(0).Selected = True
        End If

        Try
            With OleCn
                If .State <> ConnectionState.Open Then
                    .ConnectionString = StrConnection()
                    .Open()
                    With myDA

                        .SelectCommand = New SqlCommand()
                        .SelectCommand.CommandText = "SELECT PGStaffSalaryCode,SlNo,EmpID,EmpName,Designation,BSalary,AGP,Total,DA,HRA,GSalary,IT,ProfTax,TotDeduction,NetSalary FROM PGStaffSalaryDetails where SalaryDT1 = '" & dtpSalaryDateFrom.Text & "' and SalaryDT2 = '" & dtpSalaryDateTo.Text & "'order by PGStaffSalaryCode"
                        .SelectCommand.Connection = OleCn
                        Dim Ole As New SqlCommandBuilder(myDA)

                        Dst.Clear()
                        .Fill(Dst, "PGStaffSalaryDetails")
                        Call GridViewPGStaffSalaryDetails(Me.DataGridView1)
                        Me.LblRecord3.Text = Me.BindingContext(Dst, "PGStaffSalaryDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffSalaryDetails").Count
                    End With
                End If
            End With

            Dim BSalary As Int64 = 0
            Dim AGP As Int64 = 0
            Dim Total As Int64 = 0
            Dim DA As Int64 = 0
            Dim HRA As Int64 = 0
            Dim GSalary As Int64 = 0
            Dim IT As Int64 = 0
            Dim ProfTax As Int64 = 0
            Dim TotDeduction As Int64 = 0
            Dim NetSalary As Int64 = 0


            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                BSalary = BSalary + r.Cells(5).Value
                AGP = AGP + r.Cells(6).Value
                Total = Total + r.Cells(7).Value
                DA = DA + r.Cells(8).Value
                HRA = HRA + r.Cells(9).Value
                GSalary = GSalary + r.Cells(10).Value
                IT = IT + r.Cells(11).Value
                ProfTax = ProfTax + r.Cells(12).Value
                TotDeduction = TotDeduction + r.Cells(13).Value
                NetSalary = NetSalary + r.Cells(14).Value
            Next

            GroupBox4.Visible = True
            txtBSalary3.Text = BSalary
            txtAGP3.Text = AGP
            txtTotal3.Text = Total
            txtDA3.Text = DA
            txtHRA3.Text = HRA
            txtGSalary3.Text = GSalary
            txtIT3.Text = IT
            txtProfTax3.Text = ProfTax
            txtTotalDeduction3.Text = TotDeduction
            txtNetSalary3.Text = NetSalary

        Catch ex As Exception
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Error")
        Finally
            OleCn.Close()
        End Try
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        Try
            cmbSalaryCode.Text = ""
            cmdEmpName.Text = ""
            dtpSalaryDateFrom.Text = Today
            dtpSalaryDateTo.Text = Today
            LblRecord1.Text = ""
            LblRecord2.Text = ""
            LblRecord3.Text = ""

            DataGridView1.DataSource = Nothing
            DataGridView1.Columns.Clear()
            DataGridView2.DataSource = Nothing
            DataGridView2.Columns.Clear()
            DataGridView3.DataSource = Nothing
            DataGridView3.Columns.Clear()

            GroupBox3.Visible = False
            GroupBox4.Visible = False
            GroupBox7.Visible = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Try
            GroupBox3.Visible = False
            DataGridView3.DataSource = Nothing
            DataGridView3.Columns.Clear()
            cmdEmpName.Text = ""
            LblRecord2.Text = ""
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Try
            GroupBox7.Visible = False
            DataGridView2.DataSource = Nothing
            DataGridView2.Columns.Clear()
            cmbSalaryCode.Text = ""
            LblRecord1.Text = ""
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            DataGridView1.DataSource = Nothing
            DataGridView1.Columns.Clear()
            dtpSalaryDateFrom.Text = Today
            dtpSalaryDateTo.Text = Today
            GroupBox4.Visible = False
            LblRecord3.Text = ""
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If DataGridView1.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application

        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView1.RowCount - 1
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If DataGridView3.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application

        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView3.RowCount - 1
            colsTotal = DataGridView3.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView3.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView3.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If DataGridView2.RowCount = Nothing Then
            MessageBox.Show("Sorry nothing to export into excel sheet.." & vbCrLf & "Please retrieve data in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application

        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = DataGridView2.RowCount - 1
            colsTotal = DataGridView2.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView2.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView2.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally

            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Me.LblRecord3.Text = Me.BindingContext(Dst, "PGStaffSalaryDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffSalaryDetails").Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Try
            Me.LblRecord2.Text = Me.BindingContext(Dst, "PGStaffSalaryDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffSalaryDetails").Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try
            Me.LblRecord1.Text = Me.BindingContext(Dst, "PGStaffSalaryDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffSalaryDetails").Count
        Catch ex As Exception
        End Try
    End Sub
End Class