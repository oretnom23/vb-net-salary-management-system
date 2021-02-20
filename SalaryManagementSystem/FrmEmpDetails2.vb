Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FrmEmpDetails2

    Private Sub FrmEmpDetails2_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            OleCn.Close()
            Me.Dispose()
            EmployeeSalaryDetails.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Me.LblRecord.Text = Me.BindingContext(Dst, "PGStaffEmpDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffEmpDetails").Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick

        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)

            Dim designation As String = dr.Cells(2).Value.ToString()
            Dim agp As Double = dr.Cells(4).Value.ToString()
            Dim bsalary As Double = dr.Cells(3).Value.ToString()
            Dim total As Double
            Dim da As Double
            Dim HRA As Double
            Dim Gsalary As Double
            Dim it As Double
            Dim profTax As Double
            Dim totalDeduction As Double
            Dim netSalary As Double


            EmployeeSalaryDetails.txtEmpID.Text = dr.Cells(0).Value.ToString()
            EmployeeSalaryDetails.txtEmpName.Text = dr.Cells(1).Value.ToString()
            EmployeeSalaryDetails.txtDesignation.Text = dr.Cells(2).Value.ToString()
            EmployeeSalaryDetails.txtBSalary.Text = dr.Cells(3).Value.ToString()
            EmployeeSalaryDetails.txtAGP.Text = dr.Cells(4).Value.ToString()


            '''''''''''''Caculation'''''''''''''''''''''

            If dr.Cells(4).Value.ToString() > 5000 Then
                total = bsalary + agp
            Else
                total = 0
            End If

            If dr.Cells(2).Value.ToString() = "Professor" Then
                da = total * 0.61
            ElseIf dr.Cells(2).Value.ToString() = "Associate  Prof." Then
                da = total * 0.01
            ElseIf dr.Cells(2).Value.ToString() = "Asst. Prof" Then
                da = total * 0.45
            End If

            HRA = total * 0.1

            If dr.Cells(4).Value.ToString() > 5000 Then
                Gsalary = total + da + HRA
            End If

            If Gsalary > 10000 Then
                profTax = 200
            Else
                profTax = 0
            End If


            If designation = "Professor" Then
                it = 7000
            ElseIf designation = "Associate  Prof." Then
                it = 3400
            ElseIf designation = "Asst. Prof" Then
                it = 1624
            End If


            totalDeduction = it + profTax

            netSalary = Gsalary - it - profTax

            ''''''''''''''''''''''''''''''''''


            EmployeeSalaryDetails.txtTotal.Text = total
            EmployeeSalaryDetails.txtDA.Text = da
            EmployeeSalaryDetails.txtHRA.Text = HRA
            EmployeeSalaryDetails.txtGrossSalary.Text = Gsalary
            EmployeeSalaryDetails.txtIT.Text = it
            EmployeeSalaryDetails.txtProfTax.Text = profTax
            EmployeeSalaryDetails.txtTotalDeduction.Text = totalDeduction
            EmployeeSalaryDetails.txtNetSalary.Text = netSalary

            EmployeeSalaryDetails.Show()
            Me.Dispose()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
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
                        .SelectCommand.CommandText = "SELECT EmpID,EmpName,Designation,BSalary,AGP,Address,City,State,zipcode,MobileNo,EmailID FROM PGStaffEmpDetails order by EmpID"
                        .SelectCommand.Connection = OleCn

                        Dim Ole As New SqlCommandBuilder(myDA) 'For Delete... 

                        Dst.Clear()
                        .Fill(Dst, "PGStaffEmpDetails")
                        Call GridViewEmpDetails1(Me.DataGridView1)
                        Me.LblRecord.Text = Me.BindingContext(Dst, "PGStaffEmpDetails").Position + 1 & " of " & Me.BindingContext(Dst, "PGStaffEmpDetails").Count

                    End With
                End If
            End With

        Catch ex As Exception
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Error")
        Finally
            OleCn.Close()
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Try
            DataGridView1.DataSource = Nothing
            DataGridView1.Columns.Clear()
            LblRecord.Text = ""
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmEmpDetails2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Button3.PerformClick()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
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
End Class