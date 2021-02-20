Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class FrmNonTeachEmpDetails

    Private Sub FrmNonTeachEmpDetails_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            OleCn.Close()
            Me.Dispose()
            FrmNonTeachStaffEmpDetails.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Me.LblRecord.Text = Me.BindingContext(Dst, "NonTechStaffEmpDetails").Position + 1 & " of " & Me.BindingContext(Dst, "NonTechStaffEmpDetails").Count
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        FrmNonTeachStaffEmpDetails.Delete.Enabled = True
        FrmNonTeachStaffEmpDetails.Update_record.Enabled = True
        FrmNonTeachStaffEmpDetails.Save.Enabled = False
        FrmNonTeachStaffEmpDetails.Save.Visible = False
        FrmNonTeachStaffEmpDetails.NewRecord.Enabled = True
        FrmNonTeachStaffEmpDetails.NewRecord.Visible = True
        FrmNonTeachStaffEmpDetails.cmdCancel.Enabled = False
        FrmNonTeachStaffEmpDetails.cmdCancel.Visible = False
        FrmNonTeachStaffEmpDetails.cmdClose.Enabled = True
        FrmNonTeachStaffEmpDetails.cmdClose.Visible = True
        Try
            Dim dr As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim dr1 As SqlDataReader
            Dim com As New SqlCommand
            Try
                With OleCn
                    If .State <> ConnectionState.Open Then
                        .ConnectionString = StrConnection()
                        .Open()
                    End If
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            com.CommandText = "select [EmpID],[Notes] from NonTechStaffEmpDetails where EmpID = @UName"

            ' UserName
            Dim UName As SqlParameter = New SqlParameter("@UName", SqlDbType.VarChar, 15)
            UName.Value = dr.Cells(0).Value.ToString()
            com.Parameters.Add(UName)
            com.Connection = OleCn

            dr1 = com.ExecuteReader
            If dr1.Read Then
                FrmNonTeachStaffEmpDetails.txtEmpID.Text = dr.Cells(0).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtEmpName.Text = dr.Cells(1).Value.ToString()
                FrmNonTeachStaffEmpDetails.ComboBox1.Text = dr.Cells(2).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtBSalary.Text = dr.Cells(3).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtAddress.Text = dr.Cells(4).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtCity.Text = dr.Cells(5).Value.ToString()
                FrmNonTeachStaffEmpDetails.cmbState.Text = dr.Cells(6).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtZipCode.Text = dr.Cells(7).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtMobileNo.Text = dr.Cells(8).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtEmail.Text = dr.Cells(9).Value.ToString()
                FrmNonTeachStaffEmpDetails.txtNotes.Text = dr1("Notes").ToString()
                FrmNonTeachStaffEmpDetails.Show()
                Me.Dispose()

            Else
                MessageBox.Show("Some errors occuar please try again!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
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
                        .SelectCommand.CommandText = "SELECT EmpID,EmpName,Designation,BSalary,Address,City,State,zipcode,MobileNo,EmailID FROM NonTechStaffEmpDetails order by EmpID"
                        .SelectCommand.Connection = OleCn

                        Dim Ole As New SqlCommandBuilder(myDA) 'For Delete... 

                        Dst.Clear()
                        .Fill(Dst, "NonTechStaffEmpDetails")
                        Call GridViewEmpDetails2(Me.DataGridView1)
                        Me.LblRecord.Text = Me.BindingContext(Dst, "NonTechStaffEmpDetails").Position + 1 & " of " & Me.BindingContext(Dst, "NonTechStaffEmpDetails").Count

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

    Private Sub FrmNonTeachEmpDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Button3.PerformClick()
        Catch ex As Exception
        End Try
    End Sub
End Class