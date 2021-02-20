Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.Text

Public Class FrmNonTeachStaffEmpDetails
    Dim rdr As SqlDataReader = Nothing
    Dim con As SqlConnection = Nothing
    Dim cmd As SqlCommand = Nothing
    Dim cs As String = "Data Source=HASSAN-PC\SQLEXPRESS;Initial Catalog=Collage;User ID=sa;Password=ali123;"

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
        FrmNonTeachEmpDetails.Show()
    End Sub

    Private Sub FrmNonTeachStaffEmpDetails_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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

    Sub clear()
        txtAddress.Text = ""
        txtCity.Text = ""
        txtEmail.Text = ""
        txtEmpName.Text = ""
        ComboBox1.Text = ""
        txtBSalary.Text = ""
        txtMobileNo.Text = ""
        txtNotes.Text = ""
        txtEmpID.Text = ""
        txtZipCode.Text = ""
        cmbState.Text = ""
    End Sub
    Private Sub auto()
        txtEmpID.Text = "Emp-" & GetUniqueKey(4)
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

    Private Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            Me.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub NewRecord_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NewRecord.Click
        Try
            Call clear()
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
            auto()
            txtEmpName.Focus()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Try
            Call clear()
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
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Save_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Save.Click
        If Len(Trim(txtEmpName.Text)) = 0 Then
            MessageBox.Show("Please enter employee name", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtEmpName.Focus()
            Exit Sub
        End If
        If Len(Trim(ComboBox1.Text)) = 0 Then
            MessageBox.Show("Please enter designation", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ComboBox1.Focus()
            Exit Sub
        End If
        If Len(Trim(txtBSalary.Text)) = 0 Then
            MessageBox.Show("Please enter basic salary ", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtBSalary.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCity.Text)) = 0 Then
            MessageBox.Show("Please enter city", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtCity.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbState.Text)) = 0 Then
            MessageBox.Show("Please select state", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cmbState.Focus()
            Exit Sub
        End If

        Try

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            Dim ct As String = "select EmpID from NonTechStaffEmpDetails where EmpID=@find"

            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.Add(New SqlParameter("@find", System.Data.SqlDbType.VarChar, 15, "EmpID"))
            cmd.Parameters("@find").Value = txtEmpID.Text
            rdr = cmd.ExecuteReader()

            If rdr.Read Then
                MessageBox.Show("Employee ID Already Exists", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
            Else

                con = New SqlConnection(cs)
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Open()

                Dim cb As String = "insert into NonTechStaffEmpDetails(EmpID,EmpName,Designation,BSalary,Address,City,State,zipcode,MobileNo,EmailID,Notes) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11)"
                cmd = New SqlCommand(cb)
                cmd.Connection = con

                cmd.Parameters.Add(New SqlParameter("@d1", System.Data.SqlDbType.VarChar, 15, "EmpID"))

                cmd.Parameters.Add(New SqlParameter("@d2", System.Data.SqlDbType.VarChar, 50, "EmpName"))

                cmd.Parameters.Add(New SqlParameter("@d3", System.Data.SqlDbType.VarChar, 50, "Designation"))

                cmd.Parameters.Add(New SqlParameter("@d4", System.Data.SqlDbType.VarChar, 10, "BSalary"))

                cmd.Parameters.Add(New SqlParameter("@d5", System.Data.SqlDbType.VarChar, 100, "Address"))

                cmd.Parameters.Add(New SqlParameter("@d6", System.Data.SqlDbType.VarChar, 50, "City"))

                cmd.Parameters.Add(New SqlParameter("@d7", System.Data.SqlDbType.VarChar, 50, "State"))

                cmd.Parameters.Add(New SqlParameter("@d8", System.Data.SqlDbType.VarChar, 15, "zipcode"))

                cmd.Parameters.Add(New SqlParameter("@d9", System.Data.SqlDbType.VarChar, 15, "MobileNo"))

                cmd.Parameters.Add(New SqlParameter("@d10", System.Data.SqlDbType.VarChar, 100, "EmailID"))

                cmd.Parameters.Add(New SqlParameter("@d11", System.Data.SqlDbType.VarChar, 250, "Notes"))

                cmd.Parameters("@d1").Value = txtEmpID.Text

                cmd.Parameters("@d2").Value = txtEmpName.Text

                cmd.Parameters("@d3").Value = ComboBox1.Text

                cmd.Parameters("@d4").Value = txtBSalary.Text

                cmd.Parameters("@d5").Value = txtAddress.Text

                cmd.Parameters("@d6").Value = txtCity.Text

                cmd.Parameters("@d7").Value = cmbState.Text

                cmd.Parameters("@d8").Value = txtZipCode.Text

                cmd.Parameters("@d9").Value = txtMobileNo.Text

                cmd.Parameters("@d10").Value = txtEmail.Text

                cmd.Parameters("@d11").Value = txtNotes.Text

                Dim temp As Integer = cmd.ExecuteNonQuery()
                If temp > 0 Then

                    MessageBox.Show("Successfully saved", "Employee Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    cmdCancel.PerformClick()
                    con.Close()

                Else
                    MessageBox.Show("Insertion failed", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Update_record_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Update_record.Click
        Try

            con = New SqlConnection(cs)
            If con.State = ConnectionState.Open Then
                con.Close()
            End If

            con.Open()

            Dim cb As String = "update NonTechStaffEmpDetails set Empname=@d2,Designation=@d3,BSalary=@d4,Address=@d5,City=@d6,State=@d7,zipcode=@d8,MobileNo=@d9,EmailID=@d10,Notes=@d11 where EmpID='" & txtEmpID.Text & "'"
            cmd = New SqlCommand(cb)
            cmd.Connection = con

            cmd.Parameters.Add(New SqlParameter("@d1", System.Data.SqlDbType.VarChar, 15, "EmpID"))

            cmd.Parameters.Add(New SqlParameter("@d2", System.Data.SqlDbType.VarChar, 50, "EmpName"))

            cmd.Parameters.Add(New SqlParameter("@d3", System.Data.SqlDbType.VarChar, 50, "Designation"))

            cmd.Parameters.Add(New SqlParameter("@d4", System.Data.SqlDbType.VarChar, 10, "BSalary"))

            cmd.Parameters.Add(New SqlParameter("@d5", System.Data.SqlDbType.VarChar, 100, "Address"))

            cmd.Parameters.Add(New SqlParameter("@d6", System.Data.SqlDbType.VarChar, 50, "City"))

            cmd.Parameters.Add(New SqlParameter("@d7", System.Data.SqlDbType.VarChar, 50, "State"))

            cmd.Parameters.Add(New SqlParameter("@d8", System.Data.SqlDbType.VarChar, 15, "zipcode"))

            cmd.Parameters.Add(New SqlParameter("@d9", System.Data.SqlDbType.VarChar, 15, "MobileNo"))

            cmd.Parameters.Add(New SqlParameter("@d10", System.Data.SqlDbType.VarChar, 100, "EmailID"))

            cmd.Parameters.Add(New SqlParameter("@d11", System.Data.SqlDbType.VarChar, 250, "Notes"))

            cmd.Parameters("@d1").Value = txtEmpID.Text

            cmd.Parameters("@d2").Value = txtEmpName.Text

            cmd.Parameters("@d3").Value = ComboBox1.Text

            cmd.Parameters("@d4").Value = txtBSalary.Text

            cmd.Parameters("@d5").Value = txtAddress.Text

            cmd.Parameters("@d6").Value = txtCity.Text

            cmd.Parameters("@d7").Value = cmbState.Text

            cmd.Parameters("@d8").Value = txtZipCode.Text

            cmd.Parameters("@d9").Value = txtMobileNo.Text

            cmd.Parameters("@d10").Value = txtEmail.Text

            cmd.Parameters("@d11").Value = txtNotes.Text

            Dim temp As Integer = cmd.ExecuteNonQuery()
            If temp > 0 Then

                MessageBox.Show("Successfully updated", "Employee Details", MessageBoxButtons.OK, MessageBoxIcon.Information)
                con.Close()
            Else

                MessageBox.Show("Updation failed", "Attention", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

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

            Dim cq As String = "delete from NonTechStaffEmpDetails where EmpID=@DELETE1;"
            cmd = New SqlCommand(cq)
            cmd.Connection = con

            cmd.Parameters.Add(New SqlParameter("@DELETE1", System.Data.SqlDbType.VarChar, 15, "EmpID"))
            cmd.Parameters("@DELETE1").Value = Trim(txtEmpID.Text)
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then

                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                clear()
                txtEmpName.Focus()
                Update_record.Enabled = False
                Delete.Enabled = False
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)

                clear()
                txtEmpName.Focus()
                Update_record.Enabled = False
                Delete.Enabled = False
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Delete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Delete.Click
        Try

            If MessageBox.Show("Do you really want to delete the record?", "Vendor Record", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                delete_records()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class