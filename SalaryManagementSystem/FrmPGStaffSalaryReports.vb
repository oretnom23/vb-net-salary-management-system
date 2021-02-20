Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class FrmPGStaffSalaryReports
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
            adp.SelectCommand = New SqlCommand("SELECT distinct  (EmpName) FROM PGStaffSalaryDetails", CN)
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
    Private Sub FrmPGStaffSalaryReports_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Dispose()
            OleCn.Close()
            FrmMain.Show()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmPGStaffSalaryReports_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
         Try
           
            Call fillSalaryCode()
            Call fillEmpName()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim cryRpt As New ReportDocument
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table

        Try
            With OleCn
                If .State <> ConnectionState.Open Then
                    .ConnectionString = StrConnection()
                    .Open()
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        Try
            Me.Cursor = Cursors.WaitCursor
            cryRpt.Load(My.Application.Info.DirectoryPath & "\Reports\Report3.rpt")

            With crConnectionInfo
                .ServerName = "HASSAN-PC\SQLEXPRESS"
                .DatabaseName = "Collage"
                .UserID = "sa"
                .Password = "ali123"
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next


            Dim QueryString As String
           
            QueryString = "SELECT SalaryDT1,SalaryDT2,EmpID,EmpName,Designation,BSalary,AGP,Total,DA, HRA, GSalary,IT,ProfTax,TotDeduction,NetSalary,SlNo FROM PGStaffSalaryDetails where SalaryDT1 like @Code and SalaryDT2 like @Code2"

            Dim Cmd As New SqlCommand(QueryString, OleCn)
            Cmd.Parameters.Add("@Code", SqlDbType.Date).Value = dtpSalaryDateFrom.Text
            Cmd.Parameters.Add("@Code2", SqlDbType.Date).Value = dtpSalaryDateTo.Text

            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(Cmd)
            Dim ds As DataSet = New DataSet()
            Adapter.Fill(ds, "PGStaffSalaryDetails")
            cryRpt.SetDataSource(ds)

            Me.CrystalReportViewer3.ReportSource = cryRpt
            Me.Cursor = Cursors.Default
            OleCn.Close()
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Report Error")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        Try
            CrystalReportViewer1.ReportSource = Nothing
            CrystalReportViewer2.ReportSource = Nothing
            CrystalReportViewer3.ReportSource = Nothing

            cmbSalaryCode.Text = ""
            cmdEmpName.Text = ""
            dtpSalaryDateFrom.Text = Today()
            dtpSalaryDateTo.Text = Today()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PButton1.Click

        Dim cryRpt As New ReportDocument
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table

        Try
            With OleCn
                If .State <> ConnectionState.Open Then
                    .ConnectionString = StrConnection()
                    .Open()
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        Try
            Me.Cursor = Cursors.WaitCursor
            cryRpt.Load(My.Application.Info.DirectoryPath & "\Reports\Report3.rpt")

            With crConnectionInfo
                .ServerName = "HASSAN-PC\SQLEXPRESS"
                .DatabaseName = "Collage"
                .UserID = "sa"
                .Password = "ali123"
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next


            Dim QueryString As String

            QueryString = "SELECT PGStaffSalaryCode,SalaryDT1,SalaryDT2,EmpID,EmpName,Designation,BSalary,AGP,Total,DA, HRA, GSalary,IT,ProfTax,TotDeduction,NetSalary,SlNo FROM PGStaffSalaryDetails where PGStaffSalaryCode like @Code"

            Dim Cmd As New SqlCommand(QueryString, OleCn)
            Cmd.Parameters.Add("@Code", SqlDbType.VarChar).Value = cmbSalaryCode.Text
            
            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(Cmd)
            Dim ds As DataSet = New DataSet()
            Adapter.Fill(ds, "PGStaffSalaryDetails")
            cryRpt.SetDataSource(ds)

            Me.CrystalReportViewer1.ReportSource = cryRpt
            Me.Cursor = Cursors.Default
            OleCn.Close()
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Report Error")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub PButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PButton2.Click
        Dim cryRpt As New ReportDocument
        Dim crtableLogoninfos As New TableLogOnInfos
        Dim crtableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim CrTables As Tables
        Dim CrTable As Table

        Try
            With OleCn
                If .State <> ConnectionState.Open Then
                    .ConnectionString = StrConnection()
                    .Open()
                End If
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

        Try
            Me.Cursor = Cursors.WaitCursor
            cryRpt.Load(My.Application.Info.DirectoryPath & "\Reports\Report3.rpt")

            With crConnectionInfo
                .ServerName = "HASSAN-PC\SQLEXPRESS"
                .DatabaseName = "Collage"
                .UserID = "sa"
                .Password = "ali123"
            End With

            CrTables = cryRpt.Database.Tables
            For Each CrTable In CrTables
                crtableLogoninfo = CrTable.LogOnInfo
                crtableLogoninfo.ConnectionInfo = crConnectionInfo
                CrTable.ApplyLogOnInfo(crtableLogoninfo)
            Next


            Dim QueryString As String

            QueryString = "SELECT EmpID,EmpName,Designation,BSalary,AGP,Total,DA, HRA, GSalary,IT,ProfTax,TotDeduction,NetSalary,SlNo FROM PGStaffSalaryDetails where EmpName like @Code"

            Dim Cmd As New SqlCommand(QueryString, OleCn)
            Cmd.Parameters.Add("@Code", SqlDbType.VarChar).Value = cmdEmpName.Text
          
            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(Cmd)
            Dim ds As DataSet = New DataSet()
            Adapter.Fill(ds, "PGStaffSalaryDetails")
            cryRpt.SetDataSource(ds)

            Me.CrystalReportViewer2.ReportSource = cryRpt
            Me.Cursor = Cursors.Default
            OleCn.Close()
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            MsgBox(ex.Message(), MsgBoxStyle.Critical, "Report Error")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            Me.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            Me.Close()
        Catch ex As Exception
        End Try
    End Sub
End Class