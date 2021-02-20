Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class FrmPGStaffEmpDetailsReports
    Dim rdr As SqlDataReader = Nothing
    Dim dtable As DataTable
    Dim con As SqlConnection = Nothing
    Dim adp As SqlDataAdapter
    Dim ds As DataSet

    Dim cs As String = "Data Source=HASSAN-PC\SQLEXPRESS;Initial Catalog=Collage;User ID=sa;Password=ali123;"

    Sub fillDesignation()
        Try
            Dim CN As New SqlConnection(cs)
            If CN.State = ConnectionState.Open Then
                CN.Close()
            End If
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct  (Designation) FROM PGStaffEmpDetails", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmdDesignation.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmdDesignation.Items.Add(drow(0).ToString())

            Next
            CN.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Sub fillEmpCode()
        Try
            Dim CN As New SqlConnection(cs)
            If CN.State = ConnectionState.Open Then
                CN.Close()
            End If
            CN.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct  (EmpID) FROM PGStaffEmpDetails", CN)
            ds = New DataSet("ds")

            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbEmpID.Items.Clear()

            For Each drow As DataRow In dtable.Rows
                cmbEmpID.Items.Add(drow(0).ToString())

            Next
            CN.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub FrmPGStaffEmpDetailsReports_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Dispose()
            OleCn.Close()
            FrmMain.Show()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmPGStaffEmpDetailsReports_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            
            Call fillDesignation()
            Call fillEmpCode()
        Catch ex As Exception
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
            cryRpt.Load(My.Application.Info.DirectoryPath & "\Reports\Report1.rpt")

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

            QueryString = "SELECT EmpID, EmpName, Designation, BSalary, AGP, Address, City, MobileNo, EmailID FROM  PGStaffEmpDetails where EmpID like @Code"

            Dim Cmd As New SqlCommand(QueryString, OleCn)
            Cmd.Parameters.Add("@Code", SqlDbType.VarChar).Value = cmbEmpID.Text

            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(Cmd)
            Dim ds As DataSet = New DataSet()
            Adapter.Fill(ds, "PGStaffEmpDetails")
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
            cryRpt.Load(My.Application.Info.DirectoryPath & "\Reports\Report1.rpt")

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

            QueryString = "SELECT EmpID, EmpName, Designation, BSalary, AGP, Address, City, MobileNo, EmailID FROM  PGStaffEmpDetails where Designation like @Code"

            Dim Cmd As New SqlCommand(QueryString, OleCn)
            Cmd.Parameters.Add("@Code", SqlDbType.VarChar).Value = cmdDesignation.Text

            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(Cmd)
            Dim ds As DataSet = New DataSet()
            Adapter.Fill(ds, "PGStaffEmpDetails")
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
            cryRpt.Load(My.Application.Info.DirectoryPath & "\Reports\Report1.rpt")

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

            QueryString = "SELECT EmpID, EmpName, Designation, BSalary, AGP, Address, City, MobileNo, EmailID FROM  PGStaffEmpDetails"

            Dim Cmd As New SqlCommand(QueryString, OleCn)
         
            Dim Adapter As SqlDataAdapter = New SqlDataAdapter(Cmd)
            Dim ds As DataSet = New DataSet()
            Adapter.Fill(ds, "PGStaffEmpDetails")
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

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            Me.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TabControl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.Click
        Try
            CrystalReportViewer1.ReportSource = Nothing
            CrystalReportViewer2.ReportSource = Nothing
            CrystalReportViewer3.ReportSource = Nothing
            cmbEmpID.Text = ""
            cmdDesignation.Text = ""
        Catch ex As Exception
        End Try
    End Sub
End Class