Public Class FrmMain

    Private Sub PGStaffToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PGStaffToolStripMenuItem.Click
        EmployeeSalaryDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem4.Click
        EmployeeDetails.Show()
    End Sub

    Private Sub ToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem5.Click
        FrmNonTeachStaffEmpDetails.Show()
    End Sub

    Private Sub NonTeachingStaffToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonTeachingStaffToolStripMenuItem.Click
        FrmNonTeachStaffSalaryDetails.Show()
    End Sub

    Private Sub FrmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        End
    End Sub

    Private Sub PGStaffToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PGStaffToolStripMenuItem3.Click
        FrmEmpDetails3.Show()
    End Sub

    Private Sub NonTeachingStaffToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonTeachingStaffToolStripMenuItem2.Click
        FrmNonTeachEmpDetails3.Show()
    End Sub

    Private Sub PGStaffToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PGStaffToolStripMenuItem4.Click
        FrmPGStaffSalaryDetails2.Show()
    End Sub

    Private Sub NonTeachingStaffToolStripMenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonTeachingStaffToolStripMenuItem3.Click
        NonTeachStaffSalaryDetails2.Show()
    End Sub

    Private Sub PGStaffToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PGStaffToolStripMenuItem5.Click
        FrmPGStaffSalaryReports.Show()
    End Sub

    Private Sub NonTeachingStaffToolStripMenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonTeachingStaffToolStripMenuItem4.Click
        FrmNonTeachStaffSalaryReport.Show()
    End Sub

    Private Sub PGStaffToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PGStaffToolStripMenuItem6.Click
        FrmPGStaffEmpDetailsReports.Show()
    End Sub

    Private Sub NonTeachingStaffToolStripMenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NonTeachingStaffToolStripMenuItem5.Click
        FrmNonTeachEmpDetailsReports.Show()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lblTIme.Text = Format(Date.Now, "Long Time")
    End Sub

    Private Sub FrmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblDate.Text = Date.Now.ToString("dddd") & " - " & Date.Now.ToString("MMM dd, yyyy")
    End Sub

    Private Sub ToolStripButton1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        End
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        FrmAbout.Show()
    End Sub

    Private Sub ToolStripSeparator3_Click(sender As Object, e As EventArgs) Handles ToolStripSeparator3.Click

    End Sub
End Class