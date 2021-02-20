Imports vb = Microsoft.VisualBasic

Public Class Form1
    Dim i As Integer
    Dim b As Integer
    Dim str As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        str = "Salary Management System"
        b = Len(str) + 1
        i = 1
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Interval = 150
        Me.Label1.Text = vb.Left(str, i)
        i = i + 1
        If i = b + 1 Then
            Timer2.Enabled = True
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Me.Hide()
        FrmMain.Show()
    End Sub
End Class