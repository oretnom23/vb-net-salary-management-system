Public Class FrmAbout

    Private Sub CmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdOK.Click
        Try
            Me.Close()
    
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmAbout_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Try
            Me.Dispose()

        Catch ex As Exception
        End Try
    End Sub
End Class