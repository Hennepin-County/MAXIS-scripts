Public Class helpMenu

    Private Sub helpMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

End Class