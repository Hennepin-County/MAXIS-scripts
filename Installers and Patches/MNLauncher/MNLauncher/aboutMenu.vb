Public Class aboutMenu

    Private Sub okBtn_Click(sender As Object, e As EventArgs) Handles okBtn.Click
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub aboutMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

End Class