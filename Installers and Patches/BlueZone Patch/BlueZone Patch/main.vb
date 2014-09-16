Public Class main

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles applypatch_Btn.Click

        Dim desktop = CreateObject("WScript.Shell").specialfolders("Desktop")

        Dim sourcepath As String = "M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\BlueZone Profiles\Bluezone Scripts.zmd"
        Dim DestPath As String = desktop & "\Bluezone Scripts.zmd"
        My.Computer.FileSystem.CopyFile(sourcepath, DestPath, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

        Dim desktop = CreateObject("WScript.Shell").specialfolders("Desktop")

        Dim sourcepath As String = "M:\Income-Maintence-Share\Bluezone Scripts\Script Files\MNSure\BlueZone Profiles\MNSure Worker.zmd"
        Dim DestPath As String = desktop & "\MNSure Worker.zmd"
        My.Computer.FileSystem.CopyFile(sourcepath, DestPath, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

    Private Sub QuitESCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitESCToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        helpMenu.Show()
    End Sub
End Class
