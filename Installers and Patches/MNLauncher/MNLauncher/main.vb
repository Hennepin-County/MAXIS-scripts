Public Class main

    Public Property KeyCode As Keys

    Private Property WScript As Object

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        oldfilepathText.Text = FileSystem.CurDir
    End Sub

    Private Sub menuQuit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuQuit.Click

        Me.Close()

    End Sub

    Private Sub menuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuAbout.Click
        aboutMenu.Show()
    End Sub

    Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F1
                aboutMenu.Show()
            Case Keys.F2
                helpMenu.Show()
            Case Keys.Escape
                Me.Close()
        End Select
    End Sub

    Private Sub oldfilepathBrowse_Click(sender As Object, e As EventArgs) Handles oldfilepathBrowse.Click
        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            oldfilepathText.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub newfilepathBrowse_Click(sender As Object, e As EventArgs) Handles newfilepathBrowse.Click
        If (FolderBrowserDialog2.ShowDialog() = DialogResult.OK) Then
            newfilepathText.Text = Nothing
            newfilepathText.Text = FolderBrowserDialog2.SelectedPath
        End If
    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs) Handles FolderBrowserDialog1.HelpRequest

    End Sub

    Private Sub HelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem1.Click
        helpMenu.Show()
    End Sub

    Public Function copy_file(filename)

        My.Computer.FileSystem.CopyFile(oldfilepathText.Text & "\" & filename, newfilepathText.Text & "\" & filename, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

    End Function

    Public Function copy_dir()

        My.Computer.FileSystem.CopyDirectory(oldfilepathText.Text, newfilepathText.Text, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)

    End Function

    Public Function update_file(filename, find_text, replace_with)

        Dim objFSO, objFile, strText, strNewText

        Const ForReading = 1
        Const ForWriting = 2

        objFSO = CreateObject("Scripting.FileSystemObject")
        objFile = objFSO.OpenTextFile(newfilepathText.Text & "\" & filename, ForReading)

        strText = objFile.ReadAll
        objFile.Close()
        strNewText = Replace(strText, find_text, replace_with)

        objFile = objFSO.OpenTextFile(newfilepathText.Text & "\" & filename, ForWriting)
        objFile.WriteLine(strNewText)

        objFile.Close()

    End Function

    Private Sub runconfigBtn_Click(sender As Object, e As EventArgs) Handles runconfigBtn.Click

        Dim new_dir_root, new_dir_mnsure, new_dir_mnsure_pendingnotices, functions_file, mnsure_functions_file, new_dir_noticegenerator, new_mns_dir

        new_dir_root = newfilepathText.Text & "\"
        new_dir_mnsure = newfilepathText.Text & "\MNSure\"
        new_dir_mnsure_pendingnotices = newfilepathText.Text & "\MNSure\Pending Notices\"
        new_dir_noticegenerator = newfilepathText.Text & "\Notice Generator\"

        functions_file = newfilepathText.Text & "\FUNCTIONS FILE.vbs"
        mnsure_functions_file = newfilepathText.Text & "\MNSure\MNSURE FUNCTIONS FILE.vbs"
        new_mns_dir = newfilepathText.Text




        'Move MNSure directory to its new location
        Call copy_dir()

        'Update file: MNSURE - 2014 retro task.vbs
        Call update_file("MNSure\MNSURE - 2014 retro task.vbs", "FUNCTIONSFILE", functions_file)
        Call update_file("MNSure\MNSURE - 2014 retro task.vbs", "MNSUREFUNCFILE", mnsure_functions_file)
        Call update_file("MNSure\MNSURE - 2014 retro task.vbs", "RUN_TSK_NO_ACTV", new_dir_mnsure & "RUNTASK - 2014 retro with no active programs.vbs")
        Call update_file("MNSure\MNSURE - 2014 retro task.vbs", "RUN_TSK_ACTV", new_dir_mnsure & "RUNTASK - 2014 retro with active programs.vbs")

        'Update file: MNSURE - Pending Notice Generator.vbs
        Call update_file("MNSure\MNSURE - Pending Notice Generator.vbs", "FUNCTIONSFILE", functions_file)
        Call update_file("MNSure\MNSURE - Pending Notice Generator.vbs", "MNSUREFUNCFILE", mnsure_functions_file)
        Call update_file("MNSure\MNSURE - Pending Notice Generator.vbs", "ONE_NTC_GEN", new_dir_mnsure & "NOTICES - Single Notice Generator.vbs")
        Call update_file("MNSure\MNSURE - Pending Notice Generator.vbs", "MULTI_NTC_GEN", new_dir_mnsure & "NOTICES - Mass Notice Generator.vbs")

        'Update file: MNSURE FUNCTIONS FILE.vbs
        'Needs no updates

        'Update file: NOTE - 2014 Retro (standalone).vbs 
        Call update_file("MNSure\NOTE - 2014 Retro (standalone).vbs", "FUNCTIONSFILE", functions_file)
        Call update_file("MNSure\NOTE - 2014 Retro (standalone).vbs", "MNSUREFUNCFILE", mnsure_functions_file)

        'Update file: NOTE - 2014 Retro.vbs
        'Needs no updates

        'Update file: NOTE - MNsure 2013 Retro.vbs
        Call update_file("MNSure\NOTE - MNsure 2013 Retro.vbs", "FUNCTIONSFILE", functions_file)

        'Update file: NOTICES - Mass Notice Generator.vbs
        Call update_file("MNSure\NOTICES - Mass Notice Generator.vbs", "NOTICEGENERATORDIR", new_dir_noticegenerator)
        Call update_file("MNSure\NOTICES - Mass Notice Generator.vbs", "PNDNTCDIR", new_dir_mnsure_pendingnotices)

        'Update file: NOTICES - Single Notice Generator.vbs
        Call update_file("MNSure\NOTICES - Single Notice Generator.vbs", "FUNCTIONSFILE", functions_file)
        Call update_file("MNSure\NOTICES - Single Notice Generator.vbs", "MNSUREFUNCFILE", mnsure_functions_file)
        Call update_file("MNSure\NOTICES - Single Notice Generator.vbs", "NOTICEGENERATORDIR", new_dir_noticegenerator)
        Call update_file("MNSure\NOTICES - Single Notice Generator.vbs", "PNDNTCDIR", new_dir_mnsure_pendingnotices)

        'Update file: RUNTASK - 2014 retro with active programs.vbs
        Call update_file("MNSure\RUNTASK - 2014 retro with active programs.vbs", "ACTVPRGMCASENOTE", new_dir_mnsure)

        'Update file: RUNTASK - 2014 retro with no active programs.vbs
        Call update_file("MNSure\RUNTASK - 2014 retro with no active programs.vbs", "NOACTPRGMCASENOTE", new_dir_mnsure)

        'Update file: 
        'Call update_file("", "", "")




        MsgBox("You have sucessfully set up your MNSure BlueZone Scripts. Thank you for trying our script configuration utility.", 0, "Success!")
        Me.Close()

    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub
End Class
