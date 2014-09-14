
Imports System
Imports System.IO
Imports System.Collections



Public Class scripts_config_form

    Private Property fso As Object

    Private Property FSO_new_file_path As Object

    Private Property ObjFSO As Object

    Private Property objFile As Object

    Private Property function_file_lines As Object

    Private Property line_to_look_for_in_functions_file As String

    Private Property warning As MsgBoxResult

    Private Property oFSO As Object

    Const read_only = 1
    Const read_write = 2

    Private Property strLine As Object

    Private Property strText As Object

    Private Property name_of_file As Object

    Private Property list_of_files_array As String()

    Private Property text_file As String

    Private Property new_text_file As Object

    Private Property current_directory_path As Object

    Private Property agency_is_beta As Boolean

    Private Property local_copy_of_zip_file As Object

    Private Property temp_folder As Object

    Private Property GitHub_current_zip_archive As String

    Private Property objXMLHTTP As Object

    Private Property objADOStream As Object

    Private Property objShell As Object

    Private Property FilesInZip As Object

    'This is the function that actually modifies the files
    Function update_files(file_name)

        Dim text_file() As String = System.IO.File.ReadAllLines(file_name)
        Dim text_line As String


        For Each text_line In text_file
            text_line = Replace(text_line, old_file_path.Text, new_file_path.Text)
            If InStr(file_name, "FUNCTIONS FILE") <> 0 Then   'Shouldn't do this part for any scripts other than the functions file.
                If InStr(text_line, "worker_county_code = ") Then text_line = "worker_county_code = " & Chr(34) & "x1" & Strings.Left(county_selection.Text, 2) & Chr(34)
                If InStr(text_line, "EDMS_choice = ") Then text_line = "EDMS_choice = " & Chr(34) & EDMS_choice.Text & Chr(34)
                If InStr(text_line, "county_name = ") Then text_line = "county_name = " & Chr(34) & Strings.Replace(county_selection.Text, Strings.Left(county_selection.Text, 5), "") & Chr(34)
                If InStr(text_line, "county_address_line_01 = ") Then text_line = "county_address_line_01 = " & Chr(34) & county_address_line_01.Text & Chr(34)
                If InStr(text_line, "county_address_line_02 = ") Then text_line = "county_address_line_02 = " & Chr(34) & county_address_line_02.Text & Chr(34)
                If InStr(text_line, "case_noting_intake_dates = ") Then
                    If intake_dates_check.Checked = True Then
                        text_line = "case_noting_intake_dates = True"
                    Else
                        text_line = "case_noting_intake_dates = False"
                    End If
                End If
                If InStr(text_line, "move_verifs_needed = ") Then
                    If move_verifs_needed_check.Checked = True Then
                        text_line = "move_verifs_needed = True"
                    Else
                        text_line = "move_verifs_needed = False"
                    End If
                End If
            End If
            'INSERT COLLECTING STATS FIXES HERE WHEN ACCESS GOES LIVE
            new_text_file = new_text_file & text_line & Chr(10)
        Next

        new_text_file = Split(new_text_file, Chr(10))
        System.IO.File.WriteAllLines(file_name, new_text_file)
        new_text_file = Nothing
    End Function



    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.Click

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim frmAbout As New AboutBox
        frmAbout.ShowDialog(Me)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles county_selection.SelectedIndexChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles access_DB_check.CheckedChanged

    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles EDMS_check.CheckedChanged

    End Sub

    Private Sub EDMS_choice_TextClear(ByVal sender As Object, ByVal e As System.EventArgs) Handles EDMS_choice.Enter
        EDMS_choice.Text = ""
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles old_file_path.TextChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles new_file_path.TextChanged

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles run_configuration_button.Click
        'Warning if a county is not selected
        If county_selection.Text = "" Or county_address_line_01.Text = "" Or county_address_line_02.Text = "" Then
            MsgBox("You must select a county, and enter a complete county address.")
            Exit Sub
        End If



        'Warns user that they can back out
        warning = MsgBox("The following utility will modify all of the scripts in the current directory, replacing the DHS file path with " & _
        "the current file path. If you move your script directory, you'll have to use this tool again. Are you sure you want to do this?", 1)
        If warning = 2 Then Exit Sub

        Update_Files_Label.Visible = True
        Tab_Control_Main_Form.Enabled = False
        run_configuration_button.Enabled = False

        'Setting EDMS_choice as DHS eDocs if there is not a local EDMS.
        If EDMS_check.Checked = False Then EDMS_choice.Text = "DHS eDocs"

        'Grabbing each file
        list_of_files_array = Directory.GetFiles(current_directory_path)

        'Running the update_files sub on each VBS file
        For Each file_in_array In list_of_files_array
            If UCase(Strings.Right(file_in_array, 4)) = ".VBS" Then update_files(file_in_array)
        Next

        Update_Files_Label.Visible = False
        Tab_Control_Main_Form.Enabled = True
        run_configuration_button.Enabled = True

        'Success!
        Me.Hide()
        MsgBox("Success! All scripts modified to work in this directory.")
        Application.Exit()
    End Sub

    Private Sub FileOpen(Optional p1 As Object = Nothing, Optional file As Object = Nothing, Optional openMode As OpenMode = Nothing, Optional openAccess As OpenAccess = Nothing, Optional p5 As Object = Nothing, Optional p6 As Object = Nothing)
        Throw New NotImplementedException
    End Sub

    Private Sub CheckBox1_CheckedChanged_2(sender As Object, e As EventArgs) Handles intake_dates_check.CheckedChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Opening a FileSystemObject, and adding the current path to the new_file_path.text variable, as well as to current_directory_path for determining which directory is actually active.
        fso = CreateObject("Scripting.FileSystemObject")
        new_file_path.Text = fso.GetAbsolutePathName(".") & "\"
        current_directory_path = fso.GetAbsolutePathName(".") & "\"
        'Opening file read-only
        ObjFSO = CreateObject("Scripting.FileSystemObject")
        objFile = ObjFSO.OpenTextFile("FUNCTIONS FILE.vbs", read_only)

        'Reading each line, and modifying to replace the original file path with the new one 
        Do Until objFile.AtEndOfStream
            function_file_lines = objFile.ReadLine
            line_to_look_for_in_functions_file = "'Set fso_command = run_another_script_fso.OpenTextFile("
            If InStr(function_file_lines, line_to_look_for_in_functions_file) Then
                old_file_path.Text = Replace(Replace(Replace(function_file_lines, line_to_look_for_in_functions_file, ""), Chr(34), ""), "FUNCTIONS FILE.vbs)", "")
            End If
        Loop

        'Closing the read-only version
        objFile.Close()


    End Sub


    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles advanced_file_path_mods_tab.Click

    End Sub

    Private Sub CheckBox1_CheckedChanged_3(sender As Object, e As EventArgs) Handles move_verifs_needed_check.CheckedChanged

    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Update_Files_Label.Click

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        multiaddressform.ShowDialog(Me)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Make these two dynamic in actual usage, I just included as it already exists elsewhere in my VB app
        current_directory_path = "C:\zip testing again and again\"
        agency_is_beta = True                                           'If you don't care about beta/standard users (forks?), ditch this.
        '|-------------------------------------------------------------------|
        '|          GRAB A REPO FROM GITHUB WITH VBS FOR SOME REASON         |
        '|                              BY VKC                               |
        '|                           (C) VKC 2014                            |
        '|               REDISTRIBUTE AND REUSE AS DESIRED!!!!!!             |
        '|-------------------------------------------------------------------|

        'Variables for the thing we're doing
        local_copy_of_zip_file = current_directory_path & "temp\master.zip"
        temp_folder = current_directory_path & "temp"

        'If the agency is a beta agency, they'll have the beta version of the scripts instead of the main one.
        If agency_is_beta = True Then
            GitHub_current_zip_archive = "https://github.com/MN-DHS-BZS-County-Programmers/MAXIS-BZ-Scripts-County-Beta/archive/master.zip"
        Else
            GitHub_current_zip_archive = "https://github.com/MN-DHS-BZS-Official/MAXIS-BZ-Scripts/archive/master.zip"
        End If

        'First, we create a temp directory for all this madness.
        fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder("temp")

        '---------------------------------------------------------------------------------------
        'Now it downloads the zip file from Github. It was copied from https://gist.github.com/udawtr/2053179 on 09/13/2014, and modified for our purposes.

        'Creating a server object
        objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

        'Opening the file
        objXMLHTTP.open("GET", GitHub_current_zip_archive, False)
        objXMLHTTP.send()
        If objXMLHTTP.Status = 200 Then     'Guessing this means "found" but admittedly I'm not sure. -VKC, 09/13/2014
            objADOStream = CreateObject("ADODB.Stream")
            objADOStream.Open()
            objADOStream.Type = 1 'adTypeBinary
            objADOStream.Write(objXMLHTTP.ResponseBody)
            objADOStream.Position = 0    'Set the stream position to the start

            'Writing the file to the hard disk
            ObjFSO = CreateObject("Scripting.FileSystemObject")
            If ObjFSO.Fileexists(local_copy_of_zip_file) Then ObjFSO.DeleteFile(local_copy_of_zip_file)
            ObjFSO = Nothing
            objADOStream.SaveToFile(local_copy_of_zip_file)
            objADOStream.Close()
            objADOStream = Nothing
        End If

        objXMLHTTP = Nothing
        '------------------------------------------------------------------------------------


        'The following was inspired by code retrieved 09/13/2014 from "http://stackoverflow.com/questions/911053/how-to-unzip-a-file-in-vbscript-using-internal-windows-xp-options-in".
        'Modifications were made to map variables to the above.
        '------------------------------------------------------------------------------------
        'If the extraction location does not exist create it.
        fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(temp_folder) Then
            fso.CreateFolder(temp_folder)
        End If

        'Extract the contents of the zip file.
        objShell = CreateObject("Shell.Application")
        FilesInZip = objShell.NameSpace(local_copy_of_zip_file).items
        objShell.NameSpace(temp_folder).CopyHere(FilesInZip, 16)
        fso = Nothing
        objShell = Nothing

        '-------------------------------------------------------------------------------------
        'Now we get copy the script files directory to its rightful place in the current directory. We search for the script files directory dynamically, because some will use a different fork from GitHub.
        fso = CreateObject("Scripting.FileSystemObject")

        'Uses main_folder because there's typically only one folder in the zip file generated by GitHub
        For Each main_folder In fso.GetFolder(temp_folder).Subfolders
            For Each subfolder In fso.GetFolder(main_folder).Subfolders
                If InStr(subfolder, "Script Files") <> 0 Then fso.CopyFolder(subfolder, current_directory_path, True)
            Next
        Next

        'Now we delete the temp directory.
        fso.DeleteFolder(temp_folder)
    End Sub

End Class
