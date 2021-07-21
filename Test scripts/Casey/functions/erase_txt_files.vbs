'FUNCTION to erase old txt files from MY Documents made by scripts.
'FILE NAMES
    'interview-answers'
    'caf-answers-'

function erase_week_old_save_your_work_txt_files()
    ' Set wshshell = CreateObject("WScript.Shell")
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    'Needs to determine MyDocs directory before proceeding.
    Set wshshell = CreateObject("WScript.Shell")
    user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

    Set objFolder = objFSO.GetFolder(user_myDocs_folder)
    Set colFiles = objFolder.Files
    For Each objFile in colFiles
        delete_this_file = False
        this_file_name = objFile.Name
        this_file_type = objFile.Type
        this_file_created_date = objFile.DateCreated
        this_file_path = objFile.Path

        If InStr(this_file_name, "caf-answers-") <> 0 Then delete_this_file = True
        If InStr(this_file_name, "interview-answers-") <> 0 Then delete_this_file = True
        If this_file_type <> "Text Document" then delete_this_file = False
        If DateDiff("d", this_file_created_date, date) < 8 Then delete_this_file = False

        If delete_this_file = True Then objFSO.DeleteFile(this_file_path)
    Next
end function

erase_week_old_save_your_work_txt_files
