
function excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
'--- This function opens a specific excel file.
'~~~~~ file_url: name of the file
'~~~~~ visable_status: set to either TRUE (visible) or FALSE (not-visible)
'~~~~~ alerts_status: set to either TRUE (show alerts) or FALSE (suppress alerts)
'~~~~~ ObjExcel: leave as 'objExcel'
'~~~~~ objWorkbook: leave as 'objWorkbook'
'===== Keywords: MAXIS, PRISM, MMIS, Excel
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url) 'Opens an excel file from a specific URL
	objExcel.DisplayAlerts = alerts_status
end function



Function file_selection_dialog()

'creates a Windows Script Host object
Set Fshell = CreateObject("WScript.Shell")

'creates a FileSystemObject that is used as part of the powershell script temporary file process
Set fso = CreateObject("Scripting.FileSystemObject")


' creates a long string of powershell commands that will be written to a temporary powershell script file and executed. The commands will open a file selection dialog box, allow the user to select a file, and write the path of the selected file to a temporary text file.
shellCmd = "powershell -command " & "Add-Type -AssemblyName System.Windows.Forms; " & _
        "$dlg = New-Object System.Windows.Forms.OpenFileDialog; " & _
           "$dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop'); " & _
           "$dlg.Filter = 'Excel files (*.xlsx)|*.xlsx'; " & _
           "$dlg.ShowDialog() | Out-Null; " & _
           "$dlg.FileName | Out-File -FilePath 'C:\Temp\test.txt' -Encoding utf8"
Fshell.Run shellCmd, 1, TRUE

'Had to put this in here because VB Script's default text encoding does not support UTF-8, which is required for the filepath to not have weird characters in the edit box. ADODB.Stream converts file to binary than back to text with UTF-8 encoding.
Set objStream = CreateObject("ADODB.Stream")
objStream.Type = 1 
objStream.Open
objStream.LoadFromFile "C:\Temp\test.txt"
ObjStream.Position = 0
ObjStream.Type = 2 ' adTypeText
ObjStream.Charset = "utf-8"


'creates a variable of the contents of the temporary text file, which is the path to the selected file
strFileContents = objStream.ReadText()

selected_file = strFileContents

'fso.DeleteFile("C:\Temp\test.txt")
end Function


'============= DIALOG BOX

EMConnect ""

Set file_selection_path = Nothing
Set selected_file = Nothing

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "CBO referral"
    ButtonGroup ButtonPressed
    PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 145, 90, 50, 15
    CancelButton 200, 90, 50, 15
    EditBox 15, 45, 180, 15, file_selection_path
    GroupBox 10, 5, 250, 80, "Using the SEND MANUAL REFERRAL script"
    Text 20, 20, 235, 20, "This script should be used when E & T provides you with a list of recipeints that are working with CBO's and a manual referral is needed. "
    Text 15, 65, 230, 15, "Select the Excel file that contains the CBO information by selecting the 'Browse' button, and finding the file."
EndDialog



'================================== dialog box logic

    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog

    	Dialog Dialog1
    	If ButtonPressed = select_a_file_button then call file_selection_dialog()
        file_selection_path = selected_file
        
      
            Dialog Dialog1
            If ButtonPressed = select_a_file_button then call file_selection_dialog()
           
    
Dim objXLApp, objXLWb
Set objXLApp = CreateObject("Excel.Application")
objXLApp.Visible = True
Set objXLWb = objXLApp.Workbooks.Open(file_selection_path)

    
