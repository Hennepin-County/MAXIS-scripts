




'============= Functions

' Source - https://stackoverflow.com/a/79105914
' Posted by Rno, modified by community. See post 'Timeline' for change history
' Retrieved 2026-07-28, License - CC BY-SA 4.0

' Getting an OpenFileDialog in VBScript is not (or rather, no longer) possible directly.
' A well-known workaround is using HTA, but that is very old and very slow.
' The approach used here relies on the slightly more modern Powershell (though not the Core version):
' I write a powershell script on the fly, execute that and have it write the results to file.
' I then read the result from that file.
' This is also not very fast and decidedly clunky, but I prefer it over HTA


Function ChooseFiles(ByVal initialDir)

  Set Fshell = CreateObject("WScript.Shell")
  Set fso = CreateObject("Scripting.FileSystemObject")
  tempFile = Fshell.ExpandEnvironmentStrings("%TEMP%") & fso.GetTempName
  ' temporary powershell script file to be invoked
  powershellFile = tempFile & ".ps1"
  ' temporary file to store standard output from command
  powershellOutputFile = tempFile & ".txt"

  ' Powershell code
  psScript = psScript & "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null" & vbCRLF
  psScript = psScript & "$dlg = New-Object System.Windows.Forms.OpenFileDialog" & vbCRLF
  psScript = psScript & "$dlg.initialDirectory = """ &initialDir & """" & vbCRLF
  psScript = psScript & "$dlg.filter = 'ZIP files|*.zip|Text Documents|*.txt|Shell Scripts|*.*sh|All Files|*.*'" & vbCRLF
  ' filter index 4 would show all files by default
  ' filter index 1 would show zip files by default
  psScript = psScript & "$dlg.FilterIndex = 4" & vbCRLF
  ' allow selecting multiple files
  psScript = psScript & "$dlg.Multiselect = $True" & vbCRLF
  psScript = psScript & "$dlg.Title = ""Select files""" & vbCRLF
  psScript = psScript & "$dlg.ShowHelp = $True" & vbCRLF
  psScript = psScript & "$dlg.ShowDialog() | Out-Null" & vbCRLF
  psScript = psScript & "Set-Content """ &powershellOutputFile & """ $dlg.FileNames" & vbCRLF
  
  ' write the powersell code to a file
  Set textFile = fso.CreateTextFile(powershellFile, True)
  textFile.WriteLine(psScript)
  textFile.Close
  Set textFile = Nothing
  
  ' construct shell command
  Dim shellCmd
  ' potential privilege issue here, obviously
  shellCmd = "powershell -ExecutionPolicy unrestricted &'" & powershellFile & "'"
  ' objShell.Run (strCommand, [intWindowStyle], [bWaitOnReturn]) 
  ' 0 Hide the window and activate another window.
  ' bWaitOnReturn set to TRUE - indicating script should wait for the program 
  ' to finish executing before continuing to the next statement
  Fshell.Run shellCmd, 0, TRUE

  ' open file for reading, do not create if missing, using system default format
  Set textFile = fso.OpenTextFile(powershellOutputFile, 1, 0, -2)
  ' the important thing to know is that the outputfile now contains 
  ' the names of the selected files, one file per line
  ' How you want to process them is op to you, 
  ' in this example I will just return the file contents as a string
  ChooseFiles = "" ' return a default to prevent error if user canceled the dialog
  If Not textFile.AtEndOfStream Then ChooseFiles = textFile.ReadAll
  textFile.Close
  Set textFile = Nothing
  fso.DeleteFile(powershellFile)
  fso.DeleteFile(powershellOutputFile)
  Set fso = Nothing
  Set Fshell = Nothing

End Function






function file_selection_system_dialog(file_selected, file_extension_restriction)
'--- This function allows a user to select a file to be opened in a script
'~~~~~ file_selected: variable for the name of the file
'~~~~~ file_extension_restriction: restricts all other file type besides allowed file type. Example: ".csv" only allows a CSV file to be accessed.
'===== Keywords: MAXIS, MMIS, PRISM, file
	'Creates a Windows Script Host object
	Set wShell=CreateObject("WScript.Shell")

	'This loops until the right file extension is selected. If it isn't specified (= ""), it'll always exit here.
	Do
		'Creates an object which executes the "select a file" dialog, using a Microsoft HTML application (MSHTA.exe), and some handy-dandy HTML.
		Set oExec=wShell.Exec("mshta.exe ""about:<head><meta http-equiv='X-UA-Compatible' content='IE=9'></head><input type=file id=FILE ><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")

msgbox "any words for a message box" & vbCr & file_selected

		'Creates the file_selected variable from the exit
		file_selected = oExec.StdOut.ReadLine

		'If no file is selected the script will stop
		If file_selected = "" then stopscript

		'If the rightmost characters of the file selected don't match what was in the file_extension_restriction argument, it'll tell the user. Otherwise the loop (and function) ends.
		If right(file_selected, len(file_extension_restriction)) <> file_extension_restriction then MsgBox "You've entered an incorrect file type. The allowable file type is: " & file_extension_restriction & "."
	Loop until right(file_selected, len(file_extension_restriction)) = file_extension_restriction
end function


function cancel_without_confirmation()
'--- This function ends a script after a user presses cancel. There is no confirmation message box but the end message for statistical information that cancel was pressed.
'===== Keywords: MAXIS, PRISM, MMIS, cancel, script_end_procedure
	If ButtonPressed = 0 then
        script_end_procedure("~PT: user pressed cancel")
        'script_end_procedure text added for statistical purposes. If script was canceled prior to completion, the statistics will reflect this.
        'Left the If...End If in the tier in case we want more stats or error handling, or if we need specialty processing for workflows
    End if
end function




'============= DIALOG BOX

on error resume next

EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "CBO referral"
    ButtonGroup ButtonPressed
    PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 145, 90, 50, 15
    CancelButton 200, 90, 50, 15
    EditBox 15, 45, 180, 15, ChooseFiles
    GroupBox 10, 5, 250, 80, "Using the SEND MANUAL REFERRAL script"
    Text 20, 20, 235, 20, "This script should be used when E & T provides you with a list of recipeints that are working with CBO's and a manual referral is needed. "
    Text 15, 65, 230, 15, "Select the Excel file that contains the CBO information by selecting the 'Browse' button, and finding the file."
EndDialog



'================================== dialog box logic
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call ChooseFiles("C:\temp") 'C:\temp is an example directory
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in