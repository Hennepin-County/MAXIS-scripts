'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - MOVE PRODUCTION SCREEN TO INQUIRY.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

EMConnect "A"
row = 1
col = 1
EMSearch "Function: ", row, col
If row = 0 then
  MsgBox "Function not found."
  StopScript
End if
EMReadScreen MAXIS_function, 4, row, col + 10
If MAXIS_function = "____" then
  MsgBox "Function not found."
  StopScript
End if

row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row = 0 then
  MsgBox "Case number not found."
  StopScript
End if
EMReadScreen case_number, 8, row, col + 10

row = 1
col = 1
EMSearch "Month: ", row, col
If row = 0 then
  MsgBox "Footer month not found."
  StopScript
End if
EMReadScreen footer_month, 2, row, col + 7
EMReadScreen footer_year, 2, row, col + 10

row = 1
col = 1
EMSearch "(", row, col
If row = 0 then
  MsgBox "Command not found."
  StopScript
End if
EMReadScreen MAXIS_command, 4, row, col + 1
If MAXIS_command = "NOTE" then MAXIS_function = "CASE"

EMConnect "B"
EMFocus

attn
EMReadScreen inquiry_check, 7, 7, 15
If inquiry_check <> "RUNNING" then 
  MsgBox "Inquiry not found. The script will now stop."
  StopScript
End if

EMWriteScreen "FMPI", 2, 15
transmit

back_to_self

EMWriteScreen MAXIS_function, 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen footer_month, 20, 43
EMWriteScreen footer_year, 20, 46
EMWriteScreen MAXIS_command, 21, 70
transmit

script_end_procedure("")






