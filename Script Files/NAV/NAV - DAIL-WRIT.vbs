'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - DAIL-WRIT.vbs"
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
'<<<<<<DELETE OLD REDUNDANT FUNCTIONS BELOW
EMConnect ""

'SECTION 01
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check <> "MAXIS" then
  MsgBox "This script needs to be run from MAXIS."
  StopScript
End If

row = 1
col = 1

EMSearch "Case Nbr:", row, col
If row = 0 then
  second_row = 1
  second_col = 1
  EMSearch "Case Number:", second_row, second_col
  If second_row = 0 then
    MsgBox "A case number could not be found on this screen. The script will now stop."
    StopScript
  End If
  If second_row <> 0 then EMReadScreen case_number, 8, second_row, second_col + 13
End If
If row <> 0 then EMReadScreen case_number, 8, row, col + 10
case_number = replace(case_number, "_", "")
case_number = trim(case_number)


'SECTION 02
Do
  EMSendKey "<PF3>"
  EMWaitReady 1, 1
  EMReadScreen MAXIS_check, 5, 1, 39
  If MAXIS_check <> "MAXIS" then stopscript 'This will stop the script from acting if it passwords out.
  EMReadScreen SELF_check, 4, 2, 50
Loop until SELF_check = "SELF"

EMWriteScreen "DAIL", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen "WRIT", 21, 70
EMSendKey "<enter>"
EMWaitReady 1, 1

script_end_procedure("")






