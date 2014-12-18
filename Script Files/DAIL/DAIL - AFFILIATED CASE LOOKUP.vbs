EMConnect ""

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - AFFILIATED CASE LOOKUP.vbs"
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

  row = 1
  col = 1
  cola = 1
EMSearch "#", 6, col
EMSearch ")", 6, cola

case_number_digits = cola - col - 1
EMReadScreen case_number, case_number_digits, 6, col + 1
If IsNumeric(case_number) = False then MsgBox "An affiliated case could not be detected on this dail message. Try another dail message."
If IsNumeric(case_number) = False then stopscript

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"


EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 18, 43
EMSendKey "<eraseEOF>" + case_number
EMSetCursor 21, 70
EMSendKey "note" + "<enter>"

MsgBox "You are now in case/note for the affiliated case!"

script_end_procedure("")






