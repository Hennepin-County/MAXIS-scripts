'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - REPT-ACTV - BOTTOM.vbs"
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

EMConnect ""

'It sends an enter to force the screen to refresh, in order to check for a password prompt.
transmit

'Now it checks to make sure MAXIS production (or training) is running on this screen. If both are running the script will stop.
EMReadScreen MAXIS_check, 5, 1, 39
If MAXIS_check = "MAXIS" and MAXIS_check = "AXIS " then script_end_procedure("MAXIS not found. Are you passworded out? Navigate to MAXIS and try again.")

'This Do...loop gets back to SELF
do
  PF3
  EMReadScreen SELF_check, 27, 2, 28
loop until SELF_check = "Select Function Menu (SELF)"

'Enter keys for REPT/ACTV and transmit
EMSendKey "<home>" & "rept" & "<eraseeof>" & "<newline>" & "<newline>" & "actv" & "<enter>"
EMWaitReady 0, 0

'Presses "PF8" until the last page is found
do
  PF8
  EMReadScreen test, 21, 24, 2
loop until test = "THIS IS THE LAST PAGE"

script_end_procedure("")






