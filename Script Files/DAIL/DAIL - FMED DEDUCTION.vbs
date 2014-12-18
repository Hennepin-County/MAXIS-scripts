'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - FMED DEDUCTION.vbs"
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
'<<<<<GO THROUGH THE SCRIPT AND REMOVE REDUNDANT FUNCTIONS, THANKS TO CUSTOM FUNCTIONS THEY ARE NOT REQUIRED.

EMConnect ""

BeginDialog worker_sig_dialog, 0, 0, 141, 46, "Worker signature"
  EditBox 15, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 85, 5, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 75, 10, "Sign your case note."
EndDialog

Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMWriteScreen "p", 6, 3
EMSendKey "<enter>"
EMWaitReady 0, 0

EMWriteScreen "memo", 20, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF5>"
EMWaitReady 0, 0

EMWriteScreen "x", 5, 10
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "You are turning 60 next month, so you may be eligible for a new deduction for SNAP." + "<newline>" + "<newline>"
EMSendKey "Clients who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month." + "<newline>" + "<newline>"
EMSendKey "If you have medical bills over $35 each month, please contact your worker to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices." + "<newline>" + "<newline>"
EMSendKey "Please call your worker with questions."
EMSendKey "<PF4>"
EMWaitReady 0, 0

EMWriteScreen "case", 19, 22
EMWriteScreen "note", 19, 70
EMSendKey "<enter>"
EMWaitReady 0, 0

EMSendKey "<PF9>"
EMWaitReady 0, 0

EMSendKey "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey "* Sent MEMO to client about FMED deductions." + "<newline>"
EMSendKey "---" + "<newline>"
EMSendKey worker_sig + ", using automated script."

EMSendKey "<PF3>"
EMWaitReady 0, 0

EMSendKey "<PF3>"
EMWaitReady 0, 0

MsgBox "The script has sent a MEMO to the client about the possible FMED deduction, and case noted the action."

script_end_procedure("")






