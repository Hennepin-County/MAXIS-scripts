'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - DISA MESSAGE.vbs"
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


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.


EMConnect ""

EMSendKey "s"
transmit

EMSendKey "disa"
transmit

'HH member dialog to select who's job this is.
BeginDialog HH_memb_dialog, 0, 0, 191, 52, "HH member"
  EditBox 50, 25, 25, 15, HH_memb
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 5, 10, 125, 15, "Which HH member is this for? (ex: 01)"
EndDialog
HH_memb = "01"
dialog HH_memb_dialog
If ButtonPressed = 0 then stopscript

EMWriteScreen HH_memb, 20, 76
transmit

EMReadScreen cash_disa_status, 1, 11, 69
If cash_disa_status <> "1" then
  MsgBox "This type of DISA status is not yet supported. It could be a SMRT or some other type of verif needed. Process manually at this time."
  stopscript
End if

PF4

PF9

EMSendKey "<home>" + "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" + "<newline>"
If cash_disa_status = 1 then EMSendKey "* Client needs a new Medical Opinion Form. Created using " & EDMS_choice & " and sent to client. TIKLed for 30-day return." & "<newline>"
EMSendKey "---" + "<newline>"

BeginDialog worker_sig_dialog, 0, 0, 191, 57, "Worker signature"
  EditBox 35, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 25, 10, 75, 10, "Sign your case note."
EndDialog

dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMSendKey worker_sig
PF3
PF3
PF3

EMSendKey "w"
transmit

'The following will generate a TIKL formatted date for 30 days from now.
TIKL_month = datepart("m", dateadd("d", 30, date))
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_day = datepart("d", dateadd("d", 30, date))
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_year = datepart("yyyy", dateadd("d", 30, date))
TIKL_year = TIKL_year - 2000

EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year
EMSetCursor 9, 3
EMSendKey "Medical Opinion Form sent 30 days ago. If not responded to, send another, and TIKL to close in 30 additional days."
transmit
PF3


MsgBox "Case note and TIKL made. Send a Medical Opinion Form using " & EDMS_choice & "."
script_end_procedure("")






