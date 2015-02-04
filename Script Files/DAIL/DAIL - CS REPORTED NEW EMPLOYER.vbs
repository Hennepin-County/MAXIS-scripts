'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - CS REPORTED NEW EMPLOYER.vbs"
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

'DIALOGS----------------------------------------------------
'This is a dialog asking if the job is known to the agency.
BeginDialog job_known_to_agency_dialog, 0, 0, 276, 65, "Job known?"
  CheckBox 5, 10, 160, 10, "Check here if this job is known to the agency.", job_known_check
  EditBox 90, 25, 180, 15, employer
  EditBox 70, 45, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 45, 50, 15
    CancelButton 220, 45, 50, 15
  Text 5, 30, 85, 10, "Job on DAIL is listed as:"
  Text 5, 50, 60, 10, "Worker signature:"
EndDialog

'Connects to BlueZone
EMConnect ""

'The script needs to determine what the day is in a MAXIS friendly format. The following does that.
current_month = datepart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month
current_day = datepart("d", date)
If len(current_day) = 1 then current_day = "0" & current_day
current_year = datepart("yyyy", date)
current_year = current_year - 2000

'Checks for a two line or one line case note
EMReadScreen second_line_check, 1, 6, 80	
If second_line_check = "+" then
	EMSendKey "x" 
	transmit
End if

'Grabbing ref nbr
row = 1
col = 1
EMSearch "REF NBR: ", row, col

'If not found, script will exit
if row = 0 then script_end_procedure("A member number could not be found on this case. Process manually. If there is a household member number somewhere on this message, contact your alpha user with the case number, and leave the message on your DAIL for the time being. Staff may want to look at this message for bugfixing.")

'Reading that HH member and employer, and cleaning up
EMReadScreen HH_memb, 2, row, col + 9
EMReadScreen employer, 8, row, col + 12
employer = rtrim(employer)

'If there had been a second line, this will look at that info
If second_line_check = "+" then
	EMReadScreen second_line, 60, row + 1, 5
	employer = employer & " " & rtrim(second_line)
	PF3
End if

'Navigating to case/curr
EMSendKey "h"
transmit

'Searching for active SNAP
row = 1
col = 1
EMSearch "FS: ", row, col
If row = 0 then 
	SNAP_active = False
Else
	SNAP_active = True
End if

'Navigating to STAT/JOBS for the HH_memb in question
EMWriteScreen "stat", 20, 22
EMWriteScreen "jobs", 20, 69
EMWriteScreen HH_memb, 20, 74
transmit

'Searching for abended cases. If it's abended, it'll transmit
row = 1
col = 1
EMSearch "abended", row, col
If row <> 0 then transmit

'Checking to make sure we're in STAT. If not, script will exit.
EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure("This case couldn't get to stat. MAXIS may have slowed down or be in background. Try again in a few seconds. If this continues to happen and MAXIS is up, send the case number to the script administrator.")

'Checks for the error screen, and if found, transmits
ERRR_screen_check

'Checking for the HH memb on the message. If not found, script will exit.
EMReadScreen HH_memb_check, 31, 24, 02
If HH_memb_check = "REFERENCE NUMBER IS INVALID    " then script_end_procedure("That member number is invalid. The script will now stop. Try again from DAIL. If this keeps happening, send the case number to the script administrator.")
If HH_memb_check = "MEMBER " & HH_memb & " IS NOT IN THE HOUSEHO" then script_end_procedure("That member number is invalid. The script will now stop. Try again from DAIL. If this keeps happening, send the case number to the script administrator.")
If HH_memb_check = "OCCURRENCE NUMBER IS INVALID   " then script_end_procedure("That member number is invalid. The script will now stop. Try again from DAIL. If this keeps happening, send the case number to the script administrator.")

'Show dialog, exit if cancel is pressed
Dialog job_known_to_agency_dialog
If ButtonPressed = cancel then stopscript

'If worker selects that the job is known, script will exit.
If job_known_check = checked then 
	MsgBox "The script will stop, this job is known."
	script_end_procedure("")
End if

'Cuts the string length down to the first 30 characters, so it will fit on the line.
employer = left(employer, 30) 

'Checks to make sure we're still on JOBS. If not (ie, worker navigated away), script exits
EMReadScreen jobs_check, 4, 2, 45
EMReadScreen jobs_memb_check, 2, 4, 33
If jobs_check <> "JOBS" or jobs_memb_check <> HH_memb then script_end_procedure("You appear to have navigated away from the JOBS panel for member " & HH_memb & ". The script will now stop. Try again from your DAIL. If this keeps happening, send the case number and a description of what happened to the script administrator.")

'Now it will create a new JOBS panel for this case.
EMWriteScreen "nn", 20, 79
transmit

'Default info (wage income, no verification)
EMWriteScreen "w", 5, 38
EMWriteScreen "n", 6, 38

'Adding employer info
EMWriteScreen employer, 7, 42

'Reading footer month/year, to be used in the prospective column
EMReadScreen footer_month, 2, 20, 55
EMReadScreen footer_year, 2, 20, 58

'Writing the first day of the footer month as the prospective paydate, and 0 for both wage and hours
EMWriteScreen footer_month, 12, 54
EMWriteScreen "01", 12, 57
EMWriteScreen footer_year, 12, 60
EMWriteScreen "0", 12, 67
EMWriteScreen "0", 18, 72

'Creates a PIC if case is on SNAP, puts pay freq as "monthly" and sets a zero in both anticipated income and hours/wk. It's a PIC with the minimum requirements.
If SNAP_active = True then 
	EMWriteScreen "x", 19, 38
	transmit
	EMWriteScreen current_month, 5, 34
	EMWriteScreen current_day, 5, 37
	EMWriteScreen current_year, 5, 40
	EMWriteScreen "1", 5, 64
	EMWriteScreen "0", 8, 64
	EMWriteScreen "0", 9, 66
	transmit
	transmit
	transmit
End if

'Gets out of case
transmit

'Presses PF3 until we're back on the DAIL
Do
	EMReadScreen DAIL_check, 4, 2, 48
	If DAIL_check = "DAIL" then exit do
	PF3
Loop until DAIL_check = "DAIL"

'Navigating to case note
EMSendKey "n"
transmit
PF9

'Sending case note
EMSendKey "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: " & HH_memb & " " & employer & "<newline>" 
EMSendKey "---" & "<newline>"
EMSendKey "* Job unreported to the agency. Sending employment verification. TIKLed for 10-day return." & "<newline>" 
EMSendKey "---" & "<newline>"
EMSendKey worker_signature & ", using automated script."
PF3
PF3

'Opening a blank TIKL
EMSendKey "w"
transmit

'The following will generate a TIKL formatted date for 10 days from now.
call create_MAXIS_friendly_date(date, 10, 5, 18)

'Setting cursor on the text area
EMSetCursor 9, 3

'Sending TIKL
EMSendKey "Verification of new employer (via CS message) should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)."
transmit
PF3

'Success box
MsgBox "MAXIS updated for new employer message, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at " & employer & "."

'End
script_end_procedure("")
