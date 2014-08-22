'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - CS reported new employer"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'SECTION 01:
EMConnect ""

'The script needs to determine what the day is in a MAXIS friendly format. The following does that.
current_month = datepart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month
current_day = datepart("d", date)
If len(current_day) = 1 then current_day = "0" & current_day
current_year = datepart("yyyy", date)
current_year = current_year - 2000

EMReadScreen second_line_check, 1, 6, 80	'Checks for a two line or one line case note
If second_line_check = "+" then
	EMSendKey "x" 
	transmit
End if

row = 1
col = 1
EMSearch "REF NBR: ", row, col
if row = 0 then
  MsgBox "A member number could not be found on this case. Process manually. If there is a household member number somewhere on this message, contact the script administrator with the case number, and leave the message on your DAIL for the time being."
  StopScript
End if
EMReadScreen HH_memb, 2, row, col + 9
EMReadScreen employer, 8, row, col + 12
employer = rtrim(employer)
If second_line_check = "+" then
	EMReadScreen second_line, 60, row + 1, 5
	employer = employer & " " & rtrim(second_line)
	PF3
End if

'SECTION 02:
EMSendKey "h"
transmit

row = 1
col = 1

EMSearch "FS: ", row, col
If row = 0 then SNAP_active = False
If row <> 0 then SNAP_active = True

EMWriteScreen "stat", 20, 22
EMWriteScreen "jobs", 20, 69
EMWriteScreen HH_memb, 20, 74
transmit
row = 1
col = 1
EMSearch "abended", row, col
If row <> 0 then transmit

EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure("This case couldn't get to stat. MAXIS may have slowed down or be in background. Try again in a few seconds. If this continues to happen and MAXIS is up, send the case number to the script administrator.")

'SECTION 03
EMReadScreen HH_memb_check, 31, 24, 02
If HH_memb_check = "REFERENCE NUMBER IS INVALID    " then script_end_procedure("That member number is invalid. The script will now stop. Try again from DAIL. If this keeps happening, send the case number to the script administrator.")
If HH_memb_check = "MEMBER " & HH_memb & " IS NOT IN THE HOUSEHO" then script_end_procedure("That member number is invalid. The script will now stop. Try again from DAIL. If this keeps happening, send the case number to the script administrator.")
If HH_memb_check = "OCCURRENCE NUMBER IS INVALID   " then script_end_procedure("That member number is invalid. The script will now stop. Try again from DAIL. If this keeps happening, send the case number to the script administrator.")


'This is a dialog asking if the job is known to the agency.
BeginDialog job_known, 0, 0, 191, 76, "Job known?"
  Text 10, 5, 110, 10, "Is this job known to the agency?"
  DropListBox 35, 20, 40, 15, "No"+chr(9)+"Yes", job_known
  Text 10, 40, 85, 10, "Job on DAIL is listed as:"
  EditBox 5, 55, 180, 15, employer
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
EndDialog


Dialog job_known
If ButtonPressed = 0 then stopscript


If job_known = "Yes" then 
  MsgBox "The script will stop, this job is known."
  script_end_procedure("")
End if

employer = left(employer, 30) 'Cuts the string length down to the first 30 characters, so it will fit on the line.

EMReadScreen jobs_check, 4, 2, 45
EMReadScreen jobs_memb_check, 2, 4, 33
If jobs_check <> "JOBS" or jobs_memb_check <> HH_memb then script_end_procedure("You appear to have navigated away from the JOBS panel for member " & HH_memb & ". The script will now stop. Try again from your DAIL. If this keeps happening, send the case number and a description of what happened to the script administrator.")

'Now it will create a new JOBS panel for this case.
EMSetCursor 20, 79
EMSendKey "nn" 
transmit
EMWriteScreen "w", 5, 38
EMWriteScreen "n", 6, 38
EMWriteScreen employer, 7, 42
EMReadScreen footer_month, 2, 20, 55
EMReadScreen footer_year, 2, 20, 58
EMWriteScreen footer_month, 9, 35
EMWriteScreen "01", 9, 38
EMWriteScreen footer_year, 9, 41
EMWriteScreen footer_month, 12, 54
EMWriteScreen "01", 12, 57
EMWriteScreen footer_year, 12, 60
EMWriteScreen "0", 12, 67
EMWriteScreen "0", 18, 72
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
transmit
Do
  EMReadScreen DAIL_check, 4, 2, 48
  If DAIL_check = "DAIL" then exit do
  PF3
Loop until DAIL_check = "DAIL"

'SECTION 04
'Now we're back to the dail, and the script will case note what's happened.
EMSendKey "n"
transmit
PF9
transmit
EMSendKey "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR: " & HH_memb & " " & employer & "<newline>" 
EMSendKey "---" & "<newline>"
EMSendKey "* Job unreported to the agency. Sending employment verification. TIKLed for 10-day return." & "<newline>" 
EMSendKey "---" & "<newline>"

BeginDialog worker_sig_dialog, 0, 0, 191, 57, "Worker signature"
  EditBox 35, 35, 50, 15, worker_sig
  ButtonGroup ButtonPressed
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 10, 5, 105, 30, "Sign your case note. The script will now TIKL out for 10-day return of the requested proofs."
EndDialog

Do
  dialog worker_sig_dialog
  If ButtonPressed = 0 then stopscript
  If worker_sig = "" then MsgBox "You must sign your case note!!"
Loop until worker_sig <> ""

EMSendKey worker_sig + ", using automated script."
PF3
PF3

'SECTION 05
'Now it will TIKL out for this case.

EMSendKey "w"
transmit

'The following will generate a TIKL formatted date for 10 days from now.

If DatePart("d", Now + 10) = 1 then TIKL_day = "01"
If DatePart("d", Now + 10) = 2 then TIKL_day = "02"
If DatePart("d", Now + 10) = 3 then TIKL_day = "03"
If DatePart("d", Now + 10) = 4 then TIKL_day = "04"
If DatePart("d", Now + 10) = 5 then TIKL_day = "05"
If DatePart("d", Now + 10) = 6 then TIKL_day = "06"
If DatePart("d", Now + 10) = 7 then TIKL_day = "07"
If DatePart("d", Now + 10) = 8 then TIKL_day = "08"
If DatePart("d", Now + 10) = 9 then TIKL_day = "09"
If DatePart("d", Now + 10) > 9 then TIKL_day = DatePart("d", Now + 10)

If DatePart("m", Now + 10) = 1 then TIKL_month = "01"
If DatePart("m", Now + 10) = 2 then TIKL_month = "02"
If DatePart("m", Now + 10) = 3 then TIKL_month = "03"
If DatePart("m", Now + 10) = 4 then TIKL_month = "04"
If DatePart("m", Now + 10) = 5 then TIKL_month = "05"
If DatePart("m", Now + 10) = 6 then TIKL_month = "06"
If DatePart("m", Now + 10) = 7 then TIKL_month = "07"
If DatePart("m", Now + 10) = 8 then TIKL_month = "08"
If DatePart("m", Now + 10) = 9 then TIKL_month = "09"
If DatePart("m", Now + 10) > 9 then TIKL_month = DatePart("m", Now + 10)

TIKL_year = DatePart("yyyy", Now + 10)

EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year - 2000
EMSendKey "Verification of new employer (via CS message) should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)."
transmit
PF3
MsgBox "MAXIS updated for new employer message, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at " & employer & "."

script_end_procedure("")






