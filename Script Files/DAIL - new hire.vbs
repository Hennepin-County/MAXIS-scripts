'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "DAIL - new hire"
start_time = timer


'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-BZ-Scripts-County-Beta\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

EMConnect ""

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'The script needs to determine what the day is in a MAXIS friendly format. The following does that.
current_month = datepart("m", date)
If len(current_month) = 1 then current_month = "0" & current_month
current_day = datepart("d", date)
If len(current_day) = 1 then current_day = "0" & current_day
current_year = datepart("yyyy", date)
current_year = current_year - 2000


'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "x"
transmit
row = 1
col = 1
EMSearch "NEW JOB DETAILS", row, col 'Has to search, because every once in a while the rows and columns can slide one or two positions.
If row = 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the script administrator.")

EMReadScreen new_hire_first_line, 61, row, col 'Reads each line for the case note.
EMReadScreen new_hire_second_line, 61, row + 1, col
EMReadScreen new_hire_third_line, 61, row + 2, col
EMReadScreen new_hire_fourth_line, 61, row + 3, col
row = 1 'Now it's searching for info on the hire date as well as employer
col = 1
EMSearch "DATE HIRED:", row, col
EMReadScreen date_hired, 10, row, col + 12
If date_hired = "  -  -  EM" then date_hired = current_month & "-" & current_day & "-" & current_year
date_hired = CDate(date_hired)
month_hired = Datepart("m", date_hired)
If len(month_hired) = 1 then month_hired = "0" & month_hired
day_hired = Datepart("d", date_hired)
If len(day_hired) = 1 then day_hired = "0" & day_hired
year_hired = Datepart("yyyy", date_hired)
year_hired = year_hired - 2000
EMSearch "EMPLOYER:", row, col
EMReadScreen employer, 25, row, col + 10
row = 1 'Now it's searching for the SSN
col = 1
EMSearch "SSN #", row, col
EMReadScreen new_HIRE_SSN, 11, row, col + 5
PF3

'CHECKING CASE CURR. MFIP AND SNAP HAVE DIFFERENT RULES. 
EMWriteScreen "h", 6, 3
transmit
row = 1
col = 1
EMSearch "FS: ", row, col
If row <> 0 then FS_case = True
If row = 0 then FS_case = False
row = 1
col = 1
EMSearch "MFIP: ", row, col
If row <> 0 then MFIP_case = True
If row = 0 then MFIP_case = False
PF3


'GOING TO STAT
EMSendKey "s" 
transmit
EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")

'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
EMWriteScreen "memb", 20, 71
transmit
Do
  EMReadScreen MEMB_current, 1, 2, 73
  EMReadScreen MEMB_total, 1, 2, 78
  EMReadScreen MEMB_SSN, 11, 7, 42
  If new_HIRE_SSN = replace(MEMB_SSN, " ", "-") then
    EMReadScreen HH_memb, 2, 4, 33
    EMReadScreen memb_age, 2, 8, 76
    If cint(memb_age) < 19 then MsgBox "This client is under 19, so make sure to check that school verification is on file."
  End if
  transmit
Loop until (MEMB_current = MEMB_total) or (new_HIRE_SSN = replace(MEMB_SSN, " ", "-"))


'GOING TO JOBS
EMWriteScreen "jobs", 20, 71
EMWriteScreen HH_memb, 20, 76
transmit

'This is a dialog asking if the job is known to the agency.
BeginDialog new_HIRE_dialog, 0, 0, 291, 107, "New HIRE dialog"
  EditBox 80, 5, 25, 15, HH_memb
  DropListBox 120, 25, 40, 15, "Yes"+chr(9)+"No", job_known
  EditBox 95, 45, 185, 15, employer
  CheckBox 5, 70, 190, 10, "Check here to have the script make a new JOBS panel.", create_JOBS_check
  EditBox 70, 85, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 235, 65, 50, 15
    CancelButton 235, 85, 50, 15
    PushButton 175, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 175, 25, 45, 10, "next panel", next_panel_button
    PushButton 235, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 235, 25, 45, 10, "next memb", next_memb_button
  Text 5, 10, 70, 10, "HH member number:"
  GroupBox 170, 5, 115, 35, "STAT-based navigation"
  Text 5, 30, 110, 10, "Is this job known to the agency?:"
  Text 5, 50, 85, 10, "Job on DAIL is listed as:"
  Text 5, 90, 60, 10, "Worker signature:"
EndDialog



If MFIP_case = False then create_JOBS_check = 1 'MFIP cases need to manually add the JOBS panel for ES purposes.

HH_memb_row = 5 'Setting the variable

Do
  Dialog new_HIRE_dialog
  If ButtonPressed = 0 then stopscript
  EMReadScreen STAT_check, 4, 20, 21
  If STAT_check = "STAT" then
    If ButtonPressed = prev_panel_button then call panel_navigation_prev
    If ButtonPressed = next_panel_button then call panel_navigation_next
    If ButtonPressed = prev_memb_button then call memb_navigation_prev
    If ButtonPressed = next_memb_button then call memb_navigation_next
  End if
  transmit
Loop until ButtonPressed = -1

If job_known = "Yes" then script_end_procedure("The script will stop as this job is known.")

'Now it will create a new JOBS panel for this case.
If create_JOBS_check = 1 then
  EMSetCursor 20, 79
  EMSendKey "nn"
  transmit
  EMWriteScreen "w", 5, 38
  EMWriteScreen "n", 6, 38
  EMWriteScreen employer, 7, 42
  EMWriteScreen month_hired, 9, 35
  EMWriteScreen day_hired, 9, 38
  EMWriteScreen year_hired, 9, 41
  EMReadScreen footer_month, 2, 20, 55
  EMReadScreen footer_year, 2, 20, 58
  EMWriteScreen footer_month, 12, 54
  EMWriteScreen current_day, 12, 57
  EMWriteScreen footer_year, 12, 60
  EMWriteScreen "0", 12, 67
  EMWriteScreen "0", 18, 72
  If FS_case = "True" then 
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
End if
Do
  EMReadScreen DAIL_check, 4, 2, 48
  If DAIL_check = "DAIL" then exit do
  PF3
Loop until DAIL_check = "DAIL"
'Now we're back to the dail, and the script will case note what's happened.
EMSendKey "n"
transmit
PF9
transmit


EMSendKey replace(new_hire_first_line, new_HIRE_SSN, "XXX-XX-XXXX" & right(new_HIRE_SSN, 0)) & "<newline>" & new_hire_second_line & "<newline>" & new_hire_third_line + "<newline>" & new_hire_fourth_line & "<newline>" & "---" & "<newline>"
EMSendKey "* Job unreported to the agency. Sending employment verification. TIKLed for 10-day return." & "<newline>" & "---" & "<newline>"
EMSendKey worker_signature & ", using automated script."
PF3
PF3

'Now it will TIKL out for this case.
EMSendKey "w"
transmit

'The following will generate a TIKL formatted date for 10 days from now.
TIKL_month = datepart("m", dateadd("d", 10, date))
If len(TIKL_month) = 1 then TIKL_month = "0" & TIKL_month
TIKL_day = datepart("d", dateadd("d", 10, date))
If len(TIKL_day) = 1 then TIKL_day = "0" & TIKL_day
TIKL_year = datepart("yyyy", dateadd("d", 10, date))
TIKL_year = TIKL_year - 2000

EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year
EMSetCursor 9, 3
EMSendKey "Verification of NEW HIRE should have returned by now. If not received and process, take appropriate action. (TIKL auto-generated from script)." + "<enter>"
PF3
MsgBox "MAXIS updated for new HIRE message, a case note made, and a TIKL has been sent for 10 days from now. An EV should now be sent. The job is at " & employer & "."


script_end_procedure("")






