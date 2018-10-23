'GATHERING STATS===========================================================================================
name_of_script = "NOTES - APPLICATION CHECK.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/14/2018", "Updated dialog and case note to address requested enhancements.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 131, 50, "Case number dialog"
  EditBox 65, 5, 60, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 20, 30, 50, 15
    CancelButton 75, 30, 50, 15
  Text 10, 10, 45, 10, "Case number:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & case number
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)

'initial case number dialog
Do
	DO
		err_msg = ""
	    dialog case_number_dialog
      	cancel_confirmation
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'information gathering to auto-populate the application date
'pending programs information
back_to_self
EMWriteScreen MAXIS_case_number, 18, 43
Call navigate_to_MAXIS_screen("REPT", "PND2")

'Ensuring that the user is in REPT/PND2
Do
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check <> "PND2" then
		back_to_SELF
		Call navigate_to_MAXIS_screen("REPT", "PND2")
	End if
LOOP until PND2_check = "PND2"

'checking the case to make sure there is a pending case.  If not script will end & inform the user no pending case exists in PND2
EMReadScreen not_pending_check, 5, 24, 2
If not_pending_check = "CASE " THEN script_end_procedure("There is not a pending program on this case, or case is not in PND2 status." & vbNewLine & vbNewLine & "Please make sure you have the right case number, and/or check your case notes to ensure that this application has been completed.")

'grabs row and col number that the cursor is at
EMGetCursor MAXIS_row, MAXIS_col
EMReadScreen app_month, 2, MAXIS_row, 38
EMReadScreen app_day, 2, MAXIS_row, 41
EMReadScreen app_year, 2, MAXIS_row, 44
EMReadScreen days_pending, 3, MAXIS_row, 50
EMReadScreen additional_application_check, 14, MAXIS_row + 1, 17
EMReadScreen add_app_month, 2, MAXIS_row + 1, 38
EMReadScreen add_app_day, 2, MAXIS_row + 1, 41
EMReadScreen add_app_year, 2, MAXIS_row + 1, 44

'Creating new variable for application check date and additional application date.
application_check_date = app_month & "/" & app_day & "/" & app_year
additional_application_date = add_app_month & "/" & add_app_day & "/" & add_app_year

'checking for multiple application dates.  Creates message boxes giving the user an option of which app date to choose
If additional_application_check = "ADDITIONAL APP" THEN multiple_apps = MsgBox("Do you want this application date: " & application_check_date, VbYesNoCancel)
If multiple_apps = vbCancel then stopscript
If multiple_apps = vbYes then additional_date_found = False
IF multiple_apps = vbNo then
	additional_apps = Msgbox("Do you want this application date: " & additional_application_date, VbYesNoCancel)
	If additional_apps = vbCancel then stopscript
	If additional_apps = vbNo then script_end_procedure("No more application dates exist. Please review the case, and start the script again if applicable.")
	If additional_apps = vbYes then
		additional_date_found = TRUE
		application_check_date = additional_application_date
		MAXIS_row = MAXIS_row + 1
	END IF
End if

EMReadScreen PEND_CASH_check,	1, MAXIS_row, 54
EMReadScreen PEND_SNAP_check, 1, MAXIS_row, 62
EMReadScreen PEND_HC_check, 1, MAXIS_row, 65
EMReadScreen PEND_EMER_check,	1, MAXIS_row, 68
EMReadScreen PEND_GRH_check, 1, MAXIS_row, 72

'this information auto-checks programs pending into main dialog if one app date is found
pending_progs = ""
IF PEND_CASH_check 	= "A" or PEND_CASH_check = "P" THEN
 	CASH_CHECKBOX = CHECKED
	pending_progs = pending_progs & "CASH" & ", "
END IF
IF PEND_SNAP_check 	= "A" or PEND_SNAP_check = "P" THEN
	FS_CHECKBOX = CHECKED
	pending_progs = pending_progs & "SNAP" & ", "
END IF
IF PEND_HC_check 	= "P" THEN
	HC_CHECKBOX = CHECKED
	pending_progs = pending_progs & "HC" & ", "
END IF
IF PEND_EMER_check 	= "A" or PEND_EMER_check = "P" THEN
	EA_CHECKBOX = CHECKED
	pending_progs = pending_progs & "EMER" & ", "
END IF
IF PEND_GRH_check 	= "A" or PEND_GRH_check  = "P" THEN
	GRH_CHECKBOX = CHECKED
	pending_progs = pending_progs & "GRH"
END IF

'trims excess spaces of pending_progs
pending_progs = trim(pending_progs)
'takes the last comma off of pending_progs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(pending_progs, 1) = "," THEN pending_progs = left(pending_progs, len(pending_progs) - 1)

'Determines which application check the user is at----------------------------------------------------------------------------------------------------
If DateDiff("d", application_check_date, date) = 0 then
	application_check = "Day 1"
	reminder_date = dateadd("d", 5, application_check_date)
	reminder_text = "Day 5"
Elseif DateDiff("d", application_check_date, date) = 1 then
	application_check = "Day 1"
	reminder_date = dateadd("d", 5, application_check_date)
	reminder_text = "Day 5"
Elseif (DateDiff("d", application_check_date, date) > 1 AND DateDiff("d", application_check_date, date) < 9) then
	application_check = "Day 5"
	reminder_date = dateadd("d", 10, application_check_date)
	reminder_text = "Day 10"
Elseif (DateDiff("d", application_check_date, date) => 10 AND DateDiff("d", application_check_date, date) < 20) then
	application_check = "Day 10"
	reminder_date = dateadd("d", 20, application_check_date)
	reminder_text = "Day 20"
Elseif (DateDiff("d", application_check_date, date) => 20 AND DateDiff("d", application_check_date, date) < 30) then
	application_check = "Day 20"
	reminder_date = dateadd("d", 30, application_check_date)
	reminder_text = "Day 30"
Elseif (DateDiff("d", application_check_date, date) => 30 AND DateDiff("d", application_check_date, date) < 45) then
	application_check = "Day 30"
	reminder_date = dateadd("d", 45, application_check_date)
	reminder_text = "Day 45"
Elseif (DateDiff("d", application_check_date, date) => 45 AND DateDiff("d", application_check_date, date) < 60) then
	application_check = "Day 45"
	reminder_date = dateadd("d", 60, application_check_date)
	reminder_text = "Day 60"
Elseif DateDiff("d", application_check_date, date) = 60 then
	application_check = "Day 60"
	reminder_date = dateadd("d", 10, date)
	reminder_text = "Post day 60"
Elseif DateDiff("d", application_check_date, date) > 60 then
	application_check = "Over 60 days"
	reminder_date = dateadd("d", 10, date)
	reminder_text = "Post day 60"
End if

BeginDialog application_check_dialog, 0, 0, 391, 180, "Application Check:  & application_check"
  DropListBox 75, 15, 80, 15, "Select one..."+chr(9)+"Apply MN"+chr(9)+"CAF"+chr(9)+"CAF addendum"+chr(9)+"HC - certain populations"+chr(9)+"HC - LTC"+chr(9)+"HC - EMA Mnsure ", application_type_droplist
  DropListBox 75, 40, 155, 15, "Select One:"+chr(9)+"Case is ready to approve or deny"+chr(9)+"Requested verifications not received"+chr(9)+"Partial verfications received, more are needed"+chr(9)+"Interview still needed"+chr(9)+"Other", application_status_droplist
  EditBox 175, 20, 50, 15, application_check_date
  EditBox 100, 60, 170, 15, other_app_notes
  EditBox 100, 80, 170, 15, verifs_rcvd
  EditBox 100, 100, 280, 15, verifs_needed
  EditBox 100, 120, 280, 15, actions_taken
  EditBox 100, 140, 280, 15, other_notes
  EditBox 100, 160, 125, 15, worker_signature
  CheckBox 295, 55, 30, 10, "CASH", CASH_CHECKBOX
  CheckBox 295, 65, 25, 10, "EA", EA_CHECKBOX
  CheckBox 295, 75, 25, 10, "FS", FS_CHECKBOX
  CheckBox 340, 60, 30, 10, "GRH", GRH_CHECKBOX
  CheckBox 340, 70, 20, 10, "HC", HC_CHECKBOX
  ButtonGroup ButtonPressed
    PushButton 240, 15, 30, 10, "AREP", AREP_button
    PushButton 275, 15, 30, 10, "DISA", DISA_button
    PushButton 310, 15, 30, 10, "HCRE", HCRE_button
    PushButton 345, 15, 30, 10, "JOBS", JOBS_button
    PushButton 240, 25, 30, 10, "PROG", PROG_button
    PushButton 275, 25, 30, 10, "REVW", REVW_button
    PushButton 310, 25, 30, 10, "SHEL", SHEL_button
    PushButton 345, 25, 30, 10, "UNEA", UNEA_button
    OkButton 275, 165, 50, 15
    CancelButton 330, 165, 50, 15
  Text 10, 20, 55, 10, "Application type:"
  Text 175, 10, 55, 10, "Application date"
  Text 10, 45, 60, 10, "Application status:"
  Text 10, 65, 90, 10, "If status is 'other' explain:"
  Text 20, 85, 75, 10, "Verifications Received:"
  Text 10, 105, 85, 10, "Verifications Still Needed:"
  Text 50, 125, 50, 10, "Actions Taken:"
  Text 55, 145, 45, 10, "Other Notes:"
  Text 35, 165, 60, 10, "Worker Signature:"
  GroupBox 5, 5, 160, 30, "Day 1 application check only"
  GroupBox 235, 5, 145, 35, "MAXIS navigation"
  GroupBox 280, 45, 100, 45, "Pending Programs"
EndDialog


'main dialog
Do
	Do
		err_msg = ""
		dialog application_check_dialog
		cancel_confirmation
		MAXIS_dialog_navigation
		If application_status_droplist = "Select One:"  then err_msg = err_msg & vbNewLine & "* You must enter the application status."
		IF actions_taken = ""  then err_msg = err_msg & vbNewLine & "* You must enter your case actions."
		If application_status_droplist = "Other" AND other_app_notes = ""  then err_msg = err_msg & vbNewLine & "* You must enter more information about the 'other' application status."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If application_status_droplist <> "Case is ready to approve or deny" THEN
	'Outlook appointment is created in prior to the case note being created
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "Application check: " & reminder_text & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
	Outlook_remider = True
End if

If other_app_notes <> "" Then application_status_droplist = application_status_droplist & ", " & other_app_notes

'THE CASENOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("-------------------------" & application_check & " application check")
If application_type_droplist <> "Select One:" then Call write_bullet_and_variable_in_CASE_NOTE("Type of application rec'd", application_type_droplist)
Call write_bullet_and_variable_in_CASE_NOTE("Program applied for", pending_progs)
Call write_bullet_and_variable_in_CASE_NOTE("Application date", application_check_date)
Call write_variable_in_CASE_NOTE("---")
Call write_bullet_and_variable_in_CASE_NOTE("Application status", application_status_droplist)
Call write_bullet_and_variable_in_CASE_NOTE("Verification(s) received", verifs_rcvd)
Call write_bullet_and_variable_in_CASE_NOTE("Verification(s) still needed - request sent via ECF", verifs_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
Call write_bullet_and_variable_in_CASE_NOTE("Other application notes", other_notes)
If Outlook_remider = True then call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

'message boxes based on the application status chosen instructing workers which scripts to use next
If application_status_droplist = "Case is ready to approve or deny" Then
	Msgbox "Success!  You have identified that the case is either ready to approve or deny." & vbNewLine & vbNewLine & _
	"If your case is ready to approve, please use the ""NOTES - APPROVED PROGRAMS"" script." & vbNewLine & vbNewLine & _
	"If your case is ready to be denied, please use the ""NOTES - DENIED PROGRAMS"" script."
ELSEIF application_status_droplist = "No verifs rec'd yet(verification request has been sent)" Then
	Msgbox "Success!  You have identified that no verifications have been received yet, and a verification request has been sent." & vbNewLine & vbNewLine & _
	"Please check to see that there is a verification requested case note, and if not, please use the ""NOTES - VERIFICATIONS REQUESTED"" script."
ELSEIF application_status_droplist = "Some verifs rec'd & more verification are needed" Then
	Msgbox "Success!  You have identified that the your case has received some verifications, but others are needed." & vbNewLine & vbNewLine & _
	"Please check to see that the documents received have been case noted, as well as which verifications are still needed, and if a new verification request was sent." & vbNewLine & _
	"Please use the ""NOTES - DOCUMENTS RECEIVED"" script and/or the ""NOTES - VERIFICATIONS REQUESTED"" as needed."
END IF

script_end_procedure("")
