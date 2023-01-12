'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - PF11 ACTIONS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 180                	'manual run time in seconds
STATS_denomination = "C"       		'M is for each MEMBER
'END OF stats block=========================================================================================================

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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("01/12/2023", "Removed HH member selection as this script supports case-based actions.", "Ilse Ferris, Hennepin County")
call changelog_update("09/29/2021", "Removed PMI merge requests from script. As of 10/2021 all PMI merge requests need to go through a SIR webform.", "Ilse Ferris, Hennepin County")
call changelog_update("09/29/2021", "Removed MFIP New Spouse Income supports from script. PF11 is no longer needed for NSI cases.", "Ilse Ferris, Hennepin County")
call changelog_update("06/04/2021", "Updated to add option for HC REVW Dates need update GitHub Issue #168.", "MiKayla Handley, Hennepin County")
call changelog_update("08/07/2020", "Updated to review CASE/NOTE for previous PF11 request.", "MiKayla Handley, Hennepin County")
call changelog_update("07/20/2019", "Per DHS Bulletin #18-69-02C, updated New Spouse Income handling and case note.", "MiKayla Handley, Hennepin County")
call changelog_update("05/13/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT=================================================================================================================
'Connecting to MAXIS, and grabbing the case number and footer month'
EMConnect ""
CALL check_for_maxis(FALSE) 'checking for passord out, brings up dialog'
CALL MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
 BeginDialog Dialog1, 0, 0, 201, 215, "PF11 Actions"
    Text 15, 10, 45, 10, "Case number:"
    EditBox 65, 5, 45, 15, MAXIS_case_number
    Text 20, 30, 45, 10, "Select action:"
    DropListBox 65, 25, 125, 15, "Select One:"+chr(9)+"Non-Actionable DAIL Removal"+chr(9)+"Case Note Removal Request"+chr(9)+"Unable to update HC REVW dates"+chr(9)+"Other", PF11_actions
    Text 5, 50, 60, 10, "Worker signature:"
    EditBox 65, 45, 125, 15, worker_signature
    Text 5, 70, 190, 25, "Note: The MAXIS system being down, issuance problems, or any type of emergency are examples of what should NOT be reported via a PF11."
    ButtonGroup ButtonPressed
    OkButton 100, 100, 45, 15
    CancelButton 150, 100, 45, 15
    GroupBox 5, 120, 195, 75, "How to check a PF11 status:"
    Text 10, 135, 185, 50, "On the SELF menu and type TASK. If you have the task number enter it and it will take you directly into the PF11. If you do not have the task number or wish to look at a list of all the PF11s you have created, change the Option in TASK from T-task to a C-creator.  By placing an X next to a PF11 listed you will be able to view it."
EndDialog

'Running the dialog for case number and client

Do
    Do
        err_msg = ""
        Dialog Dialog1
        cancel_without_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
    	IF PF11_actions = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Select the action you wish to take."
    	IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."
    	IF err_msg <> "" Then MsgBox "Please resolve the following to continue:" & vbNewLine & err_msg
    Loop until err_msg = ""
    Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

'Out of county and PRIV handling
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True then
    script_end_procedure("Case is privileged. The script will now end.")
Else
    EmReadscreen county_code, 4, 21, 21
End if
If UCASE(worker_county_code) <> county_code then script_end_procedure("Case is out-of-county. The script will now end.")

IF PF11_actions = "Non-Actionable DAIL Removal" THEN
	Call Navigate_to_MAXIS_screen ("DAIL", "DAIL")
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 316, 85, "Non-Actionable DAIL Removal"
	  EditBox 45, 5, 35, 15, message_date
	  EditBox 110, 5, 200, 15, message_to_use
	  EditBox 110, 25, 200, 15, request_reason
	  EditBox 55, 45, 255, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 205, 65, 50, 15
	    CancelButton 260, 65, 50, 15
	  Text 5, 10, 40, 10, "DAIL Date:"
	  Text 5, 30, 105, 10, "Reason DAIL is non-actionable:"
	  Text 90, 10, 20, 10, "DAIL:"
	  Text 5, 50, 45, 10, "Other Notes:"
	EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF message_date = "" THEN err_msg = err_msg & vbNewLine & "Please enter a dail date."
			IF message_to_use = "" THEN err_msg = err_msg & vbNewLine & "Please enter a dail. This can be done via copy and paste."
			IF request_reason = "" THEN err_msg = err_msg & vbNewLine & "Please enter a request reason."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "Case Note Removal Request" THEN ' this does not leave a case note'
	Call Navigate_to_MAXIS_screen ("CASE", "NOTE")
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 316, 85, "Case Note Removal Request"
      EditBox 45, 5, 35, 15, message_date
      EditBox 120, 5, 190, 15, message_to_use
      EditBox 75, 25, 235, 15, request_reason
      EditBox 50, 45, 260, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 205, 65, 50, 15
        CancelButton 260, 65, 50, 15
      Text 5, 10, 35, 10, "Note Date:"
      Text 5, 30, 70, 10, "Reason for removal:"
      Text 90, 10, 30, 10, "Header:"
      Text 5, 50, 45, 10, "Other Notes:"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF message_date = "" THEN err_msg = err_msg & vbNewLine & "* Enter the case note date."
			IF message_to_use = "" THEN err_msg = err_msg & vbNewLine & "* Enter the case note header. This can be copy and paste."
			IF request_reason = "" THEN err_msg = err_msg & vbNewLine & "* Enter enter a request reason."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "Other" THEN
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 326, 110, "Other"
      EditBox 70, 10, 240, 15, request_reason
      EditBox 70, 30, 240, 15, other_notes
      EditBox 70, 50, 240, 15, action_taken
      ButtonGroup ButtonPressed
        OkButton 205, 70, 50, 15
        CancelButton 260, 70, 50, 15
      Text 5, 15, 60, 10, "Describe Problem:"
      Text 25, 35, 45, 10, "Other Notes:"
      Text 15, 55, 50, 10, "Actions Taken:"
      Text 5, 95, 230, 10, "While the dialog box is open navigate to the panel you wish to report."
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
			IF trim(request_reason) = "" THEN err_msg = err_msg & vbNewLine & "* Enter a request reason."
            If trim(action_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter your actions taken on the case."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

IF PF11_actions = "Unable to update HC REVW dates" THEN
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 316, 110, "Unable to update HC REVW dates"
        EditBox 60, 50, 40, 15, message_date
        EditBox 185, 50, 125, 15, last_review_date_edit
        EditBox 60, 70, 250, 15, other_notes
        ButtonGroup ButtonPressed
        OkButton 215, 90, 45, 15
        CancelButton 265, 90, 45, 15
        Text 10, 55, 45, 10, "Review Date:"
        Text 110, 55, 75, 10, "Last Review Date Edit:"
        Text 15, 75, 45, 10, "Other Notes:"
        Text 10, 20, 290, 15, "This option will send a PF11 for DHS to update the review date. Additional information found in OneSource. A SIR Email will inform you when issue is resolved. "
        GroupBox 5, 5, 305, 40, "What will this PF11 action do?"
    EndDialog

    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
    		IF message_date = "" THEN err_msg = err_msg & vbNewLine & "* Enter the review date."
			IF last_review_date_edit = "" THEN err_msg = err_msg & vbNewLine & "* Enter the edit you recieved when trying to change the REVW date. This can be done via copy and paste."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
     Call check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False
END If

'----------------------------------------------------------------------------------------------------Sending the PF11
EMReadScreen current_screen, 4, 2, 48
If current_screen <> "MEMB" then Call navigate_to_MAXIS_screen("STAT", "MEMB")

PF11

EMReadScreen nav_check, 4, 1, 27
IF nav_check = "Prob" THEN
    EMWriteScreen PF11_actions & " for case number: " & MAXIS_case_number, 05, 07
    EMWriteScreen "Date: " & message_date, 06, 07
    IF PF11_actions = "Case Note Removal Request" THEN EMWriteScreen "Case Note: " & message_to_use, 07, 07

    IF PF11_actions = "Non-Actionable DAIL Removal" THEN
        EMWriteScreen "DAIL: " & message_to_use, 07, 07
        EMWriteScreen "Other error to report:", 08, 07
    End if
    IF PF11_actions = "Unable to update HC REVW dates" THEN
        EMWriteScreen "Last Review Date Edit: " & last_review_date_edit, 08, 07
        EMWriteScreen "Other notes: " & other_notes, 09, 07
    END IF
    IF request_reason <> "" THEN EMWriteScreen "Reason for request: " & request_reason, 09, 07

	TRANSMIT
	EMReadScreen task_number, 7, 3, 27
	TRANSMIT
	PF3 ''-self'
	PF3 '- MEMB'
ELSE
	script_end_procedure_with_error_report("Could not reach PF11." & PF11_actions & " has not been sent.")
END IF

reminder_date = dateadd("d", 5, date)
Call change_date_to_soonest_working_day(reminder_date, "BACK")

IF PF11_actions <> "Case Note Removal Request" then
    'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, reminder_in_minutes, appt_category)
    Call create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "PF11 check: " & PF11_actions & " for " & MAXIS_case_number, "", "", TRUE, 5, "")

    CALL start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	CALL write_variable_in_case_note("PF11 Requested:" & PF11_actions)
	CALL write_bullet_and_variable_in_CASE_NOTE("Reason for request", request_reason)
	CALL write_bullet_and_variable_in_CASE_NOTE("Task number", task_number)
	CALL write_bullet_and_variable_in_CASE_NOTE("Message", message_to_use)
	CALL write_bullet_and_variable_in_CASE_NOTE("Review Date", message_date)
	CALL write_bullet_and_variable_in_CASE_NOTE("Last Review Date Edit", last_review_date_edit)
	IF PF11_actions = "Unable to update HC REVW dates"  THEN
		CALL write_variable_in_CASE_NOTE("Sent a PF11 for MNIT to update the review date")
		CALL write_variable_in_CASE_NOTE("This information can be found directly in OneSource")
		CALL write_variable_in_CASE_NOTE("SIR Email will be reviewed for a response when this process is complete.")
	END IF
	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	CALL write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_taken)
	call write_bullet_and_variable_in_CASE_NOTE("Outlook reminder set for", reminder_date)
	CALL write_variable_in_CASE_NOTE ("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
END IF

script_end_procedure_with_error_report("Success! " & PF11_actions & " has been sent and/or case noted.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/29/2021
'--Tab orders reviewed & confirmed----------------------------------------------09/29/2021
'--Mandatory fields all present & Reviewed--------------------------------------09/29/2021
'--All variables in dialog match mandatory fields-------------------------------09/29/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/29/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------09/29/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/29/2021
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/29/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------09/29/2021
'--PRIV Case handling reviewed -------------------------------------------------09/29/2021
'--Out-of-County handling reviewed----------------------------------------------09/29/2021
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/29/2021
'--BULK - review output of statistics and run time/count (if applicable)--------09/29/2021-----------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/29/2021 ------------updated to 180 seconds
'--Incrementors reviewed (if necessary)-----------------------------------------09/29/2021-----------------N/A
'--Denomination reviewed -------------------------------------------------------09/29/2021-----------------N/A
'--Script name reviewed---------------------------------------------------------09/29/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------09/29/2021-----------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------09/29/2021
'--comment Code-----------------------------------------------------------------09/29/2021
'--Update Changelog for release/update------------------------------------------09/29/2021
'--Remove testing message boxes-------------------------------------------------09/29/2021
'--Remove testing code/unnecessary code-----------------------------------------09/29/2021
'--Review/update SharePoint instructions----------------------------------------10/20/2021 ---------------Updated instructions to remove PMI Merge
'--Review Best Practices using BZS page ----------------------------------------09/29/2021
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/29/2021----------------May need to update if PMI merges cannot be via PF11 any longer
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/29/2021
'--Complete misc. documentation (if applicable)---------------------------------09/29/2021-----------------N/A
'--Update project team/issue contact (if applicable)----------------------------09/29/2021-----------------N/A
