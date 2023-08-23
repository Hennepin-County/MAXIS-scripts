'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IV-E-TRIAL HOME VISIT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
call changelog_update("11/25/2019", "Updated backend functionality, and added changelog.", "Ilse Ferris, Hennepin County")
call changelog_update("11/25/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog dialog1, 0, 0, 181, 75, "Select a trial home visit option"
  EditBox 95, 10, 60, 15, MAXIS_case_number
  DropListBox 95, 30, 60, 10, "Select one.."+chr(9)+"Begins"+chr(9)+"Change"+chr(9)+"Ends", THV_option
  ButtonGroup ButtonPressed
    OkButton 50, 50, 50, 15
    CancelButton 105, 50, 50, 15
  Text 15, 35, 75, 10, "Trail home visit option:"
  Text 45, 15, 45, 10, "Case number:"
EndDialog
DO
	DO
		err_msg = ""
		Dialog dialog1
        cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF THV_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a trial home visit option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If THV_option = "Begins" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 291, 160, "Trial home visit begins"
      EditBox 65, 10, 90, 15, effective_date
      DropListBox 215, 10, 70, 15, "Select one..."+chr(9)+"Yes "+chr(9)+"No", court_ordered
      EditBox 65, 30, 220, 15, THV_verif
      EditBox 65, 50, 220, 15, SSIS
      DropListBox 65, 70, 90, 15, "Select one..."+chr(9)+"Yes "+chr(9)+"No", basic_IVE
      EditBox 215, 70, 70, 15, reim_ended
      EditBox 65, 90, 220, 15, other_notes
      CheckBox 65, 110, 60, 10, "MEMI updated", MEMI_checkbox
      CheckBox 135, 110, 145, 10, "Navigate to DAIL/WRIT after case note", dail_checkbox
      EditBox 65, 130, 110, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 180, 130, 50, 15
        CancelButton 235, 130, 50, 15
      Text 5, 135, 60, 10, "Worker signature: "
      Text 20, 95, 40, 10, "Other notes: "
      Text 15, 15, 50, 10, "Effective date:"
      Text 10, 35, 45, 10, "How verified:"
      Text 160, 15, 50, 10, "Court ordered:"
      Text 45, 55, 20, 10, "SSIS:"
      Text 20, 75, 40, 10, "Basic IV-E?:"
      Text 165, 75, 50, 10, "Reimb. ended:"
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_without_confirmation
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
            IF court_ordered = "Select one..." then err_msg = err_msg & vbNewLine & "* Was the trial home visit court ordered?"
			If THV_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the trial home visit verification."
			If SSIS = "" then err_msg = err_msg & vbNewLine & "* Enter the SSIS information."
			IF basic_IVE = "Select one..." then err_msg = err_msg & vbNewLine & "* Is this Basic IV-E?"
			If isDate(reim_ended) = False then err_msg = err_msg & vbNewLine & "* Enter a valid reimbursement end date."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 	Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("~~Trial home visit begins~~")
    Call write_bullet_and_variable_in_CASE_NOTE("Effective date", effective_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Court ordered", goal_two)
    Call write_bullet_and_variable_in_CASE_NOTE("Verification", THV_verif)
    call write_bullet_and_variable_in_CASE_NOTE("Basic IV-E", basic_IVE)
    Call write_bullet_and_variable_in_CASE_NOTE("SSIS", SSIS)
    Call write_bullet_and_variable_in_CASE_NOTE("Reimbursement ended", reim_ended)
    If MEMI_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MEMI has been updated.")
    If dail_checkbox = 1 then Call write_variable_in_CASE_NOTE("* TIKL created to recheck case.")
END IF

If THV_option = "Change" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 301, 115, "Trial home visit change"
      EditBox 75, 10, 220, 15, change_reason
      EditBox 75, 30, 80, 15, effective_date
      EditBox 75, 50, 220, 15, actions_taken
      EditBox 75, 70, 220, 15, other_notes
      EditBox 75, 90, 110, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 190, 90, 50, 15
        CancelButton 245, 90, 50, 15
      Text 20, 35, 50, 10, "Effective date:"
      Text 20, 55, 50, 10, "Actions taken:"
      Text 10, 95, 60, 10, "Worker signature: "
      Text 5, 15, 65, 10, "Reason for change:"
      Text 30, 75, 40, 10, "Other notes: "
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
			IF change_reason = "" then err_msg = err_msg & vbNewLine & "* Enter the trial home visit change reason."
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
			If actions_taken = "" then err_msg = err_msg & vbNewLine & "* Enter the actions taken."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("~~Trial home visit changed~~")
    Call write_bullet_and_variable_in_CASE_NOTE("Reason for change", change_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("Effective date", effective_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
END IF

If THV_option = "Ends" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 291, 160, "Trial home visit ends"
      EditBox 65, 10, 220, 15, reason_ending
      EditBox 65, 30, 90, 15, effective_date
      DropListBox 215, 30, 70, 15, "Select one..."+chr(9)+"Yes "+chr(9)+"No", court_ordered
      EditBox 65, 50, 220, 15, THV_verif
      EditBox 65, 70, 220, 15, SSIS
      EditBox 65, 90, 70, 15, reim_updated
      EditBox 65, 110, 220, 15, other_notes
      EditBox 65, 130, 110, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 180, 130, 50, 15
        CancelButton 235, 130, 50, 15
      Text 20, 115, 40, 10, "Other notes: "
      Text 15, 35, 50, 10, "Effective date:"
      Text 20, 55, 45, 10, "How verified:"
      Text 165, 35, 50, 10, "Court ordered:"
      Text 45, 75, 20, 10, "SSIS:"
      Text 10, 95, 55, 10, "Reimb. updated:"
      Text 5, 135, 60, 10, "Worker signature: "
      Text 10, 15, 55, 10, "Reason ending:"
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
			IF reason_ending = "" then err_msg = err_msg & vbNewLine & "* Enter the ending reason."
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
			IF court_ordered = "Select one..." then err_msg = err_msg & vbNewLine & "* Was the trial home visit court ordered?"
			If THV_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the trial home visit verification."
            If SSIS = "" then err_msg = err_msg & vbNewLine & "* Enter the SSIS information."
            If isDate(reim_updated) = False then err_msg = err_msg & vbNewLine & "* Enter a valid reimbursement updated date."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("~~Trial home visit ends~~")
    Call write_bullet_and_variable_in_CASE_NOTE("Reason ending", reason_ending)
    Call write_bullet_and_variable_in_CASE_NOTE("Effective date", effective_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Court ordered", goal_two)
    Call write_bullet_and_variable_in_CASE_NOTE("Verification", THV_verif)
    Call write_bullet_and_variable_in_CASE_NOTE("SSIS", SSIS)
    Call write_bullet_and_variable_in_CASE_NOTE("Reimbursement updated", reim_updated)
END IF

Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

If dail_checkbox = 1 then
    PF3     'to save the case note
    Call navigate_to_MAXIS_screen("DAIL", "WRIT")
    msgbox "Please enter your DAIL message and date now."
END IF

script_end_procedure("")
