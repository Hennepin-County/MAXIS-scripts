'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IV-E-MA NEEDY CHILD"
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
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
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
BeginDialog dialog1, 0, 0, 181, 75, "Select a MA needy child option"
  EditBox 95, 10, 60, 15, MAXIS_case_number
  DropListBox 95, 30, 60, 10, "Select one..."+chr(9)+"Close"+chr(9)+"ER"+chr(9)+"Open", action_option
  ButtonGroup ButtonPressed
    OkButton 50, 50, 50, 15
    CancelButton 105, 50, 50, 15
  Text 10, 35, 80, 10, "Select the action to take:"
  Text 45, 15, 45, 10, "Case number:"
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog dialog1
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF action_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a MA needy child option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

If action_option = "Close" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 291, 195, "MA needy child closed"
      EditBox 65, 10, 70, 15, effective_date
      CheckBox 165, 5, 115, 10, "MAXIS/ECF case to closed files", closed_files_checkbox
      CheckBox 165, 20, 60, 10, "MMIS updated", MMIS_updated_checkbox
      EditBox 65, 35, 215, 15, reason_close
      EditBox 65, 55, 215, 15, placement_ended
      EditBox 65, 75, 215, 15, notified_by
      EditBox 65, 95, 215, 15, over_income
      EditBox 65, 115, 215, 15, fail_to_provide
      EditBox 65, 135, 215, 15, other
      EditBox 65, 155, 215, 15, other_notes
      EditBox 65, 175, 105, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 175, 175, 50, 15
        CancelButton 230, 175, 50, 15
      Text 25, 40, 35, 10, "Reason(s):"
      Text 15, 120, 50, 10, "Fail to provide:"
      Text 25, 160, 40, 10, "Other notes:"
      Text 25, 80, 40, 10, "Notified by:"
      Text 20, 100, 45, 10, "Over income:"
      Text 5, 60, 60, 10, "Placement ended:"
      Text 15, 15, 50, 10, "Effective date:"
      Text 40, 140, 25, 10, "Other:"
      Text 5, 180, 60, 10, "Worker signature:"
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
            IF reason_close = "" then err_msg = err_msg & vbNewLine & "* Enter the reason for closure."
			If placement_ended = "" then err_msg = err_msg & vbNewLine & "* Enter information about the placement ending."
			If notified_by = "" then err_msg = err_msg & vbNewLine & "* Enter the 'notified by' information."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 	Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("**MA Needy child closed effective " & effective_date & "**")
    Call write_bullet_and_variable_in_CASE_NOTE("Reason(s)", reason_close)
    Call write_bullet_and_variable_in_CASE_NOTE("placement ended", placement_ended)
    Call write_bullet_and_variable_in_CASE_NOTE("Notified by", notified_by)
    call write_bullet_and_variable_in_CASE_NOTE("Over income", over_income)
    Call write_bullet_and_variable_in_CASE_NOTE("Fail to provide", fail_to_provide)
    Call write_bullet_and_variable_in_CASE_NOTE("Other", Other)
    If closed_files_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MAXIS/case file sent to closed files.")
    If MMIS_updated_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MMIS updated.")
END IF

If action_option = "ER" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 286, 175, "MA needy child ER"
      EditBox 45, 10, 40, 15, ER_date
      EditBox 155, 10, 60, 15, HCAPP_date
      EditBox 240, 10, 40, 15, child_age
      EditBox 45, 35, 235, 15, AREP
      EditBox 45, 55, 235, 15, income
      EditBox 45, 75, 235, 15, OHC
      EditBox 60, 95, 220, 15, placement_info
      EditBox 45, 115, 170, 15, Rule_five
      CheckBox 220, 120, 60, 10, "MMIS updated", MMIS_updated
      EditBox 45, 135, 235, 15, other_notes
      EditBox 70, 155, 100, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 175, 155, 50, 15
        CancelButton 230, 155, 50, 15
      Text 15, 60, 30, 10, "Income:"
      Text 90, 15, 60, 10, "HCAPP rec'd date:"
      Text 15, 15, 30, 10, "ER date:"
      Text 5, 140, 40, 10, "Other notes:"
      Text 25, 80, 20, 10, "OHC:"
      Text 15, 120, 25, 10, "Rule 5:"
      Text 5, 100, 50, 10, "Placement info:"
      Text 20, 40, 20, 10, "AREP:"
      Text 10, 160, 60, 10, "Worker signature:"
      Text 220, 15, 15, 10, "Age:"
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
			If isDate(HCAPP_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid HCAPP date."
			If ER_date = "" then err_msg = err_msg & vbNewLine & "* Enter the ER date."
			if IsNumeric(child_age) = False then err_msg = err_msg & vbNewLine & "* Enter a valid age for the client."
			IF income = "" then err_msg = err_msg & vbNewLine & "* Enter the income information."
			If placement_info = "" then err_msg = err_msg & vbNewLine & "* Enter the placement information."
			If Rule_five = "" then err_msg = err_msg & vbNewLine & "* Enter the Rule 5 information."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("**MA Needy child ER rec'd for " & ER_date & "**")
	Call write_bullet_and_variable_in_CASE_NOTE("HCAPP rec'd date", HCAPP_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Client age", child_age)
	Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
    Call write_bullet_and_variable_in_CASE_NOTE("Income", income)
    Call write_bullet_and_variable_in_CASE_NOTE("OHC", OHC)
    call write_bullet_and_variable_in_CASE_NOTE("Placement info", placement_info)
    Call write_bullet_and_variable_in_CASE_NOTE("Rule 5", Rule_five)
    If MMIS_updated_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MMIS updated.")
END IF

If action_option = "Open" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 296, 240, "MA needy child open"
      EditBox 70, 10, 70, 15, HCAPP_date
      EditBox 200, 10, 90, 15, HH_comp
      EditBox 70, 30, 70, 15, effective_date
      EditBox 200, 30, 90, 15, ER_date
      EditBox 45, 55, 245, 15, income
      EditBox 45, 75, 245, 15, OHC
      EditBox 45, 95, 190, 15, MMIS
      CheckBox 240, 100, 55, 10, "TPL updated", TPL_updated
      EditBox 45, 115, 190, 15, AREP
      EditBox 45, 135, 245, 15, placed
      EditBox 45, 155, 245, 15, results
      EditBox 45, 175, 115, 15, Rule_five
      EditBox 205, 175, 85, 15, due_date
      EditBox 45, 195, 245, 15, other_notes
      EditBox 70, 215, 110, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 185, 215, 50, 15
        CancelButton 240, 215, 50, 15
      Text 10, 220, 60, 10, "Worker signature:"
      Text 15, 60, 30, 10, "Income:"
      Text 20, 100, 20, 10, "MMIS:"
      Text 5, 15, 60, 10, "HCAPP rec'd date:"
      Text 170, 35, 30, 10, "ER date:"
      Text 5, 200, 40, 10, "Other notes:"
      Text 15, 160, 30, 10, "Results:"
      Text 25, 80, 20, 10, "OHC:"
      Text 15, 180, 25, 10, "Rule 5:"
      Text 160, 15, 35, 10, "HH comp:"
      Text 15, 140, 25, 10, "Placed: "
      Text 15, 35, 50, 10, "Effective date:"
      Text 20, 120, 20, 10, "AREP:"
      Text 170, 180, 35, 10, "Due date:"
    EndDialog
	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
			If isDate(HCAPP_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid HCAPP date."
			IF HH_comp = "" then err_msg = err_msg & vbNewLine & "* Enter the case's HH composition."
			If isDate(effective_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid effective date."
			If ER_date = "" then err_msg = err_msg & vbNewLine & "* Enter the ER date."
			IF income = "" then err_msg = err_msg & vbNewLine & "* Enter the income information."
            If MMIS = "" then err_msg = err_msg & vbNewLine & "* Enter the MMIS information."
            If placed = "" then err_msg = err_msg & vbNewLine & "* Enter the placement information."
			If Results = "" then err_msg = err_msg & vbNewLine & "* Enter the results information."
			If Rule_five = "" then err_msg = err_msg & vbNewLine & "* Enter the Rule 5 information."
			If isDate(due_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid due date."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("**MA Needy child opened effective " & effective_date & "**")
	Call write_bullet_and_variable_in_CASE_NOTE("HCAPP rec'd date", HCAPP_date)
	Call write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
	Call write_bullet_and_variable_in_CASE_NOTE("ER date", ER_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Income", income)
	Call write_bullet_and_variable_in_CASE_NOTE("OHC", OHC)
	Call write_bullet_and_variable_in_CASE_NOTE("MMIS", MMIS)
    Call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
    call write_bullet_and_variable_in_CASE_NOTE("Placed", Placed)
	Call write_bullet_and_variable_in_CASE_NOTE("Results", Results)
    Call write_bullet_and_variable_in_CASE_NOTE("Rule 5", Rule_five)
	Call write_bullet_and_variable_in_CASE_NOTE("Due date", due_date)
    If TPL_updated = 1 then Call write_variable_in_CASE_NOTE("* TPL updated.")
END IF

Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("")'
