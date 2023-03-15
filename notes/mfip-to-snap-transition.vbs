'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MFIP TO SNAP TRANSITION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 420          'manual run time in seconds
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
call changelog_update("04/11/2018", "Removed the option to add a WCOM to snap approval.", "MiKayla Handley, Hennepin County")
call changelog_update("04/11/2018", "Updated the 'Notes on Income' field to a mandatory field.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
''================END DIALOG SECTION
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 151, 75, "MFIP To SNAP Transition Note"
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 95, 50, 50, 15
  EditBox 70, 5, 75, 15, MAXIS_case_number
  EditBox 70, 25, 75, 15, closure_date
  Text 5, 10, 60, 10, "Case Number:"
  Text 5, 30, 55, 20, "Date MFIP closes:"
EndDialog
'pulls up case number dialog
DO
    DO
    	err_msg = ""
    	Dialog Dialog1
        Cancel_without_confirmation
    	IF MAXIS_case_number = "" THEN err_msg = err_msg & vbCr & "Please enter a case number"
    	IF isdate(closure_date) = false THEN err_msg = err_msg & vbCr & "You must enter a valid MFIP closure date."
    	IF isdate(closure_date) = true THEN
    		IF datepart("d", dateadd("d", 1, closure_date)) <> 1 THEN err_msg = err_msg & vbCr & "The MFIP closure date should equal the last day of the month."
    	END IF
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    LOOP Until err_msg = ""
LOOP UNTIL are_we_passworded_out = false
'Setting the correct footer month and year (so snap budget is pulled from month snap is approved in)
MAXIS_footer_month = datepart("m", dateadd("d", 1, closure_date))
if len(MAXIS_footer_month) = 1 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", dateadd("d", 1, closure_date)), 2)

Call check_for_maxis(true)
call HH_member_custom_dialog(HH_member_array)

'This section performs case accuracy checks

'First, checking that MFIP closure was approved today
call navigate_to_MAXIS_screen("ELIG", "MFIP")
EMReadScreen MFIP_version_check, 10, 24, 2 'need to make sure there is an MFIP version out there, or this won't work
IF MFIP_version_check = "NO VERSION" THEN
	msgbox "There is currently no version of MFIP on this case.  Please check your case and try again.  The script will now stop."
	script_end_procedure("")
END IF
EMReadscreen total_versions, 1, 2, 18
For i = total_versions to 1 step -1 'Finding the most recent approved version and reading the approval date
	EMReadscreen approved_check, 8, 3, 3
	IF approved_check = "APPROVED" THEN
		EMReadscreen approval_date, 8, 3, 14
		EXIT FOR
	END IF
	EMWritescreen (i - 1), 20, 79
	transmit
Next

'comparing dates, and giving a warning message if the closure wasn't approved today
IF datediff("d", date, approval_date) <> 0 THEN msgbox "Warning! It appears the most recent MFIP version was not approved today. Approval of SNAP when closing MFIP must occur on the same day as closure."

'Next, check that there isn't a REVW due for at least 1 full month
call navigate_to_MAXIS_screen("stat", "revw")
EMReadscreen cash_revw_date, 8, 9, 37
cash_revw_date = replace(cash_revw_date, " ", "/")
IF datediff("m", date, cash_revw_date) < 1 THEN msgbox "Warning! The next cash review date on this case is not at least 1 full month after the closure date.  Please review your results."

'This grabs the information for the editboxes
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)  'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp) 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")

call autofill_editbox_from_MAXIS(HH_member_array, "SHEL", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "HEST", SHEL_HEST)
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "COEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "DCEX", COEX_DCEX)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

WCOM_check = checked 'setting checkbox default
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 451, 310, "CAF dialog part 2"
  EditBox 70, 40, 375, 15, HH_comp
  EditBox 70, 60, 375, 15, earned_income
  EditBox 70, 80, 375, 15, unearned_income
  EditBox 105, 100, 340, 15, notes_on_income
  EditBox 65, 125, 380, 15, SHEL_HEST
  EditBox 65, 145, 380, 15, COEX_DCEX
  EditBox 65, 165, 380, 15, MFIP_closure_reason
  EditBox 65, 185, 380, 15, other_notes
  EditBox 65, 205, 380, 15, Actions_taken
  'CheckBox 5, 225, 130, 15, "Add WCOM to SNAP approval notice.", WCOM_check
  CheckBox 145, 225, 155, 15, "All verifs and forms needed for MFIP on file.", verifs_check
  EditBox 335, 260, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 335, 290, 50, 15
    CancelButton 390, 290, 50, 15
    PushButton 10, 15, 15, 10, "FS", ELIG_FS_button
    PushButton 30, 15, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 150, 15, 25, 10, "BUSI", BUSI_button
    PushButton 175, 15, 25, 10, "JOBS", JOBS_button
    PushButton 200, 15, 25, 10, "PBEN", PBEN_button
    PushButton 225, 15, 25, 10, "RBIC", RBIC_button
    PushButton 250, 15, 25, 10, "UNEA", UNEA_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 130, 25, 10, "SHEL/", SHEL_button
    PushButton 30, 130, 25, 10, "HEST:", HEST_button
    PushButton 5, 145, 25, 10, "COEX/", COEX_button
    PushButton 30, 145, 25, 10, "DCEX:", DCEX_button
  Text 5, 190, 50, 10, "Other Notes:"
  Text 5, 65, 55, 10, "Earned income:"
  Text 5, 210, 50, 10, "Actions Taken:"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 45, 45, 10, "HH Comp:"
  Text 5, 85, 65, 10, "Unearned income:"
  Text 5, 105, 95, 10, "Notes on income and budget:"
  Text 255, 260, 70, 10, "Worker Signature:"
  GroupBox 145, 5, 135, 25, "Income panels"
  Text 5, 165, 55, 20, "MFIP closure reason:"
  GroupBox 5, 5, 55, 25, "ELIG panels:"
EndDialog
'Need to add MFIP closure date (and use this to enter SNAP open date)
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
        If trim(notes_on_income) = "" THEN err_msg = err_msg & vbCr & "Enter the 'notes on income' field."
		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "You must complete the actions taken field."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must sign your case note."
		IF verifs_check = unchecked THEN err_msg = err_msg & vbCr & "All needed verifications for MFIP must be on file before approving SNAP. Please update the checkbox."
		IF MFIP_closure_reason = "" THEN err_msg = err_msg & vbCr & "You must explain the MFIP closure reason."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Writing the case note
call start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE("*MFIP CLOSING " & closure_date & ", SNAP APPROVED " & dateadd("d", 1, closure_date))
CALL write_variable_in_CASE_NOTE("Reason for MFIP closure: " & MFIP_closure_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned inc.", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("UNEA", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Income Notes", notes_on_income)
CALL write_bullet_and_variable_in_CASE_NOTE("COEX/DCEX", COEX_DCEX)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", Actions_taken)
IF verifs_check = checked THEN call write_variable_in_CASE_NOTE("* All required verifications and documents needed for MFIP provided.")
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
