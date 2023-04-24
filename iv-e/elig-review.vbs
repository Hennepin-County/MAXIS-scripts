'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IV-E ELIG REVIEW.vbs"
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
BeginDialog dialog1, 0, 0, 341, 345, "IV-E eligibility review"
  EditBox 70, 10, 55, 15, MAXIS_case_number
  EditBox 70, 30, 55, 15, review_month
  EditBox 275, 10, 55, 15, child_age
  EditBox 275, 30, 55, 15, school_verif
  EditBox 70, 50, 260, 15, AFDC_receipt
  EditBox 70, 70, 260, 15, Rule_five
  EditBox 70, 105, 110, 15, earned_income
  EditBox 70, 125, 110, 15, SSI_income
  EditBox 220, 105, 110, 15, RSDI_income
  EditBox 220, 125, 110, 15, other_income
  EditBox 70, 155, 260, 15, child_assets
  EditBox 70, 175, 260, 15, dep_factor
  EditBox 70, 195, 260, 15, basic_elig
  DropListBox 100, 235, 65, 10, "Select one..."+chr(9)+"Yes"+chr(9)+"No", reimb_list
  EditBox 100, 255, 65, 15, non_reimb_months
  EditBox 265, 235, 65, 15, perm_hearing_date
  EditBox 265, 255, 65, 15, new_date
  CheckBox 65, 285, 65, 10, "Checked MAXIS", MAXIS_checkbox
  CheckBox 140, 285, 60, 10, "SSIS checked", SSIS_checkbox
  CheckBox 210, 285, 120, 10, "HC approved, 25X. MMIS updated", HC_approved_checkobx
  EditBox 65, 305, 265, 15, other_notes
  EditBox 65, 325, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 325, 50, 15
    CancelButton 280, 325, 50, 15
  Text 20, 35, 50, 10, "Review month: "
  GroupBox 10, 90, 325, 60, "Child's income:"
  Text 5, 200, 65, 10, "Title IV-E basic elig:"
  Text 15, 160, 50, 10, "Child's assets:"
  Text 5, 330, 60, 10, "Worker signature: "
  Text 5, 180, 60, 10, "Deprivation factor:"
  Text 40, 75, 25, 10, "Rule 5:"
  Text 20, 310, 40, 10, "Other notes: "
  Text 175, 240, 90, 10, "If yes: Perm. hearing date:"
  Text 10, 260, 90, 10, "If no: Non reimb. month(s):"
  Text 195, 260, 70, 10, "New date from FCLD:"
  GroupBox 0, 220, 335, 55, "Title IV-E reimbursable"
  Text 180, 15, 90, 10, "Age at end of review period:"
  Text 20, 15, 50, 10, "Case number:"
  Text 40, 110, 25, 10, "Earned:"
  Text 50, 130, 15, 10, "SSI:"
  Text 195, 110, 20, 10, "RSDI:"
  Text 195, 130, 20, 10, "Other:"
  Text 15, 240, 85, 10, "Title IV-E reimburseable: "
  Text 5, 55, 60, 10, "In receipt of AFDC: "
  Text 170, 35, 105, 10, "If 18, list the school verification:"
EndDialog
DO
	DO
		err_msg = ""
		Dialog dialog1
        cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF review_month = "" then err_msg = err_msg & vbNewLine & "* Enter the review month."
        If IsNumeric(child_age) = False then err_msg = err_msg & vbNewLine & "* Enter a valid age for child."
        If AFDC_receipt = "" then err_msg = err_msg & vbNewLine & "* Enter AFDC information."
        If rule_five = "" then err_msg = err_msg & vbNewLine & "* Enter the Rule 5 information."
        If dep_factor = "" then err_msg = err_msg & vbNewLine & "* Enter the deprivation factor information."
        If basic_elig = "" then err_msg = err_msg & vbNewLine & "* Enter the TITLE IV-E basic eligibilty information."
        If reimb_list = "Select one..." then err_msg = err_msg & vbNewLine & "* Is Title IV-E reimburseable?"
        If reimb_list = "Yes" and perm_hearing_date = "" then err_msg = err_msg & vbNewLine & "* Enter the permanent hearing date."
        If reimb_list = "No" and non_reimb_months = "" then err_msg = err_msg & vbNewLine & "* Enter the non reimburseable months."
        If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'creating incremental variable for income for the case note
child_income = ""
IF earned_income <> "" then child_income = child_income & earned_income & ", "
IF SSI_income <> "" then child_income = child_income & SSI_income & ", "
IF RSDI_income <> "" then child_income = child_income & RSDI_income & ", "
IF other_income <> "" then child_income = child_income & other_income & ", "
'trims excess spaces of child_income
child_income = trim(child_income)
'takes the last comma off of child_income when autofilled into dialog if more more than one app date is found and additional app is selected
If right(child_income, 1) = "," THEN child_income = left(child_income, len(child_income) - 1)

'The case note----------------------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("**TITLE IV-E eligibilty review for " & review_month & "**")
Call write_bullet_and_variable_in_CASE_NOTE("Age of child at end of review period", child_age)
Call write_bullet_and_variable_in_CASE_NOTE("Child is 18, school verification", school_verif)
Call write_bullet_and_variable_in_CASE_NOTE("In receipt of AFDC", AFDC_receipt)
call write_bullet_and_variable_in_CASE_NOTE("Rule 5", rule_five)
Call write_bullet_and_variable_in_CASE_NOTE("Child's income", child_income)
Call write_bullet_and_variable_in_CASE_NOTE("Child's assets", child_assets)
Call write_bullet_and_variable_in_CASE_NOTE("Deprivation factor", dep_factor)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("Title IV-E basic elig", basic_elig)
Call write_bullet_and_variable_in_CASE_NOTE("TITLE IV-E Reimbursable status", reimb_list)
Call write_bullet_and_variable_in_CASE_NOTE("Permanent hearing date", perm_hearing_date)
Call write_bullet_and_variable_in_CASE_NOTE("Non reimburseable months", non_reimb_months)
Call write_bullet_and_variable_in_CASE_NOTE("New date from FCLD", new_date)
If MAXIS_checkbox = 1 then Call write_variable_in_CASE_NOTE("* MAXIS updated.")
If SSIS_checkbox = 1 then Call write_variable_in_CASE_NOTE("* SSIS checked.")
If HC_approved_checkobx = 1 then Call write_variable_in_CASE_NOTE("* MMIS approved, 25X. MMIS updated.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("")
