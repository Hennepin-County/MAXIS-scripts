'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - IV-E.vbs"
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
call changelog_update("11/25/2019", "Updated backend functionality and added changelog.", "Ilse Ferris, Hennepin County")
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
BeginDialog dialog1, 0, 0, 181, 75, "Select a IV-E option"
  EditBox 95, 10, 60, 15, MAXIS_case_number
  DropListBox 95, 30, 75, 15, "Select one..."+chr(9)+"Approved"+chr(9)+"Closing"+chr(9)+"Denied", action_option
  ButtonGroup ButtonPressed
    OkButton 65, 50, 50, 15
    CancelButton 120, 50, 50, 15
  Text 10, 35, 80, 10, "Select the action to take:"
  Text 45, 15, 45, 10, "Case number:"
EndDialog
DO
	DO
		err_msg = ""
		Dialog dialog1
        cancel_without_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF action_option = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an IV-E option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If action_option = "Approved" then
    dialog1 = ""
    BeginDialog dialog1, 0, 0, 386, 310, "IV-E Approved"
      EditBox 70, 10, 55, 15, app_date
      EditBox 180, 10, 55, 15, elig_month
      EditBox 325, 10, 55, 15, date_pet_filed
      EditBox 70, 30, 55, 15, MA_app
      CheckBox 160, 35, 60, 10, "SSIS checked", SSIS_checkbox
      EditBox 325, 30, 55, 15, placement_date
      EditBox 325, 50, 55, 15, hearing_date
      EditBox 125, 70, 255, 15, placement_gaps
      EditBox 65, 95, 110, 15, best_interest
      EditBox 65, 115, 110, 15, resonable_efforts
      EditBox 270, 95, 110, 15, mother_info
      EditBox 270, 115, 110, 15, father_info
      EditBox 65, 140, 100, 15, HH_income
      EditBox 240, 140, 100, 15, income_verif
      EditBox 65, 160, 100, 15, HH_assets
      EditBox 240, 160, 100, 15, asset_verif
      EditBox 65, 180, 100, 15, HH_comp
      EditBox 240, 180, 100, 15, HH_verif
      EditBox 65, 200, 100, 15, rule_five
      EditBox 240, 200, 100, 15, overpayments
      EditBox 240, 220, 100, 15, county_transfer
      EditBox 75, 240, 265, 15, results
      EditBox 75, 260, 265, 15, other_notes
      EditBox 75, 280, 150, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 235, 280, 50, 15
        CancelButton 290, 280, 50, 15
      Text 5, 70, 115, 20, "Explain reasons for gaps between court order and placement:"
      Text 5, 145, 60, 10, "AFDC HH income:"
      Text 140, 15, 40, 10, "Elig month:"
      Text 15, 100, 50, 10, "Best interest:"
      Text 260, 15, 60, 10, "Date petition filed:"
      Text 190, 185, 50, 10, "HH comp verif:"
      Text 35, 15, 35, 10, "Appl date:"
      Text 190, 100, 75, 10, "Mother's name/Case #:"
      Text 10, 185, 55, 10, "AFDC HH comp:"
      Text 190, 120, 75, 10, "Father's name/Case #:"
      Text 15, 285, 60, 10, "Worker signature: "
      Text 5, 120, 60, 10, "Resonable efforts:"
      Text 195, 145, 45, 10, "Income verif:"
      Text 35, 205, 25, 10, "Rule 5:"
      Text 45, 245, 30, 10, "Results:"
      Text 185, 205, 50, 10, "Overpayments:"
      Text 35, 265, 40, 10, "Other notes: "
      Text 35, 35, 35, 10, "MA App'd:"
      Text 200, 165, 40, 10, "Asset verif:"
      Text 5, 165, 60, 10, "AFDC HH assets:"
      Text 110, 225, 125, 10, "Custody of child transferred to county:"
      Text 240, 55, 85, 10, "Court order hearing date:"
      Text 240, 35, 85, 10, "Physical placement date:"
    EndDialog

	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
            If isDate(app_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid application date."
			If elig_month = "" then err_msg = err_msg & vbNewLine & "* Enter the eligibilty month."
			If isDate(date_pet_filed) = False then err_msg = err_msg & vbNewLine & "* Enter a valid date the petition was filed."
			If isDate(MA_app) = False then err_msg = err_msg & vbNewLine & "* Enter a valid MA approved date."
            If isdate(placement_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid physical placement date."
            If isdate(hearing_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid court order hearing date."
            If placement_gaps = "" then err_msg = err_msg & vbnewline & "* Enter the gaps in placement information."
            If best_interest = "" then err_msg = err_msg & vbnewline & "* Enter the best interest information."
            If resonable_efforts = "" then err_msg = err_msg & vbnewline & "* Enter the reasonable effort information."
            If mother_info = "" then err_msg = err_msg & vbnewline & "* Enter the mother's information."
            If father_info = "" then err_msg = err_msg & vbnewline & "* Enter the father's information."
            If HH_income <> "" and income_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the income verification."
            If HH_assets <> "" and asset_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the asset verification."
            If HH_comp <> "" and HH_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the HH comp verification."
            If rule_five = "" then err_msg = err_msg & vbNewLine & "* Enter the Rule 5 information."
            If results = "" then err_msg = err_msg & vbNewLine & "* Enter the results information."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	Call write_variable_in_CASE_NOTE("**IV-E Approved on " & app_date & "**")
	Call write_bullet_and_variable_in_CASE_NOTE("Application date", app_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Elig month", elig_month)
    Call write_bullet_and_variable_in_CASE_NOTE("Date petition filed", date_pet_filed)
    Call write_bullet_and_variable_in_CASE_NOTE("MA approved", MA_app)
    Call write_bullet_and_variable_in_CASE_NOTE("Physical placement date", placement_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Court order hearing date", hearing_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason for gaps between court order and placement", placement_gaps)
    Call write_bullet_and_variable_in_CASE_NOTE("Best interest", best_interest)
    Call write_bullet_and_variable_in_CASE_NOTE("Reasonable efforts", resonable_efforts)
    Call write_bullet_and_variable_in_CASE_NOTE("Mother's name/case #", mother_info)
	Call write_bullet_and_variable_in_CASE_NOTE("Father's name/case #", father_info)
    Call write_variable_in_CASE_NOTE("* HH income: " & HH_income & ". Verif: " & income_verif)
    Call write_variable_in_CASE_NOTE("* HH assets: " & HH_assets & ". Verif: " & asset_verif)
    Call write_variable_in_CASE_NOTE("* HH comp: " & HH_comp & ". Verif: " & HH_verif)
    call write_bullet_and_variable_in_CASE_NOTE("Rule 5", rule_five)
	Call write_bullet_and_variable_in_CASE_NOTE("Overpayments", Overpayments)
    Call write_bullet_and_variable_in_CASE_NOTE("Custody of child transferred to county", county_transfer)
    Call write_bullet_and_variable_in_CASE_NOTE("Results", Results)
    If SSIS_checkbox = 1 then Call write_variable_in_CASE_NOTE("* SSIS checked.")
END IF

If action_option = "Closing" then
    dialog1 = ""
	BeginDialog dialog1, 0, 0, 341, 125, "IV-E closing"
		EditBox 115, 5, 60, 15, IVE_closure_date
		EditBox 275, 5, 60, 15, MA_closure_date
		EditBox 115, 25, 220, 15, reason_closing
		EditBox 115, 45, 220, 15, notified_by
		EditBox 115, 65, 220, 15, reim_months
		EditBox 115, 85, 220, 15, other_notes
		EditBox 115, 105, 110, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 230, 105, 50, 15
			CancelButton 285, 105, 50, 15
		Text 25, 10, 90, 10, "IV-E closure effective date:"
		Text 185, 10, 90, 10, "MA closure effective date:"
		Text 45, 30, 65, 10, "Reason for closing:"
		Text 50, 110, 60, 10, "Worker signature:"
		Text 5, 70, 110, 10, "Checked reimbursability months:"
		Text 70, 50, 40, 10, "Notified by:"
		Text 70, 90, 40, 10, "Other notes:"
	EndDialog

	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
            If isDate(IVE_closure_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid IV-E closure date."
			If isDate(MA_closure_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid MA closure date."
            IF reason_closing = "" then err_msg = err_msg & vbNewLine & "* Enter the reason the IV-E is closing."
			If notified_by = "" then err_msg = err_msg & vbNewLine & "* Enter the notified by information."
            If reim_months = "" then err_msg = err_msg & vbNewLine & "* Enter the reimbursability information."
            If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("<<IV-E closing effective " & IVE_closure_date & ">>")
    Call write_bullet_and_variable_in_CASE_NOTE("MA closing date", MA_closure_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason(s) for closure", reason_closing)
    Call write_bullet_and_variable_in_CASE_NOTE("Notified by", notified_by)
    Call write_bullet_and_variable_in_CASE_NOTE("Check reimbursability months", reim_months)
END IF

If action_option = "Denied" then
    dialog1 = ""
	BeginDialog dialog1, 0, 0, 341, 220, "IV-E denied"
		EditBox 100, 10, 55, 15, IVE_denied_date
		EditBox 205, 10, 55, 15, app_date
		CheckBox 275, 15, 60, 10, "SSIS checked", SSIS_checkbox
		EditBox 100, 30, 230, 15, denial_reason
		EditBox 100, 50, 55, 15, date_pet_filed
		EditBox 275, 50, 55, 15, court_date
		EditBox 100, 70, 55, 15, placement_date
		EditBox 70, 90, 100, 15, HH_income
		EditBox 230, 90, 100, 15, income_verif
		EditBox 70, 110, 100, 15, HH_assets
		EditBox 230, 110, 100, 15, asset_verif
		EditBox 70, 130, 100, 15, HH_comp
		EditBox 230, 130, 100, 15, HH_verif
		EditBox 70, 155, 260, 15, results
		EditBox 70, 175, 260, 15, other_notes
		EditBox 70, 195, 150, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 225, 195, 50, 15
			CancelButton 280, 195, 50, 15
		Text 10, 200, 60, 10, "Worker signature: "
		Text 165, 15, 35, 10, "App date:"
		Text 185, 95, 45, 10, "Income verif:"
		Text 40, 160, 30, 10, "Results:"
		Text 10, 115, 60, 10, "AFDC HH assets:"
		Text 190, 115, 40, 10, "Asset verif:"
		Text 10, 95, 60, 10, "AFDC HH income:"
		Text 15, 135, 55, 10, "AFDC HH comp:"
		Text 30, 180, 40, 10, "Other notes: "
		Text 180, 135, 50, 10, "HH comp verif:"
		Text 35, 55, 60, 10, "Date petition filed:"
		Text 10, 75, 85, 10, "Physical placement date:"
		Text 40, 15, 60, 10, "Date IV-E denied:"
		Text 190, 55, 85, 10, "Court order hearing date:"
		Text 30, 35, 70, 10, "Reason IV-E denied:"
	EndDialog

	DO
		DO
			err_msg = ""
			Dialog dialog1
			cancel_confirmation
			If isDate(IVE_denied_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid IV-E denied date."
            If isDate(app_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid application date."
			If denial_reason = "" then err_msg = err_msg & vbNewLine & "* Enter the denial reason."
			If date_pet_filed = "" then err_msg = err_msg & vbNewLine & "* Enter the date the petition was filed."
            If isdate(court_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid court order hearing date."
            If isdate(placement_date) = False then err_msg = err_msg & vbNewLine & "* Enter a valid physical placement date."
 			If HH_income <> "" and income_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the income verification."
            If HH_assets <> "" and asset_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the asset verification."
            If HH_comp <> "" and HH_verif = "" then err_msg = err_msg & vbNewLine & "* Enter the HH comp verification."
            If results = "" then err_msg = err_msg & vbNewLine & "* Enter the results information."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in

	'The case note
    start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
    Call write_variable_in_CASE_NOTE("<<IV-E denied effective " & IVE_denied_date & ">>")
    Call write_bullet_and_variable_in_CASE_NOTE("Application date", app_date)
    Call write_bullet_and_variable_in_CASE_NOTE("Reason(s) for closure", denial_reason)
    Call write_bullet_and_variable_in_CASE_NOTE("Date petition filed", date_pet_filed)
	Call write_bullet_and_variable_in_CASE_NOTE("Court order hearing date", court_date)
	Call write_bullet_and_variable_in_CASE_NOTE("Physical placement date", placement_date)
    Call write_variable_in_CASE_NOTE("* HH income: " & HH_income & ". Income verif: " & income_verif)
    Call write_variable_in_CASE_NOTE("* HH assets: " & HH_assets & ". Asset verif: " & asset_verif)
    Call write_variable_in_CASE_NOTE("* HH comp: " & HH_comp & ". HH comp verif: " & HH_verif)
	Call write_bullet_and_variable_in_CASE_NOTE("Results", Results)
	If SSIS_checkbox = 1 then Call write_variable_in_CASE_NOTE("* SSIS checked.")
END IF

Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("")
