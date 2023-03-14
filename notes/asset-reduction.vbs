'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - ASSET REDUCTION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
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
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("01/19/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The script----------------------------------------------------------------------------------------------------
'connecting to BlueZone and grabbing the case number & footer month/year
EMConnect ""
Call MAXIS_case_number_finder(maxis_case_number)
call check_for_MAXIS(False)	'checking for an active MAXIS session

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 146, 75, "Case number dialog"
  EditBox 85, 10, 50, 15, MAXIS_case_number
  DropListBox 85, 30, 50, 15, "Select one..."+chr(9)+"Required"+chr(9)+"Completed", reduction_status
  ButtonGroup ButtonPressed
    OkButton 30, 50, 50, 15
    CancelButton 85, 50, 50, 15
  Text 30, 15, 45, 10, "Case number:"
  Text 5, 35, 75, 10, "Asset reduction status:"
EndDialog
Do
	Do
		err_msg = ""
		DIALOG Dialog1
		cancel_without_confirmation
		if IsNumeric(MAXIS_case_number) = false or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbCr & "* Enter a valid case number."
		If reduction_status = "Select one..."THEN err_msg = err_msg & vbCr & "* Select an asset reduction status."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'Asset reduction required coding----------------------------------------------------------------------------------------------------
If reduction_status = "Required" then
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 306, 285, "Asset reduction required/pending"
      EditBox 90, 45, 60, 15, due_date
      EditBox 240, 45, 60, 15, asset_limit
      CheckBox 10, 85, 30, 10, "DWP", DWP_checkbox
      CheckBox 50, 85, 35, 10, "EMER", EMER_checkbox
      CheckBox 90, 85, 25, 10, "GA", GA_checkbox
      CheckBox 125, 85, 30, 10, "GRH", GRH_checkbox
      CheckBox 160, 85, 25, 10, "MA", MA_checkbox
      CheckBox 195, 85, 30, 10, "MFIP", MFIP_checkbox
      CheckBox 235, 85, 30, 10, "MSP", MSP_checkbox
      CheckBox 270, 85, 30, 10, "MSA", MSA_checkbox
      EditBox 65, 105, 235, 15, income
      EditBox 65, 135, 235, 15, assets
      EditBox 65, 155, 60, 15, current_asset_total
      EditBox 235, 155, 65, 15, amt_to_reduce
      EditBox 65, 180, 235, 15, other_notes
      EditBox 65, 200, 235, 15, actions_taken
      CheckBox 65, 220, 175, 10, "Sent DHS-3341 Asset reduction worksheet to client", client_3341_checkbox
      CheckBox 65, 235, 175, 10, "Sent DHS-3341 Asset reduction worksheet to AREP", AREP_3341_checkbox
      CheckBox 65, 250, 145, 10, "Set TIKL for the asset reduction due date", TIKL_checkbox
      EditBox 65, 265, 125, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 195, 265, 50, 15
        CancelButton 250, 265, 50, 15
        PushButton 10, 15, 25, 10, "BUSI", BUSI_button
        PushButton 35, 15, 25, 10, "JOBS", JOBS_button
        PushButton 10, 25, 25, 10, "RBIC", RBIC_button
        PushButton 35, 25, 25, 10, "UNEA", UNEA_button
        PushButton 80, 15, 25, 10, "ACCT", ACCT_button
        PushButton 105, 15, 25, 10, "CARS", CARS_button
        PushButton 130, 15, 25, 10, "CASH", CASH_button
        PushButton 155, 15, 25, 10, "OTHR", OTHR_button
        PushButton 80, 25, 25, 10, "REST", REST_button
        PushButton 105, 25, 25, 10, "SECU", SECU_button
        PushButton 130, 25, 25, 10, "TRAN", TRAN_button
        PushButton 155, 25, 25, 10, "HCMI", HCMI_button
        PushButton 200, 15, 45, 10, "prev. panel", prev_panel_button
        PushButton 250, 15, 45, 10, "prev. memb", prev_memb_button
        PushButton 200, 25, 45, 10, "next panel", next_panel_button
        PushButton 250, 25, 45, 10, "next memb", next_memb_button
      Text 70, 120, 230, 10, "(Income in the month received is counted as income, not as an asset)"
      GroupBox 5, 5, 60, 35, "Income panels"
      Text 35, 140, 25, 10, "Assets:"
      GroupBox 75, 5, 110, 35, "Asset panels"
      Text 5, 110, 55, 10, "Monthly income:"
      Text 15, 160, 50, 10, "Total assets: $"
      GroupBox 195, 5, 105, 35, "STAT-based navigation"
      Text 160, 50, 80, 10, "Prog(s) Asset limit(s): $"
      Text 5, 50, 85, 10, "Asset reduction due date:"
      Text 20, 185, 40, 10, "Other notes:"
      Text 15, 205, 50, 10, "Actions taken:"
      Text 5, 270, 60, 10, "Worker signature:"
      Text 150, 160, 80, 10, "Amount to be reduced: $"
      GroupBox 5, 70, 295, 30, "Asset reduction needed for the following programs:"
    EndDialog
    Do
		Do
            err_msg = ""
        	DIALOG Dialog1
        	cancel_confirmation
			MAXIS_dialog_navigation
        	IF isdate(due_date) = false THEN err_msg = err_msg & vbNewLine & "* Enter a valid asset reduction due date"
        	IF trim(asset_limit) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the asset limit amount(s)."
			IF (DWP_checkbox = 0 AND EMER_checkbox = 0 AND GA_checkbox = 0 AND GRH_checkbox = 0 AND MA_checkbox = 0 AND MFIP_checkbox = 0 AND MSP_checkbox = 0 AND MSA_checkbox = 0) then err_msg = err_msg & vbNewLine & "* Enter at least one program."
        	IF trim(income) = "" THEN err_msg = err_msg & vbNewLine & "* Enter income information for the case."
        	IF trim(current_asset_total) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the current total of counted assets."
        	IF trim(amt_to_reduce) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the amount assets need to be reduced."
			If trim(assets) = "" THEN err_msg = err_msg & vbNewLine & "* Enter asset information."
        	If trim(actions_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
			IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
        	IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
     	Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

'Asset reduction completed coding----------------------------------------------------------------------------------------------------
If reduction_status = "Completed" then
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 306, 255, "Asset reduction complete"
      EditBox 90, 50, 60, 15, within_limit_date
      EditBox 240, 50, 60, 15, asset_limit
      CheckBox 10, 80, 30, 10, "DWP", DWP_checkbox
      CheckBox 50, 80, 35, 10, "EMER", EMER_checkbox
      CheckBox 90, 80, 25, 10, "GA", GA_checkbox
      CheckBox 125, 80, 30, 10, "GRH", GRH_checkbox
      CheckBox 160, 80, 25, 10, "MA", MA_checkbox
      CheckBox 195, 80, 30, 10, "MFIP", MFIP_checkbox
      CheckBox 235, 80, 30, 10, "MSP", MSP_checkbox
      CheckBox 270, 80, 30, 10, "MSA", MSA_checkbox
      EditBox 60, 100, 240, 15, income
      EditBox 60, 130, 240, 15, assets
      EditBox 60, 150, 60, 15, current_asset_total
      EditBox 95, 170, 205, 15, how_assets_reduced
      EditBox 65, 195, 235, 15, other_assets_notes
      EditBox 65, 215, 235, 15, actions_taken
      EditBox 65, 235, 125, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 195, 235, 50, 15
        CancelButton 250, 235, 50, 15
        PushButton 10, 15, 25, 10, "BUSI", BUSI_button
        PushButton 35, 15, 25, 10, "JOBS", JOBS_button
        PushButton 10, 25, 25, 10, "RBIC", RBIC_button
        PushButton 35, 25, 25, 10, "UNEA", UNEA_button
        PushButton 80, 15, 25, 10, "ACCT", ACCT_button
        PushButton 105, 15, 25, 10, "CARS", CARS_button
        PushButton 130, 15, 25, 10, "CASH", CASH_button
        PushButton 155, 15, 25, 10, "OTHR", OTHR_button
        PushButton 80, 25, 25, 10, "REST", REST_button
        PushButton 105, 25, 25, 10, "SECU", SECU_button
        PushButton 130, 25, 25, 10, "TRAN", TRAN_button
        PushButton 155, 25, 25, 10, "HCMI", HCMI_button
        PushButton 200, 15, 45, 10, "prev. panel", prev_panel_button
        PushButton 250, 15, 45, 10, "prev. memb", prev_memb_button
        PushButton 200, 25, 45, 10, "next panel", next_panel_button
        PushButton 250, 25, 45, 10, "next memb", next_memb_button
      GroupBox 5, 5, 60, 35, "Income panels"
      Text 30, 135, 25, 10, "Assets:"
      GroupBox 75, 5, 110, 35, "Asset panels"
      Text 5, 105, 55, 10, "Monthly income:"
      Text 5, 155, 50, 10, "Total assets $:"
      GroupBox 195, 5, 105, 35, "STAT-based navigation"
      Text 165, 55, 75, 10, "Prog(s) Asset limit(s): $"
      Text 10, 55, 80, 10, "Date assets within limit:"
      Text 20, 200, 40, 10, "Other notes:"
      Text 15, 220, 50, 10, "Actions taken:"
      Text 5, 240, 60, 10, "Worker signature:"
      Text 70, 115, 230, 10, "(Income in the month received is counted as income, not as an asset)"
      Text 5, 175, 90, 10, "How were assets reduced:"
      Text 125, 155, 180, 10, "For HC processing steps: see POLI TEMP TE02.07.246"
      GroupBox 5, 70, 295, 25, "Within asset limit for following programs:"
    EndDialog
    Do
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_confirmation
    		MAXIS_dialog_navigation
    		IF isdate(within_limit_date) = false THEN err_msg = err_msg & vbNewLine & "* Enter the date case met asset limit."
    		IF trim(asset_limit) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the asset limit amount(s)."
    		IF (DWP_checkbox = 0 AND EMER_checkbox = 0 AND GA_checkbox = 0 AND GRH_checkbox = 0 AND MA_checkbox = 0 AND MFIP_checkbox = 0 AND MSP_checkbox = 0 AND MSA_checkbox = 0) then err_msg = err_msg & vbNewLine & "* Enter at least one program."
    		IF trim(income) = "" THEN err_msg = err_msg & vbNewLine & "* Enter income information for the case."
    		IF trim(assets) = "" THEN err_msg = err_msg & vbNewLine & "* Enter asset information."
    		IF trim(current_asset_total) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the current total of counted assets."
    		IF trim(how_assets_reduced) = "" THEN err_msg = err_msg & vbNewLine & "* How were assets reduced?"
    		If trim(actions_taken) = "" THEN err_msg = err_msg & vbNewLine & "* Enter the actions taken."
    		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

'Sets TIKL for the pending/reduction option
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
IF TIKL_checkbox = 1 then Call create_TIKL("Asset reduction verification is due. Please review case and case documents.", 0, due_date, False, TIKL_note_text)

'turns program checkboxes into a variable for the case note
reduction_progs = ""	'establishing variable as ""
IF DWP_checkbox = 1 THEN reduction_progs = reduction_progs & "DWP" & ", "
IF EMER_checkbox = 1 THEN reduction_progs = reduction_progs & "EMER" & ", "
IF GA_checkbox = 1 THEN reduction_progs = reduction_progs & "GA" & ", "
IF GRH_checkbox = 1 THEN reduction_progs = reduction_progs & "GRH"  & ", "
IF MA_checkbox = 1 THEN reduction_progs = reduction_progs & "MA" & ", "
IF MFIP_checkbox = 1 THEN reduction_progs = reduction_progs & "MFIP" & ", "
IF MSP_checkbox = 1 THEN reduction_progs = reduction_progs & "MSP" & ", "
IF MSA_checkbox = 1 THEN reduction_progs = reduction_progs & "MSA" & ", "
'trims excess spaces of reduction_progs
reduction_progs = trim(reduction_progs)
'takes the last comma off of reduction_progs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(reduction_progs, 1) = "," THEN reduction_progs = left(reduction_progs, len(reduction_progs) - 1)

'The case note----------------------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("--Asset reduction "& reduction_status & "--")
call write_bullet_and_variable_in_CASE_NOTE("Asset reduction due date", due_date)
Call write_bullet_and_variable_in_CASE_NOTE("Within asset limit on", within_limit_date)
call write_bullet_and_variable_in_CASE_NOTE("Program(s) asset limit", asset_limit)
call write_bullet_and_variable_in_CASE_NOTE("Asset reduction required for", reduction_progs)
call write_bullet_and_variable_in_CASE_NOTE("Monthly income", income)
call write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
Call write_bullet_and_variable_in_CASE_NOTE("Total of all assets", current_asset_total)
call write_bullet_and_variable_in_CASE_NOTE("Amount to be reduced", amt_to_reduce)
Call write_bullet_and_variable_in_CASE_NOTE("How assets were reduced", how_assets_reduced)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
If client_3341_checkbox = 1 then Call write_variable_in_CASE_NOTE("* DHS-3341 asset reduction worksheet sent to client.")
If AREP_3341_checkbox = 1 then Call write_variable_in_CASE_NOTE("* DHS-3341 asset reduction worksheet sent to AREP.")
If TIKL_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Set TIKL for the asset reduction due date.")
Call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
