'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - APPLICATION RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 420          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 156, 70, "Case number dialog"
  EditBox 60, 5, 90, 15, MAXIS_case_number
  EditBox 60, 25, 30, 15, MAXIS_footer_month
  EditBox 120, 25, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 40, 50, 50, 15
    CancelButton 100, 50, 50, 15
  Text 10, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 95, 30, 20, 10, "Year:"
EndDialog

BeginDialog LTC_app_recd_dialog, 0, 0, 286, 415, "LTC application received dialog"
  EditBox 75, 35, 65, 15, appl_date
  EditBox 75, 55, 65, 15, appl_type
  CheckBox 150, 45, 105, 10, "A transfer has been reported", transfer_reported_check
  CheckBox 150, 60, 140, 10, "Spousal allocation has been requested", spousal_allocation_check
  EditBox 160, 75, 120, 15, forms_needed
  EditBox 30, 95, 30, 15, CFR
  EditBox 110, 95, 170, 15, HH_comp
  EditBox 65, 115, 215, 15, pre_FACI_ADDR
  DropListBox 65, 135, 215, 15, "Select one..."+chr(9)+"Age 65 or older"+chr(9)+"Adult without children"+chr(9)+"Blind/disabled"+chr(9)+"Child under 21"+chr(9)+"Parent/Caretaker"+chr(9)+"Pregnant", basis_of_elig_droplist
  EditBox 35, 155, 245, 15, FACI
  EditBox 60, 175, 220, 15, retro_request
  EditBox 35, 195, 245, 15, AREP
  EditBox 60, 215, 220, 15, SWKR
  EditBox 60, 235, 220, 15, INSA
  EditBox 60, 255, 220, 15, adult_signatures
  EditBox 50, 275, 230, 15, veteran_info
  EditBox 50, 295, 230, 15, LTCC
  EditBox 55, 315, 225, 15, actions_taken
  CheckBox 5, 345, 220, 10, "Check here to have the script update PND2 to show client delay.", update_PND2_check
  CheckBox 5, 360, 280, 10, "Check here to have the script create a TIKL to deny at the 45 day mark (NON-DISA).", TIKL_45_day_check
  CheckBox 5, 375, 265, 10, "Check here to have the script create a TIKL to deny at the 60 day mark (DISA).", TIKL_60_day_check
  EditBox 90, 395, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 395, 50, 15
    CancelButton 230, 395, 50, 15
    PushButton 180, 25, 45, 10, "next panel", next_panel_button
    PushButton 230, 25, 45, 10, "next memb", next_memb_button
    PushButton 5, 160, 25, 10, "FACI:", FACI_button
    PushButton 5, 200, 25, 10, "AREP:", AREP_button
    PushButton 25, 220, 30, 10, "SWKR:", SWKR_button
    PushButton 5, 240, 25, 10, "INSA/", INSA_button
    PushButton 30, 240, 25, 10, "MEDI:", MEDI_button
    PushButton 15, 15, 25, 10, "TYPE", TYPE_button
    PushButton 40, 15, 25, 10, "PROG", PROG_button
    PushButton 65, 15, 25, 10, "HCRE", HCRE_button
    PushButton 90, 15, 25, 10, "REVW", REVW_button
    PushButton 115, 15, 25, 10, "MEMB", MEMB_button
  Text 5, 100, 20, 10, "CFR:"
  Text 70, 100, 40, 10, "HH Comp:"
  Text 5, 120, 60, 10, "Pre FACI address:"
  Text 5, 140, 60, 10, "Basis of eligibilty:"
  Text 5, 180, 55, 10, "Retro requested:"
  Text 5, 220, 20, 10, "PHN/"
  Text 5, 260, 55, 10, "Adult signatures:"
  Text 5, 300, 40, 10, "LTCC info:"
  Text 5, 320, 50, 10, "Actions taken:"
  Text 30, 400, 60, 10, "Worker signature:"
  Text 5, 40, 55, 10, "Application date:"
  Text 5, 80, 150, 10, "Forms needed? 1503, 3543, 3050, 5181, AA:"
  GroupBox 10, 5, 135, 25, "General STAT navigation:"
  GroupBox 175, 5, 105, 35, "STAT-based navigation"
  Text 5, 60, 65, 10, "Appl type received:"
  Text 5, 280, 45, 10, "Veteran info:"
  GroupBox 0, 335, 285, 55, "Actions"
  ButtonGroup ButtonPressed
    PushButton 230, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 180, 15, 45, 10, "prev. panel", prev_panel_button
EndDialog


'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
HH_memb_row = 05


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Searching for case number.
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Showing the case number dialog
Do
  Dialog case_number_dialog
  cancel_confirmation
  If MAXIS_case_number = "" then MsgBox "You must type a case number!"
Loop until MAXIS_case_number <> ""

'Now it checks to make sure MAXIS is running on this screen.
Call check_for_MAXIS(False)

'Navigating to STAT/HCRE so we can grab the app date
call navigate_to_MAXIS_screen("stat", "hcre")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Grabs autofill info from STAT
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)
call autofill_editbox_from_MAXIS(HH_member_array, "HCRE", appl_date)
call autofill_editbox_from_MAXIS(HH_member_array, "AREP", AREP)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", FACI)
call autofill_editbox_from_MAXIS(HH_member_array, "INSA", INSA)
call autofill_editbox_from_MAXIS(HH_member_array, "MEDI", MEDI)
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)
call autofill_editbox_from_MAXIS(HH_member_array, "SWKR", SWKR)

'Now, because INSA and MEDI will go on the same variable, we're going to add INSA to MEDI. To separate them in the case note, we have to add a semicolon (assuming both have data).
If INSA <> "" and MEDI <> "" then
  INSA = INSA & "; " & MEDI
Else
  INSA = INSA & MEDI
End if

'The main dialog
Do
	err_msg = ""
	Dialog LTC_app_recd_dialog
		cancel_confirmation
		If buttonpressed <> -1 then Call MAXIS_dialog_navigation
	IF appl_date = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter an application date."
	IF basis_of_elig_droplist = "Select one..." AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please select an MA basis of eligibility."
	IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please discuss the actions taken."
	IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
	IF (TIKL_45_day_check = 1 AND TIKL_60_day_check = 1) AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* You cannot TIKL for both 45 and 60 days. Please select one or neither."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
Loop UNTIL err_msg = "" AND ButtonPressed = -1

'checking for an active MAXIS session
Call check_for_MAXIS(False)


'UPDATING PND2----------------------------------------------------------------------------------------------------
If update_PND2_check = 1 then
	call navigate_to_MAXIS_screen("rept", "pnd2")
	EMGetCursor PND2_row, PND2_col
	EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
	If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
	EMReadScreen PND2_HC_status_check, 1, PND2_row, 65
	If PND2_HC_status_check = "P" then
		EMWriteScreen "x", PND2_row, 3
		transmit
		person_delay_row = 7
		Do
			EMReadScreen person_delay_check, 1, person_delay_row, 39
			If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
			person_delay_row = person_delay_row + 2
		Loop until person_delay_check = " " or person_delay_row > 20
		PF3
	End if
	PF3
	EMReadScreen PND2_check, 4, 2, 52
	If PND2_check = "PND2" then
		MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
		PF10
		client_delay_check = 0
	End if
End if


'THE TIKL's----------------------------------------------------------------------------------------------------
If TIKL_45_day_check = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(appl_date, 45, 5, 18)
	EMSetCursor 9, 3
	Call write_variable_in_TIKL("HC pending 45 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
	transmit
	PF3
End if

If TIKL_60_day_check = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(appl_date, 60, 5, 18)
	EMSetCursor 9, 3
	Call write_variable_in_TIKL("HC pending 60 days. Evaluate for possible denial. If any members are elderly/disabled, allow an additional 15 days and reTIKL out.")
	transmit
	PF3
End if


'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("***LTC intake***")
call write_bullet_and_variable_in_CASE_NOTE("Application date", appl_date)
call write_bullet_and_variable_in_CASE_NOTE("Application type received", appl_type)
call write_bullet_and_variable_in_CASE_NOTE("Forms Needed", forms_needed)
call write_bullet_and_variable_in_CASE_NOTE("HH comp", HH_comp)
call write_bullet_and_variable_in_CASE_NOTE("CFR", CFR)
call write_bullet_and_variable_in_CASE_NOTE("Pre FACI address", pre_FACI_ADDR)
call write_bullet_and_variable_in_CASE_NOTE("Basis of eligibility", basis_of_elig_droplist)
call write_bullet_and_variable_in_CASE_NOTE("FACI", FACI)
call write_bullet_and_variable_in_CASE_NOTE("Retro request", retro_request)
call write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
call write_bullet_and_variable_in_CASE_NOTE("PHN/SWKR", SWKR)
call write_bullet_and_variable_in_CASE_NOTE("INSA/MEDI", INSA)
call write_bullet_and_variable_in_CASE_NOTE("Adult signatures", adult_signatures)
call write_bullet_and_variable_in_CASE_NOTE("LTCC info", LTCC)
call write_bullet_and_variable_in_CASE_NOTE("Veteran information", veteran_info)
call write_bullet_and_variable_in_CASE_NOTE("Actions taken", actions_taken)
If transfer_reported_check = 1 THEN call write_variable_in_CASE_NOTE("* A transfer has been reported.")
IF spousal_allocation_check = 1 THEN Call write_variable_in_CASE_NOTE("* Spousal allocation has been requested.")
If update_PND2_check = 1 THEN Call write_variable_in_CASE_NOTE("* PND2 updated to show client delay.")
IF TIKL_45_day_check = 1 Then call write_variable_in_CASE_NOTE("* Set TIKL for 45 days to recheck case.")
IF TIKL_60_day_check = 1 Then call write_variable_in_CASE_NOTE("* Set TIKL for 60 days to recheck case.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure ("")
